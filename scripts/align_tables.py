"""grid table 显示宽度对齐器：按 East Asian 显示宽度重排竖线与填充空格。

背景（README「已知边界」）：pandoc 的 grid table 按**显示宽度**解析——中文占 2 列，
「等宽编辑器里看着齐」仍可能解析成单列坏表。本工具按 unicodedata 显示宽度
（W/F 记 2，其余记 1）重排 `+---+/+===+` 分隔行与 `|` 内容行的竖线位置与填充。

保证：
- **单元格内容字节不变**——仅归一化首尾空白（pandoc 本就不感知单元格首尾空白），
  内部空白逐字保留；写回前逐单元格自校验，任何差异即放弃写回并报错。
- 跳过代码围栏（```）内的内容；pipe table（无 `+` 分隔行的 `|---|`）不碰。
- 保留块内缩进（列表里的表格）与 `===` 表头分隔行。

已知边界：单元格内的代码围栏、含 `|` 的单元格内容不支持——按解析失败报错退出，
绝不猜测性修复。写回统一 LF（与仓库 i/lf 一致）。

用法:
  python scripts/align_tables.py chapters/*.md           # 报告模式：显示将改什么，不写回
  python scripts/align_tables.py --fix chapters/*.md     # 就地写回
  python scripts/align_tables.py --check chapters/*.md   # 静默检查：未对齐 exit 1（pre-commit 用）
"""

import argparse
import os
import re
import sys
import unicodedata

GRID_SEP_RE = re.compile(r"^\++[-=+]+\+$")
BLOCK_LINE_RE = re.compile(r"^\s*[+|]")


def dw(s: str) -> int:
    """显示宽度：East Asian W/F 记 2，其余记 1。"""
    return sum(2 if unicodedata.east_asian_width(c) in ("W", "F") else 1 for c in s)


def _is_fence(line: str) -> bool:
    return line.lstrip().startswith("```")


def _split_row(line: str):
    """去掉首尾竖线后按 | 拆分单元格；非合法内容行返回 None。"""
    t = line.strip()
    if not (t.startswith("|") and t.endswith("|")):
        return None
    return t[1:-1].split("|")


def _sep_fills(line: str):
    """分隔行 → 每列的填充字符（'---'→'-'，'==='→'='）。混用取首字符。"""
    cells = line.strip().strip("+").split("+")
    fills = []
    for c in cells:
        ch = c[0] if c else "-"
        if ch not in "-=":
            raise ValueError("分隔行出现意外字符 %r" % ch)
        fills.append(ch)
    return fills


def find_blocks(lines):
    """扫描出 grid table 块 → [(start, end)]（行号，排他）；跳过代码围栏内。

    块 = 连续的 +/| 行，且至少含一条 +---+/+===+ 分隔行（pipe table 不算）。
    """
    blocks = []
    i, n = 0, len(lines)
    in_fence = False
    while i < n:
        if _is_fence(lines[i]):
            in_fence = not in_fence
            i += 1
            continue
        if in_fence or not BLOCK_LINE_RE.match(lines[i]):
            i += 1
            continue
        j = i
        while j < n and not _is_fence(lines[j]) and BLOCK_LINE_RE.match(lines[j]):
            j += 1
        if any(GRID_SEP_RE.match(b.lstrip()) for b in lines[i:j]):
            blocks.append((i, j))
        i = j
    return blocks


def align_block(block):
    """对齐一个块，返回新行列表；解析失败抛 ValueError。"""
    indent = re.match(r"\s*", block[0]).group(0)
    parsed = []  # ("sep", fills) | ("row", cells)
    col_count = None
    for raw in block:
        s = raw.lstrip()
        if GRID_SEP_RE.match(s):
            fills = _sep_fills(s)
            if col_count is None:
                col_count = len(fills)
            elif len(fills) != col_count:
                raise ValueError("分隔行列数不一致（%d != %d）" % (len(fills), col_count))
            parsed.append(("sep", fills))
        else:
            cells = _split_row(raw)
            if cells is None:
                raise ValueError("内容行缺少首尾竖线: %r" % raw)
            if col_count is not None and len(cells) != col_count:
                raise ValueError(
                    "内容行列数与分隔行不一致（%d != %d）——表格已错位到无法安全解析"
                    % (len(cells), col_count)
                )
            if col_count is None:
                raise ValueError("块首行不是分隔行")
            parsed.append(("row", cells))

    widths = [0] * col_count
    for kind, payload in parsed:
        if kind == "row":
            for idx, cell in enumerate(payload):
                widths[idx] = max(widths[idx], dw(cell.strip()))

    out = []
    for kind, payload in parsed:
        if kind == "sep":
            out.append(
                indent + "+" + "+".join(ch * (w + 2) for ch, w in zip(payload, widths)) + "+"
            )
        else:
            cells = []
            for cell, w in zip(payload, widths):
                c = cell.strip()
                cells.append(" " + c + " " * (w - dw(c)) + " ")
            out.append(indent + "|" + "|".join(cells) + "|")

    # 自校验：每个单元格 strip 后必须与原内容逐字一致（首尾空白归一是唯一许可的变化）
    row_lines = [ln for ln in out if ln.lstrip().startswith("|")]
    old_rows = [payload for kind, payload in parsed if kind == "row"]
    for old_payload, new_line in zip(old_rows, row_lines):
        new_cells = _split_row(new_line)
        for oc, nc in zip(old_payload, new_cells):
            if oc.strip() != nc.strip():
                raise ValueError("自校验失败：单元格内容被改变 %r -> %r" % (oc, nc))
    return out


def process_text(text):
    """处理整篇文本 → (new_text, changed_blocks, errors)。changed_blocks 记 1 起始行号。"""
    lines = text.split("\n")
    changed, errors = [], []
    for start, end in reversed(find_blocks(lines)):
        try:
            new_block = align_block(lines[start:end])
        except ValueError as e:
            errors.append("第 %d 行起的 grid table： %s" % (start + 1, e))
            continue
        if new_block != lines[start:end]:
            changed.append(start + 1)
            lines[start:end] = new_block
    return "\n".join(lines), changed, errors


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="grid table 显示宽度对齐器（单元格内容字节不变，仅重排竖线与填充）"
    )
    ap.add_argument("files", nargs="+", help="Markdown 文件路径")
    ap.add_argument("--fix", action="store_true", help="就地写回对齐结果（默认只报告）")
    ap.add_argument(
        "--check", action="store_true", help="静默检查：发现未对齐即 exit 1（pre-commit 用）"
    )
    a = ap.parse_args(argv)

    rc = 0
    for path in a.files:
        if not os.path.isfile(path):
            print("文件不存在: %s" % path, file=sys.stderr)
            rc = 1
            continue
        with open(path, encoding="utf-8") as fh:
            text = fh.read()
        new_text, changed, errors = process_text(text)
        for err in errors:
            print("%s: %s" % (path, err), file=sys.stderr)
            rc = 1
        if a.check:
            if changed:
                print("%s: %d 处 grid table 未对齐" % (path, len(changed)), file=sys.stderr)
                rc = 1
            continue
        if changed:
            print("%s: %d 处 grid table 需对齐（行：%s）" % (path, len(changed), changed))
        if a.fix and changed and not errors:
            with open(path, "w", encoding="utf-8", newline="\n") as fh:
                fh.write(new_text)
            print("%s: 已写回" % path)
    return rc


if __name__ == "__main__":
    sys.exit(main())
