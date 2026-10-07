"""compile 能力的纯逻辑核（可导入、无 CLI 副作用）。

从 render.py 外提把 Markdown 编成 pandoc 命令、判断是否需要人工提示的纯助手：
_should_prompt（是否当场问验收）、_outside_code_fences（剥围栏代码块）、
CJK 直引号处理（is_cjk_dominant / pair_cjk_quotes，R4 #1）、
_build_pandoc_cmd（拼 pandoc 命令行，返回 list 可断言）。render.py 作为 thin
CLI 壳 re-export 同名符号，`run` / `preflight` / `doctor` / `_export_and_check`
等含 sys.exit / 子进程 / 版式的编排仍留壳。裸名 sibling import，禁止反向 import render。
"""

import os
import re

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _should_prompt(can_read, can_write):
    """该不该当场问验收？只有「人在终端前」才问。

    两个流都得是 TTY：脚本 / CI / pytest 里 capture_output=True 时 stdout 是管道，
    那时子进程的 stdin 可能仍然继承了终端——只查 stdin 会挂在那里等输入。
    """
    return bool(can_read and can_write)


def _outside_code_fences(md_text: str) -> str:
    """去掉 ``` / ~~~ 围栏代码块，只留围栏外的正文。"""
    out, fence = [], None
    for line in md_text.splitlines():
        s = line.strip()
        if fence is None and (s.startswith("```") or s.startswith("~~~")):
            fence = s[:3]
            continue
        if fence is not None:
            if s.startswith(fence):
                fence = None
            continue
        out.append(line)
    return "\n".join(out)


# --- CJK 直引号处理（用户实测反馈：中文文档里 ASCII 引号被 smart 全收成右引号）---
# pandoc 的 smart 按西文「空格分词」的习惯判断引号开闭：CJK 文本里 ASCII 直引号
# 前后没有空格，会被整体判成闭引号（pandoc 3.11 实测 6/6 错配，段首也不幸免）。
# 所以中文文档要两件事一起做：关 smart（否则错配照旧）+ 围栏外直引号按段配对成
# 全角 “”（只关 smart 会把直引号原样留在正文里，排版同样不对）。全角引号本就
# 是 U+201C/201D，直接打出全角引号的源文件不受影响。
_HAN_RE = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]")
_CJK_MIN_HAN = 30  # 汉字少于这个数不算中文文档（英文报告里出现几个中文词不触发）
_CJK_HAN_PER_LATIN = 0.3  # 汉字数低于拉丁字母数的 30% 视为西文主导
_INLINE_CODE_RE = re.compile(r"(`+[^`]*`+)")
_UNESCAPED_DQUOTE_RE = re.compile(r'(?<!\\)"')


def is_cjk_dominant(md_text: str) -> bool:
    """合并后的 Markdown 是否以中文为主导（围栏代码块不参与统计）。"""
    text = _outside_code_fences(md_text)
    han = len(_HAN_RE.findall(text))
    if han < _CJK_MIN_HAN:
        return False
    latin = sum(1 for ch in text if ch.isascii() and ch.isalpha())
    return han >= _CJK_HAN_PER_LATIN * latin


def _pair_cjk_quotes_block(block: str):
    """单个段落内的配对：行内代码 span 不动，转义引号 \\" 不算；奇数个则整段放弃。

    返回 (新文本, 配对对数, 是否因奇数放弃)。
    """
    segments = _INLINE_CODE_RE.split(block)
    quote_count = sum(
        len(_UNESCAPED_DQUOTE_RE.findall(s)) for i, s in enumerate(segments) if i % 2 == 0
    )
    if quote_count == 0:
        return block, 0, 0
    if quote_count % 2:
        return block, 0, 1
    made = 0

    def _sub(_m):
        nonlocal made
        made += 1
        return "“" if made % 2 else "”"

    for i in range(0, len(segments), 2):
        segments[i] = _UNESCAPED_DQUOTE_RE.sub(_sub, segments[i])
    return "".join(segments), quote_count // 2, 0


def pair_cjk_quotes(md_text: str):
    """围栏外的 ASCII 直引号按段配对成全角 “”；返回 (新文本, 总对数, 放弃段数)。

    按空行分段、段内独立配对（每段都从 “ 重新起算）；围栏代码块逐字保留，
    围栏判定规则与 _outside_code_fences 同一条。引号不成对（奇数个）的段落
    整段不动——多半是笔误或特殊用法，猜错方向比不猜更糟。
    """
    out, fence, buf = [], None, []
    n_pairs = n_odd = 0

    def flush():
        nonlocal n_pairs, n_odd
        if buf:
            block, pairs, odd = _pair_cjk_quotes_block("\n".join(buf))
            out.append(block)
            n_pairs += pairs
            n_odd += odd
            buf.clear()

    for line in md_text.splitlines():
        s = line.strip()
        if fence is None and (s.startswith("```") or s.startswith("~~~")):
            flush()
            fence = s[:3]
            out.append(line)
            continue
        if fence is not None:
            out.append(line)
            if s.startswith(fence):
                fence = None
            continue
        if not s:
            flush()
            out.append(line)
            continue
        buf.append(line)
    flush()
    return "\n".join(out), n_pairs, n_odd


def _build_pandoc_cmd(all_md, body, ref, uniq, cfg, cjk=False):
    """拼 pandoc 命令行；captions.lua 存在时把题注关键字一并喂给它。

    cjk=True：中文主导的文档关掉 smart 扩展——smart 的引号开闭判定以西文空格
    分词为前提，CJK 文本里必然错配（实测 6/6 全收成右引号）；直引号由
    pair_cjk_quotes 在合并阶段按中文习惯配好，不依赖 smart。
    """
    lua_filter = os.path.join(KIT, "scripts", "filters", "captions.lua")
    reader = "markdown+pipe_tables+raw_html" + ("-smart" if cjk else "")
    cmd = [
        "pandoc",
        all_md,
        "-o",
        body,
        "--reference-doc=" + ref,
        "--resource-path=" + os.pathsep.join(uniq),
        "-f",
        reader,
        "--wrap=none",
    ]
    if os.path.exists(lua_filter):
        # AST 层标记表题/图注，post.py 就不用再靠正则猜
        cmd.append("--lua-filter=" + lua_filter)
        # 题注关键字可配：同一份配置同时喂给 lua filter 与 post.py
        cw = cfg.get("caption_words") or {}
        for key, meta_name in (
            ("table", "dk-table-words"),
            ("figure", "dk-figure-words"),
        ):
            if cw.get(key):
                cmd += ["-M", "%s=%s" % (meta_name, ",".join(cw[key]))]
    else:
        print("[warn] 缺少 filters/captions.lua，题注退回文本正则判定")
    return cmd
