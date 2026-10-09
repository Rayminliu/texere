"""一键渲染：Markdown 章节目录 -> 正式中文 docx（可选 PDF 与目视验收）。

用法:
  python scripts/render.py --src <md目录> --out <out.docx> [--config config.json] [--pdf] [--check]
  python scripts/render.py --sample     # 用自带 sample.md 做冒烟测试

依赖: pandoc、python-docx；--pdf 需本机 Word + pywin32；--check 需 PyMuPDF。
说明: 若招标方/甲方提供强制格式模板 docx，可在 config 中设 "reference_doc" 指向它，
      正文样式即继承对方模板（封面/目录/页眉页脚仍由 post.py 统一注入）。
"""

import argparse
import atexit
import glob
import importlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime

from _compile import (
    _build_pandoc_cmd,
    _outside_code_fences,
    _should_prompt,
    is_cjk_dominant,
    pair_cjk_quotes,
)
from _shared import UTF8_ENV, __version__, force_utf8_stdio
from _shared import display_width as dw
from renderers import SUPPORTED_RENDERERS, get_renderer

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# --resource-path 的字符预算（Windows 命令行上限 32767，这里取保守值留余量）
RESOURCE_PATH_BUDGET = 6000


def cleanup_old_temp(prefix="texere_"):
    """清理超过 24 小时的旧临时目录，防止磁盘空间泄漏"""
    temp_base = tempfile.gettempdir()
    cutoff = datetime.now().timestamp() - 86400  # 24 小时前

    cleaned = 0
    for name in os.listdir(temp_base):
        if name.startswith(prefix):
            path = os.path.join(temp_base, name)
            try:
                if os.path.isdir(path):
                    mtime = os.stat(path).st_mtime
                    if mtime < cutoff:
                        shutil.rmtree(path, ignore_errors=True)
                        cleaned += 1
            except OSError:
                pass

    if cleaned > 0:
        print(f"[cleanup] 删除 {cleaned} 个旧临时目录")


class _Workdir:
    """中间产物的生命周期：失败一定回收，成功后按「人是否在场」分流。

    为什么成功后不删：以前跑完立刻 rmtree，用户看过 PDF 说"这页要改"时，
    pandoc 的 body.docx、合并后的 all.md、页面截图都没了——只能整链重跑
    （272 页文档上 pandoc + 真机导 PDF 是分钟级），而 patch.py / edit.py 本来
    能直接改那份 body.docx 或对照截图定位问题。

    为什么不再向用户提问（UX 反馈：渲染成功后突然被问「验收确认」很突兀）：
    人在终端 → 静默保留 + 一行提示，24 小时自动回收兜底磁盘；非交互（脚本 /
    CI）→ 跑完即回收，不涨临时目录。失败 / 缺图不走这里：那时目录里没有可
    交付的东西，留着只会涨磁盘（实测一天攒 12 个临时目录才立的规矩）。
    """

    KEEP = "keep"
    DISCARD = "discard"
    AUTO = "auto"

    def __init__(self, path):
        self.path = path
        self.settled = False
        atexit.register(self._fallback)

    def _fallback(self):
        """没走到 settle（异常、子进程报错、sys.exit）：照旧回收。"""
        if not self.settled:
            shutil.rmtree(self.path, ignore_errors=True)

    def contents(self):
        items = []
        for name in ("all.md", "body.docx", "check_pages"):
            p = os.path.join(self.path, name)
            if os.path.isdir(p):
                items.append(name + "/")
            elif os.path.exists(p):
                items.append(name)
        return items

    def settle(self, mode=AUTO):
        """交付物已产出，决定中间产物去留；返回最终状态 kept / discarded。"""
        self.settled = True
        if mode == self.DISCARD:
            shutil.rmtree(self.path, ignore_errors=True)
            print("[work] 中间产物已按要求回收（--discard-work）")
            return "discarded"
        if mode == self.AUTO:
            if _should_prompt(sys.stdin.isatty(), sys.stdout.isatty()):
                # 人在终端：不猜「验没验收」，默认先留着——24h 自动回收兜底，
                # 返修不必重跑；确认不要了用 --discard-work（不再打断用户提问）。
                return self._hold(brief=True)
            # 非交互（脚本 / CI / 管道）：维持「跑完即回收」，常驻服务器的批量
            # 渲染不能依赖「下次启动才回收」；要留现场显式 --keep-work。
            shutil.rmtree(self.path, ignore_errors=True)
            print("[work] 非交互运行：中间产物已回收；需人工验收返修请加 --keep-work")
            return "discarded"
        return self._hold()

    def _hold(self, brief=False):
        if brief:
            print(
                "[work] 中间产物（%s）已保留 24 小时供返修：%s（--discard-work 立即删）"
                % (", ".join(self.contents()), self.path)
            )
            return "kept"
        print("[work] 已保留中间产物： %s（%s）" % (self.path, ", ".join(self.contents()) or "空"))
        print(
            "[work] 返修不必重跑：拿 body.docx 直接用 patch.py / edit.py 改，"
            "或对照 check_pages/ 里的截图定位问题。"
        )
        # 保留的目录超过 24 小时会在下次 render 启动时被 cleanup_old_temp 回收
        if os.name == "nt":
            drop = 'rmdir /s /q "%s"' % self.path
        else:
            drop = "rm -rf '%s'" % self.path
        print("[work] 确认无需返修后删掉它： %s（或等 24 小时后自动回收）" % drop)
        return "kept"


def run(cmd, cwd=None, timeout=300):
    r = subprocess.run(
        cmd,
        cwd=cwd,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
        env=UTF8_ENV,
    )
    if r.stdout.strip():
        print(r.stdout.strip()[:1200])
    if r.stderr.strip():
        # 不能只在失败时看 stderr：pandoc 找不到图片只是 WARNING，
        # 静默吞掉会让人拿到一份没图的文档还以为没问题。
        line_list = [line for line in r.stderr.strip().splitlines() if line.strip()]
        print("[warn] %s 输出 %d 行诊断信息" % (os.path.basename(cmd[0]), len(line_list)))
        for line in line_list[:5]:
            print("   ", line[:160])
        if len(line_list) > 5:
            print("    ...（另 %d 行）" % (len(line_list) - 5))
    if r.returncode != 0:
        print("STDERR:", r.stderr.strip()[:2500])
        sys.exit(r.returncode)
    return r


def _extract_py_list(path, name="CONTENT_FIXES"):
    """从 .py 里取出 `name = [...]` 字面量。

    用 ast 解析而不是 import/exec——只读值，不执行用户的代码。
    """
    import ast

    tree = ast.parse(open(path, encoding="utf-8-sig").read())
    for node in tree.body:
        if isinstance(node, ast.Assign):
            for tgt in node.targets:
                if isinstance(tgt, ast.Name) and tgt.id == name:
                    return ast.literal_eval(node.value)
    sys.exit("在 %s 里没找到 %s = [...]" % (path, name))


def load_content_fixes(cfg, cfg_dir):
    """载入「编辑性替换表」：上游管线常用它删掉注释性括号、统一措辞。

    支持三种写法（可叠加）：
      "content_fixes":      [[旧, 新], ...]                  内联
      "content_fixes_file": "content_fixes.json"             JSON
      "content_fixes_file": "content_fixes.py"               取其中 CONTENT_FIXES 字面量
    相对路径按 config 文件所在目录解析。
    """
    fixes = [tuple(x) for x in (cfg.get("content_fixes") or [])]
    path = cfg.get("content_fixes_file")
    if not path:
        return fixes
    p = path if os.path.isabs(path) else os.path.join(cfg_dir, path)
    if not os.path.exists(p):
        sys.exit("content_fixes_file 不存在: " + p)
    if p.lower().endswith(".json"):
        with open(p, encoding="utf-8-sig") as f:
            data = json.load(f)
    else:
        data = _extract_py_list(p)
    fixes += list(data.items()) if isinstance(data, dict) else [tuple(x) for x in data]
    return fixes


def _pandoc():
    """返回 (可执行文件路径, 版本号)；未安装则 (None, None)。"""
    exe = shutil.which("pandoc")
    if not exe:
        return None, None
    try:
        out = subprocess.run(
            ["pandoc", "--version"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            # 版本探测必须有时限：pandoc 挂起时这里会卡住整条渲染链路（包括 --doctor）。
            # 超时与「拿不到版本」同口径回落 "?"，措辞不变（与 _evidence._pandoc_version 同档）。
            timeout=30,
        ).stdout
    except (OSError, subprocess.TimeoutExpired):
        return exe, "?"
    m = re.search(r"(\d+\.\d+(?:\.\d+)?)", out.splitlines()[0] if out else "")
    return exe, m.group(1) if m else "?"


def doctor():
    """环境自检：pandoc / Python 依赖 / Word。缺核心依赖时退出码 1。"""
    # 僵尸 Word 检查必须在渲染器探测**之前**（外部审计 R2 #3：available() 会起 COM，
    # 既有 WINWORD 实例恰恰可能干扰探测本身——告警晚了就没意义）
    if os.name == "nt":
        try:
            tl = subprocess.run(
                ["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
            if "WINWORD.EXE" in (tl.stdout or ""):
                print("[warn] 检测到已在运行的 WINWORD.EXE——既有实例可能干扰 COM 探测与导出；")
                print("       若 PDF 导出异常，先关闭所有 Word 再试。")
        except Exception:
            pass  # tasklist 不可用（非 Windows / 精简环境）时静默跳过

    rows = []

    exe, ver = _pandoc()
    rows.append(("pandoc", "%s  %s" % (ver, exe) if ver else "缺失（核心）", bool(ver)))

    for mod, label, core in (
        ("docx", "python-docx", True),
        ("lxml", "lxml", True),
        ("win32com.client", "pywin32（--pdf）", False),
        ("fitz", "PyMuPDF（--check）", False),
    ):
        try:
            importlib.import_module(mod)
            rows.append((label, "已安装", True))
        except Exception:
            rows.append((label, "缺失" + ("（核心）" if core else "（可选）"), not core))

    word_ok, word_why = False, "未探测"
    for _name in SUPPORTED_RENDERERS:
        _r = get_renderer(_name)
        _ok, _why = type(_r).available()
        rows.append(("渲染器:%s" % _name, _why, _ok))
        if _name == "word":
            # 复用循环里已探的结果：另起一次 type(...).available() 会重复拉起
            # 一次 Office COM 进程（doctor 慢的主因）
            word_ok, word_why = _ok, _why

    width = max(dw(r[0]) for r in rows)
    for name, state, _ in rows:
        print("%s%s  %s" % (name, " " * (width - dw(name)), state))

    # pandoc 版本实测范围提示（外部审计 R2 #4）：超出 3.1–3.11 的行为没有验收背书
    if ver:
        _v = re.match(r"(\d+)\.(\d+)", ver)
        if _v and not ((3, 1) <= (int(_v.group(1)), int(_v.group(2))) <= (3, 11)):
            print("[warn] pandoc %s 未经实测（实测范围 3.1–3.11）：渲染行为可能有差异" % ver)

    missing = [n for n, s, ok in rows if not ok and "缺失" in s]
    if missing:
        print("\n缺失项处理：")
        # 本仓库没有构建后端，pip install "." 一定失败；只能按依赖清单装。
        print("  pip install -r requirements.txt   # 或用 uv：uv sync --all-extras")
        if "pandoc" in missing:
            print("  winget install --id JohnMacFarlane.Pandoc   # 已装则把所在目录加进 PATH")
        return 1

    # 能力感知 preflight：不再只吐一个全局 OK，而是按管线给出 READY / NOT READY，
    # 避免用户把「docx 可用」误读成「完整 Texere 环境就绪」。
    print("\n管线就绪度：")
    print("  DOCX 渲染      %s" % ("READY" if not missing else "NOT READY"))
    print("  PDF 导出       %s" % ("READY" if word_ok else "NOT READY"))
    print("  验收(--check)  %s" % ("READY" if word_ok else "NOT READY"))
    if not word_ok:
        print(
            "\n注意：默认渲染器（Word）不可用（%s）；--pdf/--check 将失败，仅 docx 渲染可用。"
            % word_why
        )
    print("\nOK：核心依赖齐备。")
    return 0


def preflight(want_pdf, want_check, renderer_name="word"):
    """渲染前的快速预检，把裸异常换成可行动提示。"""
    if not shutil.which("pandoc"):
        sys.exit(
            "未找到 pandoc（只出 docx 也需要它）。\n"
            "  winget install --id JohnMacFarlane.Pandoc\n"
            "  已安装请将其所在目录加入 PATH，然后重跑；用 --doctor 复查。"
        )
    try:
        importlib.import_module("docx")
    except ImportError:
        sys.exit("缺少 python-docx：pip install -r requirements.txt")
    if want_pdf:
        rndr = get_renderer(renderer_name)
        ok, why = type(rndr).available()
        if not ok:
            sys.exit(
                "所选渲染器不可用（%s）：%s\n"
                "  --renderer word 需本机 Word + pywin32；libreoffice 需 soffice；"
                "wps 需本机 WPS Office。或去掉 --pdf（仍可正常产出 docx）" % (renderer_name, why)
            )
    if want_check and not want_pdf:
        print("[warn] --check 依赖 --pdf 产出的 PDF，已忽略 --check")
    if want_check:
        try:
            importlib.import_module("fitz")
        except ImportError:
            sys.exit("缺少 PyMuPDF，无法做 PDF 目视验收：\n  pip install PyMuPDF   或去掉 --check")


KNOWN_CONFIG_KEYS = frozenset(
    {
        "title",
        "author",
        "subject",
        "comments",
        "header",
        "cover",
        "toc",
        "toc_heading",
        "mode",
        "page_break_h1",
        "style",
        "caption_words",
        "content_fixes",
        "content_fixes_file",
        "reference_doc",
        "resource_paths",
    }
)


def _load_config(config_path):
    """读 config.json。

    显式传了路径却不存在 → 直接退出：静默退回默认值会做出一份「看着对」的
    错版式文档，比崩溃更难发现。
    """
    import json

    cfg = {}
    if config_path and os.path.exists(config_path):
        with open(config_path, encoding="utf-8-sig") as f:
            cfg = json.load(f)
    elif config_path:
        sys.exit(f"文件不存在：--config {config_path}")
    # 未知顶层键告警：键写错层级/拼错会被静默忽略（外部实测踩坑：
    # page_number/toc_title 写在顶层毫无作用）。白名单与 docs/CONFIG 字段表同步，
    # 并与 config.schema.json 由 tests/test_docs_sync.py 守住三方一致。
    for k in sorted(set(cfg) - KNOWN_CONFIG_KEYS):
        print(
            "[warn] config 顶层键 %r 不被识别——多半是应放进 style 段或拼写有误，"
            "键表见 docs/CONFIG.zh-CN.md" % k
        )
    return cfg


def _reject_output_colliding_with_source(out_docx, md_files, src_root):
    """--out 与源 Markdown 同路径：拒绝渲染（外部实测确凿的毁数据路径）。

    写出去的是 docx 二进制，源 .md 会被永久覆盖且无提示。src 是目录时，out 落
    到目录里**任何一个** .md 上都算撞——它们全部会被合并进正文。Windows 大小写
    不敏感（report.md 与 REPORT.MD 同文件），比较走 normcase(realpath)。
    """
    out = os.path.normcase(os.path.realpath(out_docx))
    candidates = [src_root] + list(md_files)
    for c in candidates:
        if os.path.normcase(os.path.realpath(c)) == out:
            sys.exit(
                "--out 与源 Markdown 同路径（%s）：渲染会把源文件覆盖成 docx 二进制，已拒绝。\n"
                "  给 --out 换一个文件名。" % c
            )


def _merge_sources(src_dir, config_path, cfg, tmp, out_docx=""):
    """合并 src 下全部 .md、套用 content_fixes 替换、写出 all.md。

    中文主导的文档在此做直引号配对（smart 已关，见 _compile）：all.md 落盘的
    就是 pandoc 实际吃进去的文本，证据链不打折。

    返回 (合并正文, 图片搜索锚点目录 src_root, all.md 路径, 是否中文主导)。
    """
    all_md = os.path.join(tmp, "all.md")
    # --src 既可以是目录也可以是单个 .md 文件（一页的通知不必先建目录）。
    # src_root 是图片搜索的锚点目录：单文件时取其所在目录。
    if os.path.isfile(src_dir):
        md_files = [src_dir]
        src_root = os.path.dirname(os.path.abspath(src_dir))
    else:
        md_files = sorted(glob.glob(os.path.join(src_dir, "*.md")))
        src_root = src_dir
    _reject_output_colliding_with_source(out_docx, md_files, src_root)
    parts = []
    for f in md_files:
        # utf-8-sig：源 md 带 BOM 时不至于让第一个字符变成乱码（记事本默认写 BOM）
        with open(f, encoding="utf-8-sig") as fh:
            parts.append(fh.read().rstrip() + "\n")
    if not parts:
        sys.exit("src 目录下没有 .md 文件: " + src_dir)
    merged = "\n".join(parts)

    # 编辑性替换：上游管线常用来删掉注释性括号、统一措辞。
    # 不做这一步，正文内容就和定稿版本不一致（真实项目实测差 122 处、约 1581 字）。
    cfg_dir = os.path.dirname(os.path.abspath(config_path)) if config_path else os.getcwd()
    fixes = load_content_fixes(cfg, cfg_dir)
    n_fix = 0
    for old, new in fixes:
        c = merged.count(old)
        if c:
            n_fix += c
            merged = merged.replace(old, new)
    cjk = is_cjk_dominant(merged)
    if cjk:
        fixed, n_pair, n_odd = pair_cjk_quotes(merged)
        merged = fixed.rstrip("\n") + "\n"
    with open(all_md, "w", encoding="utf-8") as fh:
        fh.write(merged)
    print("[1/3] merged %d md files (%d chars)" % (len(parts), len(merged)))
    if fixes:
        print("      content fixes: %d 处（%d 条规则）" % (n_fix, len(fixes)))
    if cjk and n_pair:
        print("      CJK 直引号配对: %d 对（smart 已关闭，全角引号不经 pandoc smart）" % n_pair)
    if cjk and n_odd:
        print(
            "[warn] %d 个段落的直引号不成对（奇数个），已整段放弃配对——请检查是否漏写引号" % n_odd
        )
    return merged, src_root, all_md, cjk


def _resource_paths(src_root, cfg):
    """算出 pandoc 的 --resource-path 去重列表。

    图片常放在 src 的子目录或**兄弟**目录里（真实项目里 md 在 src/、图在 media/），
    pandoc 只按给出的路径查找，故把 src、其全部子目录、src 的父目录及其一级子目录
    都加进 resource-path；还可用 config 的 resource_paths 补充。

    Windows 整条命令行上限 32767 字符：src 放在 /tmp 这类位置时，父目录的一级
    子目录能膨胀到上千个，直接把 pandoc 挤炸（WinError 206，实测踩过）。故设字符
    预算，超支时从尾部裁剪——父目录扩展本来就是兜底（排在最后），src 自身/子目录/
    显式 resource_paths 永不裁；真有图被裁到，pandoc 的 WARNING 会兜住。
    """
    core = [src_root]
    for root, dirs, _files in os.walk(src_root):
        core.extend(os.path.join(root, d) for d in dirs)
    core.extend(cfg.get("resource_paths") or [])
    parent = os.path.dirname(os.path.abspath(src_root))
    ext = []
    if os.path.isdir(parent):
        ext.append(parent)
        ext.extend(
            sorted(
                os.path.join(parent, d)
                for d in os.listdir(parent)
                if os.path.isdir(os.path.join(parent, d))
            )
        )

    seen, uniq = set(), []
    for p, droppable in [(x, False) for x in core] + [(x, True) for x in ext]:
        p = os.path.abspath(p)
        if p not in seen:
            seen.add(p)
            uniq.append((p, droppable))

    kept, used, dropped = [], 0, 0
    for p, droppable in uniq:
        if droppable and used + len(p) + 1 > RESOURCE_PATH_BUDGET:
            dropped += 1
            continue
        kept.append(p)
        used += len(p) + 1
    if dropped:
        print(
            "[warn] --resource-path 超出预算（%d 字符），已裁掉 %d 个父目录扩展；"
            "图片搜不到时用 config 的 resource_paths 显式指定" % (RESOURCE_PATH_BUDGET, dropped)
        )
    return kept


def _export_and_check(out_docx, renderer_name, want_check, keep_pages, tmp):
    """导 PDF（刷域写回），并按需跑 check_pdf.py 目视验收；返回 pdf 路径。"""
    pdf = os.path.splitext(out_docx)[0] + ".pdf"
    rndr = get_renderer(renderer_name)
    res = rndr.render(out_docx, pdf, save_updated_fields=True)
    if not res.ok:
        sys.exit(
            "PDF 导出失败（%s）：%s"
            % (renderer_name, (res.errors[0] if res.errors else "未知错误"))
        )
    for w in res.warnings:
        print("[warn] %s" % w)
    if want_check:
        run([sys.executable, os.path.join(KIT, "scripts", "check_pdf.py"), pdf])
        # check_pages/ 默认随临时目录清理（外部反馈：29 张 PNG 散落工作目录）；
        # 位置在最后的 ✓ 完成行里说明（--keep-pages 留在 PDF 同目录）
        pages_dir = os.path.join(os.path.dirname(os.path.abspath(pdf)), "check_pages")
        if os.path.isdir(pages_dir) and not keep_pages:
            shutil.move(pages_dir, os.path.join(tmp, "check_pages"))
    return pdf


def _page_count(pdf):
    """PDF 页数，读不到（缺 PyMuPDF 等）返回 None——摘要行就不报页数。"""
    try:
        import pymupdf

        with pymupdf.open(pdf) as d:
            return d.page_count
    except Exception:
        return None


def _print_summary(out_docx, pdf, want_check, keep_pages, renderer_name="word"):
    """最后一行以用户视角回答「成功了吗、在哪、几页」；管线细节在上面各步。

    docx-only 报文件大小（python-docx 数不了页数，不假装知道）；
    --pdf 报页数并指路预览；--check --keep-pages 顺带说明截图在哪。
    LibreOffice 是唯一不能把刷新后的域写回 docx 的渲染器（Word/WPS 走 COM）——
    那是交付级事实（TOC 域没刷新），在交付行里说明，别缩在步骤日志中间。
    """
    if pdf:
        head = "%s + %s" % (out_docx, pdf)
        n = _page_count(pdf)
        head += "（%d 页，可打开 PDF 预览）" % n if n else "（可打开 PDF 预览）"
        if want_check and keep_pages:
            pages_dir = os.path.join(os.path.dirname(os.path.abspath(pdf)), "check_pages")
            if os.path.isdir(pages_dir):
                head += "；页面截图在 %s" % pages_dir
        if renderer_name == "libreoffice":
            head += (
                "；注意：LibreOffice 不会把刷新后的域写回 docx，交付 docx 前请用 Word/WPS 打开刷新"
            )
    else:
        size_kb = max(1, os.path.getsize(out_docx) // 1024)
        head = "%s（%d KB）；需要 PDF/真机验收时追加 --pdf --check" % (out_docx, size_kb)
    print("✓ 完成：" + head)


def render(
    src_dir,
    out_docx,
    config_path,
    want_pdf,
    want_check,
    renderer_name="word",
    keep_pages=False,
    work_mode="auto",
):
    """渲染编排：读配置 → 合并 md → pandoc 出正文 → post 排版 → 可选导 PDF 验收。"""
    preflight(want_pdf, want_check, renderer_name)
    cfg = _load_config(config_path)
    ref = cfg.get("reference_doc") or os.path.join(KIT, "assets", "ref.docx")
    # --out 指向不存在的目录时直接建好（外部审计：pandoc/SaveAs 遇缺父目录直接挂）
    out_parent = os.path.dirname(os.path.abspath(out_docx))
    if out_parent:
        os.makedirs(out_parent, exist_ok=True)

    wd = _Workdir(tempfile.mkdtemp(prefix="texere_"))
    tmp = wd.path

    merged, src_root, all_md, cjk = _merge_sources(src_dir, config_path, cfg, tmp, out_docx)
    body = os.path.join(tmp, "body.docx")
    uniq = _resource_paths(src_root, cfg)
    run(_build_pandoc_cmd(all_md, body, ref, uniq, cfg, cjk=cjk))
    print("[2/3] pandoc -> body.docx")

    run(
        [
            sys.executable,
            os.path.join(KIT, "scripts", "post.py"),
            body,
            out_docx,
            config_path or "",
        ]
    )
    print("[3/3] postprocess ->", out_docx)
    images_ok = check_images(merged, out_docx)

    pdf = None
    if want_pdf:
        pdf = _export_and_check(out_docx, renderer_name, want_check, keep_pages, tmp)
    # 缺图时不在 check_images 里立即退出：先走完 PDF 导出与 --check（用户恰恰需要
    # 这些产物肉眼确认丢了哪几张图），再让退出码诚实反映「这份交付物没图」。
    # 与 validate 的 image_embedding FAIL→exit(1) 契约对齐（外部实测：只跑 render
    # 不接 validate 的 CI 场景，此前缺图仍拿到退出码 0）。
    if images_ok:
        # 交付物齐备：中间产物去留按「人是否在场」分流（见 _Workdir）
        wd.settle(work_mode)
    else:
        # 这份件本来就不合格，不留现场（诊断靠上面的 [ERROR] 与补救建议）
        wd.settle(_Workdir.DISCARD)
    if not images_ok:
        sys.exit(1)
    # 最后一行永远回答用户真正想问的：成功了吗、文件在哪、几页
    _print_summary(out_docx, pdf, want_check, keep_pages, renderer_name)


def check_images(md_text, docx_path):
    """源 md 里写了图、文档里却没嵌进去时，必须大声报错。

    这是真实项目里最容易翻车的一环：pandoc 找不到图只给 WARNING，
    静默过去就会交付一份没图的标书。

    返回 True 表示图片齐备（或源里根本没引图），False 表示有图缺失——
    调用方（render）据此在产出全部交付物后以非零码结束。
    """
    # 围栏里的 `![...]` 是示例代码不是引图（与 validate._source_md_segments 同一条
    # 规则）：不排掉就把围栏算进引用数，凭空报缺图并以退出码 1 结束。
    n_ref = len(re.findall(r"!\[", _outside_code_fences(md_text)))
    if not n_ref:
        return True
    from docx import Document

    n_img = len(Document(docx_path).inline_shapes)
    if n_img >= n_ref:
        print("images: %d/%d ok" % (n_img, n_ref))
        return True
    print("[ERROR] 源 md 引用 %d 张图，文档里只嵌进 %d 张" % (n_ref, n_img))
    print("        图片目录不在搜索范围内时就会这样；")
    print('        在 config 里加 "resource_paths": ["图片目录"] 补充搜索路径。')
    print("        本次运行将以退出码 1 结束。")
    return False


def _work_mode(a):
    """中间产物的处置模式；两个开关互斥，都不给就是 auto。"""
    if a.keep_work and a.discard_work:
        sys.exit("--keep-work 与 --discard-work 互斥：一个留现场等验收，一个跑完即删")
    if a.keep_work:
        return _Workdir.KEEP
    if a.discard_work:
        return _Workdir.DISCARD
    return _Workdir.AUTO


def main():
    # 副作用只在入口执行：被 import（工具复用/测试）时不碰 stdio、不扫临时目录
    force_utf8_stdio()
    cleanup_old_temp()

    ap = argparse.ArgumentParser(
        prog="render.py",
        description="一键渲染：Markdown -> 正式中文 docx（可选 PDF 导出与目视验收）。",
        epilog=(
            "最短可跑：\n"
            "  python scripts/render.py --src report.md --out report.docx"
            "          # 秒级出 docx，零配置\n"
            "  python scripts/render.py --src report.md --out report.docx --pdf --check\n"
            "  python scripts/render.py --sample           # 自带样例冒烟测试\n"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument("--src", help="Markdown 源：单个 .md 文件或目录（目录按文件名序合并全部 .md）")
    ap.add_argument(
        "--out", help="输出 docx 路径（父目录自动创建；同名文件会被覆盖；与源 md 同路径会被拒绝）"
    )
    ap.add_argument(
        "--config", help="config.json 路径（可省：全部键都有默认值，键表见 docs/CONFIG.md）"
    )
    ap.add_argument(
        "--pdf",
        action="store_true",
        help="同时导出 PDF（需本机渲染器：默认 Word，可 --renderer 换 LibreOffice/WPS）",
    )
    ap.add_argument(
        "--check",
        action="store_true",
        help="导出后做 PDF 目视验收：报空白页并渲染截图（需 PyMuPDF；自动启用 --pdf）",
    )
    ap.add_argument(
        "--keep-pages",
        action="store_true",
        help="--check 的页面截图保留在 PDF 同目录 check_pages/（默认移到中间产物目录）",
    )
    ap.add_argument(
        "--keep-work",
        action="store_true",
        help="跑完保留中间产物（all.md / body.docx / 截图）不删，返修时能直接改不必重跑",
    )
    ap.add_argument(
        "--discard-work",
        action="store_true",
        help="跑完立即回收中间产物（脚本 / CI 用）；默认终端里保留 24 小时，非交互时立即回收",
    )
    ap.add_argument(
        "--renderer",
        choices=list(SUPPORTED_RENDERERS),
        default="word",
        help="PDF 导出渲染器：word（默认，需本机 Word）/ libreoffice / wps",
    )
    ap.add_argument(
        "--sample", action="store_true", help="用自带样例跑一遍完整渲染（冒烟测试，零配置）"
    )
    ap.add_argument(
        "--doctor",
        action="store_true",
        help="只做环境自检：pandoc / Python 依赖 / Word，不渲染",
    )
    ap.add_argument("--version", action="version", version="texere " + __version__)
    a = ap.parse_args()
    work_mode = _work_mode(a)

    if a.doctor:
        sys.exit(doctor())

    if a.sample:
        tmp = tempfile.mkdtemp(prefix="texere_sample_")
        # 这里以前完全没清理：每跑一次 --sample 就在 %TEMP% 留一个目录
        atexit.register(shutil.rmtree, tmp, ignore_errors=True)
        shutil.copy(os.path.join(KIT, "assets", "sample.md"), os.path.join(tmp, "01_sample.md"))
        out = os.path.join(KIT, "sample_out.docx")
        render(
            tmp,
            out,
            os.path.join(KIT, "assets", "sample_config.json"),
            True,
            True,
            "word",
            work_mode=work_mode,
        )
        print("sample ok ->", out)
        return

    if not a.src or not a.out:
        sys.exit("需要 --src 与 --out，或使用 --sample")
    if not os.path.exists(a.src):
        sys.exit("找不到 --src: " + a.src)
    if a.check and not a.pdf:
        print("[note] --check 依赖 PDF，已自动启用 --pdf")
        a.pdf = True
    render(
        a.src,
        a.out,
        a.config,
        a.pdf,
        a.check,
        a.renderer,
        keep_pages=a.keep_pages,
        work_mode=work_mode,
    )


if __name__ == "__main__":
    main()
