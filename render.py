# -*- coding: utf-8 -*-
"""一键渲染：Markdown 章节目录 -> 正式中文 docx（可选 PDF 与目视验收）。

用法:
  python render.py --src <md目录> --out <out.docx> [--config config.json] [--pdf] [--check]
  python render.py --sample            # 用自带 sample.md 做冒烟测试

依赖: pandoc、python-docx；--pdf 需本机 Word + pywin32；--check 需 PyMuPDF。
说明: 若招标方/甲方提供强制格式模板 docx，可在 config 中设 "reference_doc" 指向它，
      正文样式即继承对方模板（封面/目录/页眉页脚仍由 post.py 统一注入）。
"""
import argparse
import glob
import importlib
import os
import re
import shutil
import subprocess
import sys
import tempfile

KIT = os.path.dirname(os.path.abspath(__file__))
__version__ = "0.2.0"          # 与 pyproject.toml 的 version 保持一致

# Windows 控制台默认 GBK，子进程输出里若出现 GBK 无法编码的字符（如 PyMuPDF 解出的
# U+FFFD），print 会抛 UnicodeEncodeError 让验收环节崩掉。这里保持控制台原编码不变
# （改成 utf-8 反而会让控制台显示乱码），只把无法编码的字符降级为 ?。
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(errors="replace")




def run(cmd, cwd=None):
    r = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    if r.stdout.strip():
        print(r.stdout.strip()[:1200])
    if r.stderr.strip():
        # 不能只在失败时看 stderr：pandoc 找不到图片只是 WARNING，
        # 静默吞掉会让人拿到一份没图的文档还以为没问题。
        lines = [l for l in r.stderr.strip().splitlines() if l.strip()]
        print("[warn] %s 输出 %d 行诊断信息" % (os.path.basename(cmd[0]), len(lines)))
        for l in lines[:5]:
            print("   ", l[:160])
        if len(lines) > 5:
            print("    ...（另 %d 行）" % (len(lines) - 5))
    if r.returncode != 0:
        print("STDERR:", r.stderr.strip()[:2500])
        sys.exit(r.returncode)
    return r


def _pandoc():
    """返回 (可执行文件路径, 版本号)；未安装则 (None, None)。"""
    exe = shutil.which("pandoc")
    if not exe:
        return None, None
    try:
        out = subprocess.run(["pandoc", "--version"], capture_output=True, text=True,
                             encoding="utf-8", errors="replace").stdout
    except OSError:
        return exe, "?"
    m = re.search(r"(\d+\.\d+(?:\.\d+)?)", out.splitlines()[0] if out else "")
    return exe, m.group(1) if m else "?"


def _word_engine():
    """实测验收引擎身份。

    不要用注册表 CurVer 猜：那可能是旧 Office 卸载后残留的键值
    （本机 CurVer 写着 Word.Application.11，实际却是 Microsoft Word 16.0）。
    直接起 COM 问它自己是谁。
    """
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError:
        return "未安装 pywin32，跳过探测（--pdf 需要它）"
    try:
        pythoncom.CoInitialize()
        try:
            w = win32.DispatchEx("Word.Application")
            name = "%s %s" % (w.Name, w.Version)
            path = w.Path
            w.Quit()
        finally:
            pythoncom.CoUninitialize()
    except Exception as e:
        return "启动失败（%s：%s）" % (type(e).__name__, e)
    tag = "" if "Microsoft" in name else "  [非 Microsoft Word，验收结果仅供参考]"
    return "%s  %s%s" % (name, path, tag)


def doctor():
    """环境自检：pandoc / Python 依赖 / Word。缺核心依赖时退出码 1。"""
    rows = []

    exe, ver = _pandoc()
    rows.append(("pandoc", "%s  %s" % (ver, exe) if ver else "缺失（核心）", bool(ver)))

    for mod, label, core in (("docx", "python-docx", True), ("lxml", "lxml", True),
                             ("win32com.client", "pywin32（--pdf）", False),
                             ("fitz", "PyMuPDF（--check）", False)):
        try:
            importlib.import_module(mod)
            rows.append((label, "已安装", True))
        except Exception:
            rows.append((label, "缺失" + ("（核心）" if core else "（可选）"), not core))

    rows.append(("Word 引擎", _word_engine(), True))

    def dw(s):  # 中文按 2 列宽计算
        return len(s) + sum(1 for c in s if ord(c) > 127)

    width = max(dw(r[0]) for r in rows)
    for name, state, _ in rows:
        print("%s%s  %s" % (name, " " * (width - dw(name)), state))

    missing = [n for n, s, ok in rows if not ok and "缺失" in s]
    if missing:
        print("\n缺失项处理：")
        print("  pip install \".[pdf,check]\"   # 或 uv sync --all-extras")
        if "pandoc" in missing:
            print("  winget install --id JohnMacFarlane.Pandoc   # 已装则把所在目录加进 PATH")
        return 1
    print("\nOK：核心依赖齐备。")
    return 0


def preflight(want_pdf, want_check):
    """渲染前的快速预检，把裸异常换成可行动提示。"""
    if not shutil.which("pandoc"):
        sys.exit("未找到 pandoc（只出 docx 也需要它）。\n"
                 "  winget install --id JohnMacFarlane.Pandoc\n"
                 "  已安装请将其所在目录加入 PATH，然后重跑；用 --doctor 复查。")
    try:
        importlib.import_module("docx")
    except ImportError:
        sys.exit("缺少 python-docx：pip install \".\" 或 pip install \"python-docx>=1.1,<2.0\"")
    if want_pdf:
        try:
            importlib.import_module("win32com.client")
        except ImportError:
            sys.exit("缺少 pywin32，无法做 Word 验收与导 PDF：\n"
                     "  pip install \".[pdf]\"   或去掉 --pdf（仍可正常产出 docx）")
    if want_check and not want_pdf:
        print("[warn] --check 依赖 --pdf 产出的 PDF，已忽略 --check")
    if want_check:
        try:
            importlib.import_module("fitz")
        except ImportError:
            sys.exit("缺少 PyMuPDF，无法做 PDF 目视验收：\n"
                     "  pip install \".[check]\"  或去掉 --check")


def render(src_dir, out_docx, config_path, want_pdf, want_check):
    import json
    preflight(want_pdf, want_check)
    cfg = {}
    if config_path and os.path.exists(config_path):
        cfg = json.load(open(config_path, encoding="utf-8"))
    ref = cfg.get("reference_doc") or os.path.join(KIT, "ref.docx")

    tmp = tempfile.mkdtemp(prefix="docxkit_")
    all_md = os.path.join(tmp, "all.md")
    parts = []
    for f in sorted(glob.glob(os.path.join(src_dir, "*.md"))):
        parts.append(open(f, encoding="utf-8").read().rstrip() + "\n")
    if not parts:
        sys.exit("src 目录下没有 .md 文件: " + src_dir)
    merged = "\n".join(parts)
    with open(all_md, "w", encoding="utf-8") as fh:
        fh.write(merged)
    print("[1/3] merged %d md files (%d chars)" % (len(parts), sum(len(p) for p in parts)))

    body = os.path.join(tmp, "body.docx")
    # 图片常放在 src 的子目录或**兄弟**目录里（真实项目里 md 在 src/、图在 media/），
    # pandoc 只按给出的路径查找，故把 src、其全部子目录、src 的父目录及其子目录
    # 都加进 resource-path；还可用 config 的 resource_paths 补充。
    res_paths = [src_dir]
    for root, dirs, _files in os.walk(src_dir):
        res_paths.extend(os.path.join(root, d) for d in dirs)
    parent = os.path.dirname(os.path.abspath(src_dir))
    if os.path.isdir(parent):
        res_paths.append(parent)
        for d in sorted(os.listdir(parent)):
            p = os.path.join(parent, d)
            if os.path.isdir(p):
                res_paths.append(p)
    res_paths.extend(cfg.get("resource_paths") or [])
    seen, uniq = set(), []
    for p in res_paths:
        p = os.path.abspath(p)
        if p not in seen:
            seen.add(p)
            uniq.append(p)
    lua_filter = os.path.join(KIT, "filters", "captions.lua")
    cmd = ["pandoc", all_md, "-o", body,
           "--reference-doc=" + ref,
           "--resource-path=" + os.pathsep.join(uniq),
           "-f", "markdown+pipe_tables+raw_html", "--wrap=none"]
    if os.path.exists(lua_filter):
        # AST 层标记表题/图注，post.py 就不用再靠正则猜
        cmd.append("--lua-filter=" + lua_filter)
        # 题注关键字可配：同一份配置同时喂给 lua filter 与 post.py
        cw = cfg.get("caption_words") or {}
        for key, meta_name in (("table", "dk-table-words"),
                               ("figure", "dk-figure-words")):
            if cw.get(key):
                cmd += ["-M", "%s=%s" % (meta_name, ",".join(cw[key]))]
    else:
        print("[warn] 缺少 filters/captions.lua，题注退回文本正则判定")
    run(cmd)
    print("[2/3] pandoc -> body.docx")

    run([sys.executable, os.path.join(KIT, "post.py"), body, out_docx,
         config_path or ""])
    print("[3/3] postprocess ->", out_docx)
    check_images(merged, out_docx)

    if want_pdf:
        pdf = os.path.splitext(out_docx)[0] + ".pdf"
        run([sys.executable, os.path.join(KIT, "finalize.py"), out_docx, pdf])
        if want_check:
            run([sys.executable, os.path.join(KIT, "check_pdf.py"), pdf])
    shutil.rmtree(tmp, ignore_errors=True)


def check_images(md_text, docx_path):
    """源 md 里写了图、文档里却没嵌进去时，必须大声报错。

    这是真实项目里最容易翻车的一环：pandoc 找不到图只给 WARNING，
    静默过去就会交付一份没图的标书。
    """
    n_ref = len(re.findall(r"!\[", md_text))
    if not n_ref:
        return
    from docx import Document
    n_img = len(Document(docx_path).inline_shapes)
    if n_img >= n_ref:
        print("images: %d/%d ok" % (n_img, n_ref))
        return
    print("[ERROR] 源 md 引用 %d 张图，文档里只嵌进 %d 张" % (n_ref, n_img))
    print("        图片目录不在搜索范围内时就会这样；")
    print("        在 config 里加 \"resource_paths\": [\"图片目录\"] 补充搜索路径。")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src")
    ap.add_argument("--out")
    ap.add_argument("--config")
    ap.add_argument("--pdf", action="store_true")
    ap.add_argument("--check", action="store_true")
    ap.add_argument("--sample", action="store_true")
    ap.add_argument("--doctor", action="store_true",
                    help="只做环境自检：pandoc / Python 依赖 / Word，不渲染")
    ap.add_argument("--version", action="version",
                    version="docx-kit " + __version__)
    a = ap.parse_args()

    if a.doctor:
        sys.exit(doctor())

    if a.sample:
        tmp = tempfile.mkdtemp(prefix="docxkit_sample_")
        shutil.copy(os.path.join(KIT, "sample.md"), os.path.join(tmp, "01_sample.md"))
        out = os.path.join(KIT, "sample_out.docx")
        render(tmp, out, os.path.join(KIT, "sample_config.json"), True, True)
        print("sample ok ->", out)
        return

    if not a.src or not a.out:
        sys.exit("需要 --src 与 --out，或使用 --sample")
    if a.check and not a.pdf:
        print("[note] --check 依赖 PDF，已自动启用 --pdf")
        a.pdf = True
    render(a.src, a.out, a.config, a.pdf, a.check)


if __name__ == "__main__":
    main()
