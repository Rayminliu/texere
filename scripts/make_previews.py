"""为示例生成渲染效果预览图：python scripts/make_previews.py

把 examples/ 下每个示例渲染成 PDF（需本机 Word + pandoc），再挑选代表页
输出到 assets/previews/<示例名>.png，供主 README 画廊引用。
示例版式变更后重跑本脚本即可刷新预览。
"""

import os
import shutil
import subprocess
import sys

import fitz  # PyMuPDF

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(KIT, "assets", "previews")
TMP = os.path.join(KIT, "_previews_tmp")

EXAMPLES = ["tender", "gongwen", "form", "minutes", "report", "contract", "tables"]


def render_pdf(name):
    """渲染示例为 PDF，返回 pdf 路径。"""
    src = os.path.join(KIT, "examples", name)
    docx = os.path.join(TMP, name + ".docx")
    cfg = os.path.join(src, "config.json")
    r = subprocess.run(
        [sys.executable, os.path.join(KIT, "scripts", "render.py"),
         "--src", src, "--out", docx, "--config", cfg, "--pdf"],
        capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=300,
    )
    if r.returncode != 0:
        print("[FAIL] %s: %s" % (name, (r.stderr or r.stdout)[-300:]))
        return None
    pdf = os.path.splitext(docx)[0] + ".pdf"
    return pdf if os.path.exists(pdf) else None


def pick_page(doc):
    """挑代表页：跳过封面/目录，取正文里信息量最大的一页（非空、行数适中偏多）。"""
    best, best_score = None, -1
    for i in range(len(doc)):
        lines = [ln for ln in doc[i].get_text("text", sort=True).splitlines() if ln.strip()]
        n = len(lines)
        # 封面/目录页跳过：首页常为封面；含大量点线（TOC 引导符）的是目录页
        dots = sum(1 for ln in lines if "...." in ln)
        if i == 0 or dots > 3:
            continue
        score = min(n, 40) + (8 if any(ln.strip().startswith("表") for ln in lines) else 0)
        if n >= 6 and score > best_score:
            best, best_score = i, score
    return best if best is not None else min(1, len(doc) - 1)


def save_preview(pdf, page_idx, out_png, zoom=1.4):
    doc = fitz.open(pdf)
    try:
        pix = doc[page_idx].get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        pix.save(out_png)
        print("preview: %s (p%d, %dx%d)" % (os.path.basename(out_png), page_idx + 1, pix.width, pix.height))
    finally:
        doc.close()


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    os.makedirs(TMP, exist_ok=True)
    try:
        for name in EXAMPLES:
            pdf = render_pdf(name)
            if not pdf:
                continue
            doc = fitz.open(pdf)
            try:
                idx = pick_page(doc)
            finally:
                doc.close()
            save_preview(pdf, idx, os.path.join(OUT_DIR, name + ".png"))
    finally:
        shutil.rmtree(TMP, ignore_errors=True)


if __name__ == "__main__":
    main()
