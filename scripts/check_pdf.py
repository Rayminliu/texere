"""PDF 目视验收：python scripts/check_pdf.py <file.pdf> [--max-empty N] [页码...]

输出总页数与"近空白页"清单（正文<60字且无图），并把指定页（默认 1-3 页）
渲染为 PNG 到 PDF 同目录的 check_pages/，供肉眼检查版式。

退出码（供 CI 使用）：
  0 = 通过
  1 = PDF 无页，或近空白页数超过 --max-empty（默认 0）
  2 = 命令行参数不合法（argparse）

  python scripts/check_pdf.py out.pdf && echo OK
"""

import argparse
import os
import sys

import pymupdf
from _shared import BASELINE_DPI, force_utf8_stdio

force_utf8_stdio()

NEAR_EMPTY_CHARS = 60  # 正文少于这个字符数且无图 → 视为近空白页
DEFAULT_PAGES = (1, 2, 3)
# DPI 单一来源：口径同 snapshot.py / validate.py，收敛到 _shared.BASELINE_DPI，
# 避免同一文档在不同工具里因 DPI 不同得到相反的视觉漂移结论。
SNAPSHOT_DPI = BASELINE_DPI


def _build_parser():
    p = argparse.ArgumentParser(
        prog="check_pdf.py",
        description="PDF 目视验收：报总页数与近空白页清单，并渲染指定页为 PNG 供肉眼检查。",
        epilog="退出码：0=通过；1=PDF 无页或近空白页超过 --max-empty；2=参数不合法。",
    )
    p.add_argument("pdf", help="待检查的 PDF")
    p.add_argument(
        "pages",
        nargs="*",
        type=int,
        metavar="PAGE",
        help="要渲染成 PNG 的页码（默认 %s）" % " ".join(str(x) for x in DEFAULT_PAGES),
    )
    p.add_argument(
        "--max-empty",
        dest="max_empty",
        type=int,
        default=0,
        help="允许的近空白页数（默认 0）",
    )
    return p


def find_near_empty(doc):
    """近空白页清单：正文 <60 字且没有图片。有图的封面/附图页不算空。"""
    sparse = []
    for i, page in enumerate(doc):
        text = page.get_text().strip()
        imgs = len(page.get_images(full=True))
        if len(text) < NEAR_EMPTY_CHARS and imgs == 0:
            sparse.append((i + 1, len(text), text[:40].replace("\n", " / ")))
    return sparse


def render_pages(doc, pdf, wanted):
    """把 wanted 里的页渲染成 PNG 到 PDF 同目录的 check_pages/，返回写出的路径。"""
    out_dir = os.path.join(os.path.dirname(os.path.abspath(pdf)), "check_pages")
    os.makedirs(out_dir, exist_ok=True)
    written = []
    for pno in wanted or list(DEFAULT_PAGES):
        if 1 <= pno <= doc.page_count:
            pix = doc[pno - 1].get_pixmap(dpi=SNAPSHOT_DPI)
            path = os.path.join(out_dir, "p%03d.png" % pno)
            pix.save(path)
            written.append(path)
    return written


def main(argv=None):
    a = _build_parser().parse_args(argv)
    doc = pymupdf.open(a.pdf)
    print("pages:", doc.page_count)

    sparse = find_near_empty(doc)
    print("near-empty pages:", len(sparse), "(max allowed: %d)" % a.max_empty)
    for s in sparse:
        print("   p%-3d chars=%-4d %s" % s)

    for path in render_pages(doc, a.pdf, a.pages):
        print("rendered", path)

    if doc.page_count == 0:
        print("FAIL: PDF 没有页")
        return 1
    if len(sparse) > a.max_empty:
        print("FAIL: 近空白页 %d 页，超过阈值 %d" % (len(sparse), a.max_empty))
        return 1
    print("PASS: 空白页检查通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
