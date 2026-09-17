# -*- coding: utf-8 -*-
"""PDF 目视验收：python check_pdf.py <file.pdf> [--max-empty N] [页码...]

输出总页数与"近空白页"清单（正文<60字且无图），并把指定页（默认 1-3 页）
渲染为 PNG 到 PDF 同目录的 check_pages/，供肉眼检查版式。

退出码（供 CI 使用）：
  0 = 通过
  1 = PDF 无页，或近空白页数超过 --max-empty（默认 0）

  python check_pdf.py out.pdf && echo OK
"""
import os
import sys

import fitz

if len(sys.argv) < 2:
    sys.exit("用法: python check_pdf.py <file.pdf> [--max-empty N] [页码...]")

_args = sys.argv[1:]
PDF = _args[0]
max_empty = 0
pages_want = []
_i = 1
while _i < len(_args):
    if _args[_i] == "--max-empty":
        if _i + 1 >= len(_args):
            sys.exit("--max-empty 需要一个整数参数")
        max_empty = int(_args[_i + 1])
        _i += 2
    else:
        pages_want.append(int(_args[_i]))
        _i += 1

doc = fitz.open(PDF)
print("pages:", doc.page_count)

sparse = []
for i, page in enumerate(doc):
    text = page.get_text().strip()
    imgs = len(page.get_images(full=True))
    if len(text) < 60 and imgs == 0:
        sparse.append((i + 1, len(text), text[:40].replace("\n", " / ")))
print("near-empty pages:", len(sparse), "(max allowed: %d)" % max_empty)
for s in sparse:
    print("   p%-3d chars=%-4d %s" % s)

out_dir = os.path.join(os.path.dirname(os.path.abspath(PDF)), "check_pages")
os.makedirs(out_dir, exist_ok=True)
targets = pages_want or [1, 2, 3]
for pno in targets:
    if 1 <= pno <= doc.page_count:
        pix = doc[pno - 1].get_pixmap(dpi=100)
        path = os.path.join(out_dir, "p%03d.png" % pno)
        pix.save(path)
        print("rendered", path)

if doc.page_count == 0:
    print("FAIL: PDF 没有页")
    sys.exit(1)
if len(sparse) > max_empty:
    print("FAIL: 近空白页 %d 页，超过阈值 %d" % (len(sparse), max_empty))
    sys.exit(1)
print("PASS: 空白页检查通过")
