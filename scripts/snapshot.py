"""PDF 版式快照回归：python scripts/snapshot.py <file.pdf> [--update] [--max-diff R] [--dpi N]

把 PDF 每页渲染成 PNG，与 baselines/ 下的基线逐像素比对。
差异比例超过 --max-diff（默认 0.5%）或页数不一致即退出码 1。

用途：改了 ref.docx / post.py 之后跑一遍，确认版式没被悄悄改坏。

  python scripts/snapshot.py out.pdf              # 比对
  python scripts/snapshot.py out.pdf --update     # 重录基线（确认版式变更是有意的时候）
  python scripts/snapshot.py out.pdf --dpi 150    # 更高精度（更慢、更敏感）
"""

import json
import operator
import os
import sys

import fitz

DEFAULT_DPI = 100
# 0.1%。实测：同一文档重复导出 PDF 的差异为 0.00%，而改一个页眉文字会产生 0.16%，
# 所以阈值必须压到 0.1% 才能抓住这种"小但真实"的漂移；0.5% 会直接漏报。
DEFAULT_MAX_DIFF = 0.001


def parse_args(argv):
    pdf, dpi, max_diff, update = None, DEFAULT_DPI, DEFAULT_MAX_DIFF, False
    i = 0
    while i < len(argv):
        a = argv[i]
        if a == "--update":
            update = True
        elif a == "--dpi":
            dpi = int(argv[i + 1])
            i += 1
        elif a == "--max-diff":
            max_diff = float(argv[i + 1])
            i += 1
        elif a.startswith("-"):
            sys.exit("未知参数: " + a)
        elif pdf is None:
            pdf = a
        i += 1
    if pdf is None:
        sys.exit("用法: python scripts/snapshot.py <file.pdf> [--update] [--max-diff R] [--dpi N]")
    return pdf, dpi, max_diff, update


def diff_ratio(a, b):
    """逐字节差异比例；长度不同视为全变。"""
    if len(a) != len(b):
        return 1.0
    if a == b:
        return 0.0
    return sum(map(operator.ne, a, b)) / len(a)


def main(argv):
    pdf, dpi, max_diff, update = parse_args(argv)
    if not os.path.exists(pdf):
        sys.exit("找不到 PDF: " + pdf)

    base_dir = os.path.join(os.path.dirname(os.path.abspath(pdf)), "baselines")
    meta_path = os.path.join(base_dir, "meta.json")

    doc = fitz.open(pdf)
    samples = [doc[i].get_pixmap(dpi=dpi).samples for i in range(doc.page_count)]

    if update:
        os.makedirs(base_dir, exist_ok=True)
        for i in range(doc.page_count):
            doc[i].get_pixmap(dpi=dpi).save(os.path.join(base_dir, "p%03d.png" % (i + 1)))
        json.dump(
            {"dpi": dpi, "pages": doc.page_count},
            open(meta_path, "w", encoding="utf-8"),
        )
        print("baseline recorded: %d pages @ %d dpi -> %s" % (doc.page_count, dpi, base_dir))
        return 0

    if not os.path.exists(meta_path):
        sys.exit("没有基线，先跑：python scripts/snapshot.py %s --update" % os.path.basename(pdf))
    meta = json.load(open(meta_path, encoding="utf-8"))
    if meta.get("dpi") != dpi:
        print(
            "[warn] 基线 dpi=%s，本次 %s，结果不可比；用 --dpi %s 或重录基线"
            % (meta.get("dpi"), dpi, meta.get("dpi"))
        )
        return 1
    if meta.get("pages") != doc.page_count:
        print("FAIL: 页数 %d != 基线 %d" % (doc.page_count, meta.get("pages")))
        return 1

    worst, bad = 0.0, []
    for i in range(doc.page_count):
        bp = os.path.join(base_dir, "p%03d.png" % (i + 1))
        if not os.path.exists(bp):
            print("FAIL: 缺少基线 %s" % bp)
            return 1
        r = diff_ratio(samples[i], fitz.Pixmap(bp).samples)
        worst = max(worst, r)
        if r > max_diff:
            bad.append((i + 1, r))
    for pno, r in bad:
        print("   p%-3d diff=%.2f%%" % (pno, r * 100))
    if bad:
        print(
            "FAIL: %d 页版式漂移，最大 %.2f%%（阈值 %.2f%%）"
            % (len(bad), worst * 100, max_diff * 100)
        )
        return 1
    print("PASS: %d 页与基线一致（最大差异 %.2f%%）" % (doc.page_count, worst * 100))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
