"""PDF 版式快照回归：python scripts/snapshot.py <file.pdf> [--update] [--renderer R] [--max-diff R] [--dpi N]

把 PDF 每页渲染成 PNG，与 baselines/<renderer>/ 下的基线逐像素比对。
差异比例超过 --max-diff（默认 0.1%）或页数不一致即退出码 1。

用途：改了 ref.docx / post.py 之后跑一遍，确认版式没被悄悄改坏。

基线按渲染器分目录（baselines/word/ 等）：Word 与 LibreOffice 的渲染像素
天然不可比，混录会满屏假漂移——比对时 renderer 不一致直接拒绝。

  python scripts/snapshot.py out.pdf                        # 比对（默认 word 基线）
  python scripts/snapshot.py out.pdf --update               # 重录基线（确认版式变更是有意的时候）
  python scripts/snapshot.py out.pdf --renderer libreoffice --update
  python scripts/snapshot.py out.pdf --dpi 150              # 更高精度（更慢、更敏感）
"""

import argparse
import json
import os
import sys

import pymupdf
from _shared import BASELINE_DPI as DEFAULT_DPI
from _shared import DEFAULT_MAX_DIFF, force_utf8_stdio
from _visual_diff import diff_ratio

# 阈值单一事实源在 _shared：0.1%。实测：同一文档重复导出 PDF 的差异为 0.00%，
# 而改一个页眉文字会产生 0.16%，所以阈值必须压到 0.1% 才能抓住这种"小但真实"的漂移；0.5% 会直接漏报。
RENDERERS = ("word", "libreoffice", "wps")


def _build_parser():
    p = argparse.ArgumentParser(
        prog="snapshot.py",
        description="PDF 版式快照回归：逐页与 baselines/<renderer>/ 下的基线比像素。",
        epilog="典型用法：snapshot.py out.pdf --update 重录基线；"
        "snapshot.py out.pdf --renderer libreoffice 比对。",
    )
    p.add_argument("pdf", help="待校验（或待录入基线）的 PDF")
    p.add_argument(
        "--update", action="store_true", help="用本 PDF 重录基线（确认版式变更是有意时）"
    )
    p.add_argument(
        "--renderer", default="word", choices=RENDERERS, help="基线所属渲染器（默认 word）"
    )
    p.add_argument("--dpi", type=int, default=DEFAULT_DPI, help="渲染精度，默认 %d" % DEFAULT_DPI)
    p.add_argument(
        "--max-diff",
        dest="max_diff",
        type=float,
        default=DEFAULT_MAX_DIFF,
        # 措辞里不能带 %：argparse 会把 help 字符串做一次 %-插值，
        # 单个百分号会被当成格式开头直接 ValueError（--help 实测炸在这里）
        help="单页差异比例阈值（小数，非百分数），默认 %s" % DEFAULT_MAX_DIFF,
    )
    return p


def parse_args(argv):
    """argparse 接手全部边界校验：缺参数值、非整数 --dpi、未知选项都由它报错。"""
    a = _build_parser().parse_args(argv)
    return a.pdf, a.dpi, a.max_diff, a.update, a.renderer


def main(argv):
    # 副作用只在入口执行（且必须在 argparse 之前：--help / 报错文案也要走 UTF-8）
    force_utf8_stdio()

    pdf, dpi, max_diff, update, renderer = parse_args(argv)
    if not os.path.exists(pdf):
        sys.exit("找不到 PDF: " + pdf)

    base_dir = os.path.join(os.path.dirname(os.path.abspath(pdf)), "baselines", renderer)
    meta_path = os.path.join(base_dir, "meta.json")

    doc = pymupdf.open(pdf)
    try:
        if update:
            os.makedirs(base_dir, exist_ok=True)
            for i in range(doc.page_count):
                doc[i].get_pixmap(dpi=dpi).save(os.path.join(base_dir, "p%03d.png" % (i + 1)))
            with open(meta_path, "w", encoding="utf-8") as f:
                json.dump({"dpi": dpi, "pages": doc.page_count, "renderer": renderer}, f)
            print("baseline recorded: %d pages @ %d dpi -> %s" % (doc.page_count, dpi, base_dir))
            return 0

        if not os.path.exists(meta_path):
            sys.exit(
                "没有基线，先跑：python scripts/snapshot.py %s --update" % os.path.basename(pdf)
            )
        with open(meta_path, encoding="utf-8") as f:
            meta = json.load(f)
        if meta.get("renderer", "word") != renderer:
            print(
                "FAIL: 基线由 %s 录制，本次 --renderer %s——跨渲染器的像素不可比"
                % (meta.get("renderer", "word"), renderer)
            )
            return 1
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
            cur = doc[i].get_pixmap(dpi=dpi).samples
            r = diff_ratio(cur, pymupdf.Pixmap(bp).samples)
            del cur  # 及时释放内存
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
    finally:
        # 七条 return 路径共用这个出口：不关句柄就会一直占着那份 PDF，
        # Windows 上紧接着删除 / 重导同一份文件就撞 PermissionError
        doc.close()


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
