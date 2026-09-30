"""Word 验收 + 导出 PDF 的 CLI 壳：python scripts/finalize.py <in.docx> [out.pdf] [--save-updated-fields]

实现已抽到 scripts/renderers.py 的 WordRenderer（RendererAdapter 的一个实现）。
本文件只负责参数解析与「人读」输出，行为与原 finalize.py 完全一致：
- 验收信号：Word 打不开 = OOXML 结构有问题（立即失败）；
- 刷新目录域并重排页码后导出 PDF；
- 默认只读验收（绝不写回输入 docx），只有 --save-updated-fields 才写回。
"""

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _shared import force_utf8_stdio
from renderers import WordRenderer

force_utf8_stdio()


def _build_parser():
    p = argparse.ArgumentParser(
        prog="finalize.py",
        description="用本机 Word 验收 docx 并导出 PDF（打不开 = OOXML 结构有问题）。",
        epilog="默认只读验收，绝不写回输入 docx；只有 --save-updated-fields 才把刷新后的域存回。",
    )
    p.add_argument("src", help="待验收的 docx")
    p.add_argument("pdf", nargs="?", help="输出的 PDF（默认与 src 同名换 .pdf 后缀）")
    p.add_argument(
        "--save-updated-fields",
        dest="save_updated_fields",
        action="store_true",
        help="把 Word 刷新后的域写回原 docx（交付物需要更新过的域时才用）",
    )
    return p


def report(res, pdf):
    """保留与原 finalize.py 完全一致的人读输出，避免破坏任何下游解析。"""
    print("Engine    :", res.renderer_name, res.renderer_version, "|", res.engine_path)
    s = res.stats
    print("Pages      :", s.get("pages"))
    print("Words      :", s.get("words"))
    print("Tables     :", s.get("tables"))
    print("InlineShapes:", s.get("inline_shapes"))
    print("Sections   :", s.get("sections"))
    print("PDF written:", os.path.exists(pdf))


def main(argv=None):
    a = _build_parser().parse_args(argv)
    pdf = a.pdf or os.path.splitext(a.src)[0] + ".pdf"

    res = WordRenderer().render(a.src, pdf, save_updated_fields=a.save_updated_fields)

    report(res, pdf)
    if res.errors:
        print("Errors     :", "; ".join(res.errors))
        return 1
    if not res.ok:
        # 渲染器可能一条 error 都不填却返回 ok=False（PDF 没产出）——
        # 这也必须失败退出，否则 edit.py --verify 会把没出 PDF 当成验收通过
        print("FAIL       : PDF 未产出")
        return 1
    print("OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
