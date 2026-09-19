"""读一份已有 docx（甲方模板 / 自己攒的模板），蒸馏出 texere 的 config 建议。

目的：接一份新模板时，不用手工去 Word 里量页边距、查字号、抄页眉。
脚本只读不写源模板，输出一份「建议 config」+ 一份体检报告，人工确认后即可使用。

用法:
  python scripts/distill.py 甲方模板.docx                  # 打印报告 + 建议 config
  python scripts/distill.py 甲方模板.docx --out cfg.json    # 同时把建议 config 写成文件

能确定的：页面设置、正文/标题的字体字号、页眉页脚文字、分节数、表格数。
确定不了的：封面结构、表格细节、页码格式——这些在报告里提示人工确认。

为什么叫"蒸馏"而不是"转换"：转换是让工具去理解模板（做不到，也不可靠）；
蒸馏是把模板拆成 config 字段，人来拍板。见 README「复用已有模板」。
"""

import argparse
import json
import os
import sys

from docx import Document
from docx.oxml.ns import qn

# Windows 控制台默认 GBK，遇到无法编码的字符会抛 UnicodeEncodeError。
# 与 render.py / edit.py 同一取舍：保持原编码，只把无法编码的字符降级为 ?。
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(errors="replace")

EMU_PER_CM = 360000

# 中文 Word 与英文 Word 的标题样式名不同，两边都认
HEADING_NAMES = [("Heading %d", "标题 %d")]


def cm(v):
    return None if v is None else round(v / EMU_PER_CM, 2)


def style_font(st):
    """取 (中文字体, 西文字体, 字号pt)。中文字体在 rFonts 的 eastAsia 上。"""
    latin = st.font.name
    size = st.font.size.pt if st.font.size else None
    east = None
    rPr = st.element.find(qn("w:rPr"))
    if rPr is not None:
        rf = rPr.find(qn("w:rFonts"))
        if rf is not None:
            east = rf.get(qn("w:eastAsia"))
    return east, latin, size


def pick_style(doc, names):
    for n in names:
        try:
            return doc.styles[n]
        except KeyError:
            continue
    return None


def part_text(container, linked):
    if linked:
        return None
    return " / ".join(p.text for p in container.paragraphs if p.text.strip()) or None


def distill(path):
    try:
        doc = Document(path)
    except Exception as e:
        sys.exit("打不开 %s（%s）\n若是 .dotx 模板，请先用 Word 另存为 .docx 再蒸馏。" % (path, e))

    report, cfg = [], {}
    cfg["reference_doc"] = os.path.basename(path)

    # --- 页面设置 ---
    sec = doc.sections[0]
    report.append("== 页面设置 ==")
    report.append(
        "  纸型     : %.1f x %.1f cm（%s）"
        % (
            cm(sec.page_width),
            cm(sec.page_height),
            "横向" if str(sec.orientation).endswith("LANDSCAPE") else "纵向",
        )
    )
    report.append(
        "  页边距   : 上%.2f 下%.2f 左%.2f 右%.2f cm"
        % (
            cm(sec.top_margin),
            cm(sec.bottom_margin),
            cm(sec.left_margin),
            cm(sec.right_margin),
        )
    )
    report.append("  分节数   : %d" % len(doc.sections))
    report.append("  表格数   : %d" % len(doc.tables))

    # --- 样式 ---
    report.append("")
    report.append("== 样式（字体取自 eastAsia / 字号 pt）==")
    body_east = None
    for label, names in (
        ("正文", ("Normal", "正文")),
        ("一级标题", ("Heading 1", "标题 1")),
        ("二级标题", ("Heading 2", "标题 2")),
        ("三级标题", ("Heading 3", "标题 3")),
    ):
        st = pick_style(doc, names)
        if st is None:
            report.append("  %-6s : （模板里没有这个样式）" % label)
            continue
        east, latin, size = style_font(st)
        report.append(
            "  %-6s : 中文=%s  西文=%s  %s"
            % (
                label,
                east or "-",
                latin or "-",
                ("%.1fpt" % size) if size else "字号未指定",
            )
        )
        if label == "正文" and east:
            body_east = (east, size)

    # --- 页眉页脚 ---
    report.append("")
    report.append("== 页眉页脚 ==")
    headers, footers = [], []
    for i, s in enumerate(doc.sections):
        h = part_text(s.header, s.header.is_linked_to_previous)
        f = part_text(s.footer, s.footer.is_linked_to_previous)
        report.append("  节%d 页眉: %s" % (i, h if h else "（空）"))
        report.append("  节%d 页脚: %s" % (i, f if f else "（空）"))
        if h:
            headers.append(h)
        if f:
            footers.append(f)

    # --- 建议 config ---
    if headers:
        uniq = sorted(set(headers))
        cfg["header"] = uniq[0]
        if len(uniq) > 1:
            report.append("")
            report.append(
                "  [注意] 各节页眉不一致，已取第一条，其余需手工处理：%s" % " | ".join(uniq[1:])
            )
    style = {}
    if footers:
        # 模板页脚有内容 -> 默认行为会把页码追加在后面，建议显式关掉
        style["page_number"] = None
    if style:
        cfg["style"] = style

    report.append("")
    report.append("== 建议 config ==")
    report.append(json.dumps(cfg, ensure_ascii=False, indent=2))

    # --- 人工确认项 ---
    report.append("")
    report.append("== 需人工确认（脚本判断不了）==")
    report.append("  1. 封面：模板首页的文字要自己填进 config 的 cover")
    report.append("  2. 页码格式：若甲方要求「第 X 页 共 Y 页」，配 style.page_number")
    report.append("  3. 表格样式：表头底纹、框线、字号在 style 段里调")
    if body_east and body_east[0]:
        cmd = "python scripts/make_ref.py --body-font %s" % body_east[0]
        if body_east[1]:
            cmd += " --body-size %g" % body_east[1]
        report.append("")
        report.append("== 想把这套字体做成你自己的默认模板（而不是引用甲方文件）==")
        report.append("  " + cmd)

    return "\n".join(report), cfg


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("--out", help="把建议 config 写成 json")
    a = ap.parse_args()
    if not os.path.exists(a.docx):
        sys.exit("找不到文件: " + a.docx)

    report, cfg = distill(a.docx)
    print(report)
    if a.out:
        with open(a.out, "w", encoding="utf-8") as fh:
            json.dump(cfg, fh, ensure_ascii=False, indent=2)
            fh.write("\n")
        print("\nwritten:", a.out)


if __name__ == "__main__":
    main()
