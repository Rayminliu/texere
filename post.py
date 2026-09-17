# -*- coding: utf-8 -*-
"""通用 docx 后处理器：把 pandoc 产出的正文 docx 加工成正式中文文档。

用法: python post.py <body.docx> <out.docx> [config.json]

config.json 字段（均可省）:
  cover   : [[样式名, 文本], ...]  封面行，样式可用 CoverTop/CoverTitle/CoverSub/CoverInfo/CoverDate
  header  : 页眉文字（正文节）
  title / author / subject / comments : 文档属性
  toc_heading : 目录标题，默认 "目　　录"
行为: 去掉前置书名页与空段 -> 注入封面 -> 注入目录域 -> 分节(封面目录/正文各自页码)
      -> 正文节页眉页脚(居中页码) -> 表格 100% 宽/表头加粗灰底居中/单元格 10.5pt
      -> 表题与图注居中灰色去斜体 -> settings 加 updateFields。
所有手写 OOXML 均按 ECMA-376 子元素顺序插入，避免 Word 报"文档已损坏"。
"""
import copy
import json
import re
import sys

from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

SONG, LATIN = "宋体", "Times New Roman"
GRAY = (0x40, 0x40, 0x40)
# 兼容中文模板（reference_doc 来自中文 Word 时一级标题样式名为「标题 1」）
H1_STYLES = {"Heading 1", "标题 1"}
TITLE_STYLES = {"Title", "Subtitle", "Author", "Date", "标题", "副标题"}
# 题注编号：半角/全角数字 + 半角/全角分隔符
# 覆盖 表 1-1 / 表1.1 / 表１－１ / 图 2-3 / Table 1-1 / Figure 1-2
_D = r"[0-9０-９]"
_SEP = r"[.\-－—．]"
CAPTION_RE = re.compile(
    r"^(?:表|圖|图|表格|图片|Table|Figure|Fig\.?)\s*%s+(?:\s*%s\s*%s+)?" % (_D, _SEP, _D),
    re.IGNORECASE)

PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
             "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
             "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
             "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
             "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
             "suppressOverlap", "jc", "textDirection", "textAlignment",
             "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr",
             "pPrChange"]
TCPR_ORDER = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd",
              "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark",
              "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange"]
TRPR_ORDER = ["cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter",
              "cantSplit", "trHeight", "tblHeader", "tblCellSpacing", "jc", "hidden"]
TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
               "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd",
               "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook",
               "tblCaption", "tblDescription"]
SETTINGS_ORDER = ["updateFields", "hdrShapeDefaults", "footnotePr", "endnotePr",
                  "compat", "rsids", "mathPr", "themeFontLang", "clrSchemeMapping",
                  "shapeDefaults", "decimalSymbol", "listSeparator"]


def insert_ordered(parent, child, order):
    tag = child.tag.split("}")[-1]
    idx = order.index(tag) if tag in order else len(order)
    for existing in parent:
        etag = existing.tag.split("}")[-1]
        if etag not in order:
            continue
        if order.index(etag) > idx:
            existing.addprevious(child)
            return child
    parent.append(child)
    return child


def set_run_font(run, size=None, bold=None, color=None, italic=None):
    run.font.name = LATIN
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if color is not None:
        run.font.color.rgb = RGBColor(*color)
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
    rf.set(qn("w:ascii"), LATIN)
    rf.set(qn("w:hAnsi"), LATIN)
    rf.set(qn("w:eastAsia"), SONG)


def add_field(paragraph, instr, placeholder="", size=10.5):
    run = paragraph.add_run()
    b = OxmlElement("w:fldChar"); b.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = instr
    sp = OxmlElement("w:fldChar"); sp.set(qn("w:fldCharType"), "separate")
    t = OxmlElement("w:t"); t.text = placeholder
    en = OxmlElement("w:fldChar"); en.set(qn("w:fldCharType"), "end")
    for el in (b, it, sp, t, en):
        run._element.append(el)
    set_run_font(run, size=size)
    return run


def shade(cell, fill="EDEDED"):
    tcpr = cell._tc.get_or_add_tcPr()
    sh = tcpr.find(qn("w:shd"))
    if sh is None:
        sh = insert_ordered(tcpr, OxmlElement("w:shd"), TCPR_ORDER)
    sh.set(qn("w:val"), "clear")
    sh.set(qn("w:color"), "auto")
    sh.set(qn("w:fill"), fill)


def set_pgnum_start(sectPr, start=1):
    pg = sectPr.find(qn("w:pgNumType"))
    if pg is None:
        pg = OxmlElement("w:pgNumType")
        anchor = sectPr.find(qn("w:cols")) or sectPr.find(qn("w:docGrid"))
        if anchor is not None:
            anchor.addprevious(pg)
        else:
            sectPr.append(pg)
    pg.set(qn("w:start"), str(start))


def set_repeat_header(row):
    """表头行跨页重复（w:tblHeader）。缺失时表格跨页后第 2 页起没有表头。"""
    trPr = row._tr.find(qn("w:trPr"))
    if trPr is None:
        trPr = OxmlElement("w:trPr")
        row._tr.insert(0, trPr)
    el = trPr.find(qn("w:tblHeader"))
    if el is None:
        el = insert_ordered(trPr, OxmlElement("w:tblHeader"), TRPR_ORDER)
    el.set(qn("w:val"), "true")


def style_name(p):
    try:
        return p.style.name or ""
    except Exception:
        return ""


def find_h1(paras):
    """定位第一章标题；兼容中文模板的「标题 1」。找不到返回 None。"""
    for k, p in enumerate(paras):
        if style_name(p) in H1_STYLES:
            return k
    return None


def main(body_path, out_path, cfg_path):
    cfg = json.load(open(cfg_path, encoding="utf-8")) if cfg_path else {}
    doc = Document(body_path)
    body = doc.element.body

    # 1. 定位第一章标题；删除前置书名页与空段
    h1_idx = find_h1(doc.paragraphs)
    if h1_idx is None:
        sys.exit("未找到任何一级标题（Markdown 的 # 标题）。post.py 靠它定位正文起点"
                 "以插入封面与目录；请检查源 md 是否含 # 标题，或 reference_doc 的"
                 "一级标题样式是否为 Heading 1 / 标题 1。")
    for p in list(doc.paragraphs[:h1_idx]):
        if style_name(p) in TITLE_STYLES:
            p._element.getparent().remove(p._element)
    while True:
        paras = doc.paragraphs
        h1_idx = find_h1(paras)
        if h1_idx == 0:
            break
        prev = paras[h1_idx - 1]
        if prev.text.strip() == "" and not prev._p.findall(".//" + qn("w:drawing")):
            prev._element.getparent().remove(prev._element)
        else:
            break
    first_h1 = doc.paragraphs[find_h1(doc.paragraphs)]
    first_h1.paragraph_format.page_break_before = False

    # 2. 封面 + 目录 + 分节段（先追加到末尾再整体前移）
    created = []

    def np(text="", style=None, align=None, page_break=False):
        p = doc.add_paragraph(text, style=style) if style else doc.add_paragraph(text)
        if align is not None:
            p.alignment = align
        if page_break:
            p.paragraph_format.page_break_before = True
        created.append(p)
        return p

    for style, text in cfg.get("cover", []):
        np(text, style)
    if cfg.get("cover"):
        np("", "CoverInfo")
    toc_head = np(cfg.get("toc_heading", "目　　录"), "TOC Heading",
                  align=WD_ALIGN_PARAGRAPH.CENTER, page_break=True)
    for r in toc_head.runs:
        set_run_font(r, size=16, bold=True, color=(0, 0, 0))
    toc_para = np("")
    toc_para.paragraph_format.first_line_indent = Pt(0)
    add_field(toc_para, 'TOC \\o "1-2" \\h \\z \\u',
              "【目录将在打开文档时自动生成；若未显示请全选后按 F9】", size=12)
    sect_para = np("")
    sect_para.paragraph_format.first_line_indent = Pt(0)

    anchor = first_h1._p
    for p in created:
        anchor.addprevious(p._p)

    # 3. 分节与页码：封面+目录为第 1 节（无页眉页脚），正文为第 2 节
    body_sectPr = body.find(qn("w:sectPr"))
    sec1 = copy.deepcopy(body_sectPr)
    for tag in ("w:headerReference", "w:footerReference"):
        for el in sec1.findall(qn(tag)):
            sec1.remove(el)
    set_pgnum_start(sec1, 1)
    insert_ordered(sect_para._p.get_or_add_pPr(), sec1, PPR_ORDER)
    set_pgnum_start(body_sectPr, 1)

    sec_body = doc.sections[-1]
    sec_body.footer.is_linked_to_previous = False
    fp = sec_body.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.paragraph_format.first_line_indent = Pt(0)
    set_run_font(fp.add_run("— "), size=9)
    add_field(fp, "PAGE", "1", size=9)
    set_run_font(fp.add_run(" —"), size=9)

    header_text = cfg.get("header")
    if header_text:
        sec_body.header.is_linked_to_previous = False
        hp = sec_body.header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        hp.paragraph_format.first_line_indent = Pt(0)
        set_run_font(hp.add_run(header_text), size=9, color=(0x59, 0x59, 0x59))
        pbdr = OxmlElement("w:pBdr")
        btm = OxmlElement("w:bottom")
        btm.set(qn("w:val"), "single"); btm.set(qn("w:sz"), "4")
        btm.set(qn("w:space"), "1"); btm.set(qn("w:color"), "BFBFBF")
        pbdr.append(btm)
        insert_ordered(hp._p.get_or_add_pPr(), pbdr, PPR_ORDER)

    # 4. 表格排版
    for tbl in doc.tables:
        tblPr = tbl._tbl.tblPr
        w = tblPr.find(qn("w:tblW"))
        if w is None:
            w = insert_ordered(tblPr, OxmlElement("w:tblW"), TBLPR_ORDER)
        w.set(qn("w:type"), "pct"); w.set(qn("w:w"), "5000")
        layout = tblPr.find(qn("w:tblLayout"))
        if layout is None:
            layout = insert_ordered(tblPr, OxmlElement("w:tblLayout"), TBLPR_ORDER)
        layout.set(qn("w:type"), "autofit")
        mar = tblPr.find(qn("w:tblCellMar"))
        if mar is None:
            mar = insert_ordered(tblPr, OxmlElement("w:tblCellMar"), TBLPR_ORDER)
        for side, val in (("top", "40"), ("left", "80"), ("bottom", "40"), ("right", "80")):
            el = mar.find(qn("w:" + side))
            if el is None:
                el = OxmlElement("w:" + side); mar.append(el)
            el.set(qn("w:w"), val); el.set(qn("w:type"), "dxa")
        for ri, row in enumerate(tbl.rows):
            if ri == 0:
                set_repeat_header(row)
            for cell in row.cells:
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                if ri == 0:
                    shade(cell)
                for p in cell.paragraphs:
                    pf = p.paragraph_format
                    pf.first_line_indent = Pt(0)
                    pf.space_before = Pt(1); pf.space_after = Pt(1)
                    pf.line_spacing = 1.0
                    if ri == 0:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for r in p.runs:
                        set_run_font(r, size=10.5, bold=True if ri == 0 else None)

    # 5. 表题 / 图注 居中、灰色、去斜体
    cap_styles = {"Image Caption", "Caption", "Table Caption", "Captioned Figure"}
    n_cap = 0
    for p in doc.paragraphs:
        txt = p.text.strip()
        if CAPTION_RE.match(txt) or style_name(p) in cap_styles:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf = p.paragraph_format
            pf.first_line_indent = Pt(0); pf.left_indent = Pt(0)
            pf.space_before = Pt(6); pf.space_after = Pt(4)
            for r in p.runs:
                set_run_font(r, size=10.5, bold=False, color=GRAY, italic=False)
            n_cap += 1

    # 6. 打开时自动刷域 + 文档属性
    settings = doc.settings.element
    if settings.find(qn("w:updateFields")) is None:
        uf = OxmlElement("w:updateFields"); uf.set(qn("w:val"), "true")
        insert_ordered(settings, uf, SETTINGS_ORDER)
    cp = doc.core_properties
    for k in ("title", "author", "subject", "comments"):
        if cfg.get(k):
            setattr(cp, k, cfg[k])

    doc.save(out_path)
    print("saved:", out_path, "| captions centered:", n_cap,
          "| tables:", len(doc.tables), "| sections:", len(doc.sections))


if __name__ == "__main__":
    if len(sys.argv) < 3:
        sys.exit("用法: python post.py <body.docx> <out.docx> [config.json]")
    main(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else None)
