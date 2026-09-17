# -*- coding: utf-8 -*-
"""通用 docx 后处理器：把 pandoc 产出的正文 docx 加工成正式中文文档。

用法: python post.py <body.docx> <out.docx> [config.json]

config.json 字段（均可省）:
  cover   : [[样式名, 文本], ...]  封面行，样式可用 CoverTop/CoverTitle/CoverSub/CoverInfo/CoverDate
  header  : 页眉文字（正文节）
  title / author / subject / comments : 文档属性
  toc_heading : 目录标题，默认 "目　　录"
  style   : 版式微调（字体/颜色/间距/表格边框等，见 README 的完整键表）
行为: 去掉前置书名页与空段 -> 注入封面 -> 注入目录域 -> 分节(封面目录/正文各自页码)
      -> 正文节页眉页脚(居中页码) -> 表格 100% 宽/表头加粗灰底居中/单元格 10.5pt
      -> 表题与图注居中灰色去斜体 -> settings 加 updateFields。
契约: **只改版式，不改内容**——不触碰正文与题注的文字（图表编号由源文件手写）。
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
# 可被 config.json 的 "style" 段覆盖（见 apply_style_cfg）。默认即中文正式文档惯例。
S = {
    "east_font": SONG,        # 中文字体
    "latin_font": LATIN,      # 西文字体
    "caption_gray": GRAY,     # 题注灰
    "caption_size": 10.5,     # 题注字号 pt
    "table_shade": "EDEDED",  # 表头底纹
    "table_size": 10.5,       # 表格字号 pt
    "toc_depth": "1-2",       # 目录收录层级
    "page_number": "— {n} —",  # 页码模板，{n} 处插入页码域
    "header_rows": None,      # 表头行数；None = 自动（读 pandoc 打的 w:tblHeader）
    "table_border": "full",   # 表格边框：full 全框线 / three 三线表 / none 无框线
    # --- 页眉 / 页脚 ---
    "header_size": 9.0,           # 页眉字号 pt
    "header_gray": (0x59, 0x59, 0x59),
    "header_rule_color": "BFBFBF",  # 页眉下边框颜色
    "header_rule_size": 4,        # 页眉下边框粗细（1/8 pt）
    "page_number_size": 9.0,      # 页码字号 pt
    # --- 目录 ---
    "toc_title_size": 16.0,
    "toc_title_color": (0, 0, 0),
    "toc_placeholder": "【目录将在打开文档时自动生成；若未显示请全选后按 F9】",
    "toc_placeholder_size": 12.0,
    # --- 表格 ---
    "cell_margin_v": 40,      # 单元格上下边距 twips
    "cell_margin_h": 80,      # 单元格左右边距 twips
    "table_para_space": 1.0,  # 单元格内段前后 pt
    "border_size": 6,         # 全框线粗细（1/8 pt）
    "border_color": "808080",
    "three_line_size": 12,    # 三线表顶底线粗细（1/8 pt）
    # --- 题注 ---
    "caption_space_before": 6.0,
    "caption_space_after": 4.0,
    "caption_keep_with_next": True,   # 表题与表格同页；关掉可能省页数但会分家
}
# 兼容中文模板（reference_doc 来自中文 Word 时一级标题样式名为「标题 1」）
H1_STYLES = {"Heading 1", "标题 1"}
TITLE_STYLES = {"Title", "Subtitle", "Author", "Date", "标题", "副标题"}
# 题注识别（文本兜底用；正常路径由 filters/captions.lua 在 AST 层打样式）
# 覆盖 表 1-1 / 表1.1 / 表１－１ / 图 2-3 / Table 1-1 / Figure 1-2
_D = r"[0-9０-９]"
_SEP = r"[.\-－—．]"
CAPTION_RE = re.compile(
    r"^(?:表|圖|图|表格|图片|Table|Figure|Fig\.?)\s*%s+(?:\s*%s\s*%s+)?" % (_D, _SEP, _D),
    re.IGNORECASE)
# 由 filters/captions.lua 在 AST 层打上的语义样式（按 styleId 匹配，见 style_id 注释）
CAPTION_STYLE_IDS = {"TableCaption", "FigureCaption",
                     "ImageCaption", "Caption", "CaptionedFigure"}

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
    run.font.name = S["latin_font"]
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
    rf.set(qn("w:ascii"), S["latin_font"])
    rf.set(qn("w:hAnsi"), S["latin_font"])
    rf.set(qn("w:eastAsia"), S["east_font"])


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


def shade(cell, fill=None):
    fill = fill or S["table_shade"]
    tcpr = cell._tc.get_or_add_tcPr()
    sh = tcpr.find(qn("w:shd"))
    if sh is None:
        sh = insert_ordered(tcpr, OxmlElement("w:shd"), TCPR_ORDER)
    sh.set(qn("w:val"), "clear")
    sh.set(qn("w:color"), "auto")
    sh.set(qn("w:fill"), fill)


def set_cell_border(cell, edge, sz=None, color=None):
    sz = str(sz if sz is not None else S["border_size"])
    color = color or S["border_color"]
    tcPr = cell._tc.get_or_add_tcPr()
    tb = tcPr.find(qn("w:tcBorders"))
    if tb is None:
        tb = insert_ordered(tcPr, OxmlElement("w:tcBorders"), TCPR_ORDER)
    el = tb.find(qn("w:" + edge))
    if el is None:
        el = OxmlElement("w:" + edge)
        tb.append(el)
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), sz)
    el.set(qn("w:space"), "0")
    el.set(qn("w:color"), color)


def set_table_borders(tbl, mode, header_rows):
    """表格边框：full 全框线 / three 三线表 / none 无框线。

    三线表 = 顶线 + 表头下线 + 底线，内部无竖线无横线（论文/申报书常用）。
    """
    if mode not in ("full", "three", "none"):
        raise ValueError('table_border 只能是 "full" / "three" / "none"，收到：%r' % mode)
    tblPr = tbl._tbl.tblPr
    borders = tblPr.find(qn("w:tblBorders"))
    if borders is None:
        borders = insert_ordered(tblPr, OxmlElement("w:tblBorders"), TBLPR_ORDER)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = borders.find(qn("w:" + edge))
        if el is None:
            el = OxmlElement("w:" + edge)
            borders.append(el)          # CT_TblBorders 有固定子元素顺序
        if mode == "none":
            el.set(qn("w:val"), "nil")
        elif mode == "three":
            if edge in ("top", "bottom"):
                el.set(qn("w:val"), "single")
                el.set(qn("w:sz"), str(S["three_line_size"]))  # 顶底线比内部线粗
                el.set(qn("w:space"), "0")
                el.set(qn("w:color"), "000000")
            else:
                el.set(qn("w:val"), "nil")
        else:
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), str(S["border_size"]))
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), S["border_color"])
    if mode == "three":
        # 表头最后一行的下边框 = 三线表的中间那条线
        for cell in tbl.rows[min(header_rows, len(tbl.rows)) - 1].cells:
            set_cell_border(cell, "bottom")


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
    """确保表头行跨页重复（w:tblHeader）。

    注意：pandoc 通常**已经**给表头行设了（grid table 里 `+===+` 以上的行都算表头），
    本函数只是幂等加固——若上游没设，表格跨页后第 2 页起就没有表头。
    """
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


def style_id(p):
    """样式 styleId。比 name 可靠：内置「Table Caption」与我们的 TableCaption
    同名不同字，按 name 判会歧义，按 styleId 判唯一。"""
    try:
        return p.style.style_id or ""
    except Exception:
        return ""


def build_caption_matchers(words):
    """按 config 的 caption_words 重建题注关键字（默认 表/图/Table/Figure）。

    words 形如 {"table": ["表", "表格"], "figure": ["图", "图片"]}。
    同时供 filters/captions.lua 使用（由 render.py 以 -M 传入），两边保持一致。
    """
    global CAPTION_RE
    ws = (words.get("table") or []) + (words.get("figure") or [])
    ws = [w for w in ws if w]
    if not ws:
        return
    ws.sort(key=len, reverse=True)                # 长词优先，避免「表」抢「表格」
    alt = "|".join(re.escape(w) for w in ws)
    CAPTION_RE = re.compile(
        r"^(?:%s)\s*%s+(?:\s*%s\s*%s+)?" % (alt, _D, _SEP, _D), re.IGNORECASE)


def uses_caption_styles(doc):
    """文档里是否已有 AST 层打的题注样式。

    有的话就**只信样式**，不再用文本正则兜底——否则「**表层（边缘轻算力）：**基于…」
    这类正文会被当成表题，插入编号并毁掉加粗（真实项目实测发生过）。
    """
    for p in doc.paragraphs:
        if style_id(p) in CAPTION_STYLE_IDS:
            return True
    return False


def detect_header_rows(tbl):
    """数出开头连续带 w:tblHeader 的行数。

    pandoc 已经把表头行标好了（grid table 的 `+===+` 以上全部算表头），
    直接读它，比让用户填一个全局 header_rows 靠谱得多——同一文档里
    不同表的表头行数往往不一样。
    """
    n = 0
    for row in tbl.rows:
        trPr = row._tr.find(qn("w:trPr"))
        if trPr is None or trPr.find(qn("w:tblHeader")) is None:
            break
        n += 1
    return n


def to_rgb(v):
    """颜色解析：接受 "404040" / "#404040" / "0x404040" / [r,g,b]。"""
    if isinstance(v, (list, tuple)):
        return tuple(int(x) for x in v)
    s = str(v).strip().lstrip("#")
    if s.lower().startswith("0x"):
        s = s[2:]
    if len(s) != 6:
        raise ValueError("颜色应为 6 位十六进制（如 404040），收到：%r" % v)
    return tuple(int(s[i:i + 2], 16) for i in (0, 2, 4))


# style 段各键的类型（决定怎么解析），见 apply_style_cfg
_STYLE_STR = ("east_font", "latin_font", "toc_depth", "page_number", "table_border",
              "toc_placeholder", "table_shade", "header_rule_color", "border_color")
_STYLE_FLOAT = ("caption_size", "table_size", "header_size", "page_number_size",
                "toc_title_size", "toc_placeholder_size", "table_para_space",
                "caption_space_before", "caption_space_after")
_STYLE_INT = ("cell_margin_v", "cell_margin_h", "border_size", "three_line_size",
              "header_rule_size")
_STYLE_COLOR = ("caption_gray", "header_gray", "toc_title_color")


def apply_style_cfg(cfg):
    """把 config.json 的 "style" 段套到全局设置 S 上；没写的保持默认。"""
    st = cfg.get("style") or {}
    for k in _STYLE_STR:
        if k in st:
            S[k] = st[k]
    for k in _STYLE_FLOAT:
        if k in st:
            S[k] = float(st[k])
    for k in _STYLE_INT:
        if k in st:
            S[k] = int(st[k])
    for k in _STYLE_COLOR:
        if k in st:
            S[k] = to_rgb(st[k])
    if "header_rows" in st:
        S["header_rows"] = max(1, int(st["header_rows"]))
    if "caption_keep_with_next" in st:
        S["caption_keep_with_next"] = bool(st["caption_keep_with_next"])
    # caption_words 是顶层键（它不是"样式"，是语义），也兼容写在 style 里
    cw = cfg.get("caption_words") or st.get("caption_words")
    if cw:
        build_caption_matchers(cw)
    return S


def find_h1(paras):
    """定位第一章标题；兼容中文模板的「标题 1」。找不到返回 None。"""
    for k, p in enumerate(paras):
        if style_name(p) in H1_STYLES:
            return k
    return None


def main(body_path, out_path, cfg_path):
    cfg = json.load(open(cfg_path, encoding="utf-8-sig")) if cfg_path else {}
    apply_style_cfg(cfg)
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

    # 源文件在第一个 # 之前还写了东西时，那些段落会落在目录之后。
    # 只提示、不删除——本工具的契约是不改内容。
    leftover = [p.text.strip() for p in doc.paragraphs[:find_h1(doc.paragraphs)]
                if p.text.strip()]
    if leftover:
        print("[warn] 第一个一级标题之前还有 %d 段内容，会排在目录之后：" % len(leftover))
        for t in leftover[:3]:
            print("       " + t[:60])
        print("       建议：从源文件删掉，或写进 config 的 cover 由封面承载")


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
        set_run_font(r, size=S["toc_title_size"], bold=True,
                     color=S["toc_title_color"])
    toc_para = np("")
    toc_para.paragraph_format.first_line_indent = Pt(0)
    add_field(toc_para, 'TOC \\o "%s" \\h \\z \\u' % S["toc_depth"],
              S["toc_placeholder"], size=S["toc_placeholder_size"])
    sect_para = np("")
    sect_para.paragraph_format.first_line_indent = Pt(0)

    # 封面必须是文档第一页：插到 body 最前面，而不是「第一个标题之前」。
    # 否则源文件在第一个 # 之前写的内容会排到封面之前，单独占一页（实测踩过）。
    anchor = next((c for c in body.iterchildren()
                   if c.tag in (qn("w:p"), qn("w:tbl"))), None)
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
    left, _sep, right = S["page_number"].partition("{n}")
    if left:
        set_run_font(fp.add_run(left), size=S["page_number_size"])
    add_field(fp, "PAGE", "1", size=S["page_number_size"])
    if right:
        set_run_font(fp.add_run(right), size=S["page_number_size"])

    header_text = cfg.get("header")
    if header_text:
        sec_body.header.is_linked_to_previous = False
        hp = sec_body.header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        hp.paragraph_format.first_line_indent = Pt(0)
        set_run_font(hp.add_run(header_text), size=S["header_size"],
                     color=S["header_gray"])
        pbdr = OxmlElement("w:pBdr")
        btm = OxmlElement("w:bottom")
        btm.set(qn("w:val"), "single")
        btm.set(qn("w:sz"), str(S["header_rule_size"]))
        btm.set(qn("w:space"), "1")
        btm.set(qn("w:color"), S["header_rule_color"])
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
        for side, val in (("top", S["cell_margin_v"]), ("left", S["cell_margin_h"]),
                          ("bottom", S["cell_margin_v"]),
                          ("right", S["cell_margin_h"])):
            el = mar.find(qn("w:" + side))
            if el is None:
                el = OxmlElement("w:" + side); mar.append(el)
            el.set(qn("w:w"), str(val)); el.set(qn("w:type"), "dxa")
        n_head = S["header_rows"] or detect_header_rows(tbl)
        header_rows = max(1, min(int(n_head), len(tbl.rows)))
        set_table_borders(tbl, S["table_border"], header_rows)
        for ri, row in enumerate(tbl.rows):
            is_head = ri < header_rows          # 多级表头时前 N 行都算表头
            if is_head:
                set_repeat_header(row)
            for cell in row.cells:
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                if is_head:
                    shade(cell)
                for p in cell.paragraphs:
                    pf = p.paragraph_format
                    pf.first_line_indent = Pt(0)
                    pf.space_before = Pt(S["table_para_space"])
                    pf.space_after = Pt(S["table_para_space"])
                    pf.line_spacing = 1.0
                    if is_head:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for r in p.runs:
                        set_run_font(r, size=S["table_size"],
                                     bold=True if is_head else None)

    # 5. 表题 / 图注 居中、灰色、去斜体
    n_cap = 0
    strict = uses_caption_styles(doc)      # 同上：有样式就不靠正则猜
    for p in doc.paragraphs:
        txt = p.text.strip()
        if style_id(p) in CAPTION_STYLE_IDS or (not strict and CAPTION_RE.match(txt)):
            sid = style_id(p)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf = p.paragraph_format
            pf.first_line_indent = Pt(0); pf.left_indent = Pt(0)
            pf.space_before = Pt(S["caption_space_before"])
            pf.space_after = Pt(S["caption_space_after"])
            # 「与下一段同页」只对**位于表格之前**的表题、和**载有图片**的段落有意义。
            # 图注本身在图片之后，粘住它会把后面的内容整块推走：实测 47 图文档多出 4 页。
            keep = bool(S["caption_keep_with_next"]) and (
                sid in ("TableCaption", "CaptionedFigure")
                or bool(re.match(r"^\s*(表|表格|Table)", txt, re.IGNORECASE)))
            pf.keep_with_next = keep
            for r in p.runs:
                set_run_font(r, size=S["caption_size"], bold=False,
                             color=S["caption_gray"], italic=False)
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
