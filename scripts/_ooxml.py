"""OOXML 底层原语：schema 子元素顺序表、按序插入、样式字体读取。

ECMA-376 对各 *Pr 元素的子节点顺序有硬约束，插错位置 Word 会报"文档已损坏"。
post.py 与 make_ref.py 都在手写这些顺序表，此前各存一份、insert_ordered 语义还
不同（对不在表里的既有元素：post 跳过、make_ref 视为最大序参与比较）——
这里落地唯一实现，采用 post 的保守语义（未知元素不参与比较，只在已知元素间插位）。

style_fonts() 合并了三份相似的"读样式的 eastAsia/latin/字号"逻辑
（validate._east_asia_of_style / distill.style_font / 手解 rFonts 的各处兜底）。
"""

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ---------------------------------------------------------------------------
# schema 子元素顺序表（ECMA-376 CT_*Pr 的序列）
# ---------------------------------------------------------------------------

PPR_ORDER = [
    "pStyle",
    "keepNext",
    "keepLines",
    "pageBreakBefore",
    "framePr",
    "widowControl",
    "numPr",
    "suppressLineNumbers",
    "pBdr",
    "shd",
    "tabs",
    "suppressAutoHyphens",
    "kinsoku",
    "wordWrap",
    "overflowPunct",
    "topLinePunct",
    "autoSpaceDE",
    "autoSpaceDN",
    "bidi",
    "adjustRightInd",
    "snapToGrid",
    "spacing",
    "ind",
    "contextualSpacing",
    "mirrorIndents",
    "suppressOverlap",
    "jc",
    "textDirection",
    "textAlignment",
    "textboxTightWrap",
    "outlineLvl",
    "divId",
    "cnfStyle",
    "rPr",
    "sectPr",
    "pPrChange",
]
TCPR_ORDER = [
    "cnfStyle",
    "tcW",
    "gridSpan",
    "hMerge",
    "vMerge",
    "tcBorders",
    "shd",
    "noWrap",
    "tcMar",
    "textDirection",
    "tcFitText",
    "vAlign",
    "hideMark",
    "headers",
    "cellIns",
    "cellDel",
    "cellMerge",
    "tcPrChange",
]
TRPR_ORDER = [
    "cnfStyle",
    "divId",
    "gridBefore",
    "gridAfter",
    "wBefore",
    "wAfter",
    "cantSplit",
    "trHeight",
    "tblHeader",
    "tblCellSpacing",
    "jc",
    "hidden",
]
TBLPR_ORDER = [
    "tblStyle",
    "tblpPr",
    "tblOverlap",
    "bidiVisual",
    "tblStyleRowBandSize",
    "tblStyleColBandSize",
    "tblW",
    "jc",
    "tblCellSpacing",
    "tblInd",
    "tblBorders",
    "shd",
    "tblLayout",
    "tblCellMar",
    "tblLook",
    "tblCaption",
    "tblDescription",
]
SETTINGS_ORDER = [
    "updateFields",
    "hdrShapeDefaults",
    "footnotePr",
    "endnotePr",
    "compat",
    "rsids",
    "mathPr",
    "themeFontLang",
    "clrSchemeMapping",
    "shapeDefaults",
    "decimalSymbol",
    "listSeparator",
]


def insert_ordered(parent, child, order):
    """按 schema 顺序把 child 插进 parent；未知标签不碍事（保守跳过）。

    返回 child 便于链式设属性。
    """
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


# ---------------------------------------------------------------------------
# 样式字体读取
# ---------------------------------------------------------------------------


def style_rfonts(style):
    """取 style 的 w:rFonts 元素（没有则 None）。"""
    rpr = style.element.find(qn("w:rPr"))
    if rpr is None:
        return None
    return rpr.find(qn("w:rFonts"))


def style_fonts(style):
    """取 (中文字体, 西文字体, 字号pt)。中文字体在 rFonts 的 eastAsia 上。

    style 为 None 或属性缺失时对应位返回 None，调用方自行决定回退。

    注：本函数只做信息提取（读侧）。写入路径（make_ref.set_font：找到/新建
    w:rFonts 再 setattr）语义不同，故意不复用本函数，避免把写耦合进读原语。
    """
    if style is None:
        return None, None, None
    latin = style.font.name
    size = style.font.size.pt if style.font.size else None
    rf = style_rfonts(style)
    east = rf.get(qn("w:eastAsia")) if rf is not None else None
    return east, latin, size


# ---------------------------------------------------------------------------
# 写侧原语（S-independent：不依赖 post.py 运行时样式态 global S，可安全外提）
# 与读侧 style_fonts 对称。依赖 S 的写原语（set_run_font/shade/set_cell_border/
# set_table_borders/add_field）故意留在 post.py——本阶段绝不碰 global S。
# ---------------------------------------------------------------------------


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


def clear_repeat_header(row):
    """去掉表头行的「跨页重复」标记（header_rows 显式设为 0 时用）。"""
    trPr = row._tr.find(qn("w:trPr"))
    if trPr is None:
        return
    el = trPr.find(qn("w:tblHeader"))
    if el is not None:
        trPr.remove(el)
