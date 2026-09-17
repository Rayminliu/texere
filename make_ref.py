# -*- coding: utf-8 -*-
"""定制 pandoc reference.docx：A4 页面、中文宋体/黑体排版、标题层级、表格与图注样式。

用法:
  python make_ref.py                                  # 生成包内 ref.docx（默认，改完立即生效）
  python make_ref.py --src 甲方模板.docx --dst x.docx  # 以指定模板为基准

默认输出即包内的 ref.docx，杜绝"改了别处的副本、包里那份悄悄漂移"。
基准模板缺失时自动用 pandoc 自带默认模板生成，不依赖任何本机绝对路径。
"""
import argparse
import copy
import os
import subprocess
import sys
import warnings

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

KIT = os.path.dirname(os.path.abspath(__file__))

# python-docx 1.2 在按 style_id 命中样式时会告警。本模板里 Caption 的 name 与 styleId 同名，
# 属于误报：取到的确实是目标样式（name=Caption / styleId=Caption）。过滤掉噪音，避免
# 每次重建模板都被这条告警干扰判断。
warnings.filterwarnings("ignore", message="style lookup by style_id is deprecated")


def ensure_base(path):
    """基准模板不存在时，用 pandoc 自带默认模板生成。"""
    if os.path.exists(path):
        return path
    r = subprocess.run(["pandoc", "--print-default-data-file", "reference.docx"],
                       capture_output=True)
    if r.returncode != 0 or not r.stdout:
        sys.exit("缺少基准模板，且无法从 pandoc 生成默认模板：\n"
                 "  pandoc --print-default-data-file reference.docx > %s" % path)
    with open(path, "wb") as fh:
        fh.write(r.stdout)
    print("generated base template:", path)
    return path


_ap = argparse.ArgumentParser()
_ap.add_argument("--src", default=os.path.join(KIT, "ref_default.docx"),
                 help="基准模板（默认包内 ref_default.docx，缺失时由 pandoc 生成）")
_ap.add_argument("--dst", default=os.path.join(KIT, "ref.docx"),
                 help="输出模板（默认覆盖包内 ref.docx）")
_a = _ap.parse_args()

SRC = ensure_base(_a.src)
DST = _a.dst

HEI = "黑体"
SONG = "宋体"
LATIN = "Times New Roman"

doc = Document(SRC)

# ---------- 页面设置：A4 + 中文常用页边距 ----------
for sec in doc.sections:
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.54)
    sec.bottom_margin = Cm(2.54)
    sec.left_margin = Cm(3.0)
    sec.right_margin = Cm(2.6)


def set_font(style, latin=LATIN, east=SONG, size=None, bold=None, color=None):
    f = style.font
    f.name = latin
    if size is not None:
        f.size = Pt(size)
    if bold is not None:
        f.bold = bold
    if color is not None:
        f.color.rgb = RGBColor(*color)
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), latin)
    rfonts.set(qn("w:hAnsi"), latin)
    rfonts.set(qn("w:eastAsia"), east)
    rfonts.set(qn("w:cs"), latin)


def para(style, align=None, before=None, after=None, line=1.0, indent_chars=None,
         keep_next=None, page_break_before=None, left=None, right=None):
    pf = style.paragraph_format
    if align is not None:
        pf.alignment = align
    if before is not None:
        pf.space_before = Pt(before)
    if after is not None:
        pf.space_after = Pt(after)
    pf.line_spacing = line
    if keep_next is not None:
        pf.keep_with_next = keep_next
    if page_break_before is not None:
        pf.page_break_before = page_break_before
    if indent_chars:
        # 小四（12pt）下 2 字符 = 24pt，等价于 Word 的“首行缩进 2 字符”
        pf.first_line_indent = Pt(12 * indent_chars)
    if left is not None:
        pf.left_indent = Pt(left)
    if right is not None:
        pf.right_indent = Pt(right)


S = doc.styles

# 正文
for name in ("Normal", "Body Text", "First Paragraph", "Compact"):
    if name in [s.name for s in S]:
        try:
            st = S[name]
        except KeyError:
            continue
        set_font(st, size=12, east=SONG)
        if name in ("Body Text", "First Paragraph"):
            para(st, align=WD_ALIGN_PARAGRAPH.JUSTIFY, before=0, after=6, line=1.0,
                 indent_chars=2)
        else:
            para(st, align=WD_ALIGN_PARAGRAPH.JUSTIFY, before=0, after=4, line=1.0)

# 标题
head_cfg = {
    "Heading 1": dict(size=16, align=WD_ALIGN_PARAGRAPH.CENTER, before=18, after=14,
                      page_break=True),
    "Heading 2": dict(size=14, align=WD_ALIGN_PARAGRAPH.LEFT, before=14, after=8),
    "Heading 3": dict(size=12.5, align=WD_ALIGN_PARAGRAPH.LEFT, before=10, after=6),
    "Heading 4": dict(size=12, align=WD_ALIGN_PARAGRAPH.LEFT, before=8, after=4),
    "Heading 5": dict(size=12, align=WD_ALIGN_PARAGRAPH.LEFT, before=6, after=3),
}
for name, cfg in head_cfg.items():
    try:
        st = S[name]
    except KeyError:
        continue
    set_font(st, east=HEI, latin=LATIN, size=cfg["size"], bold=True,
             color=(0x00, 0x00, 0x00))
    para(st, align=cfg["align"], before=cfg["before"], after=cfg["after"], line=1.15,
         keep_next=True, page_break_before=cfg.get("page_break"))

# 封面标题
for name, size in (("Title", 26), ("Subtitle", 16)):
    try:
        st = S[name]
    except KeyError:
        continue
    set_font(st, east=HEI, size=size, bold=(name == "Title"),
             color=(0x00, 0x00, 0x00))
    para(st, align=WD_ALIGN_PARAGRAPH.CENTER, before=6, after=10, line=1.3)

# 图注 / 表注
for name in ("Caption", "Table Caption", "Image Caption", "Captioned Figure", "Figure"):
    try:
        st = S[name]
    except KeyError:
        continue
    set_font(st, east=SONG, size=10.5, bold=False, color=(0x40, 0x40, 0x40))
    st.font.italic = False
    para(st, align=WD_ALIGN_PARAGRAPH.CENTER, before=3, after=10, line=1.0)

# 目录条目
for name in ("TOC Heading", "TOC 1", "TOC 2", "TOC 3"):
    try:
        st = S[name]
    except KeyError:
        continue
    set_font(st, east=SONG, size=12, bold=(name == "TOC Heading"),
             color=(0x00, 0x00, 0x00))

# 引用/说明段（pandoc blockquote）
for name in ("Blockquote", "Block Text", "Quote"):
    try:
        st = S[name]
    except KeyError:
        continue
    set_font(st, east=SONG, size=10.5, bold=False, color=(0x40, 0x40, 0x40))
    para(st, align=WD_ALIGN_PARAGRAPH.JUSTIFY, before=3, after=8, line=1.0,
         left=21, right=12)

# 表格文字（pandoc 表格用 Table 样式，单元格继承 Normal）
TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
               "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd",
               "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook",
               "tblCaption", "tblDescription"]


def insert_ordered(parent, child, order):
    tag = child.tag.split("}")[-1]
    idx = order.index(tag) if tag in order else len(order)
    for existing in parent:
        etag = existing.tag.split("}")[-1]
        eidx = order.index(etag) if etag in order else len(order)
        if eidx > idx:
            existing.addprevious(child)
            return child
    parent.append(child)
    return child


try:
    st = S["Table"]
    tblpr = st.element.find(qn("w:tblPr"))
    if tblpr is None:
        tblpr = OxmlElement("w:tblPr")
        st.element.insert(0, tblpr)
    borders = tblpr.find(qn("w:tblBorders"))
    if borders is None:
        borders = insert_ordered(tblpr, OxmlElement("w:tblBorders"), TBLPR_ORDER)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = borders.find(qn("w:" + edge))
        if el is None:
            el = OxmlElement("w:" + edge)
            borders.append(el)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "6")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "808080")
except KeyError:
    pass

# ---------- 自定义样式（封面与提示文字） ----------
custom = [
    # name, east, size, bold, align, before, after, line, indent_chars, color
    ("CoverTop", HEI, 16, True, WD_ALIGN_PARAGRAPH.CENTER, 0, 6, 1.5, None, None),
    ("CoverTitle", HEI, 30, True, WD_ALIGN_PARAGRAPH.CENTER, 40, 20, 1.4, None, None),
    ("CoverSub", HEI, 20, True, WD_ALIGN_PARAGRAPH.CENTER, 10, 30, 1.4, None, None),
    ("CoverInfo", SONG, 14, False, WD_ALIGN_PARAGRAPH.CENTER, 4, 4, 1.8, None, None),
    ("CoverDate", SONG, 14, False, WD_ALIGN_PARAGRAPH.CENTER, 30, 0, 1.5, None, None),
    ("NoIndent", SONG, 12, False, WD_ALIGN_PARAGRAPH.JUSTIFY, 0, 6, 1.0, None, None),
    ("Lead", SONG, 12, False, WD_ALIGN_PARAGRAPH.LEFT, 4, 8, 1.0, None, (0x33, 0x33, 0x33)),
    ("SmallNote", SONG, 10.5, False, WD_ALIGN_PARAGRAPH.LEFT, 2, 8, 1.0, None, (0x59, 0x59, 0x59)),
    ("FigurePara", SONG, 12, False, WD_ALIGN_PARAGRAPH.CENTER, 6, 2, 1.0, None, None),
    # filters/captions.lua 打上的题注语义样式（TableCaption 复用内置「Table Caption」，
    # 其 styleId 同为 TableCaption，故这里只补 FigureCaption，避免 styleId 撞车）
    ("FigureCaption", SONG, 10.5, False, WD_ALIGN_PARAGRAPH.CENTER, 3, 10, 1.0,
     None, (0x40, 0x40, 0x40)),
]
for name, east, size, bold, align, before, after, line, ind, color in custom:
    try:
        st = S[name]
    except KeyError:
        st = S.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        st.base_style = S["Normal"]
        st.quick_style = True
    set_font(st, east=east, size=size, bold=bold, color=color)
    para(st, align=align, before=before, after=after, line=line, indent_chars=ind)

doc.save(DST)
print("saved", DST)
print("styles:", len([s for s in doc.styles]))
