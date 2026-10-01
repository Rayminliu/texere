"""合成 README 首屏的「证明型 Hero」图：python scripts/make_hero.py

把已入库的示例渲染图（assets/previews/tender.png）与一份「通过态」验证凭据
（与 validate.py 的 report.json 逐字同名）合成成一张 Intent → Document → Proof
的对比面板，输出 assets/hero-verified.png 供两份 README 引用。

只依赖 PyMuPDF（requirements.txt 里的 check extra），跨平台、**不需要 Word / pandoc**：
输入是仓库里已有的 PNG，产出是确定性的——重跑即刷新，不引入外部环境耦合。
文本只用 ASCII（helv/cour 是 base-14 字体，不含 CJK/符号），勾与圆点用矢量图元；
真实中文由右侧成品位图承担。示例版式或凭据结构变更后连同 make_previews.py 一起重跑。
"""

import os

import pymupdf

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_PAGE = os.path.join(KIT, "assets", "previews", "tender.png")
OUT = os.path.join(KIT, "assets", "hero-verified.png")

# 画布（3:2 横向）
W, H = 1600, 1000

# 配色（RGB 0-1）
BG = (0.067, 0.086, 0.122)  # #111620 深色底
PANEL = (0.106, 0.129, 0.180)  # #1B2130 面板
PANEL_EDGE = (0.235, 0.282, 0.376)  # #3C4860 面板描边
TITLE = (0.910, 0.925, 0.949)  # #E8ECF2 主标题
SUB = (0.541, 0.604, 0.722)  # #8A9AB8 副标题/标签
DIM = (0.42, 0.48, 0.58)  # 更弱的说明文字
GREEN = (0.180, 0.745, 0.427)  # #2EBD6D 通过绿
CYAN = (0.541, 0.788, 0.988)  # markdown 标题行的青色
FOREG = (0.788, 0.831, 0.902)  # markdown 正文前景
CODE = (0.80, 0.84, 0.90)  # 等宽代码前景

SANS = "helv"  # Helvetica（latin-1 足够）
MONO = "cour"  # Courier 等宽，凭据像 CI 日志

# 与 validate.py report.json 逐字同名的 9 项检查（不摆拍，字段即真）
CHECKS = [
    "package_integrity",
    "source_content",
    "image_embedding",
    "section_count",
    "toc_field",
    "page_numbering",
    "blank_pages",
    "renderer_acceptance",
    "visual_drift",
]


def pt(x, y):
    return pymupdf.Point(x, y)


def text(page, x, y, s, size=14, font=SANS, color=TITLE, bold=False):
    page.insert_text(
        pt(x, y), s, fontname=font, fontsize=size, color=color, render_mode=2 if bold else 0
    )


def check(page, x, y, color=GREEN, width=2):
    """在基线 y 附近画一个绿色对勾（矢量，不依赖字体）。"""
    page.draw_polyline(
        [pt(x, y - 4), pt(x + 4, y + 1), pt(x + 12, y - 9)], color=color, width=width
    )


def hollow(page, cx, cy, r=4, color=SUB, width=1.5):
    page.draw_circle(pt(cx, cy), r, color=color, width=width)


def panel(page, rect, radius=0.03):
    # pymupdf 的 radius 是边长比例（0~0.5），不是像素
    page.draw_rect(rect, color=PANEL_EDGE, fill=PANEL, radius=radius, width=1.2)


def place_image(page, path, box):
    """在 box 内按比例居中放置图片，返回实际绘制矩形。"""
    img = pymupdf.open(path)
    try:
        iw, ih = img[0].rect.width, img[0].rect.height
    finally:
        img.close()
    scale = min(box.width / iw, box.height / ih)
    dw, dh = iw * scale, ih * scale
    x0 = box.x0 + (box.width - dw) / 2
    y0 = box.y0 + (box.height - dh) / 2
    rect = pymupdf.Rect(x0, y0, x0 + dw, y0 + dh)
    page.insert_image(rect, filename=path)
    page.draw_rect(rect, color=PANEL_EDGE, width=1)
    return rect


def draw_arrow(page, x, y):
    page.draw_line(pt(x, y - 16), pt(x + 20, y), color=SUB, width=2.5)
    page.draw_line(pt(x + 20, y), pt(x, y + 16), color=SUB, width=2.5)


def main():
    if not os.path.exists(SRC_PAGE):
        raise SystemExit("缺少源渲染图 %s，先跑 python scripts/make_previews.py" % SRC_PAGE)

    doc = pymupdf.open()
    page = doc.new_page(width=W, height=H)
    page.draw_rect(pymupdf.Rect(0, 0, W, H), color=BG, fill=BG)

    # 顶部标题条
    text(page, 48, 66, "texere", size=34, bold=True)
    text(page, 48, 98, "generate a document,  then prove it", size=16, color=SUB)
    text(page, W - 48 - 232, 66, "Markdown + template  ->  DOCX", size=15, color=SUB)

    # 左上：Intent（示意 Markdown 源码，ASCII 结构）
    in_rect = pymupdf.Rect(48, 130, 748, 560)
    panel(page, in_rect)
    text(page, 76, 168, "INPUT", size=15, color=SUB, bold=True)
    text(page, 176, 168, "requirements.md", size=14, font=MONO, color=DIM)
    md_src = [
        ("head", "# Bid Document"),
        ("gap", ""),
        ("head", "## Chapter 1  Bid Letter"),
        ("gap", ""),
        ("head", "Table 1-1  Commercial terms"),
        ("gap", ""),
        ("body", "| Clause   | Requirement | Response |"),
        ("body", "|:---------|:------------|:---------|"),
        ("body", "| Duration | 90 days     | Full     |"),
    ]
    ly = 210
    for kind, ln in md_src:
        if ln:
            text(page, 76, ly, ln, size=16, font=MONO, color=CYAN if kind == "head" else FOREG)
        ly += 30
    text(page, 76, 536, "plain Markdown - no styling, no layout", size=13, color=DIM)

    # 左下：Proof 凭据（CI-receipt 风格）
    rc_rect = pymupdf.Rect(48, 580, 748, 960)
    panel(page, rc_rect)
    text(page, 76, 616, "PROOF", size=15, color=SUB, bold=True)
    text(page, 176, 616, "report.json", size=14, font=MONO, color=DIM)
    ly = 650
    for name in CHECKS:
        check(page, 96, ly, width=2)
        text(page, 122, ly, name, size=15, font=MONO, color=CODE)
        text(page, 470, ly, "PASS", size=15, font=MONO, color=GREEN, bold=True)
        ly += 27
    # 汇总行
    sy = ly + 12
    text(page, 96, sy, "9 checks", size=15, font=MONO, color=CODE)
    check(page, 214, sy, width=2)
    text(page, 236, sy, "0 failed", size=15, font=MONO, color=CODE)
    hollow(page, 356, sy - 5, r=4)
    text(page, 368, sy, "2 skipped", size=15, font=MONO, color=CODE)
    text(page, 500, sy, "drift 0.00%", size=15, font=MONO, color=GREEN, bold=True)

    # 中：箭头（连接 INPUT 与 OUTPUT）
    draw_arrow(page, 766, 340)

    # 右：Document（真实渲染成品图）
    out_rect = pymupdf.Rect(816, 130, 1552, 960)
    panel(page, out_rect)
    text(page, 844, 168, "OUTPUT", size=15, color=SUB, bold=True)
    text(page, 944, 168, "professional.docx - rendered by real Word", size=14, color=DIM)
    inner = pymupdf.Rect(844, 196, 1524, 848)
    drawn = place_image(page, SRC_PAGE, inner)

    # VERIFIED 印章：放在成品图下方的暗边距，不遮挡文档内容
    stamp = pymupdf.Rect(drawn.x1 - 256, out_rect.y1 - 92, drawn.x1 - 16, out_rect.y1 - 24)
    page.draw_rect(stamp, color=GREEN, fill=(0.075, 0.129, 0.098), radius=0.12, width=2)
    check(page, stamp.x0 + 22, stamp.y0 + 30, width=2.5)
    text(page, stamp.x0 + 44, stamp.y0 + 34, "VERIFIED", size=22, color=GREEN, bold=True)
    text(
        page,
        stamp.x0 + 22,
        stamp.y0 + 60,
        "9 checks  /  0.00% drift",
        size=13,
        font=MONO,
        color=(0.72, 0.86, 0.78),
    )

    pix = page.get_pixmap(matrix=pymupdf.Matrix(1, 1))
    pix.save(OUT)
    doc.close()
    print(
        "hero: %s (%dx%d)" % (os.path.relpath(OUT, KIT).replace(os.sep, "/"), pix.width, pix.height)
    )


if __name__ == "__main__":
    main()
