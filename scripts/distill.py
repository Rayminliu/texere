"""读一份已有 docx（甲方模板 / 自己攒的模板），蒸馏出 texere 的 config 建议。

目的：接一份新模板时，不用手工去 Word 里量页边距、查字号、抄页眉。
脚本只读不写源模板，输出一份「建议 config」+ 一份体检报告，人工确认后即可使用。

用法:
  python scripts/distill.py 甲方模板.docx                  # 打印报告 + 建议 config
  python scripts/distill.py 甲方模板.docx --out cfg.json    # 同时把建议 config 写成文件

能确定的：页面设置、正文/标题的字体字号、页眉页脚文字、分节数、表格数、
模板字体在本机是否可用（引用式排版能否落地的关键一环）。
确定不了的：封面结构、表格细节、页码格式——这些在报告里提示人工确认。

为什么叫"蒸馏"而不是"转换"：转换是让工具去理解模板（做不到，也不可靠）；
蒸馏是把模板拆成 config 字段，人来拍板。见 README「复用已有模板」。
"""

import argparse
import json
import os
import sys

from _ooxml import style_fonts
from _shared import force_utf8_stdio
from docx import Document

EMU_PER_CM = 360000


def cm(v):
    if v is None:
        return "n/a"
    return round(v / EMU_PER_CM, 2)


def _font_available(name):
    """查一个字体名在本机是否可用，返回 (可用?, 规范英文名|None)。

    docx 里只写字体名字，渲染时由系统按名解析；名字在本机不存在时 Word 会
    静默回退到别的字体——版式契约（reference_doc / profile 断言的字体与字号）
    就在无声中失效。此函数把「解析得到什么」提前亮出来，放进蒸馏报告。

    实测口径（Windows GDI+ flat API）：
      - GdipCreateFontFamilyFromName 按名解析走 GDI+ 字体集合（与 Word 同源），
        中文别名在内部处理：'宋体'/'SimSun'、'黑体'/'SimHei'、'楷体'/'KaiTi'
        均解析到同一族；真正缺失时返回状态码 FontFamilyNotFound(17)——
        不抛异常、不靠回落值猜，状态码即真相；
      - GdipGetFamilyName 的第三参是 LANGID：传 0 取进程 locale 名（zh-CN
        机器上请求 SimSun 也返回"宋体"），传 0x0409 (en-US) 才返回规范
        英文名。**报告统一展示英文名**——它是跨机器稳定不变的注册名，
        中文 locale 名换个机器就变，做不了对拍锚点。

    非 Windows 或探测自身失败 → (True, None)：查不了就不拦路（探测是提示
    不是门禁，宁缺毋假）。
    """
    if not name or os.name != "nt":
        return True, None
    try:
        import ctypes

        gdiplus = ctypes.windll.gdiplus
        token = ctypes.c_ulong()

        import ctypes.wintypes as wt

        class _StartupInput(ctypes.Structure):
            _fields_ = [
                ("GdiplusVersion", wt.UINT),
                ("DebugEventCallback", ctypes.c_void_p),
                ("SuppressBackgroundThread", ctypes.c_int),
                ("SuppressExternalCodecs", ctypes.c_int),
            ]

        si = _StartupInput(1, None, 0, 0)
        if gdiplus.GdiplusStartup(ctypes.byref(token), ctypes.byref(si), None) != 0:
            return True, None
        try:
            family = ctypes.c_void_p()
            status = gdiplus.GdipCreateFontFamilyFromName(
                ctypes.c_wchar_p(name), None, ctypes.byref(family)
            )
            if status == 0:  # Ok
                buf = ctypes.create_unicode_buffer(64)
                gdiplus.GdipGetFamilyName(family, buf, 0x0409)  # en-US：规范英文名
                gdiplus.GdipDeleteFontFamily(family)
                return True, buf.value
            return False, None  # FontFamilyNotFound(17) 等：名字在本机解析不到
        finally:
            gdiplus.GdiplusShutdown(token)
    except Exception:
        return True, None  # 探测失败不拦路：宁缺毋假


def style_font(st):
    """取 (中文字体, 西文字体, 字号pt)。中文字体在 rFonts 的 eastAsia 上。"""
    return style_fonts(st)


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
        "  纸型     : %s x %s cm（%s）"
        % (
            cm(sec.page_width),
            cm(sec.page_height),
            "横向" if str(sec.orientation).endswith("LANDSCAPE") else "纵向",
        )
    )
    report.append(
        "  页边距   : 上%s 下%s 左%s 右%s cm"
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
        line = "  %-6s : 中文=%s  西文=%s  %s" % (
            label,
            east or "-",
            latin or "-",
            ("%.1fpt" % size) if size else "字号未指定",
        )
        # 字体可用性：名字只写进 docx，渲染时系统按名解析，缺了就静默回退——
        # 版式契约无声失效。这里提前亮「解析得到什么」，缺的指路安装。
        for fname in (east, latin):
            if not fname:
                continue
            ok, resolved = _font_available(fname)
            if not ok:
                line += "\n           [字体缺失] %s 在本机解析不到（Word 将静默回退，" % fname
                line += "版式会悄悄变）— 安装该字体，或改用本机有的等价字体"
                break
            if resolved and resolved != fname:
                line += "（%s→%s）" % (fname, resolved)
        report.append(line)
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
        uniq = list(dict.fromkeys(headers))  # 保序去重：uniq[0] 就是文档里最先出现的页眉
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
    # 副作用只在入口执行（且必须在 argparse 之前：--help / 报错文案也要走 UTF-8）
    force_utf8_stdio()

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
