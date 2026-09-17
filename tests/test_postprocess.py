# -*- coding: utf-8 -*-
"""后处理冒烟断言：只需 pandoc + python-docx，不需要 Word / PDF，秒级可跑。

  pytest -q

覆盖 README「排版六条军规」里能被机器判定的部分：
跨页重复表头、题注居中、封面与目录注入、分节页码。
"""
import os
import json
import shutil
import subprocess
import sys

import pytest

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LUA_FILTER = os.path.join(KIT, "filters", "captions.lua")

pytestmark = pytest.mark.skipif(
    shutil.which("pandoc") is None, reason="需要 pandoc")


def _pandoc_cmd(md, body, res, cfg=None):
    """与 render.py 保持一致：带上 lua filter，并在配置题注关键字时传 -M。"""
    cmd = ["pandoc", md, "-o", body,
           "--reference-doc=" + os.path.join(KIT, "ref.docx"),
           "--resource-path=" + res,
           "-f", "markdown+pipe_tables+raw_html", "--wrap=none"]
    if os.path.exists(LUA_FILTER):
        cmd.append("--lua-filter=" + LUA_FILTER)
        cw = (cfg or {}).get("caption_words") or {}
        for key, meta_name in (("table", "dk-table-words"),
                               ("figure", "dk-figure-words")):
            if cw.get(key):
                cmd += ["-M", "%s=%s" % (meta_name, ",".join(cw[key]))]
    return cmd


def _build(tmp_path):
    src = tmp_path / "src"
    src.mkdir()
    shutil.copy(os.path.join(KIT, "sample.md"), src / "01_sample.md")
    body = str(tmp_path / "body.docx")
    out = str(tmp_path / "out.docx")
    subprocess.run(_pandoc_cmd(str(src / "01_sample.md"), body, str(src)),
                   check=True, capture_output=True)
    subprocess.run(
        [sys.executable, os.path.join(KIT, "post.py"), body, out,
         os.path.join(KIT, "sample_config.json")],
        check=True, capture_output=True)
    return out


def _build_md(tmp_path, md_text, cfg):
    """用自定义 md 与 config 走一遍 pandoc + post.py，返回输出 docx 路径。"""
    src = tmp_path / "src"
    src.mkdir(exist_ok=True)
    (src / "01.md").write_text(md_text, encoding="utf-8")
    cfg_path = tmp_path / "cfg.json"
    cfg_path.write_text(json.dumps(cfg, ensure_ascii=False), encoding="utf-8")
    body, out = str(tmp_path / "body.docx"), str(tmp_path / "out.docx")
    subprocess.run(_pandoc_cmd(str(src / "01.md"), body, str(src), cfg),
                   check=True, capture_output=True)
    subprocess.run(
        [sys.executable, os.path.join(KIT, "post.py"), body, out, str(cfg_path)],
        check=True, capture_output=True)
    return out


@pytest.fixture(scope="module")
def docx_path(tmp_path_factory):
    return _build(tmp_path_factory.mktemp("build"))


def test_tables_repeat_header(docx_path):
    """军规 3：每个表格首行必须带 w:tblHeader，跨页才会重复表头。"""
    from docx import Document
    from docx.oxml.ns import qn
    doc = Document(docx_path)
    assert doc.tables, "样例里应当有表格"
    for i, tbl in enumerate(doc.tables):
        trPr = tbl.rows[0]._tr.find(qn("w:trPr"))
        assert trPr is not None, "表 %d 首行缺少 trPr" % i
        assert trPr.find(qn("w:tblHeader")) is not None, "表 %d 首行缺少 tblHeader" % i


def test_captions_centered(docx_path):
    """军规 4：「表 x-y 标题」应被识别为表题并居中（sample.md 里有 2 处）。"""
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    doc = Document(docx_path)
    caps = [p for p in doc.paragraphs
            if p.text.strip().startswith("表 ") and "响应表" in p.text or
            p.text.strip().startswith("表 2-1")]
    assert len(caps) == 2, "应识别到 2 条表题，实际 %d" % len(caps)
    for p in caps:
        assert p.alignment == WD_ALIGN_PARAGRAPH.CENTER, "表题未居中: %s" % p.text


def test_cover_and_toc_injected(docx_path):
    """封面行与目录域必须注入，且目录在正文之前。"""
    from docx import Document
    doc = Document(docx_path)
    texts = [p.text.strip() for p in doc.paragraphs]
    assert "投 标 文 件" in texts, "封面标题未注入"
    toc_idx = next((i for i, t in enumerate(texts) if "目录" in t), None)
    assert toc_idx is not None, "目录标题未注入"
    body_idx = next((i for i, t in enumerate(texts) if "投标函" in t), None)
    assert body_idx is not None, "未找到正文第一章"
    assert toc_idx < body_idx, "目录必须排在正文之前"


def test_auto_number_and_crossref(tmp_path):
    """P1：表/图按章自动编号（含已手写编号的重排）+ @tab:/@fig: 交叉引用。"""
    from docx import Document
    md = (
        "# 第一章 概述\n\n"
        "表 9-9 商务条款响应表 @tab:clause\n\n"      # 故意写错编号，应被重排
        "| 条款 | 响应 |\n|:---|:---|\n| 工期 | 完全响应 |\n\n"
        "详见 @tab:clause 与 @tab:mod。\n\n"
        "# 第二章 方案\n\n"
        "表 模块清单 @tab:mod\n\n"                    # 不写编号，应自动插入
        "| 模块 | 说明 |\n|:---|:---|\n| 接入层 | 设备接入 |\n"
    )
    out = _build_md(tmp_path, md, {"auto_number": True})
    texts = [p.text.strip() for p in Document(out).paragraphs]

    cap1 = next(t for t in texts if "商务条款响应表" in t)
    cap2 = next(t for t in texts if "模块清单" in t)
    assert cap1.startswith("表 1-1"), "章内首个表题应为 表 1-1，实际：%s" % cap1
    assert cap2.startswith("表 2-1"), "第二章表题应为 表 2-1，实际：%s" % cap2
    assert "@tab" not in cap1 and "@tab" not in cap2, "标签标记应从题注里剥掉"

    ref = next(t for t in texts if "详见" in t)
    assert "表 1-1" in ref and "表 2-1" in ref, "交叉引用未替换：%s" % ref
    assert "@tab" not in ref, "引用标记未替换干净：%s" % ref


def test_auto_number_off_by_default(tmp_path):
    """未开启时不得改动原文。"""
    from docx import Document
    md = "# 第一章\n\n表 9-9 商务条款响应表\n\n| a |\n|:--|\n| 1 |\n"
    out = _build_md(tmp_path, md, {})
    texts = [p.text.strip() for p in Document(out).paragraphs]
    assert "表 9-9 商务条款响应表" in texts, "默认应保持手写编号不变"


def test_two_sections_with_page_number(docx_path):
    """封面/目录为第 1 节，正文为第 2 节，正文节页脚有页码域。"""
    from docx import Document
    from docx.oxml.ns import qn
    doc = Document(docx_path)
    assert len(doc.sections) == 2, "应为 2 节，实际 %d" % len(doc.sections)
    footer_xml = doc.sections[-1].footer.paragraphs[0]._p.xml
    assert "PAGE" in footer_xml, "正文节页脚缺少页码域"


def test_caption_variants_renumbered(tmp_path):
    """全角编号、英文关键字、写错的编号都要识别并重排，且居中。"""
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    md = (
        "# 第一章 测试\n\n"
        "表1.1 窄格式表题\n\n| a |\n|:--|\n| 1 |\n\n"
        "表 １－１ 全角编号\n\n| a |\n|:--|\n| 1 |\n\n"
        "Figure 1-1 英文图注\n\n"
        "Table 1-9 写错的表号\n\n| a |\n|:--|\n| 1 |\n"
    )
    out = _build_md(tmp_path, md, {"auto_number": True})
    doc = Document(out)
    texts = [p.text.strip() for p in doc.paragraphs]
    for want in ("表 1-1 窄格式表题", "表 1-2 全角编号",
                 "图 1-1 英文图注", "表 1-3 写错的表号"):
        assert want in texts, "未得到「%s」，实际：%s" % (want, texts)
    caps = [p for p in doc.paragraphs if p.text.strip().startswith(("表 1-", "图 1-"))]
    assert caps, "没有题注被编号"
    assert all(p.alignment == WD_ALIGN_PARAGRAPH.CENTER for p in caps), "题注未居中"


def test_figure_caption_auto_number(tmp_path):
    """图注（Lua filter 标为 FigureCaption）与正文引用。"""
    from docx import Document
    md = "# 第一章 测试\n\n图 架构示意 @fig:arch\n\n如 @fig:arch 所示。\n"
    out = _build_md(tmp_path, md, {"auto_number": True})
    texts = [p.text.strip() for p in Document(out).paragraphs]
    assert "图 1-1 架构示意" in texts, "图注未编号：%s" % texts
    ref = next(t for t in texts if "所示" in t)
    assert "图 1-1" in ref, "图注引用未替换：%s" % ref


GRID_MD = (
    "# 第一章 测试\n\n"
    "表 1-1 多级表头 @tab:t1\n\n"
    "+------------------+------------------+\n"
    "| 商务部分         | 技术部分         |\n"
    "+--------+---------+--------+---------+\n"
    "| 条款   | 响应    | 模块   | 说明   |\n"
    "+========+=========+========+=========+\n"
    "| 工期   | 完全响应| 接入层 | 设备   |\n"
    "+--------+---------+--------+---------+\n"
    "| 质保   | 优于要求| 平台层 | 治理   |\n"
    "+--------+---------+--------+---------+\n"
)


def _has_tbl_header(row):
    from docx.oxml.ns import qn
    trPr = row._tr.find(qn("w:trPr"))
    return trPr is not None and trPr.find(qn("w:tblHeader")) is not None


def test_grid_table_colspan_preserved(tmp_path):
    """grid table 的合并单元格要原样带过来。"""
    from docx import Document
    from docx.oxml.ns import qn
    out = _build_md(tmp_path, GRID_MD, {})
    row0 = Document(out).tables[0].rows[0]._tr.findall(qn("w:tc"))
    spans = []
    for tc in row0:
        tcPr = tc.find(qn("w:tcPr"))
        gs = tcPr.find(qn("w:gridSpan")) if tcPr is not None else None
        spans.append(gs.get(qn("w:val")) if gs is not None else None)
    assert spans == ["2", "2"], "合并单元格丢失：%s" % spans


def _shaded(cell):
    from docx.oxml.ns import qn
    tcPr = cell._tc.find(qn("w:tcPr"))
    sh = tcPr.find(qn("w:shd")) if tcPr is not None else None
    return sh is not None and sh.get(qn("w:fill")) == "EDEDED"


def test_pandoc_sets_tblheader_natively(tmp_path):
    """记录事实：tblHeader 是 pandoc 原生就给的（+===+ 以上全是表头行）。

    post.py 里的 set_repeat_header 只是幂等加固，不是这个功能的实现者。
    """
    from docx import Document
    out = _build_md(tmp_path, GRID_MD, {})
    rows = Document(out).tables[0].rows
    assert _has_tbl_header(rows[0]) and _has_tbl_header(rows[1])
    assert not _has_tbl_header(rows[2])


def test_auto_header_rows_from_pandoc(tmp_path):
    """默认自动识别：pandoc 标了几行表头，就给几行上灰底。"""
    from docx import Document
    out = _build_md(tmp_path, GRID_MD, {})
    rows = Document(out).tables[0].rows
    assert _shaded(rows[0].cells[0]) and _shaded(rows[1].cells[0]), "多级表头未全部灰底"
    assert not _shaded(rows[2].cells[0]), "数据行被误当成表头"


def test_auto_header_rows_per_table(tmp_path):
    """同一文档里不同表表头行数不同，必须各算各的（这是全局配置做不到的）。"""
    from docx import Document
    md = ("# 第一章\n\n"
          "| a | b |\n|:--|:--|\n| 1 | 2 |\n\n"          # 单行表头
          + GRID_MD.split("# 第一章 测试\n\n", 1)[1])       # 两行表头
    out = _build_md(tmp_path, md, {})
    t0, t1 = Document(out).tables
    assert _shaded(t0.rows[0].cells[0]) and not _shaded(t0.rows[1].cells[0]), \
        "pipe table 只该第一行灰底"
    assert _shaded(t1.rows[0].cells[0]) and _shaded(t1.rows[1].cells[0]), \
        "grid table 该前两行灰底"


def test_header_rows_override(tmp_path):
    """显式给 header_rows 时以配置为准（覆盖自动识别）。"""
    from docx import Document
    out = _build_md(tmp_path, GRID_MD, {"style": {"header_rows": 1}})
    rows = Document(out).tables[0].rows
    assert _shaded(rows[0].cells[0])
    assert not _shaded(rows[1].cells[0]), "显式设为 1 时第二行不该有灰底"


def test_three_line_table(tmp_path):
    """三线表：内部无框线，顶底线加粗，表头行有下边框。"""
    from docx import Document
    from docx.oxml.ns import qn
    out = _build_md(tmp_path, GRID_MD, {"style": {"table_border": "three"}})
    tbl = Document(out).tables[0]
    borders = tbl._tbl.tblPr.find(qn("w:tblBorders"))

    def val(edge):
        el = borders.find(qn("w:" + edge))
        return el.get(qn("w:val")) if el is not None else None

    assert val("insideH") == "nil" and val("insideV") == "nil", "三线表内部不应有框线"
    assert val("top") == "single" and borders.find(qn("w:top")).get(qn("w:sz")) == "12"
    # 表头下线画在「最后一行表头」上；本表自动识别为 2 行表头，所以查 rows[1]
    last_head = tbl.rows[1].cells[0]._tc.find(qn("w:tcPr")).find(qn("w:tcBorders"))
    assert last_head is not None and last_head.find(qn("w:bottom")) is not None, "缺表头下线"


def test_caption_keeps_with_table(tmp_path):
    """表题不能与表格分家（keep_with_next）。"""
    from docx import Document
    out = _build_md(tmp_path, "# 第一章\n\n表 1-1 清单\n\n| a |\n|:--|\n| 1 |\n", {})
    cap = next(p for p in Document(out).paragraphs if p.text.strip().startswith("表 1-1"))
    assert cap.paragraph_format.keep_with_next, "表题未设 keep_with_next，可能与表格分页"


def test_page_number_default(tmp_path):
    """默认页码样式 — N —。"""
    from docx import Document
    out = _build_md(tmp_path, "# 第一章\n\n正文。\n", {})
    footer = Document(out).sections[-1].footer.paragraphs[0].text
    assert "—" in footer and "1" in footer, "默认页码异常：%r" % footer


def test_style_cfg_overrides(tmp_path):
    """config 的 style 段能改页码模板 / 题注颜色 / 字号。"""
    from docx import Document
    from docx.shared import RGBColor
    md = "# 第一章\n\n表 清单 @tab:list\n\n| a |\n|:--|\n| 1 |\n"
    cfg = {"auto_number": True,
           "style": {"page_number": "第 {n} 页", "toc_depth": "1-3",
                     "caption_gray": "FF0000", "caption_size": 9.0}}
    out = _build_md(tmp_path, md, cfg)
    doc = Document(out)

    footer = doc.sections[-1].footer.paragraphs[0].text
    assert "第" in footer and "页" in footer, "页码模板未生效：%r" % footer
    assert "—" not in footer, "旧页码样式残留：%r" % footer

    cap = next(p for p in doc.paragraphs if "清单" in p.text)
    assert cap.runs[0].font.color.rgb == RGBColor(0xFF, 0x00, 0x00), "题注颜色未生效"
    assert cap.runs[0].font.size.pt == 9.0, "题注字号未生效"


def test_style_cfg_numeric_and_spacing(tmp_path):
    """页眉/页码字号、单元格边距、边框、题注间距都要能被 style 段改。"""
    from docx import Document
    from docx.oxml.ns import qn
    md = "# 第一章\n\n表 清单 @tab:t\n\n| a | b |\n|:--|:--|\n| 1 | 2 |\n"
    out = _build_md(tmp_path, md, {
        "auto_number": True,
        "header": "页眉文字",
        "style": {"header_size": 12.0, "page_number_size": 14.0,
                  "cell_margin_h": 200, "border_size": 12,
                  "border_color": "000000", "caption_space_before": 12.0}})
    doc = Document(out)

    footer_run = doc.sections[-1].footer.paragraphs[0].runs[0]
    assert footer_run.font.size.pt == 14.0, "页码字号未生效"
    hdr = doc.sections[-1].header.paragraphs[0].runs[0]
    assert hdr.font.size.pt == 12.0, "页眉字号未生效"

    tbl = doc.tables[0]
    mar = tbl._tbl.tblPr.find(qn("w:tblCellMar"))
    assert mar.find(qn("w:left")).get(qn("w:w")) == "200", "单元格边距未生效"
    top = tbl._tbl.tblPr.find(qn("w:tblBorders")).find(qn("w:top"))
    assert top.get(qn("w:sz")) == "12" and top.get(qn("w:color")) == "000000"

    cap = next(p for p in doc.paragraphs if p.text.strip().startswith("表 1-1"))
    assert cap.paragraph_format.space_before.pt == 12.0, "题注段前间距未生效"


def test_caption_words_configurable(tmp_path):
    """题注关键字可配（post.py 与 lua filter 同步）。"""
    from docx import Document
    md = "# 第一章\n\n照片 架构示意 @fig:a\n\n如 @fig:a 所示。\n"
    words = {"table": ["清单"], "figure": ["照片"]}

    off = _build_md(tmp_path, md, {"auto_number": True})
    assert "图 1-1" not in [p.text for p in Document(off).paragraphs], \
        "默认关键字下不该把「照片」当图注"

    on = _build_md(tmp_path, md, {"auto_number": True, "caption_words": words})
    texts = [p.text.strip() for p in Document(on).paragraphs]
    assert "图 1-1 架构示意" in texts, "自定义关键字未生效：%s" % texts


def test_to_rgb_accepts_common_forms():
    import importlib.util
    spec = importlib.util.spec_from_file_location(
        "post", os.path.join(KIT, "post.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    assert mod.to_rgb("404040") == (0x40, 0x40, 0x40)
    assert mod.to_rgb("#FF0000") == (0xFF, 0, 0)
    assert mod.to_rgb("0x010203") == (1, 2, 3)
    assert mod.to_rgb([1, 2, 3]) == (1, 2, 3)
    with pytest.raises(ValueError):
        mod.to_rgb("abc")


def test_render_version_flag():
    r = subprocess.run([sys.executable, os.path.join(KIT, "render.py"), "--version"],
                       capture_output=True)
    assert r.returncode == 0
    assert b"docx-kit" in r.stdout, r.stdout


def test_missing_h1_exits_with_hint(tmp_path):
    """缺一级标题时要给可诊断提示，而不是抛 StopIteration。"""
    src = tmp_path / "src"
    src.mkdir()
    (src / "01.md").write_text("## 只有二级标题\n\n正文。\n", encoding="utf-8")
    body, out = str(tmp_path / "body.docx"), str(tmp_path / "out.docx")
    subprocess.run(_pandoc_cmd(str(src / "01.md"), body, str(src)),
                   check=True, capture_output=True)
    r = subprocess.run(
        [sys.executable, os.path.join(KIT, "post.py"), body, out, ""],
        capture_output=True, env=dict(os.environ, PYTHONIOENCODING="utf-8"))
    assert r.returncode != 0, "缺一级标题时必须失败"
    assert "一级标题" in r.stderr.decode("utf-8", "replace")


def test_crossref_inside_table_cell(tmp_path):
    """表格单元格里的引用也要被替换。"""
    from docx import Document
    md = ("# 第一章\n\n表 清单 @tab:list\n\n"
          "| 说明 | 备注 |\n|:--|:--|\n| 见 @tab:list | ok |\n")
    out = _build_md(tmp_path, md, {"auto_number": True})
    cell = Document(out).tables[0].rows[1].cells[0].text
    assert "表 1-1" in cell, "单元格内引用未替换：%s" % cell
