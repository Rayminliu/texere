"""后处理冒烟断言：只需 pandoc + python-docx，不需要 Word / PDF，秒级可跑。

  pytest -q

覆盖 README「排版六条军规」里能被机器判定的部分：
跨页重复表头、题注居中、封面与目录注入、分节页码。
"""

import json
import os
import re
import shutil
import subprocess
import sys

import pytest

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LUA_FILTER = os.path.join(KIT, "scripts", "filters", "captions.lua")

pytestmark = pytest.mark.skipif(shutil.which("pandoc") is None, reason="需要 pandoc")


def _pandoc_cmd(md, body, res, cfg=None):
    """与 render.py 保持一致：带上 lua filter，并在配置题注关键字时传 -M。"""
    cmd = [
        "pandoc",
        md,
        "-o",
        body,
        "--reference-doc=" + os.path.join(KIT, "assets", "ref.docx"),
        "--resource-path=" + res,
        "-f",
        "markdown+pipe_tables+raw_html",
        "--wrap=none",
    ]
    if os.path.exists(LUA_FILTER):
        cmd.append("--lua-filter=" + LUA_FILTER)
        cw = (cfg or {}).get("caption_words") or {}
        for key, meta_name in (
            ("table", "dk-table-words"),
            ("figure", "dk-figure-words"),
        ):
            if cw.get(key):
                cmd += ["-M", "%s=%s" % (meta_name, ",".join(cw[key]))]
    return cmd


def _build(tmp_path):
    src = tmp_path / "src"
    src.mkdir()
    shutil.copy(os.path.join(KIT, "assets", "sample.md"), src / "01_sample.md")
    body = str(tmp_path / "body.docx")
    out = str(tmp_path / "out.docx")
    subprocess.run(
        _pandoc_cmd(str(src / "01_sample.md"), body, str(src)),
        check=True,
        capture_output=True,
    )
    subprocess.run(
        [
            sys.executable,
            os.path.join(KIT, "scripts", "post.py"),
            body,
            out,
            os.path.join(KIT, "assets", "sample_config.json"),
        ],
        check=True,
        capture_output=True,
    )
    return out


def _build_md(tmp_path, md_text, cfg):
    """用自定义 md 与 config 走一遍 pandoc + post.py，返回输出 docx 路径。"""
    src = tmp_path / "src"
    src.mkdir(exist_ok=True)
    (src / "01.md").write_text(md_text, encoding="utf-8")
    cfg_path = tmp_path / "cfg.json"
    cfg_path.write_text(json.dumps(cfg, ensure_ascii=False), encoding="utf-8")
    body, out = str(tmp_path / "body.docx"), str(tmp_path / "out.docx")
    subprocess.run(
        _pandoc_cmd(str(src / "01.md"), body, str(src), cfg),
        check=True,
        capture_output=True,
    )
    subprocess.run(
        [
            sys.executable,
            os.path.join(KIT, "scripts", "post.py"),
            body,
            out,
            str(cfg_path),
        ],
        check=True,
        capture_output=True,
    )
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
    caps = [
        p
        for p in doc.paragraphs
        if p.text.strip().startswith("表 ")
        and "响应表" in p.text
        or p.text.strip().startswith("表 2-1")
    ]
    assert len(caps) == 2, "应识别到 2 条表题，实际 %d" % len(caps)
    for p in caps:
        assert p.alignment == WD_ALIGN_PARAGRAPH.CENTER, "表题未居中: %s" % p.text


def test_english_caption_styled(tmp_path):
    """英文题注（Table 1: ...）同样被识别并居中——neutral-en 语境下的管线事实。"""
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    md = (
        "Intro paragraph before the table.\n"
        "\n"
        "Table 1: Quarterly milestone summary\n"
        "\n"
        "| Milestone | Status |\n"
        "|-----------|--------|\n"
        "| M1        | Done   |\n"
    )
    docx = _build_md(tmp_path, md, {})
    doc = Document(docx)
    caps = [p for p in doc.paragraphs if p.text.strip().startswith("Table 1:")]
    assert len(caps) == 1, "英文题注应被识别，实际 %d" % len(caps)
    assert caps[0].alignment == WD_ALIGN_PARAGRAPH.CENTER, "英文表题未居中: %s" % caps[0].text


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


def _has_toc_field(path):
    """document.xml 里是否有 TOC 域（instrText 含 'TOC \\o'）。"""
    import zipfile

    with zipfile.ZipFile(path) as z:
        return "TOC \\o" in z.read("word/document.xml").decode("utf-8")


MD_WITH_HEADINGS = """# 第一章 总则

正文一二三。

## 1.1 细则

细则内容。
"""


def test_toc_false_skips_toc_keeps_headings(tmp_path):
    """config "toc": false：不插目录域、不分节，但 Heading 样式照常保留。

    这是通知/公示类短文档的正路；旧写法只能不写 # 绕过，代价是丢标题样式。
    """
    from docx import Document

    out = _build_md(tmp_path, MD_WITH_HEADINGS, {"toc": False})
    assert not _has_toc_field(out), "toc:false 时不该有 TOC 域"

    doc = Document(out)
    styles = [p.style.name for p in doc.paragraphs if p.text.strip()]
    assert "Heading 1" in styles, "toc:false 不能把标题降级成普通段落"
    assert "Heading 2" in styles
    assert len(doc.sections) == 1, "无封面又无目录时不该分节（会多出空白首页）"


def test_toc_false_with_cover_still_sections(tmp_path):
    """toc:false + 封面：封面单独成节，页码体系照常，只是没有目录页。"""
    from docx import Document

    cfg = {"toc": False, "cover": [["CoverTitle", "通 知"]]}
    out = _build_md(tmp_path, MD_WITH_HEADINGS, cfg)
    assert not _has_toc_field(out)
    doc = Document(out)
    assert len(doc.sections) == 2, "有封面时仍应分节（封面/正文各自页码）"
    assert "通 知" in [p.text.strip() for p in doc.paragraphs]


def test_never_touches_text(tmp_path):
    """核心契约：post.py 只改版式，一个字都不改内容。

    曾经开启过自动编号功能，它改写过题注文字（`附件 8-1 …` → `图 8-1 附件 8-1 …`）
    并篡改过正文（`**表层…**` → `表 3-5 层…`）。该功能已删除，这里守住契约。
    """
    from docx import Document

    md = (
        "# 第一章\n\n表 9-9 商务条款响应表\n\n| a |\n|:--|\n| 1 |\n\n"
        "- **表层（边缘轻算力）：** 基于公开预训练模型迁移微调。\n"
    )
    out = _build_md(tmp_path, md, {})
    texts = [p.text.strip() for p in Document(out).paragraphs]
    assert "表 9-9 商务条款响应表" in texts, "题注文字被改动了"
    assert any(t.startswith("表层（边缘轻算力）") for t in texts), "正文被改动了：%s" % [
        t for t in texts if "表层" in t
    ]


def test_two_sections_with_page_number(docx_path):
    """封面/目录为第 1 节，正文为第 2 节，正文节页脚有页码域。"""
    from docx import Document

    doc = Document(docx_path)
    assert len(doc.sections) == 2, "应为 2 节，实际 %d" % len(doc.sections)
    footer_xml = doc.sections[-1].footer.paragraphs[0]._p.xml
    assert "PAGE" in footer_xml, "正文节页脚缺少页码域"


def test_caption_variants_centered(tmp_path):
    """全角编号、英文关键字都要识别为题注（居中），但文字保持原样。"""
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    md = (
        "# 第一章 测试\n\n"
        "表1.1 窄格式表题\n\n| a |\n|:--|\n| 1 |\n\n"
        "表 １－１ 全角编号\n\n| a |\n|:--|\n| 1 |\n\n"
        "Figure 1-1 英文图注\n\n"
        "Table 1-9 另一张表\n\n| a |\n|:--|\n| 1 |\n"
    )
    out = _build_md(tmp_path, md, {})
    doc = Document(out)
    texts = [p.text.strip() for p in doc.paragraphs]
    for want in (
        "表1.1 窄格式表题",
        "表 １－１ 全角编号",
        "Figure 1-1 英文图注",
        "Table 1-9 另一张表",
    ):
        assert want in texts, "题注文字被改动了：缺「%s」，实际：%s" % (want, texts)
    caps = [
        p
        for p in doc.paragraphs
        if p.text.strip().startswith(("表1.1", "表 １", "Figure 1-1", "Table 1-9"))
    ]
    assert len(caps) == 4, "题注未全部识别：%s" % [p.text[:20] for p in caps]
    assert all(p.alignment == WD_ALIGN_PARAGRAPH.CENTER for p in caps), "题注未居中"


def test_figure_caption_centered(tmp_path):
    """图注（Lua filter 标为 FigureCaption）居中，文字原样。"""
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    md = "# 第一章 测试\n\n图 架构示意\n\n正文段落。\n"
    out = _build_md(tmp_path, md, {})
    doc = Document(out)
    cap = next(p for p in doc.paragraphs if p.text.strip().startswith("图 架构示意"))
    assert cap.alignment == WD_ALIGN_PARAGRAPH.CENTER, "图注未居中"


GRID_MD = (
    "# 第一章 测试\n\n"
    "表 1-1 多级表头\n\n"
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


HEADERLESS_MD = (
    "# 第一章\n\n表 1-1 表单式表格\n\n"
    "+-----------+-----------+\n"
    "| 项目名称  | 某某项目  |\n"
    "+-----------+-----------+\n"
    "| 负责人    | 张三      |\n"
    "+-----------+-----------+\n"
)


def test_header_rows_zero_for_form_tables(tmp_path):
    """表单类表格：显式 header_rows=0 时首行不上灰底、也不跨页重复。

    （外部 docx 转 md 后 pandoc 会给首行加 tblHeader，表单的「项目名称」行
    看起来就成了列标题，视觉上不对。）
    """
    from docx import Document

    out = _build_md(tmp_path, HEADERLESS_MD, {"style": {"header_rows": 0}})
    tbl = Document(out).tables[0]
    assert not _shaded(tbl.rows[0].cells[0]), "header_rows=0 时首行不该有灰底"
    assert not _has_tbl_header(tbl.rows[0]), "header_rows=0 时不该跨页重复表头"


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

    md = (
        "# 第一章\n\n"
        "| a | b |\n|:--|:--|\n| 1 | 2 |\n\n" + GRID_MD.split("# 第一章 测试\n\n", 1)[1]  # 单行表头
    )  # 两行表头
    out = _build_md(tmp_path, md, {})
    t0, t1 = Document(out).tables
    assert _shaded(t0.rows[0].cells[0]) and not _shaded(t0.rows[1].cells[0]), (
        "pipe table 只该第一行灰底"
    )
    assert _shaded(t1.rows[0].cells[0]) and _shaded(t1.rows[1].cells[0]), "grid table 该前两行灰底"


def test_header_rows_override(tmp_path):
    """显式给 header_rows 时以配置为准（覆盖自动识别）。"""
    from docx import Document

    out = _build_md(tmp_path, GRID_MD, {"style": {"header_rows": 1}})
    rows = Document(out).tables[0].rows
    assert _shaded(rows[0].cells[0])
    assert not _shaded(rows[1].cells[0]), "显式设为 1 时第二行不该有灰底"


def test_table_zebra_and_header_color(tmp_path):
    """表头可换深色底配白字，表体可加斑马纹（借用通用 docx 技能的表格规范）。

    默认（浅灰底 + 黑字 + 无斑马纹）走的是中文正式文档惯例，保持不动。
    """
    from docx import Document
    from docx.oxml.ns import qn
    from docx.shared import RGBColor

    md = "# 第一章\n\n表 1-1 清单\n\n| 序号 | 名称 |\n|:--|:--|\n| 1 | a |\n| 2 | b |\n| 3 | c |\n"
    out = _build_md(
        tmp_path,
        md,
        {
            "style": {
                "table_shade": "4472C4",
                "table_header_color": "FFFFFF",
                "table_zebra": True,
                "table_zebra_fill": "F7F7F7",
            }
        },
    )
    tbl = Document(out).tables[0]

    def fill(ri):
        tcPr = tbl.rows[ri].cells[0]._tc.find(qn("w:tcPr"))
        shd = tcPr.find(qn("w:shd")) if tcPr is not None else None
        return shd.get(qn("w:fill")) if shd is not None else None

    assert fill(0) == "4472C4", "表头底纹未生效"
    run = tbl.rows[0].cells[0].paragraphs[0].runs[0]
    assert run.font.color.rgb == RGBColor(0xFF, 0xFF, 0xFF), "表头白字未生效"
    assert fill(1) is None, "首条数据行不该有斑马纹底"
    assert fill(2) == "F7F7F7", "第 2 条数据行应有斑马纹底"
    assert fill(3) is None, "斑马纹应隔行"


def test_table_style_defaults_unchanged(tmp_path):
    """不配 style 时保持中文惯例：浅灰表头、无斑马纹。"""
    from docx import Document
    from docx.oxml.ns import qn

    md = "# 第一章\n\n表 1-1 清单\n\n| 序号 | 名称 |\n|:--|:--|\n| 1 | a |\n| 2 | b |\n"
    tbl = Document(_build_md(tmp_path, md, {})).tables[0]

    def fill(ri):
        tcPr = tbl.rows[ri].cells[0]._tc.find(qn("w:tcPr"))
        shd = tcPr.find(qn("w:shd")) if tcPr is not None else None
        return shd.get(qn("w:fill")) if shd is not None else None

    assert fill(0) == "EDEDED", "默认表头应为浅灰"
    assert fill(1) is None, "默认不该有斑马纹"


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

    md = "# 第一章\n\n表 清单\n\n| a |\n|:--|\n| 1 |\n"
    cfg = {
        "style": {
            "page_number": "第 {n} 页",
            "toc_depth": "1-3",
            "caption_gray": "FF0000",
            "caption_size": 9.0,
        }
    }
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

    md = "# 第一章\n\n表 清单\n\n| a | b |\n|:--|:--|\n| 1 | 2 |\n"
    out = _build_md(
        tmp_path,
        md,
        {
            "header": "页眉文字",
            "style": {
                "header_size": 12.0,
                "page_number_size": 14.0,
                "cell_margin_h": 200,
                "border_size": 12,
                "border_color": "000000",
                "caption_space_before": 12.0,
            },
        },
    )
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

    cap = next(p for p in doc.paragraphs if p.text.strip().startswith("表 清单"))
    assert cap.paragraph_format.space_before.pt == 12.0, "题注段前间距未生效"


def test_temp_dir_cleaned_on_failure(tmp_path):
    """渲染失败时也要清掉临时目录。

    以前只在 render() 末尾 rmtree，任何 sys.exit（配置有误、子进程报错）都会绕过它，
    临时目录就烂在 %TEMP% 里——实测一天攒了 12 个。
    """
    import glob
    import tempfile

    pat = os.path.join(tempfile.gettempdir(), "texere_*")
    before = set(glob.glob(pat))

    src = tmp_path / "src"
    src.mkdir()
    (src / "01.md").write_text("# 第一章\n\n正文。\n", encoding="utf-8")
    cfg = tmp_path / "cfg.json"
    cfg.write_text('{"content_fixes_file": "这个文件不存在.py"}', encoding="utf-8")

    r = subprocess.run(
        [
            sys.executable,
            os.path.join(KIT, "scripts", "render.py"),
            "--src",
            str(src),
            "--out",
            str(tmp_path / "o.docx"),
            "--config",
            str(cfg),
        ],
        capture_output=True,
        env=dict(os.environ, PYTHONIOENCODING="utf-8"),
    )
    assert r.returncode != 0, "这个配置应当失败"

    leaked = sorted(set(glob.glob(pat)) - before)
    assert not leaked, "失败后临时目录未清理：%s" % leaked


def test_render_finds_sibling_media(tmp_path):
    """图片在 src 的**兄弟**目录时也要能找到（真实项目就是这种布局：

    build/src/*.md 引用了 build/media_plan/*.jpg）。
    """
    pymupdf = pytest.importorskip("pymupdf")
    proj = tmp_path / "proj"
    (proj / "src").mkdir(parents=True)
    (proj / "media").mkdir()
    pymupdf.open().new_page().get_pixmap().save(str(proj / "media" / "img.png"))
    (proj / "src" / "01.md").write_text(
        "# 第一章\n\n![图 1-1 兄弟目录图片](media/img.png)\n\n正文。\n",
        encoding="utf-8",
    )

    out = str(tmp_path / "out.docx")
    r = subprocess.run(
        [
            sys.executable,
            os.path.join(KIT, "scripts", "render.py"),
            "--src",
            str(proj / "src"),
            "--out",
            out,
        ],
        capture_output=True,
        env=dict(os.environ, PYTHONIOENCODING="utf-8"),
    )
    assert r.returncode == 0, r.stderr.decode("utf-8", "replace")

    from docx import Document

    assert len(Document(out).inline_shapes) == 1, "兄弟目录里的图片没被找到"
    assert "images: 1/1 ok" in r.stdout.decode("utf-8", "replace"), "缺图自检未通过"


def test_missing_image_exits_nonzero(tmp_path):
    """源 md 引了图却没嵌进去时，render 必须以非零码结束。

    外部实测：只跑 render.py（不接 validate）的 CI 场景，此前即便丢图退出码仍为 0，
    交付一份没图的标书还当成功。缺图先照常产出 docx（供肉眼确认），再让退出码诚实。
    """
    import glob
    import tempfile

    pat = os.path.join(tempfile.gettempdir(), "texere_*")
    before = set(glob.glob(pat))

    src = tmp_path / "src"
    src.mkdir()
    (src / "01.md").write_text(
        "# 第一章\n\n![图 1-1 缺失的图](not_exist.png)\n\n正文。\n",
        encoding="utf-8",
    )
    out = str(tmp_path / "out.docx")
    r = subprocess.run(
        [
            sys.executable,
            os.path.join(KIT, "scripts", "render.py"),
            "--src",
            str(src),
            "--out",
            out,
        ],
        capture_output=True,
        env=dict(os.environ, PYTHONIOENCODING="utf-8"),
    )
    combined = (r.stdout + r.stderr).decode("utf-8", "replace")
    assert r.returncode != 0, "缺图必须以非零码结束，实际退出码 %d" % r.returncode
    assert "[ERROR]" in combined, "缺图必须打印 [ERROR]：%s" % combined

    leaked = sorted(set(glob.glob(pat)) - before)
    assert not leaked, "缺图退出后临时目录未清理：%s" % leaked


def _render_module():
    """加载 scripts/render.py 本体（部分契约只能直接叫函数测）。"""
    import importlib.util

    scripts = os.path.join(KIT, "scripts")
    if scripts not in sys.path:  # render.py 依赖兄弟模块 _shared / renderers
        sys.path.insert(0, scripts)
    spec = importlib.util.spec_from_file_location(
        "render_under_test", os.path.join(scripts, "render.py")
    )
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


def test_code_fence_image_examples_are_not_counted_as_refs(tmp_path):
    """围栏代码块里的 `![...]` 是示例不是引图。

    render 的缺图自检此前全文数 `![`，一份带 markdown 写法教程的文档会凭空
    报「引用 3 张只嵌进 0 张」并以退出码 1 结束（与 validate._source_md_segments
    早就排围栏的规则不对齐）。
    """
    m = _render_module()

    fenced = (
        "# 标题\n\n```markdown\n![alt](img.png)\n![alt2](b.png)\n\t![tab](c.png)\n```\n\n正文。\n"
    )
    assert m._outside_code_fences(fenced).count("![") == 0
    # 全文三处 `![` 都在围栏里 → 引图计数为 0，不碰 docx 也不误报
    assert m.check_images(fenced, str(tmp_path / "unused.docx")) is True
    # 围栏外真引图时仍然数得到
    assert m._outside_code_fences(fenced + "![真实图](a.png)\n").count("![") == 1


# ------------------------------------------------ 验收确认后才回收中间产物


def _md_src(tmp_path):
    src = tmp_path / "src"
    src.mkdir(exist_ok=True)
    (src / "01.md").write_text("# 第一章\n\n正文一句话，长到不会被归一化跳过。\n", encoding="utf-8")
    return src


def _run_render(tmp_path, *extra):
    """跑一次 render（不带 --pdf，秒级）。stdin 置 DEVNULL：模拟脚本 / CI 的非交互调用。"""
    r = subprocess.run(
        [
            sys.executable,
            os.path.join(KIT, "scripts", "render.py"),
            "--src",
            str(_md_src(tmp_path)),
            "--out",
            str(tmp_path / "o.docx"),
            *extra,
        ],
        capture_output=True,
        stdin=subprocess.DEVNULL,
        env=dict(os.environ, PYTHONIOENCODING="utf-8"),
    )
    assert r.returncode == 0, r.stderr.decode("utf-8", "replace")
    return r.stdout.decode("utf-8", "replace")


def test_keep_work_leaves_intermediates_reusable_for_rework(tmp_path):
    """验收前中间产物必须在：--keep-work 留下的 body.docx 要能直接被重新排版。

    以前跑完就 rmtree，用户看过 PDF 说“这页要改”时，pandoc 中间件与合并正文
    已经没了，只能整链重跑。
    """
    out = _run_render(tmp_path, "--keep-work")
    m = re.search(r"已保留中间产物： (\S+?)（", out)
    assert m, out
    work = m.group(1)
    try:
        body = os.path.join(work, "body.docx")
        assert os.path.exists(body), "pandoc 中间件没留下，返修仍要重跑"
        assert os.path.exists(os.path.join(work, "all.md")), "合并后的正文没留下"
        # 留下的必须是可用的：拿它直接跑 post.py（返修路径，不重跑 pandoc）
        from docx import Document

        assert Document(body).paragraphs
        r = subprocess.run(
            [
                sys.executable,
                os.path.join(KIT, "scripts", "post.py"),
                body,
                str(tmp_path / "r.docx"),
            ],
            capture_output=True,
            stdin=subprocess.DEVNULL,
            env=dict(os.environ, PYTHONIOENCODING="utf-8"),
        )
        assert r.returncode == 0, r.stderr.decode("utf-8", "replace")
        assert os.path.exists(tmp_path / "r.docx")
    finally:
        shutil.rmtree(work, ignore_errors=True)


def test_noninteractive_run_still_reclaims_work(tmp_path):
    """非交互（脚本 / CI）不猜意图：维持旧的跑完即回收，并把 --keep-work 指路说清。"""
    import glob
    import tempfile

    pat = os.path.join(tempfile.gettempdir(), "texere_*")
    before = set(glob.glob(pat))
    out = _run_render(tmp_path)
    assert "非交互运行：中间产物已回收" in out, out
    assert "--keep-work" in out, "回收了但不说怎么留：" + out
    leaked = sorted(set(glob.glob(pat)) - before)
    assert not leaked, "非交互运行后临时目录未回收：%s" % leaked


def test_keep_and_discard_work_are_mutually_exclusive(tmp_path):
    r = subprocess.run(
        [
            sys.executable,
            os.path.join(KIT, "scripts", "render.py"),
            "--src",
            str(_md_src(tmp_path)),
            "--out",
            str(tmp_path / "o.docx"),
            "--keep-work",
            "--discard-work",
        ],
        capture_output=True,
        stdin=subprocess.DEVNULL,
        env=dict(os.environ, PYTHONIOENCODING="utf-8"),
    )
    assert r.returncode != 0
    assert "互斥" in r.stderr.decode("utf-8", "replace")


def test_workdir_prompt_only_asks_when_a_human_is_reading_it(tmp_path, monkeypatch):
    """确认那句只在「人在终端前」问，且 y 才删、回车不删。

    _should_prompt 必须两个流都是 TTY：pytest / CI 里 stdout 是管道，而 stdin
    可能仍然继承了终端——只查 stdin 会挂在那里等输入。
    """
    m = _render_module()
    assert m._should_prompt(True, True) is True
    assert m._should_prompt(True, False) is False
    assert m._should_prompt(False, True) is False

    keep = tmp_path / "keep"
    keep.mkdir()
    (keep / "body.docx").write_text("x", encoding="utf-8")
    wd = m._Workdir(str(keep))
    monkeypatch.setattr(m, "_should_prompt", lambda *a: True)
    monkeypatch.setattr("builtins.input", lambda *a: "")  # 直接回车 = 先留着等返修
    assert wd.settle(m._Workdir.AUTO) == "kept"
    assert keep.exists() and wd.contents() == ["body.docx"]

    drop = tmp_path / "drop"
    drop.mkdir()
    wd2 = m._Workdir(str(drop))
    monkeypatch.setattr("builtins.input", lambda *a: "y")  # 验收过了，是真过了
    assert wd2.settle(m._Workdir.AUTO) == "discarded"
    assert not drop.exists()


def test_workdir_fallback_reclaims_when_nobody_settles(tmp_path):
    """没走到 settle（异常 / sys.exit）时 atexit 兼容旧行为：回收，不涨磁盘。"""
    m = _render_module()
    d = tmp_path / "crashed"
    d.mkdir()
    (d / "all.md").write_text("x", encoding="utf-8")
    wd = m._Workdir(str(d))
    wd._fallback()
    assert not d.exists(), "失败路径没回收中间产物"


def test_images_survive_postprocess(tmp_path):
    """后处理绝不能把图片弄丢（曾被自动编号抹掉过 28 张）。"""
    pymupdf = pytest.importorskip("pymupdf")
    from docx import Document

    src = tmp_path / "src"
    src.mkdir(exist_ok=True)
    pymupdf.open().new_page().get_pixmap().save(str(src / "fig.png"))

    md = "# 第一章\n\n![图 架构示意](fig.png)\n\n正文引用。\n"
    out = _build_md(tmp_path, md, {})
    doc = Document(out)
    assert len(doc.inline_shapes) == 1, "图片被后处理弄丢了"


def test_caption_words_configurable(tmp_path):
    """题注关键字可配（post.py 与 lua filter 同步）。"""
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    md = "# 第一章\n\n照片 架构示意\n\n正文段落。\n"
    words = {"table": ["清单"], "figure": ["照片"]}

    off = _build_md(tmp_path, md, {})
    para = next(p for p in Document(off).paragraphs if "照片" in p.text)
    assert para.alignment != WD_ALIGN_PARAGRAPH.CENTER, "默认关键字下不该把「照片」当图注"

    on = _build_md(tmp_path, md, {"caption_words": words})
    para = next(p for p in Document(on).paragraphs if "照片" in p.text)
    assert para.alignment == WD_ALIGN_PARAGRAPH.CENTER, "自定义关键字未生效（应被识别为图注并居中）"


def test_config_with_bom(tmp_path):
    """带 UTF-8 BOM 的 config.json 必须能读。

    Windows 记事本 / PowerShell 写出来的 json 默认带 BOM，真实用户高频踩到。
    """
    from docx import Document

    src = tmp_path / "src"
    src.mkdir(exist_ok=True)
    (src / "01.md").write_text("# 第一章\n\n正文。\n", encoding="utf-8")
    cfg = tmp_path / "cfg.json"
    cfg.write_bytes(b"\xef\xbb\xbf" + '{"header": "带 BOM 的页眉"}'.encode())

    body, out = str(tmp_path / "body.docx"), str(tmp_path / "out.docx")
    subprocess.run(_pandoc_cmd(str(src / "01.md"), body, str(src)), check=True, capture_output=True)
    r = subprocess.run(
        [sys.executable, os.path.join(KIT, "scripts", "post.py"), body, out, str(cfg)],
        capture_output=True,
        env=dict(os.environ, PYTHONIOENCODING="utf-8"),
    )
    assert r.returncode == 0, "带 BOM 的 config 读不了：" + r.stderr.decode("utf-8", "replace")
    assert "带 BOM 的页眉" in Document(out).sections[-1].header.paragraphs[0].text


def test_to_rgb_accepts_common_forms():
    import importlib.util

    spec = importlib.util.spec_from_file_location("post", os.path.join(KIT, "scripts", "post.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    assert mod.to_rgb("404040") == (0x40, 0x40, 0x40)
    assert mod.to_rgb("#FF0000") == (0xFF, 0, 0)
    assert mod.to_rgb("0x010203") == (1, 2, 3)
    assert mod.to_rgb([1, 2, 3]) == (1, 2, 3)
    with pytest.raises(ValueError):
        mod.to_rgb("abc")


def test_render_version_flag():
    r = subprocess.run(
        [sys.executable, os.path.join(KIT, "scripts", "render.py"), "--version"],
        capture_output=True,
    )
    assert r.returncode == 0
    assert b"texere" in r.stdout, r.stdout


def test_no_h1_is_processed(tmp_path):
    """表单/附件类文档没有一级标题也要能处理：跳过目录，且不产生空白首页。"""
    from docx import Document

    src = tmp_path / "src"
    src.mkdir()
    (src / "01.md").write_text(
        "项目名称：某某项目\n\n| a | b |\n|:--|:--|\n| 1 | 2 |\n", encoding="utf-8"
    )
    body, out = str(tmp_path / "body.docx"), str(tmp_path / "out.docx")
    subprocess.run(_pandoc_cmd(str(src / "01.md"), body, str(src)), check=True, capture_output=True)
    r = subprocess.run(
        [sys.executable, os.path.join(KIT, "scripts", "post.py"), body, out, ""],
        capture_output=True,
        env=dict(os.environ, PYTHONIOENCODING="utf-8"),
    )
    assert r.returncode == 0, r.stderr.decode("utf-8", "replace")
    doc = Document(out)
    assert len(doc.sections) == 1, "无封面无标题时不该分节（跳转符会变成空白首页）"
    assert "项目名称：某某项目" in [p.text.strip() for p in doc.paragraphs], "正文被丢了"
    assert not any(p.text.strip().startswith("目") and "录" in p.text for p in doc.paragraphs), (
        "没有标题就不该插目录"
    )


def test_render_warns_on_unknown_config_key(tmp_path):
    """config 顶层未知键必须有 [warn]——外部反馈：page_number 写顶层被静默忽略。"""
    src = tmp_path / "src"
    src.mkdir()
    shutil.copy(os.path.join(KIT, "assets", "sample.md"), src / "01_sample.md")
    cfg = tmp_path / "cfg.json"
    cfg.write_text(json.dumps({"style": {}, "toplevel_typo": 1}), encoding="utf-8")
    r = subprocess.run(
        [
            sys.executable,
            os.path.join(KIT, "scripts", "render.py"),
            "--src",
            str(src),
            "--out",
            str(tmp_path / "o.docx"),
            "--config",
            str(cfg),
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    assert r.returncode == 0
    assert "[warn]" in r.stdout and "toplevel_typo" in r.stdout


def test_math_survives_postprocess(tmp_path):
    """公式守卫：行内/显示公式必须以原生 OMML 存活到 docx（外部审计实测过的边界）。"""
    import zipfile

    md = "质能方程 $E=mc^2$ 是物理学的基石。\n\n$$\n\\int_0^1 x \\, dx = \frac{1}{2}\n$$\n"
    docx = _build_md(tmp_path, md, {})
    xml = zipfile.ZipFile(docx).read("word/document.xml").decode("utf-8")
    assert xml.count("<m:oMath") >= 2, "行内/显示公式丢失——post.py 可能吞了 OMML"
    assert "oMathPara" in xml, "显示公式应成段（oMathPara）"


def test_render_creates_out_parent_dir(tmp_path):
    """--out 指向不存在的父目录时自动创建（外部审计 R2：pandoc/SaveAs 遇缺目录直接挂）。"""
    src = tmp_path / "src"
    src.mkdir()
    shutil.copy(os.path.join(KIT, "assets", "sample.md"), src / "01_sample.md")
    out = tmp_path / "nested" / "deeper" / "o.docx"
    r = subprocess.run(
        [
            sys.executable,
            os.path.join(KIT, "scripts", "render.py"),
            "--src",
            str(src),
            "--out",
            str(out),
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    assert r.returncode == 0, r.stdout + r.stderr
    assert out.exists(), "嵌套 --out 目录应被自动创建"


def _load_post():
    import importlib.util

    scripts = os.path.join(KIT, "scripts")
    if scripts not in sys.path:  # post.py 依赖兄弟模块 _shared / _ooxml
        sys.path.insert(0, scripts)
    spec = importlib.util.spec_from_file_location("post", os.path.join(KIT, "scripts", "post.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def test_keep_with_next_uses_configured_table_words(tmp_path):
    """keep_with_next 走 caption_words 的表侧 matcher，不再硬编码 (表|表格|Table)。

    历史 bug：config 配 caption_words={"table":["附表"]} 后「附表 1-1」能被
    重建的 CAPTION_RE 识别成表题，keep 判定却仍用硬编码词表——表题和表格分家。
    """
    from docx import Document

    post = _load_post()
    saved = (post.CAPTION_RE, post.TABLE_CAPTION_RE)
    try:
        post.build_caption_matchers({"table": ["附表"]})
        doc = Document()
        cap = doc.add_paragraph("附表 1-1 商务条款清单")
        doc.add_paragraph("表格后续内容")
        assert post._format_captions(doc) >= 1, "重建词表后表题未被识别"
        assert cap.paragraph_format.keep_with_next is True, "自定义表题词未贯通到 keep_with_next"
    finally:
        post.CAPTION_RE, post.TABLE_CAPTION_RE = saved


def test_keep_with_next_default_words_unchanged(tmp_path):
    """默认词表下行为与历史硬编码一致：带编号题注走正则兜底也能 keep，
    图注不 keep（历史实测 47 图文档多出 4 页的教训）。"""
    from docx import Document

    post = _load_post()
    doc = Document()
    # 无任何题注样式 → 非 strict，文本正则兜底生效；keep 由表题前缀正则判定
    t1 = doc.add_paragraph("表 1-1 capacity")
    t2 = doc.add_paragraph("Table 2-3 Items")
    f1 = doc.add_paragraph("图 1-1 架构图")
    doc.add_paragraph("结束")
    post._format_captions(doc)
    assert t1.paragraph_format.keep_with_next is True
    assert t2.paragraph_format.keep_with_next is True
    assert f1.paragraph_format.keep_with_next is False, "图注不该 keep（会把后文整块推走）"


def test_keep_with_next_styled_captions(tmp_path):
    """样式路径：styleId=TableCaption 的段落即使不带编号也 keep（历史硬编码
    前缀正则「表 商务条款响应表」的受众）；FigureCaption 样式不 keep。"""
    from docx import Document

    post = _load_post()
    doc = Document()
    # 模拟 lua 在 AST 层打的样式：styleId 才是判据，名字随意
    cap_style = doc.styles.add_style("LuaTableCaption", 1)
    cap_style.style_id = "TableCaption"
    fig_style = doc.styles.add_style("LuaFigureCaption", 1)
    fig_style.style_id = "FigureCaption"
    t3 = doc.add_paragraph("表 商务条款响应表", style=cap_style)
    f2 = doc.add_paragraph("图片", style=fig_style)
    doc.add_paragraph("结束")
    post._format_captions(doc)
    assert t3.paragraph_format.keep_with_next is True, "不带编号的样式表题应粘住表格"
    assert f2.paragraph_format.keep_with_next is False, "图注样式不该 keep"


# ---------------------------------------------------------------------------
# R4 修复：短报告模式 / H1 段前分页 / 标题块保留 / 警告门控 / 样式出口
# ---------------------------------------------------------------------------

_R4_MD = (
    "---\ntitle: 校园快递调研\nauthor: 李四\n---\n\n"
    "引言段落在第一个标题之前。\n\n"
    "# 第一节\n\n正文。\n\n# 第二节\n\n正文。\n"
)


def _h1s(doc):
    return [p for p in doc.paragraphs if p.style.name == "Heading 1"]


def test_simple_report_mode(tmp_path):
    """mode=simple-report：标题/作者居中保留、无目录、所有 H1 不另起一页。"""
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    out = _build_md(tmp_path, _R4_MD, {"mode": "simple-report"})
    doc = Document(out)
    styles = [p.style.name for p in doc.paragraphs]
    assert "Title" in styles and "Author" in styles, "标题块必须保留"
    for p in doc.paragraphs:
        if p.style.name in ("Title", "Author"):
            assert p.alignment == WD_ALIGN_PARAGRAPH.CENTER, "保留的标题/作者要居中"
    assert "TOC Heading" not in styles, "simple-report 默认不插目录"
    h1 = _h1s(doc)
    assert len(h1) == 2
    for p in h1:
        assert p.paragraph_format.page_break_before is False, "短文档不该每章一页"


def test_formal_default_first_h1_exempt_rest_inherit(tmp_path):
    """formal + toc:false：标题块保留（此前一律删除）、第一章豁免分页、其余继承模板。"""
    from docx import Document

    out = _build_md(tmp_path, _R4_MD, {"toc": False})
    doc = Document(out)
    styles = [p.style.name for p in doc.paragraphs]
    assert "Title" in styles, "无封面时标题块应保留（此前被无条件删除）"
    assert "TOC Heading" not in styles
    h1 = _h1s(doc)
    assert h1[0].paragraph_format.page_break_before is False
    assert h1[1].paragraph_format.page_break_before is None, "后续 H1 继承样式的每章一页"


def test_page_break_h1_false_disables_all(tmp_path):
    from docx import Document

    out = _build_md(tmp_path, _R4_MD, {"page_break_h1": False})
    for p in _h1s(Document(out)):
        assert p.paragraph_format.page_break_before is False


def test_cover_still_deletes_title_block(tmp_path):
    """有封面时 pandoc 的 Title/Author 段与封面行重复，维持删除行为。"""
    from docx import Document

    out = _build_md(tmp_path, _R4_MD, {"cover": [["CoverTitle", "封面标题"]]})
    doc = Document(out)
    assert not any(p.style.name in ("Title", "Author") for p in doc.paragraphs)
    assert doc.paragraphs[0].text == "封面标题", "封面必须是文档第一页"


def test_toc_lands_after_kept_title_block(tmp_path):
    """无封面 + toc：目录插在标题块之后，标题仍居文档最前。"""
    from docx import Document

    out = _build_md(tmp_path, _R4_MD, {"toc": True})
    styles = [p.style.name for p in Document(out).paragraphs]
    assert "TOC Heading" in styles
    assert styles.index("Title") < styles.index("TOC Heading")


def test_leftover_warning_gated_by_toc(tmp_path, capsys):
    """R4 #6：toc:false 时「排在目录之后」的警告必须消失（目录根本不存在）。"""
    post = _load_post()
    src = tmp_path / "src"
    src.mkdir()
    (src / "01.md").write_text("引言段落。\n\n# 第一节\n\n正文。\n", encoding="utf-8")
    body = str(tmp_path / "body.docx")
    subprocess.run(_pandoc_cmd(str(src / "01.md"), body, str(src)), check=True, capture_output=True)
    cfg_path = tmp_path / "cfg.json"
    cfg_path.write_text('{"toc": false}', encoding="utf-8")
    post.main(body, str(tmp_path / "out.docx"), str(cfg_path))
    captured = capsys.readouterr()
    assert "保留在文档开头" in captured.out
    assert "会排在目录之后" not in captured.out

    cfg_path.write_text('{"toc": true}', encoding="utf-8")
    post.main(body, str(tmp_path / "out2.docx"), str(cfg_path))
    captured = capsys.readouterr()
    assert "会排在目录之后" in captured.out, "toc 开着时假警告照旧（此时它是真的）"


def test_invalid_mode_exits(tmp_path):
    src = tmp_path / "src"
    src.mkdir()
    (src / "01.md").write_text("# 标题\n\n正文。\n", encoding="utf-8")
    body = str(tmp_path / "body.docx")
    subprocess.run(_pandoc_cmd(str(src / "01.md"), body, str(src)), check=True, capture_output=True)
    cfg_path = tmp_path / "cfg.json"
    cfg_path.write_text('{"mode": "typo"}', encoding="utf-8")
    r = subprocess.run(
        [
            sys.executable,
            os.path.join(KIT, "scripts", "post.py"),
            body,
            str(tmp_path / "out.docx"),
            str(cfg_path),
        ],
        capture_output=True,
        env=dict(os.environ, PYTHONIOENCODING="utf-8"),
    )
    assert r.returncode != 0, "非法 mode 必须立刻报错，不能静默按 formal 处理"
    # sys.exit(文案) 走 stderr
    assert "simple-report" in r.stderr.decode("utf-8", "replace")


def test_style_overrides_sizes_and_margins(tmp_path):
    """R4 #5：config.style 直接改字号/页边距，不必重建 ref.docx。"""
    from docx import Document
    from docx.shared import Pt

    cfg = {"style": {"body_size": 14, "h1_size": 18, "margin_left": 2.0}}
    out = _build_md(tmp_path, "# 标题\n\n正文段落，首行要缩进两字符。\n", cfg)
    doc = Document(out)
    assert doc.styles["Body Text"].font.size == Pt(14)
    assert doc.styles["Body Text"].paragraph_format.first_line_indent == Pt(28), (
        "首行缩进要跟字号重算（2 字符）"
    )
    assert doc.styles["Heading 1"].font.size == Pt(18)
    for sec in doc.sections:
        # sectPr 以 twips 存储转 EMU 有取整，用 cm 容差比较
        assert abs(sec.left_margin.cm - 2.0) < 0.01
        assert abs(sec.top_margin.cm - 2.54) < 0.01, "没写的键保留模板基线"


def test_margins_untouched_without_explicit_keys(tmp_path):
    """没写 margin_* 时不动 sectPr——甲方模板（reference_doc）自带页边距不被冲掉。"""
    from docx import Document

    out = _build_md(tmp_path, "# 标题\n\n正文。\n", {})
    for sec in Document(out).sections:
        assert abs(sec.left_margin.cm - 3.0) < 0.01, "应与模板基线一致"


def test_title_size_overrides_title_style_only(tmp_path):
    """0.8.0 验收残留：YAML title 落 Title 样式（模板基线 26pt），
    作业类「标题三号（16pt）」用 style.title_size 直接写；Author 不跟随。"""
    from docx import Document
    from docx.shared import Pt

    md = "---\ntitle: 课程作业标题\nauthor: 张三\n---\n\n# 第一节\n\n正文。\n"
    out = _build_md(tmp_path, md, {"style": {"title_size": 16}})
    doc = Document(out)
    assert doc.styles["Title"].font.size == Pt(16)
    title_p = next(p for p in doc.paragraphs if p.style.name == "Title")
    assert title_p.text == "课程作业标题", "标题文字必须原样保留"

    (tmp_path / "d2").mkdir()
    out2 = _build_md(tmp_path / "d2", md, {})
    doc2 = Document(out2)
    assert doc2.styles["Title"].font.size == Pt(26), "没写键时保持模板基线（make_ref 的 26pt）"
