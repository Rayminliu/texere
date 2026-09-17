# -*- coding: utf-8 -*-
"""后处理冒烟断言：只需 pandoc + python-docx，不需要 Word / PDF，秒级可跑。

  pytest -q

覆盖 README「排版六条军规」里能被机器判定的部分：
跨页重复表头、题注居中、封面与目录注入、分节页码。
"""
import os
import shutil
import subprocess
import sys

import pytest

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

pytestmark = pytest.mark.skipif(
    shutil.which("pandoc") is None, reason="需要 pandoc")


def _build(tmp_path):
    src = tmp_path / "src"
    src.mkdir()
    shutil.copy(os.path.join(KIT, "sample.md"), src / "01_sample.md")
    body = str(tmp_path / "body.docx")
    out = str(tmp_path / "out.docx")
    subprocess.run(
        ["pandoc", str(src / "01_sample.md"), "-o", body,
         "--reference-doc=" + os.path.join(KIT, "ref.docx"),
         "--resource-path=" + str(src),
         "-f", "markdown+pipe_tables+raw_html", "--wrap=none"],
        check=True, capture_output=True)
    subprocess.run(
        [sys.executable, os.path.join(KIT, "post.py"), body, out,
         os.path.join(KIT, "sample_config.json")],
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


def test_two_sections_with_page_number(docx_path):
    """封面/目录为第 1 节，正文为第 2 节，正文节页脚有页码域。"""
    from docx import Document
    from docx.oxml.ns import qn
    doc = Document(docx_path)
    assert len(doc.sections) == 2, "应为 2 节，实际 %d" % len(doc.sections)
    footer_xml = doc.sections[-1].footer.paragraphs[0]._p.xml
    assert "PAGE" in footer_xml, "正文节页脚缺少页码域"
