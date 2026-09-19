"""scripts/distill.py 的模板蒸馏，以及复用模板时能继承什么。

用 python-docx 现场造模板；渲染那两条需要 pandoc，不需要 Word。
守的是两条：
  1. 蒸馏要报出页面设置 / 字体 / 页眉页脚，并给出能直接用的 config；
  2. "page_number": null 必须留住模板自带页脚，一个字都不许追加。
"""

import json
import os
import shutil
import subprocess
import sys

import pytest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DISTILL = os.path.join(KIT, "scripts", "distill.py")
RENDER = os.path.join(KIT, "scripts", "render.py")


def run(script, *args, expect=0):
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    r = subprocess.run([sys.executable, script, *args], capture_output=True, env=env)
    out = (r.stdout + r.stderr).decode("utf-8", "replace")
    if expect is not None:
        assert r.returncode == expect, out
    return out


def set_ea(style, name, size):
    """设样式字体，并把 eastAsia 也写上（中文字体靠这个属性）。"""
    style.font.name = name
    style.font.size = Pt(size)
    rPr = style.element.get_or_add_rPr()
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rPr.append(rf)
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia"):
        rf.set(qn(a), name)


def make_tpl(path, header="甲方专属页眉", footer="文件编号 XYZ-2026"):
    d = Document()
    for s in d.sections:
        s.left_margin = Cm(3.0)
        s.right_margin = Cm(3.0)
        s.top_margin = Cm(2.5)
        s.bottom_margin = Cm(2.0)
    set_ea(d.styles["Normal"], "楷体", 14)
    set_ea(d.styles["Heading 1"], "黑体", 18)
    if header:
        d.sections[0].header.paragraphs[0].text = header
    if footer:
        d.sections[0].footer.paragraphs[0].text = footer
    d.save(str(path))
    return str(path)


@pytest.fixture
def tpl(tmp_path):
    return make_tpl(tmp_path / "tpl.docx")


# ---------------------------------------------------------------- 蒸馏报告


def test_reports_page_setup(tpl):
    out = run(DISTILL, tpl)
    assert "3.0" in out and "2.5" in out, "页边距必须报出来"
    assert "分节数" in out


def test_reports_fonts(tpl):
    out = run(DISTILL, tpl)
    assert "楷体" in out and "14.0pt" in out, "正文字体字号"
    assert "黑体" in out and "18.0pt" in out, "一级标题字体字号"


def test_reports_header_and_footer(tpl):
    out = run(DISTILL, tpl)
    assert "甲方专属页眉" in out
    assert "文件编号 XYZ-2026" in out


def test_suggests_make_ref_command(tpl):
    out = run(DISTILL, tpl)
    assert "make_ref.py --body-font 楷体" in out
    assert "--body-size 14" in out


def test_missing_file_exits(tmp_path):
    run(DISTILL, str(tmp_path / "nope.docx"), expect=1)


# ---------------------------------------------------------------- 建议 config


def test_out_writes_usable_config(tpl, tmp_path):
    dst = tmp_path / "cfg.json"
    run(DISTILL, tpl, "--out", str(dst))
    cfg = json.loads(dst.read_text(encoding="utf-8"))
    assert cfg["reference_doc"] == "tpl.docx"
    assert cfg["header"] == "甲方专属页眉"
    # 模板页脚有内容 -> 建议关掉页码，避免被追加一截
    assert cfg["style"]["page_number"] is None


def test_no_footer_keeps_page_number_default(tmp_path):
    t = make_tpl(tmp_path / "nofooter.docx", footer=None)
    dst = tmp_path / "cfg.json"
    run(DISTILL, t, "--out", str(dst))
    cfg = json.loads(dst.read_text(encoding="utf-8"))
    # 模板没有页脚，就不必动 page_number
    assert "style" not in cfg or "page_number" not in cfg.get("style", {})


def test_out_config_is_valid_json(tpl, tmp_path):
    dst = tmp_path / "cfg.json"
    run(DISTILL, tpl, "--out", str(dst))
    json.loads(dst.read_text(encoding="utf-8"))  # 解析不了会抛


# ---------------------------------------------------------------- 复用模板（集成）


def _render(tmp_path, tpl, style):
    src = tmp_path / "src"
    src.mkdir(exist_ok=True)
    shutil.copy(os.path.join(KIT, "assets", "sample.md"), src / "01.md")
    cfg_path = tmp_path / "cfg.json"
    cfg_path.write_text(
        json.dumps({"reference_doc": tpl, "style": style}, ensure_ascii=False),
        encoding="utf-8",
    )
    out = tmp_path / "out.docx"
    run(RENDER, "--src", str(src), "--out", str(out), "--config", str(cfg_path))
    return Document(str(out))


def test_reference_doc_inherits_margins_and_fonts(tmp_path, tpl):
    doc = _render(tmp_path, tpl, {})
    sec = doc.sections[-1]
    assert round(sec.left_margin.cm, 1) == 3.0
    assert round(sec.top_margin.cm, 1) == 2.5
    assert doc.styles["Normal"].font.name == "楷体"


def test_header_is_kept_when_not_configured(tmp_path, tpl):
    """config 不写 header 时，模板自带页眉必须原样留下。"""
    doc = _render(tmp_path, tpl, {})
    assert "甲方专属页眉" in doc.sections[-1].header.paragraphs[0].text


def test_default_appends_page_number_to_template_footer(tmp_path, tpl):
    doc = _render(tmp_path, tpl, {})
    text = doc.sections[-1].footer.paragraphs[0].text
    assert "文件编号 XYZ-2026" in text
    assert "1" in text, "默认会在模板页脚后面追加页码"


def test_page_number_null_keeps_template_footer_intact(tmp_path, tpl):
    """ "page_number": null —— 一个字都不许追加。"""
    doc = _render(tmp_path, tpl, {"page_number": None})
    text = doc.sections[-1].footer.paragraphs[0].text
    assert text == "文件编号 XYZ-2026", "模板页脚被改写了：%r" % text


if __name__ == "__main__":
    sys.exit(pytest.main([__file__, "-q"]))
