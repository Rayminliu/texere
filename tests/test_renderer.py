"""Renderer 抽象护栏：证明渲染器可插拔，且 PDF 派生检查与具体渲染器无关。

不启动 Word（除非显式需要）：
- 验证 RendererAdapter 抽象、WordRenderer / FakeRenderer 满足契约；
- 用 PyMuPDF 合成一个 PDF，证明 page_numbering / blank_pages / visual_drift
  只消费 pdf_path、不感知渲染器——换任何 renderer 出同一个 PDF，结论不变。
"""

import importlib.util
import os
import sys

import pytest

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(KIT, "scripts")
sys.path.insert(0, SCRIPTS)

_rspec = importlib.util.spec_from_file_location("renderers", os.path.join(SCRIPTS, "renderers.py"))
r = importlib.util.module_from_spec(_rspec)
_rspec.loader.exec_module(r)

_vspec = importlib.util.spec_from_file_location("validate", os.path.join(SCRIPTS, "validate.py"))
v = importlib.util.module_from_spec(_vspec)
_vspec.loader.exec_module(v)


def test_renderer_adapter_is_abstract():
    # 含 @abstractmethod 的基类不可直接实例化
    with pytest.raises(TypeError):
        r.RendererAdapter()


def test_word_renderer_is_adapter():
    assert isinstance(r.WordRenderer(), r.RendererAdapter)
    assert callable(getattr(r.WordRenderer, "render"))


def test_fake_renderer_satisfies_contract(tmp_path):
    src = tmp_path / "src.pdf"
    src.write_bytes(b"%PDF-1.4 fake")
    out = tmp_path / "out.pdf"
    res = r.FakeRenderer(str(src)).render(str(tmp_path / "in.docx"), str(out))
    assert isinstance(res, r.RenderResult)
    assert res.ok is True
    assert res.pdf == str(out)
    assert os.path.exists(out)


def _make_synthetic_pdf(path):
    pymupdf = pytest.importorskip("pymupdf", reason="合成 PDF 需要 PyMuPDF")
    doc = pymupdf.open()
    for i in range(1, 4):  # 3 页，页脚带连续页码 1/2/3
        page = doc.new_page()
        page.insert_text((50, 50), f"body text page {i}")
        page.insert_text((50, 800), str(i))  # 页脚页码
    doc.save(str(path))
    doc.close()


def test_pdf_derived_checks_are_renderer_agnostic(tmp_path):
    # 关键护栏：这些检查只吃 pdf_path，从不引用渲染器；
    # 因此无论 Word / WPS / LO / Fake 谁出的同一个 PDF，结论都不变。
    pdf = tmp_path / "doc.pdf"
    _make_synthetic_pdf(pdf)
    rn = v.check_page_numbering(str(pdf))
    assert rn.status == v.PASS  # 连续页码 1-3
    bp = v.check_blank_pages(str(pdf))
    assert bp.status == v.PASS  # 无空白页
    vd = v.check_visual_drift(str(pdf))  # 无基线 → SKIP，不依赖渲染器
    assert vd.status == v.SKIP


# --------------------------------------------------- 渲染器工厂 & 可插拔接缝


def test_get_renderer_default_is_word():
    assert isinstance(r.get_renderer(), r.WordRenderer)
    assert isinstance(r.get_renderer("word"), r.WordRenderer)


def test_get_renderer_returns_each_type():
    assert isinstance(r.get_renderer("libreoffice"), r.LibreOfficeRenderer)
    assert isinstance(r.get_renderer("wps"), r.WPSRenderer)


def test_get_renderer_unknown_raises():
    with pytest.raises(ValueError):
        r.get_renderer("gimp")


def test_export_pdf_once_uses_injected_renderer_not_word(tmp_path):
    # slice2 的硬护栏：validate.export_pdf_once 必须能接受任意 RendererAdapter，
    # 不回溯去 shell finalize.py / 起 Word。注入 FakeRenderer 即可在无 Word 的
    # CI 上跑通 PDF 派生检查链路。
    src = tmp_path / "src.pdf"
    src.write_bytes(b"%PDF-1.4 fake")
    docx = tmp_path / "in.docx"
    docx.write_bytes(b"docx")  # 渲染器只看路径，不解析内容
    pdf = tmp_path / "verify.pdf"
    ok, err, _ = v.export_pdf_once(str(docx), str(pdf), renderer=r.FakeRenderer(str(src)))
    assert ok is True, err
    assert pdf.exists()
    assert err == ""


def test_word_renderer_available_is_bool_tuple():
    ok, why = r.WordRenderer.available()
    assert isinstance(ok, bool)
    assert isinstance(why, str)


def test_libreoffice_renderer_skips_when_missing():
    ok, why = r.LibreOfficeRenderer.available()
    if not ok:
        pytest.skip("本环境未装 LibreOffice（soffice）：%s" % why)
    # 装了才跑到这里：探测必须给出可执行路径
    assert "soffice" in why or "libreoffice" in why


def test_wps_renderer_skips_when_missing():
    ok, why = r.WPSRenderer.available()
    if not ok:
        pytest.skip("本环境未装 WPS Office（KWPS.Application）：%s" % why)
    assert "WPS" in why


def test_renderer_identity_recorded_in_evidence(tmp_path):
    # 评审优先级 #2：renderer 身份（name/version/engine_path）必须进 report 的 metadata，
    # 否则 evidence 缺 provenance——别人看到 PASS 也不知道是 Word / WPS / LO 出的。
    # 用 FakeRenderer 即可在无 Word 的 CI 上守住这条不变量。
    import docx as docxmod

    src_pdf = tmp_path / "src.pdf"
    _make_synthetic_pdf(
        src_pdf
    )  # 合法 PDF：FakeRenderer 复制后，generate_evidence_package 的截图步才能打开
    docx_path = tmp_path / "in.docx"
    docxmod.Document().save(str(docx_path))
    rndr = r.FakeRenderer(str(src_pdf))
    report = v.generate_evidence_package(
        str(docx_path),
        str(tmp_path / "evidence"),
        {},
        None,
        0,
        None,
        None,
        False,
        False,
        renderer=rndr,
    )
    meta = report["metadata"]["renderer"]
    assert meta["name"] == "Fake"  # FakeRenderer 自报 "Fake"
    assert "version" in meta and "engine_path" in meta
