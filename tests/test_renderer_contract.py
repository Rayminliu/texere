"""Renderer 抽象契约测试：把「接口接好了」升级为「行为契约满足」。

不要等第三个渲染器才发现接口不完整——这里把三者必须满足的同一套外部语义写成
抽象基类 `RendererContractTests`，WordRenderer / LibreOfficeRenderer / WPSRenderer
各自继承。未来加 OnlyOfficeRenderer 也只需通过这套契约。

契约（与用户对齐的 guarantees）：
  - no input mutation（默认 save_updated_fields=False 时输入 docx 字节不变）
  - deterministic output location（res.pdf == 请求的 pdf_path）
  - non-empty output
  - renderer metadata（renderer_name 非空、renderer_version 为 str）
  - structured errors（失败时 ok=False、errors 非空、pdf=None、绝不抛栈）
  - no ambiguous artifact（失败不留半截/空 pdf）

另含集成层测试：超时结构化、并发无相互污染、Word 并发隔离（真机）。
"""

import hashlib
import importlib.util
import os
import sys
import threading
import time

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


def _make_docx(path):
    from docx import Document

    Document().save(str(path))


def _sha256(path):
    return hashlib.sha256(open(path, "rb").read()).hexdigest()


class _HangingRenderer(r.RendererAdapter):
    """测试用：故意睡死，逼出超时路径（不抛、返回结构化错误）。"""

    def render(self, docx_path, pdf_path, *, save_updated_fields=False):
        time.sleep(8)
        return r.RenderResult(ok=True, pdf=pdf_path)


class RendererContractTests:
    """所有渲染器必须继承的契约基类（pytest 按 Test* 子类收集）。"""

    renderer_factory = None  # 子类填：返回 RendererAdapter 实例
    available = staticmethod(lambda: (False, "未实现"))

    def _require_available(self):
        ok, why = type(self).available()
        if not ok:
            pytest.skip("渲染器不可用，跳过契约项：%s" % why)

    def test_does_not_modify_input(self, tmp_path):
        self._require_available()
        docx = tmp_path / "in.docx"
        _make_docx(docx)
        before = _sha256(docx)
        pdf = tmp_path / "out.pdf"
        res = type(self).renderer_factory().render(str(docx), str(pdf))
        assert res.ok, res.errors
        assert _sha256(docx) == before, "渲染器默认不应改输入 docx 字节"

    def test_produces_nonempty_pdf(self, tmp_path):
        self._require_available()
        docx = tmp_path / "in.docx"
        _make_docx(docx)
        pdf = tmp_path / "out.pdf"
        res = type(self).renderer_factory().render(str(docx), str(pdf))
        assert res.ok, res.errors
        assert os.path.exists(pdf)
        assert os.path.getsize(pdf) > 0

    def test_reports_version(self, tmp_path):
        self._require_available()
        docx = tmp_path / "in.docx"
        _make_docx(docx)
        pdf = tmp_path / "out.pdf"
        res = type(self).renderer_factory().render(str(docx), str(pdf))
        assert res.ok, res.errors
        assert res.renderer_name, "renderer_name 必须非空"
        assert isinstance(res.renderer_version, str), "renderer_version 必须是 str"

    def test_respects_output_path(self, tmp_path):
        self._require_available()
        docx = tmp_path / "in.docx"
        _make_docx(docx)
        pdf = tmp_path / "out.pdf"
        res = type(self).renderer_factory().render(str(docx), str(pdf))
        assert res.ok, res.errors
        assert res.pdf == str(pdf), "res.pdf 必须等于请求的 pdf_path"

    def test_failure_is_structured(self, tmp_path):
        # 不存在的输入：必须 ok=False、errors 非空、pdf=None，且绝不抛栈
        pdf = tmp_path / "x.pdf"
        res = type(self).renderer_factory().render("nope.docx", str(pdf))
        assert res.ok is False
        assert res.errors, "失败必须给出结构化 errors"
        assert res.pdf is None

    def test_no_ambiguous_artifact_on_failure(self, tmp_path):
        pdf = tmp_path / "x.pdf"
        res = type(self).renderer_factory().render("nope.docx", str(pdf))
        assert res.ok is False
        # 失败不允许留下「存在但为空」的半成品 pdf
        if os.path.exists(pdf):
            assert os.path.getsize(pdf) > 0, "失败留下了空 pdf，状态含混"


class TestWordRendererContract(RendererContractTests):
    renderer_factory = r.WordRenderer
    available = staticmethod(r.WordRenderer.available)


class TestLibreOfficeRendererContract(RendererContractTests):
    renderer_factory = r.LibreOfficeRenderer
    available = staticmethod(r.LibreOfficeRenderer.available)


class TestWPSRendererContract(RendererContractTests):
    renderer_factory = r.WPSRenderer
    available = staticmethod(r.WPSRenderer.available)


# ----------------------------------------------------------- 集成层：超时 / 并发


def test_export_pdf_once_timeout_is_structured(tmp_path):
    # 线程级超时必须返回结构化 (False, 含「超时」)，不抛、不留半成品
    docx = tmp_path / "in.docx"
    docx.write_bytes(b"x")
    pdf = tmp_path / "out.pdf"
    ok, err, _ = v.export_pdf_once(str(docx), str(pdf), timeout=1, renderer=_HangingRenderer())
    assert ok is False
    assert "超时" in err
    assert not (os.path.exists(pdf) and os.path.getsize(pdf) == 0)


def test_export_pdf_once_concurrent_no_pollution(tmp_path):
    # 并发跑多个导出，每个输出必须来自自己的 source，不被别的线程污染
    n = 4
    srcs, docs, pdfs = [], [], []
    for i in range(n):
        s = tmp_path / ("src%d.pdf" % i)
        s.write_bytes(("pdf-content-%d" % i).encode())
        d = tmp_path / ("in%d.docx" % i)
        d.write_bytes(b"x")
        p = tmp_path / ("out%d.pdf" % i)
        srcs.append(str(s))
        docs.append(str(d))
        pdfs.append(str(p))
    results = {}

    def work(i):
        results[i] = v.export_pdf_once(docs[i], pdfs[i], renderer=r.FakeRenderer(srcs[i]))

    threads = [threading.Thread(target=work, args=(i,)) for i in range(n)]
    for t in threads:
        t.start()
    for t in threads:
        t.join(timeout=30)
    for i in range(n):
        ok, err, _ = results[i]
        assert ok, err
        assert open(pdfs[i], "rb").read() == open(srcs[i], "rb").read()


def test_word_concurrent_renders_isolated(tmp_path):
    # 真机并发 canary：同时起 2 个 Word 实例渲染不同 docx，验证无共享状态/输出污染。
    # 默认跳过：Word COM 跨线程 teardown 已知不稳定（见 SCRIPT_HELP 边界），会偶发
    # 解释器退出期访问违规回溯（pytest 退出码仍为 0，但 stderr 有噪音）。
    # 想探测 COM 稳定性时显式开启：TEXERE_COM_CONCURRENCY=1 pytest ...
    if not os.environ.get("TEXERE_COM_CONCURRENCY"):
        pytest.skip("设 TEXERE_COM_CONCURRENCY=1 才跑 Word 并发隔离 canary（COM teardown 噪声）")
    ok_w, _ = r.WordRenderer.available()
    if not ok_w:
        pytest.skip("本环境无 Word，跳过并发隔离实测")
    n = 2
    docs = [str(tmp_path / ("in%d.docx" % i)) for i in range(n)]
    pdfs = [str(tmp_path / ("out%d.pdf" % i)) for i in range(n)]
    for d in docs:
        _make_docx(d)
    results = {}

    def work(i):
        results[i] = r.WordRenderer().render(docs[i], pdfs[i])

    threads = [threading.Thread(target=work, args=(i,)) for i in range(n)]
    for t in threads:
        t.start()
    for t in threads:
        t.join(timeout=120)
    for i in range(n):
        res = results[i]
        assert res.ok, res.errors
        assert os.path.getsize(pdfs[i]) > 0
        assert res.pdf == pdfs[i]
