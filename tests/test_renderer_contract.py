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

import ast
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


def _ast_param_default(rel_path, class_name, func_name, param):
    """ast 读某个函数参数的默认值字面量；参数不存在或默认值非字面量时返回 None。

    位置参数与关键字-only 参数都要看：LibreOfficeRenderer.render 的 timeout 就在
    `*` 之后（契约里的 save_updated_fields 同样是 kw-only）。
    不 import：这类守卫要在没装齐依赖（docx / pymupdf / soffice）的 CI 上也能跑。
    """
    with open(os.path.join(KIT, rel_path), encoding="utf-8") as f:
        tree = ast.parse(f.read())
    body = tree.body
    if class_name:
        matched = [n for n in tree.body if isinstance(n, ast.ClassDef) and n.name == class_name]
        assert matched, "找不到类 %s" % class_name
        body = matched[0].body
    fn = [n for n in body if isinstance(n, ast.FunctionDef) and n.name == func_name]
    assert fn, "找不到函数 %s" % func_name
    a = fn[0].args
    pairs = list(zip(a.args[len(a.args) - len(a.defaults) :], a.defaults))
    pairs += [(arg, d) for arg, d in zip(a.kwonlyargs, a.kw_defaults) if d is not None]
    found = [d for arg, d in pairs if arg.arg == param]
    if not found:
        return None
    value = ast.literal_eval(found[0])
    return value if isinstance(value, int) else None


def _ast_module_literal(rel_path, var):
    """ast 读模块级字面量常量。"""
    with open(os.path.join(KIT, rel_path), encoding="utf-8") as f:
        tree = ast.parse(f.read())
    for node in tree.body:
        if isinstance(node, ast.Assign) and any(
            isinstance(t, ast.Name) and t.id == var for t in node.targets
        ):
            return ast.literal_eval(node.value)
    raise AssertionError("找不到模块级常量 %s" % var)


def test_outer_watchdog_exceeds_inner_engine_timeout():
    """超时层级不许倒挂：validate 的外层预算必须大于渲染器内层 watchdog。

    倒挂时的真实形态：外层写死 300 < LibreOffice 内层 600，外层先弃等并报
    「PDF 导出超时 (300s)」，soffice 却接着跑到 600s 才被强杀 —— 报的时限是假的，
    内层强杀逻辑形同虚设。这里同时守住：内层放大值不许被缩小、外层不许
    再钉死一个具体秒数、COM 路径（无内层超时）必须有兜底预算而不是无限等。
    """
    inner = _ast_param_default("scripts/renderers.py", "LibreOfficeRenderer", "render", "timeout")
    assert isinstance(inner, int) and inner > 0, "LibreOfficeRenderer.render 必须保留内层强杀超时"
    # 600 不是随手写的：它是当前唯一有实证必要的放大值（272 页 / 252 表 / 112 图
    # 端到端 66.9–69.7s，300s 已有 4.3× 余量），把它改小会误杀大文档的合法导出
    assert inner >= 600, "LO 内层强杀超时不得被缩小（当前口径 600s）"

    outer_default = _ast_param_default("scripts/validate.py", None, "export_pdf_once", "timeout")
    assert outer_default is None, "外层预算必须按渲染器算（timeout=None），不许再写死秒数"

    margin = _ast_module_literal("scripts/validate.py", "_OUTER_MARGIN")
    fallback = _ast_module_literal("scripts/validate.py", "_OUTER_FALLBACK")
    assert margin > 0
    assert v._outer_budget(r.LibreOfficeRenderer()) == inner + margin > inner

    # Word / WPS 的 render 没有超时形参（COM 阻塞不可控），只能靠外层兜底
    assert _ast_param_default("scripts/renderers.py", "_ComRenderer", "render", "timeout") is None
    assert v._outer_budget(r.WordRenderer()) == fallback
    assert v._outer_budget(r.WPSRenderer()) == fallback


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


def test_com_renderers_share_one_skeleton():
    """Word / WPS 的 COM 链必须共用 _ComRenderer 骨架，不许回到两份平行副本。

    平行副本的实际危害是改一边忘另一边（僵尸进程清理、陈旧 PDF 判据、超时看门狗
    都曾只修在 Word 那一侧）。这里锁结构：render/available 只有骨架一份，
    差异全部落在声明式的 ProgID 与钩子上。
    """
    base = r._ComRenderer

    assert issubclass(r.WordRenderer, base) and issubclass(r.WPSRenderer, base)
    assert r.WordRenderer.render is base.render
    assert r.WPSRenderer.render is base.render
    assert r.WordRenderer.available.__func__ is base.available.__func__
    assert r.WPSRenderer.available.__func__ is base.available.__func__

    # 差异只剩声明
    assert r.WordRenderer.PROGID == "Word.Application"
    assert r.WPSRenderer.PROGID == "KWPS.Application"
    assert r.WordRenderer.TMP_PREFIX != r.WPSRenderer.TMP_PREFIX

    # Word 独有的步骤不许悄悄跑到 WPS 上（WPS 没有 TOC 域 API / DisplayAlerts）
    assert r.WordRenderer.prepare_document is not base.prepare_document
    assert r.WPSRenderer.prepare_document is base.prepare_document
    assert r.WordRenderer.collect_stats is not base.collect_stats
    assert r.WPSRenderer.collect_stats is base.collect_stats
    assert r.WordRenderer.configure_app is not base.configure_app
    assert r.WPSRenderer.configure_app is base.configure_app


# ---------------------------------------------------------- finally 不能赢过结果


def test_quit_failure_does_not_invert_success(monkeypatch, tmp_path):
    """T1.2 回归：app.Quit() 招 RPC 异常时，已成功导出的 PDF 不能被报成失败。

    Python 语义下 finally 里的异常会丢弃即将 return 的结果，被外层 except 吞成
    ok=False。本机实测过 Word COM 的 Windows fatal exception 0x800706be 正好走
    这条路径——验收器不能把「不确定成功」当成「确定失败」。
    """
    import types

    class FakeDoc:
        def Close(self, *args):  # noqa: N802  COM 接口名，逐字对齐 Word API
            pass

        def ExportAsFixedFormat(self, path, fmt):  # noqa: N802  同上
            with open(path, "wb") as f:
                f.write(b"%PDF-1.4 fake export")

    class FakeDocuments:
        def Open(self, *args):  # noqa: N802  同上
            return FakeDoc()

    class FakeApp:
        Version = "0.0-fake"
        Path = str(tmp_path)
        Documents = FakeDocuments()

        def Quit(self):  # noqa: N802  COM 接口名
            raise RuntimeError("(-2147023170, '远程过程调用失败')")

    app = FakeApp()
    pythoncom = types.ModuleType("pythoncom")
    pythoncom.CoInitialize = lambda: None
    pythoncom.CoUninitialize = lambda: None
    client = types.ModuleType("win32com.client")
    client.DispatchEx = lambda progid: app
    win32com = types.ModuleType("win32com")
    win32com.client = client
    monkeypatch.setitem(sys.modules, "pythoncom", pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", client)

    class ProbeRenderer(r._ComRenderer):
        PROGID = "Probe.Application"
        LABEL = "probe"
        ENGINE = "Probe"
        TMP_PREFIX = "probe-test-"

        def configure_app(self, app):
            pass

        def prepare_document(self, doc):
            pass

        def collect_stats(self, app, doc):
            return {"pages": 1}

    docx = tmp_path / "in.docx"
    docx.write_bytes(b"docx-bytes")
    pdf = tmp_path / "out.pdf"

    res = ProbeRenderer().render(str(docx), str(pdf))

    assert res.ok is True, "导出成功不得因 Quit 失败而被误判：%s" % res.errors
    assert res.pdf == str(pdf)
    assert os.path.exists(pdf)
    # Quit 异常不能默默吞掉——要留在 warnings 里
    assert any("Quit" in w for w in res.warnings), res.warnings


def test_quit_failure_does_not_mask_a_real_error(monkeypatch, tmp_path):
    """真失败时仍得是 ok=False，且 errors 里是原始原因，不是 Quit 的噪声。"""
    import types

    class FakeDocuments:
        def Open(self, *args):  # noqa: N802  COM 接口名
            raise ValueError("文档打开失败")

    class FakeApp:
        Version = "0.0-fake"
        Path = str(tmp_path)
        Documents = FakeDocuments()

        def Quit(self):  # noqa: N802  COM 接口名
            raise RuntimeError("(-2147023170, '远程过程调用失败')")

    app = FakeApp()
    pythoncom = types.ModuleType("pythoncom")
    pythoncom.CoInitialize = lambda: None
    pythoncom.CoUninitialize = lambda: None
    client = types.ModuleType("win32com.client")
    client.DispatchEx = lambda progid: app
    win32com = types.ModuleType("win32com")
    win32com.client = client
    monkeypatch.setitem(sys.modules, "pythoncom", pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", client)

    class ProbeRenderer(r._ComRenderer):
        PROGID = "Probe.Application"
        LABEL = "probe"
        ENGINE = "Probe"
        TMP_PREFIX = "probe-test-"

        def configure_app(self, app):
            pass

    docx = tmp_path / "in.docx"
    docx.write_bytes(b"docx-bytes")
    res = ProbeRenderer().render(str(docx), str(tmp_path / "out.pdf"))

    assert res.ok is False
    assert any("文档打开失败" in e for e in res.errors), res.errors
    assert not any("Quit" in e for e in res.errors), "不该拿 Quit 异常掩盖真因"
