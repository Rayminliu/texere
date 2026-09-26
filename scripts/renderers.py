"""Renderer 抽象：把「docx → PDF」从核心里抽出来。

Texere 的核心（compile / OOXML / source / image / profile / metadata / evidence）
不依赖任何 Office。只有「真机验收 + 出 PDF」这一步需要具体渲染器，这一层就是
RendererAdapter：

    WordRenderer             —— 本机 Microsoft Word（COM）        【当前默认】
    LibreOfficeRenderer      —— 本机 LibreOffice headless（soffice，CI / Linux 友好）
    WPSRenderer              —— 本机 WPS Writer（COM: KWPS.Application）

核心只认统一能力：

    render(docx) -> RenderResult

一旦这层成立，Texere 就从「Word automation tool」变成「Document acceptance
engine」，Word / WPS / LibreOffice 只是不同的「事实渲染器」。

设计约束（不变式）：
- PDF 派生检查（page_numbering / blank_pages / visual_drift）只读 RenderResult.pdf，
  绝不感知具体渲染器——这是 renderer-agnostic 的硬护栏，见 tests/test_renderer.py。
- 渲染器只负责 docx → pdf（+ 身份 / 告警），page_images 等下游处理由 PyMuPDF
  完成，不进 renderer 契约。
- 默认只读验收：在临时副本上刷新域 / 重排 / 导 PDF，绝不写回输入 docx；
  只有 save_updated_fields=True 才写回（render --pdf 的交付物需要）。

注意：本模块只依赖标准库；pywin32 等 Office 依赖在具体渲染器的方法内惰性导入，
因此没有安装 Word 的机器也能 import 本模块（只是 WordRenderer.render 会返回
ok=False 的结构化错误）。
"""

import abc
import os
import shutil
import subprocess
import tempfile
from dataclasses import dataclass, field


@dataclass
class RenderResult:
    """一次渲染的产物。核心 / 检查只消费这里的东西，不直接碰 Office。"""

    ok: bool
    pdf: str | None = None
    renderer_name: str = ""
    renderer_version: str = ""
    engine_path: str = ""
    warnings: list = field(default_factory=list)
    errors: list = field(default_factory=list)
    stats: dict = field(default_factory=dict)


class RendererAdapter(abc.ABC):
    """docx → PDF 的统一契约。具体渲染器（Word / LO / WPS / test）实现 render()。"""

    @abc.abstractmethod
    def render(
        self,
        docx_path: str,
        pdf_path: str,
        *,
        save_updated_fields: bool = False,
    ) -> RenderResult:
        """把 docx_path 渲染 / 导出为 pdf_path。

        save_updated_fields=True 时允许把刷新后的域写回原 docx（交付物需要）；
        默认 False 表示只读验收，原 docx 字节不可变。
        """
        raise NotImplementedError


class WordRenderer(RendererAdapter):
    """本机 Microsoft Word（COM）渲染器——即原 finalize.py 的 Word / COM 逻辑。"""

    @classmethod
    def available(cls):
        try:
            import pythoncom  # noqa: F401
            import win32com.client as win32  # noqa: F401
        except ImportError:
            return False, "未安装 pywin32（Word 渲染器需要它）"
        try:
            pythoncom.CoInitialize()
            try:
                w = win32.DispatchEx("Word.Application")
                name = "%s %s" % (w.Name, w.Version)
                w.Quit()
            finally:
                pythoncom.CoUninitialize()
        except Exception as e:
            return False, "Word 启动失败（%s：%s）" % (type(e).__name__, e)
        return True, name

    def render(
        self,
        docx_path: str,
        pdf_path: str,
        *,
        save_updated_fields: bool = False,
    ) -> RenderResult:
        WD_PAGES, WD_WORDS = 2, 0
        WD_PDF = 17
        co_initialized = False
        tmp_dir = tempfile.mkdtemp(prefix="texere_render_")
        tmp_src = os.path.join(tmp_dir, os.path.basename(docx_path))
        try:
            # 惰性 import 留在 try 内：pywin32 缺失也要走结构化失败，
            # 而不是裸抛 ModuleNotFoundError（契约：失败必须结构化，绝不抛栈）
            import pythoncom
            import win32com.client as win32

            shutil.copy2(docx_path, tmp_src)
            pythoncom.CoInitialize()
            co_initialized = True
            word = win32.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            try:
                doc = word.Documents.Open(os.path.abspath(tmp_src), False, False, False)
                for i in range(1, doc.TablesOfContents.Count + 1):
                    doc.TablesOfContents(i).Update()
                doc.Fields.Update()
                doc.Repaginate()
                stats = {
                    "pages": doc.ComputeStatistics(WD_PAGES),
                    "words": doc.ComputeStatistics(WD_WORDS),
                    "tables": doc.Tables.Count,
                    "inline_shapes": doc.InlineShapes.Count,
                    "sections": doc.Sections.Count,
                }
                doc.ExportAsFixedFormat(os.path.abspath(pdf_path), WD_PDF)
                ok = os.path.exists(pdf_path)
                if save_updated_fields:
                    doc.SaveAs(os.path.abspath(docx_path))
                doc.Close(0)
                return RenderResult(
                    ok=ok,
                    pdf=pdf_path if ok else None,
                    renderer_name=word.Name,
                    renderer_version=word.Version,
                    engine_path=word.Path,
                    warnings=[],
                    errors=[],
                    stats=stats,
                )
            finally:
                word.Quit()
        except Exception as e:  # 捕获一切 Office 异常，转成结构化 Result（不抛栈）
            return RenderResult(
                ok=False,
                pdf=None,
                renderer_name="Microsoft Word",
                renderer_version="",
                engine_path="",
                warnings=[],
                errors=[f"{type(e).__name__}: {e}"],
                stats={},
            )
        finally:
            if co_initialized:
                pythoncom.CoUninitialize()
            shutil.rmtree(tmp_dir, ignore_errors=True)


class FakeRenderer(RendererAdapter):
    """测试 / 演示用渲染器：不调用任何 Office，仅把 source_pdf 复制到 pdf_path。

    用途 ONLY：证明 RendererAdapter 接口可被「无 Office」的实现满足
    （renderer-agnostic 护栏），以及在没有 Word 的 CI 上跑通管线。
    它不是真实保真 oracle，不可用于实际文档验收。
    """

    def __init__(self, source_pdf: str):
        self.source_pdf = source_pdf

    def render(
        self,
        docx_path: str,
        pdf_path: str,
        *,
        save_updated_fields: bool = False,
    ) -> RenderResult:
        if not os.path.exists(self.source_pdf):
            return RenderResult(
                ok=False,
                pdf=None,
                renderer_name="Fake",
                renderer_version="0",
                engine_path="",
                warnings=[],
                errors=[f"source_pdf 不存在: {self.source_pdf}"],
                stats={},
            )
        shutil.copy2(self.source_pdf, pdf_path)
        return RenderResult(
            ok=True,
            pdf=pdf_path,
            renderer_name="Fake",
            renderer_version="0",
            engine_path="",
            warnings=[],
            errors=[],
            stats={},
        )


def _soffice_version(exe: str) -> str:
    """尽力取 LibreOffice 版本（仅供 renderer 元数据；失败返回空串）。"""
    try:
        r = subprocess.run(
            [exe, "--version"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=30,
        )
        lines = (r.stdout or r.stderr or "").splitlines()
        return lines[0].strip() if lines else ""
    except Exception:
        return ""


def _kill_process_tree(pid):
    """跨平台强杀进程及其子进程（LibreOffice / WPS 会拉起多个子进程）。

    Windows 用 taskkill /T 杀整棵进程树；POSIX 用 killpg 杀独立会话。
    捕获一切异常并兜底单进程 SIGKILL——目的是「绝不留下 zombie」。
    """
    if pid is None:
        return
    try:
        if os.name == "nt":
            subprocess.run(
                ["taskkill", "/F", "/T", "/PID", str(pid)],
                capture_output=True,
                timeout=15,
            )
        else:
            import signal

            os.killpg(os.getsid(pid), signal.SIGKILL)
    except Exception:
        try:
            import signal

            os.kill(pid, signal.SIGKILL)
        except Exception:
            pass


class LibreOfficeRenderer(RendererAdapter):
    """LibreOffice headless（soffice）渲染器——跨平台，CI / Linux 友好。

    用 `soffice --headless --convert-to pdf` 把 docx 导出到 outdir，再搬到目标
    pdf_path。不依赖 Windows / COM，因此可在无 Word 的容器里跑真机验收。
    """

    @classmethod
    def available(cls):
        exe = shutil.which("soffice") or shutil.which("libreoffice")
        if exe:
            return True, exe
        return False, "未找到 soffice / libreoffice（CI 需装 libreoffice-headless）"

    def render(
        self,
        docx_path: str,
        pdf_path: str,
        *,
        save_updated_fields: bool = False,
        timeout: int = 600,
    ) -> RenderResult:
        exe = shutil.which("soffice") or shutil.which("libreoffice")
        if not exe:
            return RenderResult(
                ok=False,
                pdf=None,
                renderer_name="LibreOffice",
                renderer_version="",
                engine_path="",
                warnings=[],
                errors=["LibreOffice 不可用：未找到 soffice / libreoffice"],
                stats={},
            )
        out_dir = os.path.dirname(os.path.abspath(pdf_path)) or "."
        src = os.path.abspath(docx_path)
        proc = None
        try:
            proc = subprocess.Popen(
                [exe, "--headless", "--convert-to", "pdf", "--outdir", out_dir, src],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                start_new_session=True,
            )
            try:
                out, err = proc.communicate(timeout=timeout)
            except subprocess.TimeoutExpired:
                # 超时必须强杀整棵进程树，否则 soffice 卡死会留下 zombie（见 SCRIPT_HELP 边界）
                _kill_process_tree(proc.pid)
                try:
                    proc.wait(timeout=10)
                except Exception:
                    pass
                return RenderResult(
                    ok=False,
                    pdf=None,
                    renderer_name="LibreOffice",
                    engine_path=exe,
                    warnings=[],
                    errors=[f"LibreOffice 导出超时 ({timeout}s)，已强杀进程树"],
                    stats={},
                )
            version = _soffice_version(exe)
            produced = os.path.join(out_dir, os.path.splitext(os.path.basename(src))[0] + ".pdf")
            if os.path.exists(produced):
                if os.path.abspath(produced) != os.path.abspath(pdf_path):
                    shutil.move(produced, pdf_path)
                ok = os.path.exists(pdf_path)
                warnings = []
                if save_updated_fields:
                    warnings.append(
                        "LibreOffice 渲染器不支持把刷新后的域写回 docx；"
                        "交付 docx 请改用 Word / WPS 渲染器"
                    )
                return RenderResult(
                    ok=ok,
                    pdf=pdf_path if ok else None,
                    renderer_name="LibreOffice",
                    renderer_version=version,
                    engine_path=exe,
                    warnings=warnings,
                    errors=[],
                    stats={},
                )
            err_text = (
                (err or out or b"").decode("utf-8", "replace").strip() or "soffice 未产出 PDF"
            )[:300]
            return RenderResult(
                ok=False,
                pdf=None,
                renderer_name="LibreOffice",
                engine_path=exe,
                renderer_version=version,
                warnings=[],
                errors=[err_text],
                stats={},
            )
        except Exception as e:
            if proc is not None and proc.poll() is None:
                _kill_process_tree(proc.pid)
            return RenderResult(
                ok=False,
                pdf=None,
                renderer_name="LibreOffice",
                engine_path=exe,
                warnings=[],
                errors=[f"{type(e).__name__}: {e}"],
                stats={},
            )


class WPSRenderer(RendererAdapter):
    """本机 WPS Writer（COM: KWPS.Application）渲染器——与 WordRenderer 平行。

    仅在装了 WPS Office 的 Windows 上可用；逻辑与 WordRenderer 一致（开文档 →
    更新域 → 导 PDF），只是换了个 ProgID。未验证环境里用 available() 探测。
    """

    @classmethod
    def available(cls):
        try:
            import pythoncom  # noqa: F401
            import win32com.client as win32  # noqa: F401
        except ImportError:
            return False, "未安装 pywin32（WPS 渲染器需要它）"
        try:
            pythoncom.CoInitialize()
            try:
                app = win32.DispatchEx("KWPS.Application")
                name = getattr(app, "Name", "WPS Writer")
                app.Quit()
            finally:
                pythoncom.CoUninitialize()
        except Exception as e:
            return False, "WPS 启动失败（%s：%s）" % (type(e).__name__, e)
        return True, name

    def render(
        self,
        docx_path: str,
        pdf_path: str,
        *,
        save_updated_fields: bool = False,
    ) -> RenderResult:
        WD_PDF = 17
        co_initialized = False
        tmp_dir = tempfile.mkdtemp(prefix="texere_render_wps_")
        tmp_src = os.path.join(tmp_dir, os.path.basename(docx_path))
        try:
            # 同 Word：pywin32 缺失也要走结构化失败，不裸抛
            import pythoncom
            import win32com.client as win32

            shutil.copy2(docx_path, tmp_src)
            pythoncom.CoInitialize()
            co_initialized = True
            app = win32.DispatchEx("KWPS.Application")
            app.Visible = False
            try:
                doc = app.Documents.Open(os.path.abspath(tmp_src), False, False, False)
                doc.Fields.Update()
                doc.ExportAsFixedFormat(os.path.abspath(pdf_path), WD_PDF)
                ok = os.path.exists(pdf_path)
                if save_updated_fields:
                    doc.SaveAs(os.path.abspath(docx_path))
                doc.Close(0)
                return RenderResult(
                    ok=ok,
                    pdf=pdf_path if ok else None,
                    renderer_name="WPS Writer",
                    renderer_version=getattr(app, "Version", ""),
                    engine_path=getattr(app, "Path", ""),
                    warnings=[],
                    errors=[],
                    stats={},
                )
            finally:
                app.Quit()
        except Exception as e:  # 捕获一切 Office 异常，转成结构化 Result
            return RenderResult(
                ok=False,
                pdf=None,
                renderer_name="WPS Writer",
                renderer_version="",
                engine_path="",
                warnings=[],
                errors=[f"{type(e).__name__}: {e}"],
                stats={},
            )
        finally:
            if co_initialized:
                pythoncom.CoUninitialize()
            shutil.rmtree(tmp_dir, ignore_errors=True)


# 渲染器能力探测统一入口（每个渲染器各自声明 available()）
SUPPORTED_RENDERERS = ("word", "libreoffice", "wps")


def get_renderer(name: str | None = None) -> RendererAdapter:
    """按名字返回渲染器实例（默认 word）。

    只做名字→实现的映射，不在此处探测可用性——可用性请在调用方用
    type(r).available() 显式检查（doctor / preflight 已这么做）。
    """
    name = (name or "word").lower()
    if name == "word":
        return WordRenderer()
    if name == "libreoffice":
        return LibreOfficeRenderer()
    if name == "wps":
        return WPSRenderer()
    raise ValueError("未知渲染器: %s（支持 %s）" % (name, ", ".join(SUPPORTED_RENDERERS)))
