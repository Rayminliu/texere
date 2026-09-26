"""全量套件共享的样例渲染产物：整个 pytest 进程只渲染一次。

历史上 test_patch / test_validate 的每条 e2e 用例都子进程重渲染同一份样例
（pandoc + 后处理 + Word 导出，41 次 ≈ 本地全量耗时的大头）。这里渲染一次，
之后每条用例拿现成产物的拷贝——改写型用例（--apply 重写输入）不污染共享源，
只读用例拿到的是同一字节状态。

test_patch 与 test_validate 各加载一份本模块；加载器用 sys.modules 去重，
保证两个文件拿到同一个实例，缓存才真正跨文件生效。
"""

import atexit
import os
import shutil
import subprocess
import sys
import tempfile

import pytest

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

_RENDERED = None  # 渲染产物 docx 路径（进程级缓存）


def _require_renderer():
    """默认渲染器不可用就 skip：与原先两个测试文件里的守卫语义一致（托管 CI 逐用例跳过）。"""
    import importlib.util

    spec = importlib.util.spec_from_file_location(
        "renderers", os.path.join(KIT, "scripts", "renderers.py")
    )
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    ok, why = mod.WordRenderer.available()
    if not ok:
        pytest.skip("渲染器不可用（%s）：e2e 需 render.py --pdf" % why)


def rendered_copy(tmp_path):
    """返回样例 docx 的一份拷贝；进程内首次调用触发唯一一次渲染。"""
    global _RENDERED
    if _RENDERED is None:
        _require_renderer()
        base = tempfile.mkdtemp(prefix="texere_shared_sample_")
        atexit.register(shutil.rmtree, base, ignore_errors=True)
        src = os.path.join(base, "src")
        os.mkdir(src)
        sample_md = os.path.join(KIT, "assets", "sample.md")
        shutil.copy(sample_md, os.path.join(src, "01_sample.md"))
        rendered = os.path.join(base, "test.docx")
        subprocess.run(
            [
                sys.executable,
                os.path.join(KIT, "scripts", "render.py"),
                "--src",
                src,
                "--out",
                rendered,
                "--config",
                os.path.join(KIT, "assets", "sample_config.json"),
                "--pdf",
            ],
            check=True,
            capture_output=True,
        )
        _RENDERED = rendered
    return shutil.copy2(_RENDERED, str(tmp_path / "test.docx"))
