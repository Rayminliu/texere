"""版本号一致性测试：scripts/_version.py 与 pyproject.toml 不许漂移。"""

import os
import re
import subprocess
import sys

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

sys.path.insert(0, os.path.join(KIT, "scripts"))
from _version import __version__  # noqa: E402


def test_version_matches_pyproject():
    """_version.py 必须与 pyproject.toml 的 [project] version 一致。"""
    with open(os.path.join(KIT, "pyproject.toml"), encoding="utf-8") as f:
        m = re.search(r'^version\s*=\s*"([^"]+)"', f.read(), re.MULTILINE)
    assert m, "pyproject.toml 里找不到 version 字段"
    assert m.group(1) == __version__, (
        f"版本漂移：_version.py={__version__} vs pyproject.toml={m.group(1)}；两处要一起改"
    )


def test_render_cli_reports_current_version():
    r = subprocess.run(
        [sys.executable, os.path.join(KIT, "scripts", "render.py"), "--version"],
        capture_output=True,
        text=True,
    )
    assert r.returncode == 0
    assert __version__ in r.stdout, r.stdout


def test_no_hardcoded_tool_version_left():
    """validate.py / patch.py 不允许再写死版本号字面量，必须走 _version.py。"""
    for name in ("validate.py", "patch.py"):
        with open(os.path.join(KIT, "scripts", name), encoding="utf-8") as f:
            src = f.read()
        assert '"tool_version": "0' not in src, f"{name} 里还有硬编码 tool_version"
