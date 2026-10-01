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
        encoding="utf-8",
    )
    assert r.returncode == 0
    assert __version__ in r.stdout, r.stdout


def test_no_hardcoded_tool_version_left():
    """validate.py / patch.py 不允许再写死版本号字面量，必须走 _version.py。"""
    for name in ("validate.py", "patch.py"):
        with open(os.path.join(KIT, "scripts", name), encoding="utf-8") as f:
            src = f.read()
        assert '"tool_version": "0' not in src, f"{name} 里还有硬编码 tool_version"


def test_changelog_top_release_matches_version():
    """CHANGELOG 首个带版本号的 `## X.Y.Z` 标题必须等于当前版本。

    _version.py ↔ pyproject 由上面那条守住，但 CHANGELOG 是第三个会漂移的地方：
    取「首个带版本号的」标题而不是「首个 `## `」，因为顶部可以有合法的「未发布」
    段落（maintenance-mode 下改动先攒在那里）。发版时忘了把未发布段收编成
    `## <新版本号> — <日期>`，用户读到的「最新变更」就不是实际发布的那一版。
    """
    with open(os.path.join(KIT, "CHANGELOG.md"), encoding="utf-8") as f:
        text = f.read()
    m = re.search(r"^##\s+(\d+\.\d+\.\d+)\b", text, re.MULTILINE)
    assert m, "CHANGELOG.md 里找不到任何 `## X.Y.Z` 发布标题"
    assert m.group(1) == __version__, (
        f"CHANGELOG 顶部发布版本 {m.group(1)} != 当前版本 {__version__}；"
        "发版时要把「未发布」段落收编成带新版本号的标题"
    )
