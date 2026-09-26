"""align_tables 单元测试：显示宽度重排、内容不变守卫、围栏/pipe table 不误伤。

pytest -q tests/test_align_tables.py
"""

import os
import shutil
import subprocess
import sys

import pytest

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPT = os.path.join(KIT, "scripts", "align_tables.py")

pytestmark = pytest.mark.skipif(shutil.which("pandoc") is None, reason="e2e 用例需要 pandoc")

MISALIGNED = """\
+-----+--------+
| 表 1-1 | Wide Column |
+=====+========+
| A | B |
+-----+--------+
"""

ALIGNED = """\
+--------+-------------+
| 表 1-1 | Wide Column |
+========+=============+
| A      | B           |
+--------+-------------+
"""


def _run(tmp_path, text, *flags):
    p = tmp_path / "t.md"
    p.write_text(text, encoding="utf-8")
    r = subprocess.run([sys.executable, SCRIPT, *flags, str(p)], capture_output=True, text=True)
    return p, r


def test_check_misaligned_exits_1(tmp_path):
    _, r = _run(tmp_path, MISALIGNED, "--check")
    assert r.returncode == 1
    assert "未对齐" in r.stderr


def test_check_aligned_exits_0(tmp_path):
    _, r = _run(tmp_path, ALIGNED, "--check")
    assert r.returncode == 0


def test_fix_aligns_and_preserves_cells(tmp_path):
    p, r = _run(tmp_path, MISALIGNED, "--fix")
    assert r.returncode == 0
    assert p.read_text(encoding="utf-8") == ALIGNED


def test_fix_report_mode_default_no_write(tmp_path):
    p, r = _run(tmp_path, MISALIGNED)
    assert r.returncode == 0
    assert p.read_text(encoding="utf-8") == MISALIGNED  # 默认只报告
    assert "需对齐" in r.stdout


def test_pipe_table_and_fence_untouched(tmp_path):
    text = "```bash\n+---+\n| x |\n+---+\n```\n\n| a | b |\n|---|---|\n| 1 | 2 |\n"
    _, r = _run(tmp_path, text, "--fix")
    assert r.returncode == 0
    assert (tmp_path / "t.md").read_text(encoding="utf-8") == text  # 围栏与 pipe table 不碰


def test_unparseable_no_write(tmp_path):
    text = "+---+---+\n| a |\n+---+---+\n"  # 内容行列数与分隔行不一致
    _, r = _run(tmp_path, text, "--fix")
    assert r.returncode == 1
    assert "无法安全解析" in r.stderr
    assert (tmp_path / "t.md").read_text(encoding="utf-8") == text  # 不写回


def test_multi_line_rows_preserved(tmp_path):
    text = "+------+---+\n| 长 的 | b |\n| 内容 |   |\n+------+---+\n"
    _, r = _run(tmp_path, text, "--fix")
    assert r.returncode == 0
    out = (tmp_path / "t.md").read_text(encoding="utf-8")
    for cell in ["长 的", "内容", "b"]:
        assert cell in out  # 多行单元格内容逐字保留


def test_indented_block_indent_preserved(tmp_path):
    text = "  +---+---+\n  | a | b |\n  +---+---+\n"
    _, r = _run(tmp_path, text, "--fix")
    assert r.returncode == 0
    assert all(
        line.startswith("  ")
        for line in (tmp_path / "t.md").read_text().splitlines()
        if line.strip()
    )


def test_pandoc_parse_equivalent_to_handwritten(tmp_path):
    """--fix 后 pandoc 解析结果必须与手写对齐版一致（同内容即同解析）。"""
    fixed = tmp_path / "fixed.md"
    hand = tmp_path / "hand.md"
    fixed.write_text(MISALIGNED, encoding="utf-8")
    hand.write_text(ALIGNED, encoding="utf-8")
    subprocess.run([sys.executable, SCRIPT, "--fix", str(fixed)], check=True, capture_output=True)
    outs = []
    for src in (fixed, hand):
        r = subprocess.run(
            ["pandoc", "-f", "markdown", "-t", "plain", str(src)],
            capture_output=True,
            text=True,
            check=True,
        )
        outs.append(r.stdout)
    assert outs[0] == outs[1], "对齐器修复后的表格与手写对齐版解析不一致"
