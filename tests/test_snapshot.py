# -*- coding: utf-8 -*-
"""snapshot.py / check_pdf.py 的退出码与快照比对逻辑。

用 fitz 现场造 PDF，不需要 Word，秒级可跑。
"""
import os
import subprocess
import sys

import pytest

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

fitz = pytest.importorskip("fitz", reason="需要 PyMuPDF")

LOREM = ("This sentence is long enough to be treated as real body text "
         "by the near-empty page detector in check_pdf.py. " * 2)


def make_pdf(path, lines, pages=1):
    doc = fitz.open()
    for _ in range(pages):
        page = doc.new_page()
        for i, t in enumerate(lines):
            page.insert_text((72, 100 + i * 18), t, fontsize=11)
    doc.save(str(path))
    return str(path)


def run_script(name, *args):
    # 固定子进程编码，否则 Windows 下拿到的是 GBK 字节，断言会假失败
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    return subprocess.run([sys.executable, os.path.join(KIT, "scripts", name), *args],
                          capture_output=True, env=env)


def test_check_pdf_passes_on_text_page(tmp_path):
    pdf = make_pdf(tmp_path / "ok.pdf", [LOREM])
    r = run_script("check_pdf.py", pdf)
    assert r.returncode == 0, r.stdout.decode("utf-8", "replace")


def test_check_pdf_fails_on_empty_page(tmp_path):
    pdf = make_pdf(tmp_path / "empty.pdf", [], pages=1)
    r = run_script("check_pdf.py", pdf)
    assert r.returncode == 1, "空白页必须让验收失败（CI 靠退出码判定）"


def test_check_pdf_max_empty_relaxes(tmp_path):
    pdf = make_pdf(tmp_path / "empty.pdf", [])
    r = run_script("check_pdf.py", pdf, "--max-empty", "1")
    assert r.returncode == 0, "放宽阈值后应通过"


def test_snapshot_roundtrip(tmp_path):
    """录基线后比对自身必须一致。"""
    pdf = make_pdf(tmp_path / "a.pdf", [LOREM])
    assert run_script("snapshot.py", pdf, "--update").returncode == 0
    r = run_script("snapshot.py", pdf)
    assert r.returncode == 0, r.stdout.decode("utf-8", "replace")


def test_snapshot_detects_drift(tmp_path):
    """版式变了必须失败。"""
    base = tmp_path / "base.pdf"
    make_pdf(base, [LOREM])
    assert run_script("snapshot.py", str(base), "--update").returncode == 0

    # 同一路径重新生成、内容不同 -> 与基线不符
    make_pdf(base, ["COMPLETELY DIFFERENT CONTENT " * 6])
    r = run_script("snapshot.py", str(base))
    assert r.returncode == 1, "内容变了却通过，快照就白做了"


def test_snapshot_detects_page_count_change(tmp_path):
    base = tmp_path / "pages.pdf"
    make_pdf(base, [LOREM], pages=2)
    assert run_script("snapshot.py", str(base), "--update").returncode == 0
    make_pdf(base, [LOREM], pages=3)
    assert run_script("snapshot.py", str(base)).returncode == 1


def test_diff_ratio_semantics():
    """完全相同=0；长度不同=1；逐字节算比例。"""
    import importlib.util
    spec = importlib.util.spec_from_file_location(
        "snapshot", os.path.join(KIT, "scripts", "snapshot.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    assert mod.diff_ratio(b"abc", b"abc") == 0.0
    assert mod.diff_ratio(b"abc", b"abd") == pytest.approx(1 / 3)
    assert mod.diff_ratio(b"abc", b"ab") == 1.0


def test_snapshot_requires_baseline(tmp_path):
    pdf = make_pdf(tmp_path / "nobase.pdf", [LOREM])
    r = run_script("snapshot.py", pdf)
    assert r.returncode == 1
    # sys.exit 的提示走 stderr
    assert "基线" in (r.stdout + r.stderr).decode("utf-8", "replace")
