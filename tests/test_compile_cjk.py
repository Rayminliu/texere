"""CJK 直引号处理与 --resource-path 字符预算的单测（R4 #1 + 复现时踩到的 WinError 206）。

pandoc smart 的引号开闭判定以西文空格分词为前提：CJK 文本里 ASCII 直引号
6/6 错配（pandoc 3.11 实测，段首也不幸免）。这里钉死两件事的边界：
is_cjk_dominant 的中文判定口径、pair_cjk_quotes 的围栏/行内代码/奇数段豁免。

resource-path 预算用 importlib 加载 render.py（test_docs_sync 同款模式：
import 无副作用，CI 无 Word 也能跑）；末尾一条真 pandoc 端到端守住
「合并 → 配对 → 关 smart → 落 docx」整条链。
"""

import importlib.util
import os
import shutil
import sys

import pytest

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

sys.path.insert(0, os.path.join(KIT, "scripts"))

from _compile import is_cjk_dominant, pair_cjk_quotes  # noqa: E402

# ---------------------------------------------------------------------------
# is_cjk_dominant：中文判定口径
# ---------------------------------------------------------------------------


def test_chinese_document_is_dominant():
    assert is_cjk_dominant("这是一份足够长的中文正式文档，正文以汉字为主。" * 8)


def test_english_document_is_not_dominant():
    assert not is_cjk_dominant("This is an English progress report with many words. " * 20)


def test_few_han_chars_in_english_report_do_not_trigger():
    # 英文报告里出现几个中文地名，不该触发中文口径（smart 对西文是必要的）
    md = "Projects in Shenzhen and Beijing delivered on time. " * 20 + "深圳 北京 广州"
    assert not is_cjk_dominant(md)


def test_han_chars_inside_code_fences_do_not_count():
    md = "只有几个汉字。\n\n```python\n# " + "汉" * 200 + "\n```\n"
    assert not is_cjk_dominant(md)


# ---------------------------------------------------------------------------
# pair_cjk_quotes：配对规则
# ---------------------------------------------------------------------------


def test_pair_basic():
    out, pairs, odd = pair_cjk_quotes('他说"你好"了')
    assert out == "他说“你好”了"
    assert (pairs, odd) == (1, 0)


def test_pair_multiple_in_one_paragraph():
    out, pairs, odd = pair_cjk_quotes('写"甲"和"乙"两处')
    assert out == "写“甲”和“乙”两处"
    assert (pairs, odd) == (2, 0)


def test_alternation_resets_per_paragraph():
    out, pairs, odd = pair_cjk_quotes('"甲"\n\n"乙"')
    assert out == "“甲”\n\n“乙”"
    assert (pairs, odd) == (2, 0)


def test_odd_count_paragraph_is_left_untouched():
    text = '他说"你好了'
    out, pairs, odd = pair_cjk_quotes(text)
    assert out == text
    assert (pairs, odd) == (0, 1)


def test_double_backtick_fence_verbatim():
    md = '正文"配对"正常。\n\n```\nfence "不" 动\n```\n\n再"来"一段。'
    out, pairs, odd = pair_cjk_quotes(md)
    assert 'fence "不" 动' in out
    assert "“配对”" in out and "“来”" in out
    assert (pairs, odd) == (2, 0)


def test_tilde_fence_verbatim():
    md = '正文"配对"。\n\n~~~\n"不" 动\n~~~\n'
    out, pairs, odd = pair_cjk_quotes(md)
    assert '"不" 动' in out
    assert (pairs, odd) == (1, 0)


def test_inline_code_spans_are_skipped():
    out, pairs, odd = pair_cjk_quotes('像 `"x"` 和 "甲" 这样')
    assert '`"x"`' in out and "“甲”" in out
    assert (pairs, odd) == (1, 0)


def test_escaped_quote_is_not_counted():
    out, pairs, odd = pair_cjk_quotes('转义 \\" 不算，"配对"算')
    assert '\\"' in out and "“配对”" in out
    assert (pairs, odd) == (1, 0)


def test_fullwidth_quotes_pass_through():
    text = "全角“引号”本就不该动。"
    out, pairs, odd = pair_cjk_quotes(text)
    assert out == text
    assert (pairs, odd) == (0, 0)


# ---------------------------------------------------------------------------
# _resource_paths：字符预算（WinError 206 兜底）
# ---------------------------------------------------------------------------


def _load_render():
    sys.path.insert(0, os.path.join(KIT, "scripts"))
    spec = importlib.util.spec_from_file_location(
        "render_units_guard", os.path.join(KIT, "scripts", "render.py")
    )
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def test_resource_path_budget_trims_parent_expansion(tmp_path, capsys):
    mod = _load_render()
    src = tmp_path / "src"
    src.mkdir()
    for i in range(20):
        (tmp_path / ("sibling_%02d" % i)).mkdir()
    mod.RESOURCE_PATH_BUDGET = 10
    uniq = mod._resource_paths(str(src), {})
    out = capsys.readouterr().out
    assert "[warn] --resource-path 超出预算" in out
    assert uniq == [str(src)]  # src 自身永不裁，父目录扩展全部裁掉


def test_resource_path_explicit_paths_survive_budget(tmp_path, capsys):
    mod = _load_render()
    src = tmp_path / "src"
    src.mkdir()
    extra = tmp_path / "extra"
    extra.mkdir()
    mod.RESOURCE_PATH_BUDGET = 10
    uniq = mod._resource_paths(str(src), {"resource_paths": [str(extra)]})
    assert str(extra) in uniq
    assert str(src) in uniq


def test_resource_path_no_warn_under_budget(tmp_path, capsys):
    mod = _load_render()
    src = tmp_path / "src"
    src.mkdir()
    (src / "media").mkdir()
    uniq = mod._resource_paths(str(src), {})
    assert "[warn]" not in capsys.readouterr().out
    assert str(src / "media") in uniq  # src 子目录属于 core，照常保留


# ---------------------------------------------------------------------------
# 端到端：合并 → 配对 → 关 smart → pandoc 落 docx
# ---------------------------------------------------------------------------


@pytest.mark.skipif(shutil.which("pandoc") is None, reason="需要 pandoc")
def test_render_pairs_cjk_quotes_end_to_end(tmp_path):
    mod = _load_render()
    src = tmp_path / "src"
    src.mkdir()
    (src / "01.md").write_text(
        "这是一份中文正式文档的正文，篇幅足以被判成中文主导。" * 6
        + '\n\n他说"你好"了，这是"测试"。\n\n# 一级标题\n\n正文段落。\n',
        encoding="utf-8",
    )
    out = tmp_path / "out.docx"
    mod.render(str(src), str(out), None, False, False, work_mode=mod._Workdir.DISCARD)
    from docx import Document

    text = "\n".join(p.text for p in Document(str(out)).paragraphs)
    assert "“你好”" in text and "“测试”" in text
    assert '"' not in text  # smart 关闭后不应残留任何西文直引号
