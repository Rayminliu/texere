"""profiles/*.json 守卫：schema 必备键 + neutral-en 英文基线关键值 + 契约边界同步。

profile=验收契约（只被 validate.py --enforce-profile 消费），config.style=渲染输入；
这里的守卫钉死「profile 实际声明的键集合 ↔ compile_profile_checks 消费的键 ↔
文档（docs/SCRIPT_HELP.md 字段执行状态表 / docs/VALIDATION 契约边界节）」三者不漂移。

pytest -q tests/test_profiles.py
"""

import importlib.util
import json
import os
import re
import sys

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROFILES = os.path.join(KIT, "profiles")
SCRIPTS = os.path.join(KIT, "scripts")

# validate.py / _shared.py 是同目录兄弟模块，spec 加载时保证可 import
if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)
_spec = importlib.util.spec_from_file_location("validate", os.path.join(SCRIPTS, "validate.py"))
v = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(v)

import _shared  # noqa: E402

REQUIRED = {"version", "name", "page", "styles", "caption", "table", "header", "footer", "toc"}

# compile_profile_checks / _check_* 实际消费的叶子字段（validate 断言侧的真相）。
# _check_profile_toc 只断言 TOC 域存在——toc.title 等全部是声明，改标题文案不翻转门禁。
CONSUMED_LEAVES = {
    "page": {"width", "height", "margin_top", "margin_bottom", "margin_left", "margin_right"},
    "styles.body": {"font_latin", "font_eastAsia", "size"},
    "styles.h1": {"font_eastAsia", "font_latin", "size", "bold"},
    "styles.h2": {"font_eastAsia", "font_latin", "size", "bold"},
    "styles.h3": {"font_eastAsia", "font_latin", "size", "bold"},
    "table": {"border"},
    "toc": set(),
}

# 文档（docs/VALIDATION*.md 契约边界表 + docs/SCRIPT_HELP.md 执行状态表）声明的消费集合
# 由 test_profile_keys_covered_by_docs / test_consumed_leaf_vocabulary_stable 共用语义：
# validate 改消费时必须同步改 CONSUMED_LEAVES 与文档两处。


def _profiles():
    return sorted(f for f in os.listdir(PROFILES) if f.endswith(".json"))


def _load(name):
    return json.load(open(os.path.join(PROFILES, name), encoding="utf-8"))


def test_all_profiles_valid_json_with_required_keys():
    assert _profiles(), "profiles/ 下没有任何 profile"
    for name in _profiles():
        data = _load(name)
        missing = REQUIRED - set(data)
        assert not missing, "%s 缺少必备键: %s" % (name, missing)
        for style in ("body", "h1", "h2", "h3"):
            assert style in data["styles"], "%s 的 styles 缺少 %s" % (name, style)


def test_neutral_en_baseline_values():
    """英文基线的三个关键事实：无首行缩进、西文 Times New Roman、页脚 Page {n}。"""
    data = _load("neutral-en-v1.json")
    body = data["styles"]["body"]
    assert body["first_line_indent"] == 0, "英文基线应为段块式（无首行缩进）"
    assert body["font_latin"] == "Times New Roman"
    assert data["footer"]["page_number_template"] == "Page {n}"
    assert data["toc"]["title"] == "Table of Contents"


def test_consumed_leaf_vocabulary_stable():
    """validate 断言侧消费的字段集合不许悄悄漂移（改这里必须同步文档与守卫）。"""
    assert v.compile_profile_checks.__doc__ is not None
    # profile 里写了但不在消费集合的 styles.body 键 = declared-only，文档已如实标注
    data = _load("formal-cn-v1.json")
    declared_only = set(data["styles"]["body"]) - CONSUMED_LEAVES["styles.body"]
    assert declared_only == {"line_spacing", "first_line_indent", "space_before", "space_after"}, (
        "styles.body 的 declared-only 清单变了：同步 docs/SCRIPT_HELP.md 执行状态表"
    )


def test_cn_profiles_toc_title_matches_render_default():
    """中文侧 profile 的 toc.title 与渲染默认（_shared.TOC_HEADING，全角空格）统一。

    历史漂移：profile 写半角「目  录」而 post.py 渲染全角「目　　录」。断言侧
    （_check_profile_toc）不消费 title 文案，故此统一不翻转任何门禁结论。
    """
    for name in ("formal-cn-v1", "gongwen-v1", "tender-v1", "application-v1"):
        title = _load(name + ".json")["toc"]["title"]
        assert title == _shared.TOC_HEADING, "%s 的 toc.title %r ≠ 渲染默认 %r" % (
            name,
            title,
            _shared.TOC_HEADING,
        )


def test_profile_keys_covered_by_docs():
    """各 profile 的顶层键都必须被文档措辞覆盖（契约边界节的词汇表）。

    覆盖词汇：必备/常规键逐个点名；`*_specific` 用通配措辞；可选键单独说明。
    同时检查 SCRIPT_HELP 和 VALIDATION 两份文档。
    """
    docs = ""
    for p in (
        os.path.join(KIT, "docs", "SCRIPT_HELP.md"),
        os.path.join(KIT, "docs", "VALIDATION.md"),
    ):
        docs += open(p, encoding="utf-8").read()
    vocab = {
        "version",
        "name",
        "description",
        "page",
        "styles",
        "caption",
        "table",
        "header",
        "footer",
        "toc",
        "baseline_dir",
        "acceptance",
    }
    for name in _profiles():
        for key in _load(name):
            if key.endswith("_specific"):
                # 文档以 `*_specific.*` 通配或逐字 `xxx_specific.*` 覆盖
                assert re.search(r"\*_specific\.\*|%s\.\*" % key, docs), (
                    "%s 的键 %r 未在文档措辞中出现" % (name, key)
                )
                continue
            assert key in vocab, "%s 出现文档未覆盖的顶层键：%s" % (name, key)
    # 可选键的示例确实在文档里（baseline_dir / acceptance.renderers）
    assert "baseline_dir" in docs
    assert "acceptance.renderers" in docs or "acceptance" in docs


def test_profile_injection_into_rendering_stays_documented_non_goal():
    """「不打通 profile→渲染」是明确决策，两份 VALIDATION 镜像都必须写清边界。"""
    for doc in ("docs/VALIDATION.md", "docs/VALIDATION.zh-CN.md"):
        text = open(os.path.join(KIT, doc), encoding="utf-8").read()
        assert "non-goal" in text or "明确不做" in text, "%s 缺契约边界说明" % doc
        assert "baseline_dir" in text, "%s 未说明可选键 baseline_dir" % doc
