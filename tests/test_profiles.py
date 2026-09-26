"""profiles/*.json 守卫：schema 必备键 + neutral-en 英文基线关键值。

pytest -q tests/test_profiles.py
"""

import json
import os

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROFILES = os.path.join(KIT, "profiles")

REQUIRED = {"version", "name", "page", "styles", "caption", "table", "header", "footer", "toc"}


def _profiles():
    return sorted(f for f in os.listdir(PROFILES) if f.endswith(".json"))


def test_all_profiles_valid_json_with_required_keys():
    assert _profiles(), "profiles/ 下没有任何 profile"
    for name in _profiles():
        data = json.load(open(os.path.join(PROFILES, name), encoding="utf-8"))
        missing = REQUIRED - set(data)
        assert not missing, "%s 缺少必备键: %s" % (name, missing)
        for style in ("body", "h1", "h2", "h3"):
            assert style in data["styles"], "%s 的 styles 缺少 %s" % (name, style)


def test_neutral_en_baseline_values():
    """英文基线的三个关键事实：无首行缩进、西文 Times New Roman、页脚 Page {n}。"""
    data = json.load(open(os.path.join(PROFILES, "neutral-en-v1.json"), encoding="utf-8"))
    body = data["styles"]["body"]
    assert body["first_line_indent"] == 0, "英文基线应为段块式（无首行缩进）"
    assert body["font_latin"] == "Times New Roman"
    assert data["footer"]["page_number_template"] == "Page {n}"
    assert data["toc"]["title"] == "Table of Contents"
