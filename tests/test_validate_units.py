"""validate.py 纯函数单元测试：页码识别与基线比对，不启动 Word。

这些逻辑此前只在 CLI 级测试里被间接覆盖（且只测了 skip 路径），
这里对核心判定函数做直接断言。
"""

import importlib.util
import os
import sys

import pytest

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(KIT, "scripts")

# validate.py 与 snapshot.py 是同目录兄弟模块，spec 加载时保证可 import
sys.path.insert(0, SCRIPTS)
_spec = importlib.util.spec_from_file_location("validate", os.path.join(SCRIPTS, "validate.py"))
v = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(v)


# ------------------------------------------------------------- 页脚页码识别


class TestFooterPageNumber:
    def test_em_dash_format(self):
        assert v.footer_page_number("— 3 —", 100) == 3

    def test_hyphen_dash_format(self):
        assert v.footer_page_number("- 12 -", 100) == 12

    def test_cn_format(self):
        assert v.footer_page_number("第 5 页", 100) == 5

    def test_cn_format_with_total(self):
        assert v.footer_page_number("第 5 页 共 20 页", 100) == 5

    def test_page_format(self):
        assert v.footer_page_number("Page 7", 100) == 7

    def test_bare_number(self):
        assert v.footer_page_number("9", 100) == 9

    def test_prose_with_digits_is_not_page_number(self):
        # 正文里的日期/数量不能被当成页码
        assert v.footer_page_number("2026年9月", 100) is None
        assert v.footer_page_number("工期 90 日历天", 100) is None

    def test_out_of_range_rejected(self):
        # 匹配到数字但超出页数范围（如正文恰好一行裸数字 2026）
        assert v.footer_page_number("2026", 4) is None


class FakePage:
    def __init__(self, text):
        self._text = text

    def get_text(self, kind, sort=False):
        return self._text


class FakePdf:
    def __init__(self, pages):
        self.pages = pages

    def __len__(self):
        return len(self.pages)

    def __getitem__(self, i):
        return self.pages[i]


class TestCollectPageNumbers:
    def test_skips_pages_without_footer_number(self):
        # 封面（无页脚页码）+ 正文两页（— 1 — / — 2 —）
        doc = FakePdf(
            [
                FakePage("投标文件\n投 标 文 件"),
                FakePage("正文内容一二三\n— 1 —"),
                FakePage("更多正文\n— 2 —"),
            ]
        )
        assert v.collect_page_numbers(doc) == [1, 2]

    def test_continuous_sequence(self):
        doc = FakePdf([FakePage("a\n— %d —" % n) for n in (1, 2, 3, 4)])
        assert v.collect_page_numbers(doc) == [1, 2, 3, 4]

    def test_detects_gap(self):
        # 4 页文档、页码 1、2、4（末页无页码）：缺 3 —— 连续性校验应报失败
        doc = FakePdf(
            [
                FakePage("x\n— 1 —"),
                FakePage("x\n— 2 —"),
                FakePage("x\n— 4 —"),
                FakePage("封底\n无页码"),
            ]
        )
        numbers = v.collect_page_numbers(doc)
        assert numbers == [1, 2, 4]
        expected = list(range(min(numbers), max(numbers) + 1))
        assert sorted(numbers) != expected


# ------------------------------------------------------------- 基线像素比对


class TestBaselinePageDiff:
    def test_identical_image_has_zero_diff(self):
        pytest.importorskip("fitz")
        p = os.path.join(KIT, "baselines", "p001.png")
        assert v.baseline_page_diff(p, p) == 0.0

    def test_different_pages_report_drift(self):
        # 基线里的第 1 页 vs 第 2 页：必须判为漂移（> 0.1% 阈值）
        pytest.importorskip("fitz")
        p1 = os.path.join(KIT, "baselines", "p001.png")
        p2 = os.path.join(KIT, "baselines", "p002.png")
        assert v.baseline_page_diff(p1, p2) > v.DEFAULT_MAX_DIFF
