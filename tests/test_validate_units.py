"""validate.py 纯函数单元测试：页码识别与基线比对，不启动 Word。

这些逻辑此前只在 CLI 级测试里被间接覆盖（且只测了 skip 路径），
这里对核心判定函数做直接断言。
"""

import glob
import hashlib
import importlib.util
import os
import sys

import pytest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

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


# ------------------------------------------------------- 状态分级 & 反 fail-open


def _save_docx(tmp_path, name="a.docx"):
    p = tmp_path / name
    Document().save(str(p))
    return str(p)


class TestStatusConstants:
    def test_four_statuses_exist(self):
        assert (v.PASS, v.FAIL, v.SKIP, v.ERROR) == ("PASS", "FAIL", "SKIP", "ERROR")


class TestNoFailOpen:
    """「查不了」必须报 SKIP 或 ERROR，绝不能报 PASS。

    这是验收器最容易失真的地方：前置条件缺失时旧实现一律 return True，
    于是 9 项里塞着几项从没真跑过的检查，摘要却写 Passed: 9/9。
    """

    def test_page_numbering_without_pdf_is_skip(self):
        assert v.check_page_numbering(None)[0] == v.SKIP

    def test_blank_pages_without_pdf_is_skip(self):
        assert v.check_blank_pages(None)[0] == v.SKIP

    def test_visual_drift_without_baseline_is_skip(self):
        assert v.check_visual_drift(None, None)[0] == v.SKIP

    def test_source_content_without_any_evidence_is_skip(self, tmp_path):
        assert v.check_source_content_integrity(_save_docx(tmp_path))[0] == v.SKIP

    def test_image_embedding_without_reference_is_skip(self, tmp_path):
        assert v.check_image_embedding(_save_docx(tmp_path))[0] == v.SKIP

    def test_toc_field_without_toc_is_skip(self, tmp_path):
        assert v.check_toc_field(_save_docx(tmp_path))[0] == v.SKIP


class TestSamplePageIndices:
    def test_single_page(self):
        assert v.sample_page_indices(1) == [0]

    def test_short_document(self):
        assert v.sample_page_indices(3) == [0, 2]

    def test_long_document(self):
        assert v.sample_page_indices(20) == [0, 10, 19]


# ------------------------------------------------------------- TOC 域真检查


def _docx_with_toc(path, instr='TOC \\o "1-2" \\h \\z \\u'):
    """按 post.py 注入目录域的写法造一个含 TOC 域的 docx。"""
    doc = Document()
    para = doc.add_paragraph()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    para.add_run()._r.append(begin)

    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = instr
    para.add_run()._r.append(it)

    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    para.add_run()._r.append(end)

    doc.save(str(path))
    return str(path)


class TestTocField:
    def test_detects_injected_toc(self, tmp_path):
        """python-docx 没有 tables_of_contents 属性，旧实现靠 hasattr 短路后
        这一项恒定跳过——这里守住「真能从 OOXML 里数出 TOC 域」。"""
        f = _docx_with_toc(tmp_path / "toc.docx")
        assert v.count_toc_fields(Document(f)) == 1
        assert v.check_toc_field(f)[0] == v.PASS

    def test_plain_document_is_skipped_not_passed(self, tmp_path):
        status, _ = v.check_toc_field(_save_docx(tmp_path, "plain.docx"))
        assert status == v.SKIP


# ------------------------------------------------------------- 源内容契约


def _docx_with_text(tmp_path, text, name="b.docx"):
    p = tmp_path / name
    doc = Document()
    doc.add_paragraph(text)
    doc.save(str(p))
    return str(p)


class TestMarkdownNormalization:
    """这五条都是真实样例上踩出来的误报（曾一次报出 7 处假 FAIL），
    改匹配器时必须一并守住，否则 --source-md 会变成噪音源。"""

    def test_curly_and_straight_quotes_match(self):
        assert v._normalize("采用“感知—平台—应用”架构") == v._normalize('采用"感知—平台—应用"架构')

    def test_pandoc_fenced_div_is_not_body(self):
        md = ':::: {custom-style="Lead"}\n正文内容足够长的这一句话\n::::\n'
        assert v._source_md_segments(md) == ["正文内容足够长的这一句话"]

    def test_hard_line_break_is_stripped(self):
        assert v._source_md_segments("项目名称：示例智慧园区平台建设项目\\\n") == [
            "项目名称：示例智慧园区平台建设项目"
        ]

    def test_list_marker_is_stripped(self):
        assert v._source_md_segments("- 电子与智能化工程专业承包资质；\n") == [
            "电子与智能化工程专业承包资质；"
        ]

    def test_table_separator_row_is_skipped(self):
        assert v._source_md_segments("|:---|:---|:---|\n") == []

    def test_code_fence_content_is_skipped(self):
        assert v._source_md_segments("```python\n这段代码不该参与比对\n```\n") == []


class TestSourceContentIntegrity:
    def test_md_equivalence_passes(self, tmp_path):
        text = "本项目的整体架构说明如下所述的完整段落"
        f = _docx_with_text(tmp_path, text)
        md = tmp_path / "a.md"
        md.write_text("# 标题\n\n%s\n" % text, encoding="utf-8")
        status, msg = v.check_source_content_integrity(f, source_md=str(md))
        assert status == v.PASS, msg

    def test_md_equivalence_detects_dropped_text(self, tmp_path):
        f = _docx_with_text(tmp_path, "文档里只有这一句话")
        md = tmp_path / "a.md"
        md.write_text("源文档里这一段根本没有进到 docx 里\n", encoding="utf-8")
        status, _ = v.check_source_content_integrity(f, source_md=str(md))
        assert status == v.FAIL

    def test_code_fence_is_not_checked_as_body(self, tmp_path):
        """代码块里的内容不该被当成正文去比对。"""
        f = _docx_with_text(tmp_path, "正文段落内容足够长度")
        md = tmp_path / "a.md"
        md.write_text(
            "正文段落内容足够长度\n\n```python\n这段代码不该出现在docx里\n```\n", encoding="utf-8"
        )
        status, msg = v.check_source_content_integrity(f, source_md=str(md))
        assert status == v.PASS, msg

    def test_expected_hash_matches(self, tmp_path):
        import hashlib

        f = _save_docx(tmp_path, "h.docx")
        h = hashlib.sha256(open(f, "rb").read()).hexdigest()
        status, msg = v.check_source_content_integrity(f, expected_hash=h)
        assert status == v.PASS, msg

    def test_expected_hash_mismatch_fails(self, tmp_path):
        status, _ = v.check_source_content_integrity(
            _save_docx(tmp_path, "h.docx"), expected_hash="deadbeef"
        )
        assert status == v.FAIL


# ------------------------------------------------------------- 图片计数


class TestImageEmbedding:
    def test_missing_image_fails(self, tmp_path):
        status, msg = v.check_image_embedding(_save_docx(tmp_path, "img.docx"), "![a](a.png)")
        assert status == v.FAIL, msg

    def test_count_is_a_lower_bound_only(self, tmp_path):
        """n_img >= n_ref 只证明数量够，不证明对应关系——这里把这条边界写死。"""
        status, msg = v.check_image_embedding(_save_docx(tmp_path, "img.docx"))
        assert status == v.SKIP
        assert "未提供 Markdown 引用" in msg


# ------------------------------------------------------------- 图片身份校验


def _img_paths():
    return sorted(glob.glob(os.path.join(KIT, "assets", "previews", "*.png")))


class TestImageIdentity:
    """数量检查（n_img >= n_ref）抓不住串位/错图/重复占位，这里守住身份校验。"""

    @staticmethod
    def _docx_with_images(paths, out):
        from docx import Document
        from docx.shared import Pt

        doc = Document()
        for p in paths:
            doc.add_picture(p, width=Pt(100))
        doc.save(str(out))
        return str(out)

    @staticmethod
    def _md(tmp_path, paths):
        md = tmp_path / "a.md"
        md.write_text(
            "\n\n".join("![图%d](%s)" % (i, p.replace("\\", "/")) for i, p in enumerate(paths)),
            encoding="utf-8",
        )
        return md

    def test_embedded_shas_match_source_bytes(self, tmp_path):
        """pandoc 原样嵌入字节，所以 sha256 可以直接当图片指纹（实测一致）。"""
        pytest.importorskip("PIL", reason="add_picture 需要 Pillow")
        a, b = _img_paths()[:2]
        f = self._docx_with_images([a, b], tmp_path / "p.docx")
        assert v._docx_image_shas(Document(f)) == [v._sha256_file(a), v._sha256_file(b)]

    def test_md_paths_are_parsed_in_order(self, tmp_path):
        (tmp_path / "imgs").mkdir()
        for name in ("a.png", "b.png"):
            (tmp_path / "imgs" / name).write_bytes(b"x")
        md = tmp_path / "a.md"
        md.write_text("![图一](imgs/a.png)\n\n![图二](imgs/b.png){width=3cm}\n", encoding="utf-8")
        got = v._md_image_paths(md.read_text(encoding="utf-8"), str(tmp_path))
        assert [os.path.basename(p) for p in got] == ["a.png", "b.png"]

    def test_same_image_twice_is_detected(self, tmp_path):
        """图 A / 图 B → docx 里两张都是 A：数量 2/2 是对的，但 B 根本没进去。"""
        pytest.importorskip("PIL", reason="add_picture 需要 Pillow")
        a, b = _img_paths()[:2]
        md = self._md(tmp_path, [a, b])
        f = self._docx_with_images([a, a], tmp_path / "d.docx")
        status, msg = v.check_image_embedding(f, md.read_text(encoding="utf-8"), str(md))
        assert status == v.FAIL, msg

    def test_matching_images_pass(self, tmp_path):
        pytest.importorskip("PIL", reason="add_picture 需要 Pillow")
        a, b = _img_paths()[:2]
        md = self._md(tmp_path, [a, b])
        f = self._docx_with_images([a, b], tmp_path / "d.docx")
        status, msg = v.check_image_embedding(f, md.read_text(encoding="utf-8"), str(md))
        assert status == v.PASS, msg

    def test_swapped_order_is_detected(self, tmp_path):
        """两张图都在，但次序反了——数量检查对这种情况完全无感。"""
        pytest.importorskip("PIL", reason="add_picture 需要 Pillow")
        a, b = _img_paths()[:2]
        md = self._md(tmp_path, [a, b])
        f = self._docx_with_images([b, a], tmp_path / "d.docx")
        status, msg = v.check_image_embedding(f, md.read_text(encoding="utf-8"), str(md))
        assert status == v.FAIL, msg

    def test_without_md_path_falls_back_to_count(self, tmp_path):
        pytest.importorskip("PIL", reason="add_picture 需要 Pillow")
        a = _img_paths()[0]
        f = self._docx_with_images([a], tmp_path / "d.docx")
        status, msg = v.check_image_embedding(f, "![图一](a.png)")
        assert status == v.PASS
        assert "仅数量" in msg

    def test_sha256_file_helper(self):
        a = _img_paths()[0]
        assert v._sha256_file(a) == hashlib.sha256(open(a, "rb").read()).hexdigest()


class TestBaselinePageDiff:
    def test_identical_image_has_zero_diff(self):
        pytest.importorskip("pymupdf")
        p = os.path.join(KIT, "baselines", "p001.png")
        assert v.baseline_page_diff(p, p) == 0.0

    def test_different_pages_report_drift(self):
        # 基线里的第 1 页 vs 第 2 页：必须判为漂移（> 0.1% 阈值）
        pytest.importorskip("pymupdf")
        p1 = os.path.join(KIT, "baselines", "p001.png")
        p2 = os.path.join(KIT, "baselines", "p002.png")
        assert v.baseline_page_diff(p1, p2) > v.DEFAULT_MAX_DIFF
