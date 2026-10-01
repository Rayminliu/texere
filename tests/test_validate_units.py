"""validate.py 纯函数单元测试：页码识别与基线比对，不启动 Word。

这些逻辑此前只在 CLI 级测试里被间接覆盖（且只测了 skip 路径），
这里对核心判定函数做直接断言。check 函数现已返回 CheckResult，测试统一走
`.status` / `.message`，不再按下标 / 元组解包。
"""

import glob
import hashlib
import importlib.util
import json
import os
import sys
import tempfile
import types

import pytest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(KIT, "scripts")

# validate.py 与 snapshot.py 是同目录兄弟模块，spec 加载时保证可 import
sys.path.insert(0, SCRIPTS)
_spec = importlib.util.spec_from_file_location("validate", os.path.join(SCRIPTS, "validate.py"))
v = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(v)


# ------------------------------------- renderer_acceptance 语义（CI 抓到的回归）


class TestRendererAcceptanceSemantics:
    """渲染器不在只能 SKIP，不许冒充「验收失败」；渲染器在但导出失败才是真 FAIL。

    「查不了当 FAIL」与「查不了当 PASS」是同一种失真（validate.py 文档自己立的规矩），
    与缺 PyMuPDF 的 SKIP 同语义。锁 run 36251321724 抓到的真 bug。
    """

    def test_renderer_missing_is_skip_not_fail(self):
        res = v.check_renderer_acceptance(
            False,
            "ModuleNotFoundError: No module named 'pythoncom'",
            "Microsoft Word",
            False,
            "未安装 pywin32（Word 渲染器需要它）",
        )
        assert res.status == v.SKIP

    def test_renderer_present_but_export_failed_is_fail(self):
        res = v.check_renderer_acceptance(False, "COM error", "Word", True, "")
        assert res.status == v.FAIL

    def test_export_ok_is_pass(self):
        res = v.check_renderer_acceptance(True, "", "Microsoft Word", True, "")
        assert res.status == v.PASS

    def test_three_arg_call_keeps_fail_semantics(self):
        # 旧调用方（不传 renderer_ok）→ 可用性未知 → 导出结果定 FAIL，行为不回退
        res = v.check_renderer_acceptance(False, "boom")
        assert res.status == v.FAIL


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

    # --- 多语言容错（词表单一事实源 _shared.FOOTER_PAGE_RES） ---

    def test_page_of_total_format(self):
        # Word 英文页脚域「Page N of M」：取 N，不拿总页数 M 错位
        assert v.footer_page_number("Page 3 of 20", 20) == 3

    def test_fraction_format(self):
        # 整行锁定的「N / M」总页码式页脚（半角/全角斜杠都认）
        assert v.footer_page_number("3 / 20", 20) == 3
        assert v.footer_page_number("3／20", 20) == 3

    def test_roman_abbr_formats(self):
        # 北欧/德语页码缩写：S. 5 / Nr. 5
        assert v.footer_page_number("S. 5", 20) == 5
        assert v.footer_page_number("Nr. 5", 20) == 5

    def test_inline_fraction_in_prose_not_matched(self):
        # 容错模式整行锚定：正文里的「9 / 10」比例不该被当页码
        assert v.footer_page_number("比例 9 / 10 达标", 20) is None

    def test_em_dash_still_wins_over_bare_number(self):
        # 顺序守卫：「— N —」优先于裸数字（历史上顺序有意义）
        assert v.footer_page_number("— 4 —", 20) == 4


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
        assert v.check_page_numbering(None).status == v.SKIP

    def test_blank_pages_without_pdf_is_skip(self):
        assert v.check_blank_pages(None).status == v.SKIP

    def test_image_page_not_blank(self, tmp_path):
        """纯图片页（无文字但有图）不应被判为空白页。"""
        pymupdf = pytest.importorskip("pymupdf")
        pdf_path = str(tmp_path / "imgpage.pdf")
        doc = pymupdf.open()
        page = doc.new_page(width=200, height=200)
        # 插入一个红色小方块（纯图片，无文字）
        pix = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 10, 10))
        pix.set_rect(pix.irect, (255, 0, 0))
        page.insert_image(page.rect, pixmap=pix)
        doc.save(pdf_path)
        doc.close()
        res = v.check_blank_pages(pdf_path, max_empty=0)
        assert res.status != v.FAIL, f"误判空白：{res.message}"

    def test_visual_drift_without_baseline_is_skip(self):
        res = v.check_visual_drift(None, None)
        assert res.status == v.SKIP
        # SKIP 不能只说“未提供基线”，要指路怎么开启这维保护（但绝不自动建基线）
        assert "snapshot" in res.message

    def test_source_content_without_any_evidence_is_skip(self, tmp_path):
        assert v.check_source_content_integrity(_save_docx(tmp_path)).status == v.SKIP

    def test_image_embedding_without_reference_is_skip(self, tmp_path):
        assert v.check_image_embedding(_save_docx(tmp_path)).status == v.SKIP

    def test_toc_field_without_toc_is_skip(self, tmp_path):
        assert v.check_toc_field(_save_docx(tmp_path)).status == v.SKIP


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
        assert v.check_toc_field(f).status == v.PASS

    def test_plain_document_is_skipped_not_passed(self, tmp_path):
        res = v.check_toc_field(_save_docx(tmp_path, "plain.docx"))
        assert res.status == v.SKIP


# ------------------------------------------------------------- 源内容契约


def _docx_with_text(tmp_path, text, name="b.docx"):
    p = tmp_path / name
    doc = Document()
    doc.add_paragraph(text)
    doc.save(str(p))
    return str(p)


class TestMarkdownNormalization:
    """这些都是真实样例上踩出来的误报（曾一次报出 7 处假 FAIL），
    改匹配器时必须一并守住，否则 --source-md 会变成噪音源。"""

    def test_curly_and_straight_quotes_match(self):
        assert v._normalize("采用“感知—平台—应用”架构") == v._normalize('采用"感知—平台—应用"架构')

    def test_smart_dashes_and_ellipsis_match_their_ascii_source(self):
        # pandoc smart：源里的 -- / --- / ... 到 docx 变成 – / — / …
        # （实测过：修前这两行都是假 FAIL）
        assert v._normalize("统计区间 2020--2024 年") == v._normalize("统计区间 2020–2024 年")
        assert v._normalize("感知---平台---应用") == v._normalize("感知—平台—应用")
        assert v._normalize("详见附录...") == v._normalize("详见附录…")

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

    def test_tilde_fence_content_is_skipped(self):
        """~~~ 与 ``` 都是围栏。只认 ``` 会把代码当正文比对 → 假 FAIL。"""
        assert v._source_md_segments("~~~\n这段代码不该参与比对\n~~~\n") == []

    def test_closing_marker_must_match_the_opening_one(self):
        """``` 块里的 ~~~ 不是结束符 —— 判定规则必须与 _compile 同源。"""
        md = "```\n~~~\n这段代码不该参与比对\n```\n收尾正文足够长了呀\n"
        segs = v._source_md_segments(md)
        assert [s for s in segs if "代码" in s] == []
        assert any("收尾正文" in s for s in segs), segs


class TestSourceContentIntegrity:
    def test_md_equivalence_passes(self, tmp_path):
        text = "本项目的整体架构说明如下所述的完整段落"
        f = _docx_with_text(tmp_path, text)
        md = tmp_path / "a.md"
        md.write_text("# 标题\n\n%s\n" % text, encoding="utf-8")
        res = v.check_source_content_integrity(f, source_md=str(md))
        assert res.status == v.PASS, res.message

    def test_md_equivalence_detects_dropped_text(self, tmp_path):
        f = _docx_with_text(tmp_path, "文档里只有这一句话")
        md = tmp_path / "a.md"
        md.write_text("源文档里这一段根本没有进到 docx 里\n", encoding="utf-8")
        res = v.check_source_content_integrity(f, source_md=str(md))
        assert res.status == v.FAIL

    def test_code_fence_is_not_checked_as_body(self, tmp_path):
        """代码块里的内容不该被当成正文去比对。"""
        f = _docx_with_text(tmp_path, "正文段落内容足够长度")
        md = tmp_path / "a.md"
        md.write_text(
            "正文段落内容足够长度\n\n```python\n这段代码不该出现在docx里\n```\n", encoding="utf-8"
        )
        res = v.check_source_content_integrity(f, source_md=str(md))
        assert res.status == v.PASS, res.message

    def test_tilde_fence_is_not_checked_as_body(self, tmp_path):
        """T1.4 回归：围栏换成 ~~~ 时结论必须一致，不能凭空判源内容缺失。"""
        f = _docx_with_text(tmp_path, "正文段落内容足够长度")
        md = tmp_path / "a.md"
        md.write_text(
            "正文段落内容足够长度\n\n~~~\n这段代码不该出现在docx里\n~~~\n", encoding="utf-8"
        )
        res = v.check_source_content_integrity(f, source_md=str(md))
        assert res.status == v.PASS, res.message

    def test_expected_hash_matches(self, tmp_path):
        f = _save_docx(tmp_path, "h.docx")
        h = hashlib.sha256(open(f, "rb").read()).hexdigest()
        res = v.check_source_content_integrity(f, expected_hash=h)
        assert res.status == v.PASS, res.message

    def test_expected_hash_mismatch_fails(self, tmp_path):
        res = v.check_source_content_integrity(
            _save_docx(tmp_path, "h.docx"), expected_hash="deadbeef"
        )
        assert res.status == v.FAIL


# ------------------------------------------------------------- 图片计数


class TestImageEmbedding:
    def test_missing_image_fails(self, tmp_path):
        res = v.check_image_embedding(_save_docx(tmp_path, "img.docx"), "![a](a.png)")
        assert res.status == v.FAIL, res.message

    def test_fence_example_is_not_an_image_reference(self, tmp_path):
        """T1.1 回归：围栏里的 `![...]` 是文档示例不是引图。

        修前 validate 拿原始 md 计数、render 拿 _outside_code_fences(md) 计数，
        同一份文档两个工具结论相反（validate 凭空 FAIL「图片缺失」）。
        """
        md = "正文。\n\n```markdown\n![示例图](assets/none.png)\n```\n"
        assert "![" in md  # bug 的形状：不过围栏就会被数成一次引用
        ref = v._outside_code_fences(md)
        assert "![" not in ref
        res = v.check_image_embedding(_save_docx(tmp_path, "fence.docx"), ref)
        assert res.status != v.FAIL, res.message

    def test_count_is_a_lower_bound_only(self, tmp_path):
        """n_img >= n_ref 只证明数量够，不证明对应关系——这里把这条边界写死。"""
        res = v.check_image_embedding(_save_docx(tmp_path, "img.docx"))
        assert res.status == v.SKIP
        assert "未提供 Markdown 引用" in res.message


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
        res = v.check_image_embedding(f, md.read_text(encoding="utf-8"), str(md))
        assert res.status == v.FAIL, res.message

    def test_matching_images_pass(self, tmp_path):
        pytest.importorskip("PIL", reason="add_picture 需要 Pillow")
        a, b = _img_paths()[:2]
        md = self._md(tmp_path, [a, b])
        f = self._docx_with_images([a, b], tmp_path / "d.docx")
        res = v.check_image_embedding(f, md.read_text(encoding="utf-8"), str(md))
        assert res.status == v.PASS, res.message

    def test_swapped_order_is_detected(self, tmp_path):
        """两张图都在，但次序反了——数量检查对这种情况完全无感。"""
        pytest.importorskip("PIL", reason="add_picture 需要 Pillow")
        a, b = _img_paths()[:2]
        md = self._md(tmp_path, [a, b])
        f = self._docx_with_images([b, a], tmp_path / "d.docx")
        res = v.check_image_embedding(f, md.read_text(encoding="utf-8"), str(md))
        assert res.status == v.FAIL, res.message

    def test_without_md_path_falls_back_to_count(self, tmp_path):
        pytest.importorskip("PIL", reason="add_picture 需要 Pillow")
        a = _img_paths()[0]
        f = self._docx_with_images([a], tmp_path / "d.docx")
        res = v.check_image_embedding(f, "![图一](a.png)")
        assert res.status == v.PASS
        assert "仅数量" in res.message

    def test_sha256_file_helper(self):
        a = _img_paths()[0]
        assert v._sha256_file(a) == hashlib.sha256(open(a, "rb").read()).hexdigest()


class TestEvidenceEnrichment:
    """Evidence Enrichment：填充已有 CheckResult.evidence（不新增 check / 不改 status）。

    守住五件事：profile.page / body_font / heading / table / toc 的 evidence 非空且形状正确，
    且所有 evidence 必须 JSON-serializable（防止 Decimal / Length / EMU / Enum 漏进 report.json）。
    """

    def _results(self, tmp_path, conforming=True):
        prof = _formal_profile()
        docx = _make_profiled_docx(tmp_path, conforming)
        return {r.name: r for r in v.compile_profile_checks(prof, docx)}

    def test_page_evidence_has_expected_and_actual(self, tmp_path):
        ev = self._results(tmp_path)["profile.page"].evidence
        assert ev, "profile.page 的 evidence 不应为空"
        for dim in (
            "width",
            "height",
            "margin_top",
            "margin_bottom",
            "margin_left",
            "margin_right",
        ):
            assert dim in ev, f"page evidence 缺维度 {dim}"
            assert "expected_cm" in ev[dim] and "actual_cm" in ev[dim]

    def test_body_font_evidence_nonempty_and_missing_is_none(self, tmp_path):
        # conforming：PASS 时 evidence 也必须非空
        r_ok = self._results(tmp_path)["profile.body_font"]
        assert r_ok.status == v.PASS
        assert r_ok.evidence["fields"], "PASS 时 body_font evidence 不应为空"
        # nonconforming 文档的 Normal 不设 eastAsia → 逼出「缺失 = actual None」
        prof = _formal_profile()
        docx = _make_profiled_docx(tmp_path, False)
        r = next(x for x in v.compile_profile_checks(prof, docx) if x.name == "profile.body_font")
        assert r.status == v.FAIL
        ea = next(f for f in r.evidence["fields"] if f["field"] == "styles.body.font_eastAsia")
        assert ea["actual"] is None, "缺失字段的 actual 必须是 None（而非静默放过）"

    def test_heading_evidence_contains_level_fields(self, tmp_path):
        r = self._results(tmp_path)["profile.heading"]
        fields = r.evidence["fields"]
        assert fields, "heading evidence 不应为空"
        assert any(f["field"].startswith("styles.h1.") for f in fields)

    def test_table_evidence_records_bordered_total_rule(self, tmp_path):
        ev = self._results(tmp_path)["profile.table"].evidence
        assert "bordered" in ev and "total" in ev and "rule" in ev
        assert ev["rule"] == "at_least_one_visible_border"

    def test_toc_evidence_has_field_and_count(self, tmp_path):
        ev = self._results(tmp_path)["profile.toc"].evidence
        assert "has_field" in ev and "count" in ev
        assert ev["has_field"] is True and ev["count"] >= 1

    def test_all_evidence_is_json_serializable(self, tmp_path):
        # 钉死「evidence 只装可序列化观测值」，防 EMU / Decimal / Enum 漏进 report.json
        import json

        for r in self._results(tmp_path).values():
            assert isinstance(json.dumps(r.evidence, ensure_ascii=False), str)

    def test_image_evidence_reuses_single_scan(self, tmp_path):
        pytest.importorskip("PIL", reason="add_picture 需要 Pillow")
        a, b = _img_paths()[:2]
        md = TestImageIdentity._md(tmp_path, [a, b])
        f = TestImageIdentity._docx_with_images([a, b], tmp_path / "d.docx")
        res = v.check_image_embedding(f, md.read_text(encoding="utf-8"), str(md))
        assert res.status == v.PASS, res.message
        ev = res.evidence
        assert "referenced" in ev and "embedded" in ev and "resolved" in ev
        assert ev["images"], "images 必须记录文档顺序的 sha256"
        assert len(ev["images"]) == ev["embedded"]  # 不二次扫描
        import json

        json.dumps(ev, ensure_ascii=False)  # 必须可序列化


class TestBaselinePageDiff:
    def test_identical_image_has_zero_diff(self):
        pytest.importorskip("pymupdf")
        p = os.path.join(KIT, "baselines", "word", "p001.png")
        assert v.baseline_page_diff(p, p) == 0.0

    def test_different_pages_report_drift(self):
        # 基线里的第 1 页 vs 第 2 页：必须判为漂移（> 0.1% 阈值）
        pytest.importorskip("pymupdf")
        p1 = os.path.join(KIT, "baselines", "word", "p001.png")
        p2 = os.path.join(KIT, "baselines", "word", "p002.png")
        assert v.baseline_page_diff(p1, p2) > v.DEFAULT_MAX_DIFF


# ------------------------------------------------------------- Profile 可执行契约


def _formal_profile():
    with open(os.path.join(KIT, "profiles", "formal-cn-v1.json"), encoding="utf-8-sig") as f:
        return json.load(f)


def _make_profiled_docx(tmp_path, conforming=True):
    """造一个符合 / 不符合 formal-cn-v1.json 的 docx（纯 python-docx，不启 Word）。"""
    p = tmp_path / "prof.docx"
    doc = Document()
    sec = doc.sections[0]
    if conforming:
        sec.page_width = Cm(21)
        sec.page_height = Cm(29.7)
        sec.top_margin = Cm(2.5)
        sec.bottom_margin = Cm(2.5)
        sec.left_margin = Cm(3.17)
        sec.right_margin = Cm(3.17)
        n = doc.styles["Normal"].font
        n.name = "Times New Roman"
        n.size = Pt(12)
        doc.styles["Normal"].element.get_or_add_rPr().get_or_add_rFonts().set(
            qn("w:eastAsia"), "宋体"
        )
        for _wname, _size in (("Heading 1", 16), ("Heading 2", 14), ("Heading 3", 12)):
            _hf = doc.styles[_wname].font
            _hf.name = "Arial"
            _hf.size = Pt(_size)
            _hf.bold = True
            doc.styles[_wname].element.get_or_add_rPr().get_or_add_rFonts().set(
                qn("w:eastAsia"), "黑体"
            )
        # 带可见边框的表格（真实 OOXML：w:tblBorders 的子元素是 w:top/w:left/...，不是 w:border）
        t = doc.add_table(rows=1, cols=2)
        borders = t._tbl.tblPr.makeelement(qn("w:tblBorders"), {})
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            borders.append(
                borders.makeelement(
                    qn("w:" + edge),
                    {
                        qn("w:val"): "single",
                        qn("w:sz"): "4",
                        qn("w:space"): "0",
                        qn("w:color"): "auto",
                    },
                )
            )
        t._tbl.tblPr.append(borders)
        # TOC 域（与 post.py 注入写法一致）
        para = doc.add_paragraph()
        begin = OxmlElement("w:fldChar")
        begin.set(qn("w:fldCharType"), "begin")
        para.add_run()._r.append(begin)
        it = OxmlElement("w:instrText")
        it.set(qn("xml:space"), "preserve")
        it.text = 'TOC \\o "1-2" \\h \\z \\u'
        para.add_run()._r.append(it)
        end = OxmlElement("w:fldChar")
        end.set(qn("w:fldCharType"), "end")
        para.add_run()._r.append(end)
    else:
        sec.page_width = Cm(21.59)
        sec.page_height = Cm(27.94)
        sec.top_margin = Cm(1.0)
        nf = doc.styles["Normal"].font
        nf.name = "Calibri"
        nf.size = Pt(10)
    doc.save(str(p))
    return str(p)


class TestProfileContract:
    """profile 从 metadata 升级为可执行契约：声明字段 → profile.<field> 断言。"""

    def test_empty_profile_yields_no_checks(self, tmp_path):
        assert v.compile_profile_checks({}, _make_profiled_docx(tmp_path, True)) == []

    def test_conforming_docx_passes(self, tmp_path):
        prof = _formal_profile()
        docx = _make_profiled_docx(tmp_path, True)
        results = {r.name: r.status for r in v.compile_profile_checks(prof, docx)}
        for name in (
            "profile.page",
            "profile.body_font",
            "profile.heading",
            "profile.table",
            "profile.toc",
        ):
            assert name in results, "缺失 profile 断言: " + name
        # 规范文档不得有任何 FAIL/ERROR
        assert all(s in (v.PASS, v.SKIP) for s in results.values()), results

    def test_nonconforming_fails(self, tmp_path):
        prof = _formal_profile()
        docx = _make_profiled_docx(tmp_path, False)
        results = {r.name: r.status for r in v.compile_profile_checks(prof, docx)}
        assert results["profile.page"] == v.FAIL
        assert results["profile.body_font"] == v.FAIL

    def test_enforce_profile_feeds_report(self, tmp_path):
        # 契约断言要能进 report["checks"] 并影响 summary（不依赖 Word：用空 baseline 仍会触发
        # Word 导出，故这里只验证编译结果能并入报告结构——直接复用 compile 的结果做等价断言）
        prof = _formal_profile()
        docx = _make_profiled_docx(tmp_path, True)
        checks = v.compile_profile_checks(prof, docx)
        # 模拟 generate_evidence_package 的并入逻辑：每条都该有 name/status
        for r in checks:
            assert r.name.startswith("profile.")
            assert r.status in (v.PASS, v.FAIL, v.SKIP, v.ERROR)

    def test_table_border_detection(self, tmp_path):
        # 回归：旧实现误用 borders.findall(qn("w:border"))，而真实 OOXML 里
        # w:tblBorders 的子元素是 w:top/w:left/w:bottom/w:right/w:insideH/w:insideV，
        # 根本没有 w:border 这个标签 → 恒返回空 → 所有表被错判为「无可见边框」，
        # render 默认加的全框线被错杀成 FAIL。这里用真实结构钉死检测逻辑。

        # 有真实边框的表
        p1 = tmp_path / "b1.docx"
        d1 = Document()
        t = d1.add_table(rows=1, cols=1)
        borders = t._tbl.tblPr.makeelement(qn("w:tblBorders"), {})
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            borders.append(
                borders.makeelement(
                    qn("w:" + edge),
                    {
                        qn("w:val"): "single",
                        qn("w:sz"): "6",
                        qn("w:space"): "0",
                        qn("w:color"): "808080",
                    },
                )
            )
        t._tbl.tblPr.append(borders)
        d1.save(str(p1))
        assert v._check_table_borders(Document(str(p1))).status == v.PASS

        # 无边框的表（python-docx 默认不加 tblBorders）
        p2 = tmp_path / "b2.docx"
        d2 = Document()
        d2.add_table(rows=1, cols=1)
        d2.save(str(p2))
        assert v._check_table_borders(Document(str(p2))).status == v.FAIL

    def test_missing_required_property_fails(self, tmp_path):
        # contract semantics：profile 要求某字段时，实际「缺失」也应 FAIL，
        # 而不是只在「存在但不符」时才 FAIL（否则「要求宋体、实际没设东亚字体」会被放行）。
        prof = {
            "page": {"width": 21, "height": 29.7},
            "styles": {"body": {"font_eastAsia": "宋体"}},
        }
        p = tmp_path / "m.docx"
        d = Document()
        d.styles["Normal"].font.name = "Times New Roman"  # 只设西文，不声明东亚字体
        d.save(str(p))
        body = next(
            r for r in v.compile_profile_checks(prof, str(p)) if r.name == "profile.body_font"
        )
        assert body.status == v.FAIL
        assert "缺失" in body.message

        # 反向：profile 不要求东亚字体时，缺失不应额外 FAIL（只检查声明了的字段）
        prof2 = {
            "page": {"width": 21, "height": 29.7},
            "styles": {"body": {"font_latin": "Times New Roman"}},
        }
        body2 = next(
            r for r in v.compile_profile_checks(prof2, str(p)) if r.name == "profile.body_font"
        )
        assert body2.status == v.PASS


# ------------------------------------------------- 近空白页的有意稀疏豁免


class TestBlankPagesSparseExemption:
    """签字 / 盖章 / 无正文声明属于**有意稀疏**，不计入空页（README 已知边界的表单尾页误报）。

    纯附图页（有图无文）走另一条豁免：判空必须同时满足「文字极少 + 无图」。
    """

    @staticmethod
    def _pdf(tmp_path, lines):
        import pymupdf

        doc = pymupdf.open()
        page = doc.new_page()
        for i, t in enumerate(lines):
            page.insert_text((72, 100 + i * 18), t, fontsize=11, fontname="china-s")
        p = tmp_path / "t.pdf"
        doc.save(str(p))
        return str(p)

    def test_image_only_page_is_not_blank(self, tmp_path):
        """纯附图页（有图无文）不算空白页——标书的附图页 / 案例页很常见。

        判据过去只看文字长度，整页只有图的页被当成空页而假 FAIL。
        """
        import pymupdf

        doc = pymupdf.open()
        page = doc.new_page()
        pix = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 120, 90))
        pix.set_rect(pix.irect, (30, 60, 200))
        page.insert_image(pymupdf.Rect(72, 100, 192, 190), pixmap=pix)
        p = tmp_path / "fig.pdf"
        doc.save(str(p))

        res = v.check_blank_pages(str(p), max_empty=0)
        assert res.status == v.PASS, "有图的页不该算空白页：%s" % res.message

        # 无图又无文才是真空白页（双条件判空，不能反过来把真空页漏掉）
        empty = tmp_path / "empty.pdf"
        d2 = pymupdf.open()
        d2.new_page()
        d2.save(str(empty))
        assert v.check_blank_pages(str(empty), max_empty=0).status == v.FAIL

    def test_signature_page_exempt(self, tmp_path):
        pdf = self._pdf(tmp_path, ["盖章"])
        res = v.check_blank_pages(pdf, max_empty=0)
        assert res.status == v.PASS
        assert "豁免" in res.message

    def test_short_noise_without_keyword_still_counts(self, tmp_path):
        pdf = self._pdf(tmp_path, ["无"])
        res = v.check_blank_pages(pdf, max_empty=0)
        assert res.status == v.FAIL, "无关键字的真实短页仍应算空页"
        assert res.message.startswith("空白页：1/")

    def test_cn_declared_blank_page_exempt(self, tmp_path):
        # 中文声明变体：「此页有意留白」与「本页面故意留白」都豁免
        for line in ("此页有意留白", "本页面故意留白"):
            res = v.check_blank_pages(self._pdf(tmp_path, [line]), max_empty=0)
            assert res.status == v.PASS, line

    def test_en_intentionally_left_blank_variants(self):
        # 英文全称/副词序变体只改词表不走向：全称 35 字 ≥ 10 字阈值，
        # 根本进不了稀疏分支——这里钉 SPARSE_OK_RE 本身的容错（词表在 _shared）
        assert v.SPARSE_OK_RE.search("This page intentionally left blank")
        assert v.SPARSE_OK_RE.search("Page left blank intentionally")
        assert v.SPARSE_OK_RE.search("intentionally left blank")
        assert not v.SPARSE_OK_RE.search("a page with plenty of ordinary body text")

    def test_mixed_document_exempts_only_keyword_pages(self, tmp_path):
        import pymupdf

        doc = pymupdf.open()
        p1 = doc.new_page()
        p1.insert_text((72, 100), "This page has plenty of real body text to count.", fontsize=11)
        p2 = doc.new_page()
        p2.insert_text((72, 100), "盖章", fontsize=11, fontname="china-s")
        doc.new_page()  # 完全空白页
        p = tmp_path / "mix.pdf"
        doc.save(str(p))

        res = v.check_blank_pages(str(p), max_empty=0)
        assert res.status == v.FAIL, "1 空页 > 阈值 0"
        assert res.message.startswith("空白页：1/3")
        assert "豁免有意稀疏页 1" in res.message


class TestSharedDocumentParse:
    """Fix 16：一条验证链里同一个 docx 只解析一次，检查函数共用注入的 Document。

    python-docx 每次 Document(path) 都要重新解 zip + 构 lxml 树，272 页文档上
    这部分开销最大。generate_evidence_package 顶层解一次后注入，这里锁定：
    给了 doc 就绝不再碰磁盘；没给 doc 仍能自己解（回落语义不破）。
    """

    class _NoParse:
        """替身 Document：被调用即失败，用来证明检查函数没再解析文件。"""

        def __call__(self, *a, **kw):
            raise AssertionError("检查函数不应再解析 docx")

    def test_checks_use_injected_doc_without_reparsing(self, tmp_path, monkeypatch):
        f = _save_docx(tmp_path, "shared.docx")
        doc = Document(f)
        monkeypatch.setattr(v, "Document", self._NoParse())

        assert v.check_section_count(f, doc=doc).status == v.PASS
        assert v.check_toc_field(f, doc=doc).status == v.SKIP
        assert v.check_image_embedding(f, doc=doc).status == v.SKIP
        assert v.check_source_content_integrity(f, doc=doc).status == v.SKIP
        assert v.compile_profile_checks({"page": {}}, f, doc=doc) == []

    def test_falls_back_to_own_parse_when_doc_not_given(self, tmp_path):
        """不传 doc 时行为不变：自己解析并按原语义报错。"""
        f = _save_docx(tmp_path, "fallback.docx")
        assert v.check_section_count(f).status == v.PASS
        # 文件不存在 → 该检查报 ERROR（查不了不算合格），而不是静默 PASS
        assert v.check_section_count(str(tmp_path / "nope.docx")).status == v.ERROR


class TestEvidenceInProcess:
    """L2 零 Word：用 conftest 的 _NoopRenderer 进程内跑通整条 evidence 组装。

    此前 report.json 的结构 / 落盘文件 / SKIP 传播只在 test_validate 的
    subprocess + 真 Word 链里验过。这里把「报告骨架 + 计数 + 文件产出 + 反 fail-open」
    下沉到进程内：renderer 不可用→ PDF 派生检查全 SKIP（绝不冒充 PASS），
    纯 docx 结构检查仍真跑并 PASS，report.json / signature 照常落盘。不启动 Word。
    """

    def test_report_structure_and_files_without_word(self, tmp_path, make_docx, noop_renderer):
        docx = make_docx("e.docx", ["正文一段"])
        out = str(tmp_path / "ev")
        report = v.generate_evidence_package(docx, out, renderer=noop_renderer)

        # report.json 落盘 + 证据清单 signature，全程不靠 Word
        assert os.path.exists(os.path.join(out, "report.json"))
        assert os.path.exists(os.path.join(out, "signature"))

        # 9 项检查全部入报告；summary 与逐项一致，failed 恒 0
        assert len(report["checks"]) == 9
        s = report["summary"]
        assert s["total"] == 9 and s["failed"] == 0
        assert s["passed"] + s["skipped"] + s["failed"] == s["total"]

        # 反 fail-open：没有真渲染器时 PDF 派生检查必须 SKIP，不许 PASS/FAIL
        for name in ("page_numbering", "blank_pages", "visual_drift", "renderer_acceptance"):
            assert report["checks"][name]["status"] == v.SKIP, name
        # 纯 docx 结构检查仍真跑并 PASS（不依赖导出 PDF）
        assert report["checks"]["package_integrity"]["status"] == v.PASS
        assert report["checks"]["section_count"]["status"] == v.PASS

        # provenance 记录的是注入的 noop 渲染器身份，而不是假的 Word
        assert report["metadata"]["renderer"]["name"] == "_NoopRenderer"

    def test_screenshot_failure_keeps_the_evidence_package(
        self, tmp_path, make_docx, noop_renderer, monkeypatch
    ):
        """截图环节抛异常不许带走整份证据包（B-4）。

        真实触发路径：page-001.png 是入库资产，被图片查看器占用时 pix.save 抛
        PermissionError，异常穿出 try 会让 report.json / signature 都写不出来——
        验证跑完了却交不出凭据，而且恰好在用户最需要看凭据的失败时刻。
        """
        docx = make_docx("shotfail.docx", ["正文一段"])
        out = str(tmp_path / "ev")

        def boom(pdf, dst):
            raise PermissionError("page-001.png 正被其他程序占用")

        monkeypatch.setattr(v, "_write_screenshots", boom)
        report = v.generate_evidence_package(docx, out, renderer=noop_renderer)

        assert os.path.exists(os.path.join(out, "report.json"))
        assert os.path.exists(os.path.join(out, "signature"))
        # 失败要留痕（不静默吞），但只进 metadata：九项名单有四方对拍，
        # 不能为记一笔截图失败而添第十个名字
        assert "PermissionError" in report["metadata"]["screenshot_error"]
        assert len(report["checks"]) == 9 and report["summary"]["total"] == 9

    def test_evidence_lands_even_when_an_earlier_stage_raises(
        self, tmp_path, make_docx, noop_renderer, monkeypatch
    ):
        """落盘搬进 finally 的兜底：截图之外的异常同样不许吃掉证据。

        同时验证原始异常不被顶掉——finally 里再抛新异常会把真正的原因掩盖，
        所以落盘失败只能告警。
        """
        docx = make_docx("crash.docx", ["正文一段"])
        out = str(tmp_path / "ev")

        def boom(*a, **kw):
            raise RuntimeError("导出阶段炸了")

        monkeypatch.setattr(v, "export_pdf_once", boom)
        with pytest.raises(RuntimeError, match="导出阶段"):
            v.generate_evidence_package(docx, out, renderer=noop_renderer)
        assert os.path.exists(os.path.join(out, "report.json"))
        assert os.path.exists(os.path.join(out, "signature"))

    def test_one_locked_page_does_not_take_down_the_others(self, tmp_path, monkeypatch):
        """单页写盘失败只丢那一页：其余截图照写，失败页码回报上层留痕，句柄仍要关。"""
        saved, closed = [], []

        class FakePix:
            def __init__(self, idx):
                self.idx = idx

            def save(self, path):
                if self.idx == 0:
                    raise PermissionError("首页被图片查看器占用")
                saved.append(os.path.basename(path))

        class FakePage:
            def __init__(self, idx):
                self.idx = idx

            def get_pixmap(self, matrix=None):
                return FakePix(self.idx)

        class FakePdf:
            def __init__(self, n):
                self.n = n

            def __len__(self):
                return self.n

            def __getitem__(self, i):
                return FakePage(i)

            def close(self):
                closed.append(True)

        fake = types.SimpleNamespace(open=lambda p: FakePdf(12), Matrix=lambda *a: None)
        monkeypatch.setattr(v, "pymupdf", fake)

        failed = v._write_screenshots("fake.pdf", str(tmp_path))
        # 12 页 → 首页 + 中间页(6) + 尾页(11)；首页被占，另两页必须仍然落盘
        assert failed == [1]
        assert saved == ["page-007.png", "page-012.png"]
        assert closed == [True]

    def test_cli_line_does_not_leak_machine_paths(
        self, tmp_path, make_docx, noop_renderer, monkeypatch
    ):
        """证据里的 cli 只留文件名：report.json 是对外交付物，不该带用户名与临时目录。

        以前直接 `" ".join(sys.argv)`，真跑一次就把
        `C:\\Users\\<user>\\AppData\\Local\\Temp\\pytest-of-<user>\\...` 写进交付证据。
        """
        docx = make_docx("leak.docx", ["正文一段"])
        out = str(tmp_path / "ev")
        monkeypatch.setattr(
            sys,
            "argv",
            [
                os.path.join(SCRIPTS, "validate.py"),
                docx,
                "--out",
                out,
                "--config",
                os.path.join(tempfile.gettempdir(), "cfg.json"),
                "--enforce-profile=on",
            ],
        )
        report = v.generate_evidence_package(docx, out, renderer=noop_renderer)

        cli = report["metadata"]["cli"]
        assert cli, "cli 不许被剥空"
        for secret in (tempfile.gettempdir(), os.path.expanduser("~")):
            assert secret not in cli, "cli 泄露了本机路径：%s" % cli
        # 选项名与文件名都要保留：它们才回答「用哪些开关跑的哪份文档」
        assert "validate.py" in cli and "leak.docx" in cli and "cfg.json" in cli
        assert "--enforce-profile=on" in cli
