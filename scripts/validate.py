"""文档验收器：统一验收入口，生成结构化报告和证据包。

用法:
  python scripts/validate.py <document.docx> [--out evidence-dir] [--profile profile.json]

检查项:
  - Package integrity: DOCX 包结构完整
  - Source content: 源内容完整性 (--source-md 正文比对 / --expected-hash 产件 hash)
  - Image embedding: 图片嵌入数量 vs 引用数量
  - Section count: 分节数合理性
  - TOC field: 目录域 (w:instrText / w:fldSimple) 是否存在
  - Page numbering: 页码连续无断档
  - Blank pages: 空白页数量在阈值内
  - Word open/export: 真机验收 (Word 打开 + 导 PDF)
  - Visual baseline drift: 与基线逐页比对 (可选)

状态分级 (tests/test_validate.py 守住):
  PASS  检查跑通且判为合格
  FAIL  判为不合格 (退出码 1)
  SKIP  前置条件缺失，根本没检查 —— 不计入 passed
  ERROR 检查自身抛异常 —— 按失败计
「查不了就当通过」是这套验收器此前最大的失真来源，故一律降级为 SKIP/ERROR。

输出:
  - report.json: 结构化验证报告
  - diff.pdf: 差异高亮 PDF (如有基线)
  - page-XXX.png: 抽样截图
  - signature: SHA256 校验清单（checksum manifest，非密码学签名；含 docx 与 report 的 hash）
"""

import argparse
import concurrent.futures
import hashlib
import inspect
import json
import os
import re
import shutil
import sys
import tempfile

try:
    import pymupdf  # PyMuPDF：包结构/分节等基础检查不需要它，页码/空白页/视觉比对才需要
except ImportError:
    pymupdf = None
from _compile import _outside_code_fences
from _evidence import _new_report, _tally, _write_report, _write_signature
from _shared import (
    DEFAULT_MAX_DIFF,
    SPARSE_OK_RE,
    force_utf8_stdio,
)
from _shared import sha256_file as _sha256_file
from _verify import (
    ERROR,
    FAIL,
    PASS,
    SKIP,
    CheckResult,
    _check_table_borders,  # noqa: F401  re-export：test_validate_units 哨兵直访 v._check_table_borders
    _docx_text,
    _is_subsequence,
    _md_image_paths,
    _md_image_raw_refs,
    _normalize,
    _open_doc,
    _source_md_segments,
    collect_page_numbers,
    compile_profile_checks,
    count_toc_fields,
    footer_page_number,  # noqa: F401  re-export：test_validate_units 哨兵直访 v.footer_page_number
)
from _visual_diff import (
    baseline_page_diff,  # noqa: F401  re-export：test_validate_units 哨兵直访 v.baseline_page_diff
    sample_page_indices,  # noqa: F401  re-export：test_validate_units 哨兵直访 v.sample_page_indices
)
from _visual_diff import (
    compute_visual_diff as check_visual_drift,
)
from docx import Document
from renderers import SUPPORTED_RENDERERS, RenderResult, get_renderer

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


# =============================================================================
# 检查项实现 —— CheckResult / 状态常量 / _open_doc / count_toc_fields 见 _verify.py
# （已在文件顶部 re-export，保持 `python scripts/validate.py` 行为逐字不变）
# =============================================================================


def check_package_integrity(docx_path: str) -> CheckResult:
    """检查 DOCX 包结构完整性 (zip 格式 + 必要部分)。"""
    try:
        # 尝试以 zip 方式打开
        import zipfile

        with zipfile.ZipFile(docx_path, "r") as z:
            # 检查必要部分
            namelist = z.namelist()
            required = ["word/document.xml", "word/styles.xml"]
            missing = [r for r in required if r not in namelist]
            if missing:
                return CheckResult("package_integrity", FAIL, f"缺少必要部分：{missing}")

            # 检查是否有损坏
            for name in namelist:
                try:
                    z.read(name)
                except Exception as e:
                    return CheckResult("package_integrity", FAIL, f"部分 {name} 读取失败：{e}")

        return CheckResult("package_integrity", PASS, "OK")
    except zipfile.BadZipFile as e:
        return CheckResult("package_integrity", FAIL, f"ZIP 格式错误：{e}")
    except PermissionError as e:
        return CheckResult("package_integrity", FAIL, f"文件权限不足：{e}")
    except FileNotFoundError as e:
        return CheckResult("package_integrity", FAIL, f"文件不存在：{e}")
    except Exception as e:
        return CheckResult("package_integrity", ERROR, f"未知错误：{type(e).__name__} - {e}")


def _sha256_files_dedup(paths) -> list:
    """批量算 sha256，同一路径只读一次（重复用图在标书里很常见）。"""
    cache = {}
    out = []
    for p in paths:
        if p not in cache:
            cache[p] = _sha256_file(p)
        out.append(cache[p])
    return out


def check_source_content_integrity(
    docx_path: str, expected_hash: str = None, source_md: str = None, doc=None
) -> CheckResult:
    """检查源内容完整性。

    两条路径，强度完全不同，报告里必须说清走的是哪条：
      1. --source-md：正文等价性比对（Markdown 片段是否在 docx 里出现）——真契约。
      2. --expected-hash：产件文件级 SHA256——只能证明「字节没变」，
         证明不了「内容与源一致」，故措辞上是 artifact hash。
      3. 两者都没有：SKIP。以前这一档返回 PASS，是 9 项里最大的水分。
    """
    if source_md:
        if not os.path.exists(source_md):
            return CheckResult("source_content", ERROR, f"源 Markdown 不存在：{source_md}")
        try:
            md_text = open(source_md, encoding="utf-8-sig").read()
            doc_text = _normalize(_docx_text(_open_doc(docx_path, doc)))
        except PermissionError as e:
            return CheckResult("source_content", FAIL, f"文件权限不足：{e}")
        except FileNotFoundError as e:
            return CheckResult("source_content", FAIL, f"文件不存在：{e}")
        except Exception as e:
            return CheckResult("source_content", ERROR, f"未知错误：{type(e).__name__} - {e}")

        segs = _source_md_segments(md_text)
        missing = [s for s in segs if s not in doc_text]
        if missing:
            sample = " / ".join(m[:24] for m in missing[:3])
            return CheckResult(
                "source_content",
                FAIL,
                f"源内容缺失：{len(missing)}/{len(segs)} 段未在 docx 中找到（例：{sample}）",
            )
        return CheckResult("source_content", PASS, f"正文等价性：{len(segs)} 段全部命中 docx")

    if not expected_hash:
        return CheckResult("source_content", SKIP, "跳过 (未提供 --source-md 或 --expected-hash)")

    try:
        actual = _sha256_file(docx_path)
        if actual is None:
            return CheckResult("source_content", FAIL, f"文件不存在：{docx_path}")
        if actual == expected_hash:
            return CheckResult(
                "source_content", PASS, "产件 hash 一致 (仅证明字节未变，非正文等价性)"
            )
        return CheckResult(
            "source_content",
            FAIL,
            f"Hash 不匹配 (期望:{expected_hash[:16]}..., 实际:{actual[:16]}...)",
        )
    except PermissionError as e:
        return CheckResult("source_content", FAIL, f"文件权限不足：{e}")
    except FileNotFoundError as e:
        return CheckResult("source_content", FAIL, f"文件不存在：{e}")
    except Exception as e:
        return CheckResult("source_content", ERROR, f"未知错误：{type(e).__name__} - {e}")


def _docx_image_shas(doc) -> list[str]:
    """按文档顺序取每张嵌入图的 sha256。

    两个关键点：
      1. 走 a:blip 而不是 doc.inline_shapes —— 后者不包含浮动型（anchor）图片，
         浮动图既不被计数也不被发现。
      2. 走文档顺序而不是 media 文件名排序 —— 文件名排序会把 rId12 排在
         rId9 前面（实测踩到），顺序直接反。
    """
    from docx.oxml.ns import qn

    shas, rels = [], doc.part.rels
    for blip in doc.element.body.iter(qn("a:blip")):
        rid = blip.get(qn("r:embed"))
        if not rid or rid not in rels:
            continue
        try:
            shas.append(hashlib.sha256(rels[rid].target_part.blob).hexdigest())
        except Exception:
            continue
    return shas


def check_image_embedding(
    docx_path: str, md_ref_text: str = None, md_path: str = None, doc=None
) -> CheckResult:
    """检查图片：能给源 md 就做逐图身份 + 顺序校验，否则退回数量下限。

    两档强度，报告里必须说清走的是哪档：
      - 有 md_path：按 sha256 逐图比对（pandoc 原样嵌入字节，实测 sha 一致），
        能抓住「图串位 / 错图 / 拿同一张图重复占位」——数量检查对这三类全瞎。
      - 只有 md_ref_text：仍是数量下限 `n_img >= n_ref`。
      - 都没有：SKIP。

    失败闭环（fail-close）只在「身份 / 顺序」这条轴：解析不到源图路径属于验证器能力
    边界（解析规则在 render 的 resource_paths），不能把「部分未校验」算成强 PASS——
    降级为 SKIP。
    """
    try:
        doc = _open_doc(docx_path, doc)
        doc_shas = _docx_image_shas(doc)
        n_img = len(doc_shas) or len(doc.inline_shapes)
        # evidence：复用同一次扫描结果（doc_shas 即文档顺序的 sha256），绝不为了填
        # evidence 再跑一遍解析。只记录观测值，不影响下方 status 判定。
        images = [{"document_index": i, "sha256": s} for i, s in enumerate(doc_shas)]

        if not md_ref_text:
            return CheckResult(
                "image_embedding",
                SKIP,
                f"跳过 (未提供 Markdown 引用；docx 共 {n_img} 张图)",
                evidence={"available": False, "docx_images": n_img, "images": images},
            )

        n_ref = len(re.findall(r"!\[", md_ref_text))
        ev = {
            "referenced": n_ref,
            "embedded": n_img,
            "resolved": None,
            "unresolved": None,
            "unresolved_paths": [],
            "images": images,
        }
        if n_img < n_ref:
            return CheckResult(
                "image_embedding",
                FAIL,
                f"图片缺失：引用{n_ref}张，只嵌入{n_img}张",
                evidence=ev,
            )

        if not md_path:
            return CheckResult(
                "image_embedding",
                PASS,
                f"图片嵌入：{n_img}/{n_ref} ok (仅数量，不校验对应关系)",
                evidence=ev,
            )

        # ---- 逐图身份比对 ----
        md_dir = os.path.dirname(os.path.abspath(md_path))
        paths = _md_image_paths(md_ref_text, md_dir)
        resolved = [p for p in paths if p]
        unresolved = len(paths) - len(resolved)
        ev["resolved"] = len(resolved)
        ev["unresolved"] = unresolved
        raws = _md_image_raw_refs(md_ref_text)  # 与 paths 同序；None 位置的原始引用从这里取
        ev["unresolved_paths"] = [os.path.basename(r) for p, r in zip(paths, raws) if not p]

        expected = _sha256_files_dedup(resolved)
        missing = [resolved[i] for i, s in enumerate(expected) if s not in doc_shas]
        if missing:
            sample = " / ".join(os.path.basename(m) for m in missing[:3])
            ev["missing_sources"] = [os.path.basename(m) for m in missing]
            return CheckResult(
                "image_embedding",
                FAIL,
                f"图片对应错误：{len(missing)}/{len(resolved)} 张引用的图没出现在 docx 里"
                f"（例：{sample}）",
                evidence=ev,
            )

        # 源图必须按顺序出现；docx 里允许夹带模板 logo 等额外图片
        if not _is_subsequence(expected, doc_shas):
            return CheckResult(
                "image_embedding",
                FAIL,
                f"图片顺序不一致：{len(resolved)} 张引用的图与 docx 出现次序不同",
                evidence=ev,
            )

        n_extra = len(doc_shas) - len(expected)
        extra = f"，另有 {n_extra} 张非源引用图（模板 logo 等）" if n_extra > 0 else ""
        if unresolved:
            # 路径解析不到是验证器能力边界，降级为 SKIP，不计入强 PASS
            return CheckResult(
                "image_embedding",
                SKIP,
                f"图片身份校验降级：{len(resolved)}/{len(paths)} 张已定位且身份+顺序一致{extra}；"
                f"{unresolved} 张源图路径未定位，跳过身份校验",
                evidence=ev,
            )
        return CheckResult(
            "image_embedding",
            PASS,
            f"图片逐图比对：{len(resolved)}/{len(resolved)} 张身份与顺序一致{extra}",
            evidence=ev,
        )
    except PermissionError as e:
        return CheckResult("image_embedding", FAIL, f"文件权限不足：{e}")
    except FileNotFoundError as e:
        return CheckResult("image_embedding", FAIL, f"文件不存在：{e}")
    except Exception as e:
        return CheckResult("image_embedding", ERROR, f"未知错误：{type(e).__name__} - {e}")


def check_section_count(docx_path: str, doc=None) -> CheckResult:
    """检查分节数合理性。"""
    try:
        doc = _open_doc(docx_path, doc)
        n_sections = len(doc.sections)

        # 合理范围：至少 1 节，一般不超过 100 节
        if 1 <= n_sections <= 100:
            return CheckResult("section_count", PASS, f"分节数：{n_sections} (合理)")
        return CheckResult("section_count", FAIL, f"分节数异常：{n_sections}")
    except PermissionError as e:
        return CheckResult("section_count", FAIL, f"文件权限不足：{e}")
    except FileNotFoundError as e:
        return CheckResult("section_count", FAIL, f"文件不存在：{e}")
    except Exception as e:
        return CheckResult("section_count", ERROR, f"未知错误：{type(e).__name__} - {e}")


def check_toc_field(docx_path: str, doc=None) -> CheckResult:
    """检查目录域是否存在。

    文档没有 TOC 域时返回 SKIP 而不是 PASS：不配目录是合法配置（toc:false、
    表单类文档），但那意味着「这项没验证」，不该计进 passed。
    """
    try:
        doc = _open_doc(docx_path, doc)
        n = count_toc_fields(doc)
        if n == 0:
            return CheckResult(
                "toc_field", SKIP, "跳过 (文档中没有 TOC 域；未要求目录的文档属正常)"
            )
        return CheckResult("toc_field", PASS, f"目录域：{n} 个 TOC 域")
    except PermissionError as e:
        return CheckResult("toc_field", FAIL, f"文件权限不足：{e}")
    except FileNotFoundError as e:
        return CheckResult("toc_field", FAIL, f"文件不存在：{e}")
    except Exception as e:
        return CheckResult("toc_field", ERROR, f"目录域检查异常：{type(e).__name__} - {e}")


# 超时层级规则：外层预算必须大于内层引擎自己的 watchdog。
# LibreOfficeRenderer.render 的 timeout=600 是真强杀（杀 soffice 进程树），而这里的
# 外层只是「放弃等这个线程」。以前外层写死 300 < 内层 600：外层先弃等并报错称
# 「PDF 导出超时 (300s)」，soffice 却接着跑到 600s 才被强杀 —— 报的时限是假的，
# 内层强杀形同虚设。规则：读得到渲染器的内层 timeout 就取「内层 + 余量」；
# 读不到（Word / WPS 走 COM，render 没有超时形参）才用 300 的兜底预算。
# 守卫：tests/test_renderer_contract.py::test_outer_watchdog_exceeds_inner_engine_timeout
_OUTER_MARGIN = 30
_OUTER_FALLBACK = 300


def _outer_budget(rndr) -> int:
    """本渲染器的外层预算：内层引擎超时 + 余量；无内层超时时用兜底值。"""
    try:
        param = inspect.signature(rndr.render).parameters.get("timeout")
    except (TypeError, ValueError):  # 自定义可调用对象可能没有可读签名
        return _OUTER_FALLBACK
    if param is not None and isinstance(param.default, int):
        return param.default + _OUTER_MARGIN
    return _OUTER_FALLBACK


def export_pdf_once(
    docx_path: str, pdf_path: str, timeout: int = None, renderer=None
) -> tuple[bool, str, "RenderResult"]:
    """用渲染器导一次 PDF，供所有基于 PDF 的检查共享。

    旧实现里页码/空白页/Word 验收/视觉比对/截图各自启动一次 Word（一次 validate
    要起 4-5 次 Word COM，慢且容易残留孤儿进程）；这里收敛为一次导出，
    且渲染器可插拔（word / libreoffice / wps），见 scripts/renderers.py。
    timeout=None 表示按渲染器给预算（见 _outer_budget 的层级规则），不是一律 300。
    返回三元组 (ok, err, result)：result 是渲染器产出的 RenderResult——即使导出
    失败也带 renderer_name / engine_path，供 evidence 记录 provenance。
    """
    rndr = renderer or get_renderer("word")
    if timeout is None:
        timeout = _outer_budget(rndr)
    ex = concurrent.futures.ThreadPoolExecutor(max_workers=1)
    try:
        res = ex.submit(rndr.render, docx_path, pdf_path).result(timeout=timeout)
        ex.shutdown(wait=False)
    except concurrent.futures.TimeoutError:
        # 不等待卡死的 COM 线程——取消排队项后让后台自然结束
        ex.shutdown(wait=False, cancel_futures=True)
        res = RenderResult(
            ok=False,
            pdf=None,
            renderer_name=type(rndr).__name__,
            warnings=[],
            errors=[f"PDF 导出超时 ({timeout}s)"],
        )
        return False, f"PDF 导出超时 ({timeout}s)", res
    except Exception as e:
        ex.shutdown(wait=False, cancel_futures=True)
        res = RenderResult(
            ok=False,
            pdf=None,
            renderer_name=type(rndr).__name__,
            warnings=[],
            errors=[f"{type(e).__name__}: {e}"],
        )
        return False, f"PDF 导出异常：{type(e).__name__} - {e}", res
    if res.ok and os.path.exists(pdf_path):
        return True, "", res
    err = (res.errors[0] if res.errors else "导出失败")[:300]
    return False, err, res


def check_renderer_acceptance(
    export_ok: bool,
    export_err: str,
    renderer_name: str = "Word",
    renderer_ok: bool = True,
    renderer_why: str = "",
) -> CheckResult:
    """真机验收：渲染器打开 + 导出 PDF（基于步骤 1 的共享导出结果）。

    名称随渲染器走（Word / WPS / LibreOffice），不再写死 word_acceptance——
    否则 --renderer wps 时检查名仍叫 word_acceptance 就是语义撒谎。
    渲染器根本不在（托管 CI / 未装 Office）→ SKIP：查不了当 FAIL 与查不了当
    PASS 是同一种失真，与缺 PyMuPDF 的 SKIP 同语义；装了但导出失败才是真 FAIL。
    """
    if not renderer_ok:
        return CheckResult(
            "renderer_acceptance", SKIP, f"跳过 ({renderer_name} 不可用：{renderer_why})"
        )
    if export_ok:
        return CheckResult("renderer_acceptance", PASS, f"{renderer_name} 验收：OK")
    return CheckResult("renderer_acceptance", FAIL, f"{renderer_name} 验收失败：{export_err}")


# 页脚行/稀疏页识别式单一事实源在 _shared（见顶部 import；多语言容错模式在那里）。


def check_page_numbering(pdf_path, max_pages: int = 1000) -> CheckResult:
    """检查页码连续性 (基于共享导出的 PDF)。"""
    if pdf_path is None:
        return CheckResult("page_numbering", SKIP, "跳过 (Word 导出 PDF 失败，无 PDF 可比)")
    if pymupdf is None:
        return CheckResult("page_numbering", SKIP, "跳过 (未安装 PyMuPDF)")
    try:
        pdf_doc = pymupdf.open(pdf_path)
        try:
            n_pages = len(pdf_doc)
            page_numbers = collect_page_numbers(pdf_doc)
        finally:
            pdf_doc.close()

        if not page_numbers:
            # 识别不出页码格式 = 连续性根本没验证。旧实现在这里返回 PASS 且
            # 文案里写「连续」，把一个 fail-open 说成了通过。
            if 1 <= n_pages <= max_pages:
                return CheckResult(
                    "page_numbering",
                    SKIP,
                    f"跳过 (未识别到页脚页码格式，{n_pages}页；连续性未验证)",
                )
            return CheckResult("page_numbering", FAIL, f"页数异常：{n_pages}页")

        # 验证连续性：检测到的页码应构成无缺口的递增序列
        expected = list(range(min(page_numbers), max(page_numbers) + 1))
        if sorted(page_numbers) != expected:
            missing = set(expected) - set(page_numbers)
            dups = [x for x in set(page_numbers) if page_numbers.count(x) > 1]
            parts = []
            if missing:
                parts.append(f"缺失{sorted(missing)}")
            if dups:
                parts.append(f"重复{sorted(dups)}（可能分节重启编号）")
            return CheckResult("page_numbering", FAIL, f"页码不连续：{' | '.join(parts)}")

        return CheckResult(
            "page_numbering",
            PASS,
            f"页码：{n_pages}页 (连续，检测到页码{min(page_numbers)}-{max(page_numbers)})",
        )
    except PermissionError as e:
        return CheckResult("page_numbering", FAIL, f"文件权限不足：{e}")
    except FileNotFoundError as e:
        return CheckResult("page_numbering", FAIL, f"文件不存在：{e}")
    except Exception as e:
        return CheckResult("page_numbering", ERROR, f"检查失败：{type(e).__name__} - {e}")


def check_blank_pages(pdf_path, max_empty: int = 0) -> CheckResult:
    """检查空白页数量 (基于共享导出的 PDF)。"""
    if pdf_path is None:
        return CheckResult("blank_pages", SKIP, "跳过 (Word 导出 PDF 失败，无 PDF 可比)")
    if pymupdf is None:
        return CheckResult("blank_pages", SKIP, "跳过 (未安装 PyMuPDF)")
    try:
        pdf_doc = pymupdf.open(pdf_path)
        try:
            empty_count = 0
            exempt_count = 0
            for page_num in range(len(pdf_doc)):
                # 简单判断：文字极少（少于 10 个字符）且无图片时视为空白页
                txt = pdf_doc[page_num].get_text("text").strip()
                if len(txt) >= 10:
                    continue
                # 有图片的页不算空白（标书附图/案例页）
                if pdf_doc[page_num].get_images(full=True):
                    continue
                # 但签字 / 盖章 / 无正文声明属于**有意稀疏**，不算误报
                if SPARSE_OK_RE.search(txt):
                    exempt_count += 1
                    continue
                empty_count += 1
            n_total = len(pdf_doc)
        finally:
            pdf_doc.close()

        passed = empty_count <= max_empty
        status = f"空白页：{empty_count}/{n_total} (阈值：{max_empty})"
        if exempt_count:
            status += f"，豁免有意稀疏页 {exempt_count}"
        return CheckResult("blank_pages", PASS if passed else FAIL, status)
    except Exception as e:
        return CheckResult("blank_pages", ERROR, f"检查失败：{type(e).__name__} - {e}")


# 口径单一事实源在 _shared：BASELINE_DPI / DEFAULT_MAX_DIFF 与 snapshot.py 共用。
#
# 覆盖面同样要统一：snapshot.py 是逐页全量，validate 曾经只比首/中/尾 3 页。
# 272 页的标书第 137 页表格溢出时，抽样的 3 页可能全都干净，于是「视觉漂移」
# 通过了——同一份文档在两套工具里给出两个结论。默认改为全量，
# --sample-visual 才退回抽样（长文档快速预检用）。


# diff_ratio / baseline_page_diff / sample_page_indices / check_visual_drift 已外提至
# _visual_diff.py（compute_visual_diff 为内核）；validate.py 于顶部 re-export，
# 原 _import_diff_ratio 的 sys.path hack 一并消灭，行为逐字不变。


# =============================================================================
# Profile enforcement —— 已外提至 _verify.py（compile_profile_checks 及全部
# _check_* 断言内核）；validate.py 于顶部 re-export，行为逐字不变。
# =============================================================================


# =============================================================================
# 证据生成
# =============================================================================


# _new_report / _tally / _write_report / _write_signature / build_manifest 已外提至
# _evidence.py（含 pandoc provenance memo _pandoc_version）；validate.py 顶部 re-export，
# report.json / signature 输出逐字节不变。


def _run_checks(report: dict, checks: list):
    """逐项跑检查：单项异常降级为该项 ERROR，绝不拖垮整条验证链。"""
    for name, checker, args, kwargs in checks:
        try:
            res = checker(*args, **kwargs)
            if not isinstance(res, CheckResult) or res.status not in (PASS, FAIL, SKIP, ERROR):
                res = CheckResult(name, ERROR, f"检查返回了非预期结果：{res!r}")
        except Exception as e:
            res = CheckResult(name, ERROR, f"检查异常：{type(e).__name__} - {e}")
        _tally(report, res, name)


def _run_profile_asserts(report: dict, profile: dict, docx_path: str, doc):
    """profile 契约断言（仅在 --enforce-profile 时计入门禁；profile 即文档规范）。

    非法 profile 值（如 width: "abc"）会在编译断言时抛异常——整批包起来降级为
    一条 ERROR。否则异常穿出会连带 report.json / signature 都写不出来，
    整个证据包直接消失。
    """
    try:
        for res in compile_profile_checks(profile, docx_path, doc=doc):
            _tally(report, res)
    except Exception as e:
        _tally(report, CheckResult("profile_assert", ERROR, f"profile 断言执行失败：{e}"))


def _write_screenshots(shared_pdf: str, out_dir: str) -> list:
    """生成 PDF 截图证据（复用同一份导出 PDF）：首页 / 中间页 / 尾页。

    返回写盘失败的页码（1-based）列表，正常时为空。单页失败不许带走其余页：
    Windows 下 page-001.png 是入库资产，被图片查看器 / 资源管理器预览占用时
    pix.save 会抛 PermissionError，以前这会让整批截图连同 report.json 一起消失。
    """
    if shared_pdf is None or pymupdf is None:
        return []
    failed_pages = []
    pdf_doc = None
    try:
        # open 也在 try 内：拿不到句柄同样只是「没截图」，而不是证据包消失
        pdf_doc = pymupdf.open(shared_pdf)
        sample_pages = [0]
        if len(pdf_doc) > 10:
            sample_pages.append(len(pdf_doc) // 2)
        sample_pages.append(len(pdf_doc) - 1)
        for page_idx in sample_pages:
            try:
                page = pdf_doc[page_idx]
                zoom = pymupdf.Matrix(2, 2)  # 2x 缩放（证据图只给人看，不受基线口径约束）
                pix = page.get_pixmap(matrix=zoom)
                pix.save(os.path.join(out_dir, f"page-{page_idx + 1:03d}.png"))
            except Exception:
                failed_pages.append(page_idx + 1)
    finally:
        if pdf_doc is not None:
            pdf_doc.close()
    return failed_pages


def generate_evidence_package(
    docx_path: str,
    out_dir: str,
    profile: dict = None,
    expected_hash: str = None,
    max_empty: int = 0,
    baseline_dir: str = None,
    source_md: str = None,
    sample_visual: bool = False,
    enforce_profile: bool = False,
    renderer=None,
    config_path: str = None,
    reference_doc: str = None,
):
    """生成证据包：report.json + 截图 + signature。

    本函数只做编排：导出共享 PDF → 跑检查 → 落盘报告 → 写证据清单。
    各阶段的具体实现拆在同名小函数里（_new_report / _run_checks /
    _write_screenshots / _write_signature）。
    """
    os.makedirs(out_dir, exist_ok=True)

    # 1. 结构化报告骨架
    report = _new_report(docx_path, profile, config_path, reference_doc)

    # 2. Word 只启动一次：导出共享 PDF，后面的页码/空白页/视觉比对/截图都用它
    tmp_dir = tempfile.mkdtemp(prefix="texere_validate_")
    pdf_path = os.path.join(tmp_dir, "verify.pdf")
    rndr = renderer or get_renderer("word")
    # 渲染器在不在是 acceptance 的前置：不在只能 SKIP，不能冒充「验收失败」。
    # 但 available() 只有真实渲染器有——Fake 等测试注入的适配器没有它，
    # getattr 探测：没有就当「可用性未知」，回退到导出结果本身定 PASS/FAIL
    _avail = getattr(type(rndr), "available", None)
    rndr_ok, rndr_why = _avail() if callable(_avail) else (True, "")
    try:
        export_ok, export_err, render_res = export_pdf_once(docx_path, pdf_path, renderer=rndr)
        renderer_name = render_res.renderer_name
        shared_pdf = pdf_path if export_ok else None
        # renderer 身份进 evidence：别人看到 PASS/FAIL 也能知道是 Word / WPS / LO 出的，
        # 否则 evidence 缺 provenance（评审：Build Manifest / Evidence）。
        report["metadata"]["renderer"] = {
            "name": render_res.renderer_name,
            "version": render_res.renderer_version,
            "engine_path": render_res.engine_path,
        }

        # 3. 执行所有检查
        md_ref_text = None
        if source_md and os.path.exists(source_md):
            with open(source_md, encoding="utf-8-sig") as f:
                # 围栏内的 `![...]` 是示例代码不是引图（与 render.py:546 同一条规则）：
                # 不排掉会把示例算进引用数，凭空判 image_embedding FAIL。
                md_ref_text = _outside_code_fences(f.read())

        # ---- docx 只解一次：共用 Document 实例注入给下游 docx 级检查 ----
        # 272 页文档上 python-docx 的冷启动解压 + 构 lxml 树是大头，
        # 而之前同一个文件在一条验证链里被解 6 次（4 个检查 + profile 断言 + 本函数）。
        doc = None
        try:
            doc = Document(docx_path)
        except Exception:
            # 预解析失败不在此处定级：下游检查拿 doc=None 会自己重解一次，
            # 从而保留各自的 FileNotFoundError / PermissionError 准确措辞。
            pass

        # kwargs 位统一存在（不用则传 {}），避开可变长解包陷阱
        checks = [
            ("package_integrity", check_package_integrity, (docx_path,), {}),
            (
                "source_content",
                check_source_content_integrity,
                (docx_path, expected_hash, source_md),
                {"doc": doc},
            ),
            (
                "image_embedding",
                check_image_embedding,
                (docx_path, md_ref_text, source_md),
                {"doc": doc},
            ),
            ("section_count", check_section_count, (docx_path,), {"doc": doc}),
            ("toc_field", check_toc_field, (docx_path,), {"doc": doc}),
            ("page_numbering", check_page_numbering, (shared_pdf,), {}),
            ("blank_pages", check_blank_pages, (shared_pdf, max_empty), {}),
            (
                "renderer_acceptance",
                check_renderer_acceptance,
                (export_ok, export_err, renderer_name, rndr_ok, rndr_why),
                {},
            ),
            (
                "visual_drift",
                check_visual_drift,
                (shared_pdf, baseline_dir, DEFAULT_MAX_DIFF, sample_visual),
                {},
            ),
        ]

        _run_checks(report, checks)

        # 3b. profile 契约断言（仅在 --enforce-profile 时计入门禁）
        if enforce_profile:
            _run_profile_asserts(report, profile, docx_path, doc)

        # 4. 生成 PDF 截图证据（复用同一份导出 PDF）
        # 截图是证据的辅助件，不许反噬主件：异常穿出会让 report.json / signature
        # 都写不出来（与 _run_profile_asserts 同型的坑）。失败只记进 metadata——
        # 九项检查名单有四方对拍，不能为了记一笔截图失败而添第十个名字。
        try:
            shot_failed = _write_screenshots(shared_pdf, out_dir)
        except Exception as e:
            report["metadata"]["screenshot_error"] = f"{type(e).__name__}: {e}"
        else:
            if shot_failed:
                report["metadata"]["screenshot_pages_failed"] = shot_failed
    finally:
        # 5-6. 先落盘 report.json，再以它为输入写证据清单（顺序不能反：清单要盖
        # report 的 hash）。两者必须在 finally：验证跑完却交不出凭据，恰恰发生在
        # 用户最需要知道「哪里错了」的时刻。落盘自身失败时只告警，不在 finally 里
        # 抛新异常——那会把在飞的原始异常顶掉，真正的原因就查不出来了。
        try:
            report_path = _write_report(report, out_dir)
            _write_signature(report, out_dir, docx_path, report_path)
        except Exception as e:
            print(f"[warn] 证据落盘失败：{type(e).__name__}: {e}", file=sys.stderr)
        shutil.rmtree(tmp_dir, ignore_errors=True)

    return report


def print_report(report: dict, quiet: bool = False):
    """打印人类可读的验证报告。"""
    if not quiet:
        print("\nDocument Validation")
        print("─" * 40)

        # 检查项名称与状态在两种模式下都输出（quiet 只是省去标题装饰），
        # 否则调用方/CI 无法从 stdout 判断哪项检查出了问题。
    symbols = {PASS: "✅", FAIL: "❌", SKIP: "⏭️", ERROR: "⚠️"}
    for name, check in report["checks"].items():
        status = check["status"]
        message = check["message"]
        print(f"{symbols.get(status, '⚠️')} [{status}] {name}: {message}")

    if not quiet:
        print("\nSummary")
        print("─" * 40)

    # Always print summary even in quiet mode
    passed = report["summary"]["passed"]
    total = report["summary"]["total"]
    skipped = report["summary"].get("skipped", 0)
    # 「8/9 通过 + 1 跳过」和「8/9 通过 + 1 失败」在旧口径下长得一样，这里分开写
    print(f"Passed: {passed}/{total}" + (f" (skipped: {skipped})" if skipped else ""))

    if report["summary"]["failed"] > 0:
        print(f"\n❌ Validation FAILED ({report['summary']['failed']} checks failed)")
        # 这里不 sys.exit：退出码统一由 main() 末尾决定。在此之前退出会让
        # 「Evidence package saved to ...」永远打不出来 —— 而失败时恰是用户最
        # 需要知道证据包位置的时刻。
    else:
        if not quiet:
            if skipped:
                # 有跳过时不能只说 "All checks passed"：结论行也要让人一眼看到覆盖缺口
                print(f"\n✅ Required checks passed ({skipped} skipped — see Summary above)")
            else:
                print("\n✅ All checks passed")


# =============================================================================
# CLI 入口
# =============================================================================


def main():
    # 副作用只在入口执行：被 import（工具复用/测试）时不碰宿主 stdio
    force_utf8_stdio()

    ap = argparse.ArgumentParser(
        description="texere document validator — Compiler + Contract + Evidence"
    )
    ap.add_argument("docx", help="Document to validate")
    ap.add_argument(
        "--out",
        default="evidence",
        help="Output directory for evidence package (default: evidence)",
    )
    ap.add_argument("--profile", help="Profile JSON for visual baseline comparison")
    ap.add_argument("--baseline", help="Baseline directory for visual drift comparison")
    ap.add_argument("--expected-hash", help="Expected SHA256 hash of the docx artifact")
    ap.add_argument(
        "--source-md",
        help="Source Markdown (.md or a directory's merged text): enables real body-text "
        "equivalence checking instead of the weaker artifact hash",
    )
    ap.add_argument(
        "--sample-visual",
        action="store_true",
        help="Compare only first/middle/last page against the baseline (default: all pages)",
    )
    ap.add_argument(
        "--enforce-profile",
        action="store_true",
        help="把 profile 里声明的页面/字体/标题/表格/目录编译成硬断言（profile 即文档规范；"
        "默认只当 baseline 选择器，不编译成检查项）",
    )
    ap.add_argument(
        "--max-empty",
        type=int,
        default=0,
        help="Maximum allowed blank pages (default: 0)",
    )
    ap.add_argument("--quiet", action="store_true", help="Suppress detailed output")
    ap.add_argument(
        "--renderer",
        choices=list(SUPPORTED_RENDERERS),
        default="word",
        help="PDF 导出渲染器：word（默认，需本机 Word）/ libreoffice / wps",
    )
    ap.add_argument(
        "--config",
        help="渲染时使用的 config.json——其 sha256 进证据（provenance：哪份契约）",
    )
    ap.add_argument(
        "--reference",
        help="渲染时使用的 reference docx——其 sha256 进证据（provenance：哪个模板）",
    )
    a = ap.parse_args()

    if not os.path.exists(a.docx):
        sys.exit(f"File not found: {a.docx}")

    # 渲染器：解析并探测可用性（不可用也只是让 PDF 相关验收失败，不阻断 docx 检查）
    renderer = get_renderer(a.renderer)
    avail, why = type(renderer).available()
    if not avail:
        print("[warn] 所选渲染器不可用（%s）：%s；PDF 相关验收将失败" % (a.renderer, why))

    # 准备 profile
    profile = {}
    if a.profile and os.path.exists(a.profile):
        with open(a.profile, encoding="utf-8-sig") as f:
            profile = json.load(f)
    elif a.profile:
        sys.exit(f"找不到 --profile 文件：{a.profile}")

    # 生成证据包
    print(f"Validating: {os.path.basename(a.docx)}")
    # profile 里可以声明 baseline_dir；命令行未提供时逐层回退
    baseline_dir = a.baseline or (profile.get("baseline_dir") if profile else None)
    report = generate_evidence_package(
        a.docx,
        a.out,
        profile,
        a.expected_hash,
        a.max_empty,
        baseline_dir,
        a.source_md,
        a.sample_visual,
        a.enforce_profile,
        renderer=renderer,
        config_path=a.config,
        reference_doc=a.reference,
    )

    # 打印报告
    print_report(report, quiet=a.quiet)

    print(f"\nEvidence package saved to: {os.path.abspath(a.out)}")
    print("  - report.json (structured validation report)")
    print("  - page-XXX.png (sample screenshots)")
    print("  - signature (checksum manifest, not a cryptographic signature)")

    # 退出码
    sys.exit(0 if report["summary"]["failed"] == 0 else 1)


if __name__ == "__main__":
    main()
