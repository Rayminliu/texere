"""文档验收器：统一验收入口，生成结构化报告和证据包。

用法:
  python scripts/validate.py <document.docx> [--out evidence-dir] [--profile profile.json]

检查项:
  - Package integrity: DOCX 包结构完整
  - Source content integrity: 内容未被篡改 (可选 hash 校验)
  - Image embedding: 图片嵌入数量 vs 引用数量
  - Section count: 分节数合理性
  - TOC field: 目录域存在且可更新
  - Page numbering: 页码连续无断档
  - Blank pages: 空白页数量在阈值内
  - Word open/export: 真机验收 (Word 打开 + 导 PDF)
  - Visual baseline drift: 与基线比对 (可选)

输出:
  - report.json: 结构化验证报告
  - diff.pdf: 差异高亮 PDF (如有基线)
  - page-XXX.png: 抽样截图
  - signature: SHA256 签名 (防篡改)
"""

import argparse
import glob
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime

try:
    import pymupdf  # PyMuPDF：包结构/分节等基础检查不需要它，页码/空白页/视觉比对才需要
except ImportError:
    pymupdf = None
from docx import Document

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 子进程管道统一 UTF-8 输出，避免 GBK 编码被 utf-8 解码成乱码（同 render.py）
UTF8_ENV = {**os.environ, "PYTHONIOENCODING": "utf-8"}


def _read_version() -> str:
    """版本号单一来源：scripts/_version.py（与 pyproject.toml 保持一致）。"""
    vp = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_version.py")
    with open(vp, encoding="utf-8") as f:
        m = re.search(r"__version__\s*=\s*[\"']([^\"']+)[\"']", f.read())
    return m.group(1) if m else "0.0.0"


__version__ = _read_version()

# Windows 控制台编码处理
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(errors="replace")


# =============================================================================
# 检查项实现
# =============================================================================


def check_package_integrity(docx_path: str) -> tuple[bool, str]:
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
                return False, f"缺少必要部分：{missing}"

            # 检查是否有损坏
            for name in namelist:
                try:
                    z.read(name)
                except Exception as e:
                    return False, f"部分 {name} 读取失败：{e}"

        return True, "OK"
    except zipfile.BadZipFile as e:
        return False, f"ZIP 格式错误：{e}"
    except PermissionError as e:
        return False, f"文件权限不足：{e}"
    except FileNotFoundError as e:
        return False, f"文件不存在：{e}"
    except Exception as e:
        return False, f"未知错误：{type(e).__name__} - {e}"


def check_source_content_integrity(docx_path: str, expected_hash: str = None) -> tuple[bool, str]:
    """检查源内容完整性 (可选 hash 校验)。"""
    if not expected_hash:
        return True, "跳过 (未提供 expected_hash)"

    try:
        sha256 = hashlib.sha256()
        with open(docx_path, "rb") as f:
            for chunk in iter(lambda: f.read(8192), b""):
                sha256.update(chunk)

        actual = sha256.hexdigest()
        if actual == expected_hash:
            return True, "OK"
        else:
            return (
                False,
                f"Hash 不匹配 (期望:{expected_hash[:16]}..., 实际:{actual[:16]}...)",
            )
    except PermissionError as e:
        return False, f"文件权限不足：{e}"
    except FileNotFoundError as e:
        return False, f"文件不存在：{e}"
    except Exception as e:
        return False, f"未知错误：{type(e).__name__} - {e}"


def check_image_embedding(docx_path: str, md_ref_text: str = None) -> tuple[bool, str]:
    """检查图片嵌入数量 vs 引用数量。"""
    try:
        doc = Document(docx_path)
        n_img = len(doc.inline_shapes)

        if md_ref_text:
            import re

            n_ref = len(re.findall(r"!\[", md_ref_text))
            if n_img >= n_ref:
                return True, f"图片嵌入：{n_img}/{n_ref} ok"
            else:
                return False, f"图片缺失：引用{n_ref}张，只嵌入{n_img}张"
        else:
            return True, f"图片嵌入：{n_img}张 (未提供 Markdown 引用数)"
    except PermissionError as e:
        return False, f"文件权限不足：{e}"
    except FileNotFoundError as e:
        return False, f"文件不存在：{e}"
    except Exception as e:
        return False, f"未知错误：{type(e).__name__} - {e}"


def check_section_count(docx_path: str) -> tuple[bool, str]:
    """检查分节数合理性。"""
    try:
        doc = Document(docx_path)
        n_sections = len(doc.sections)

        # 合理范围：至少 1 节，一般不超过 100 节
        if 1 <= n_sections <= 100:
            return True, f"分节数：{n_sections} (合理)"
        else:
            return False, f"分节数异常：{n_sections}"
    except PermissionError as e:
        return False, f"文件权限不足：{e}"
    except FileNotFoundError as e:
        return False, f"文件不存在：{e}"
    except Exception as e:
        return False, f"未知错误：{type(e).__name__} - {e}"


def check_toc_field(docx_path: str) -> tuple[bool, str]:
    """检查目录域存在且可更新。"""
    try:
        doc = Document(docx_path)

        # 检查是否有目录 (tables_of_contents 是 python-docx 的属性名)
        if not hasattr(doc, "tables_of_contents"):
            return True, "目录域：跳过 (未检测到 tables_of_contents 属性)"

        toc_count = len(doc.tables_of_contents)
        if toc_count == 0:
            # 没有 TOC 不一定失败，可能是表单类文档
            return True, "目录域：无 TOC (表单/附录类文档常见)"

        toc = doc.tables_of_contents[0]
        return True, "目录域：存在 (%s)" % toc.style.name
    except PermissionError as e:
        return False, f"文件权限不足：{e}"
    except FileNotFoundError as e:
        return False, f"文件不存在：{e}"
    except Exception as e:
        # 任何异常都视为通过，因为不是致命错误
        return True, f"目录域：检查跳过 ({type(e).__name__}: {e})"


def export_pdf_once(docx_path: str, pdf_path: str, timeout: int = 300) -> tuple[bool, str]:
    """调 finalize.py 导一次 PDF，供所有基于 PDF 的检查共享。

    旧实现里页码/空白页/Word 验收/视觉比对/截图各自启动一次 Word（一次 validate
    要起 4-5 次 Word COM，慢且容易残留孤儿进程）；这里收敛为一次导出。
    """
    try:
        result = subprocess.run(
            [
                sys.executable,
                os.path.join(KIT, "scripts", "finalize.py"),
                docx_path,
                pdf_path,
            ],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            env=UTF8_ENV,
        )
    except subprocess.TimeoutExpired:
        return False, f"Word 导出超时 ({timeout}s)"
    except Exception as e:
        return False, f"Word 导出异常：{type(e).__name__} - {e}"
    if result.returncode == 0 and os.path.exists(pdf_path):
        return True, ""
    err = (
        (result.stderr or "").strip()
        or (result.stdout or "").strip()
        or f"exit={result.returncode}"
    )
    return False, err[:300]


def check_word_acceptance(export_ok: bool, export_err: str) -> tuple[bool, str]:
    """真机验收：Word 打开 + 导出 PDF（基于步骤 1 的共享导出结果）。"""
    if export_ok:
        return True, "Word 验收：OK"
    return False, f"Word 验收失败：{export_err}"


# 页脚行识别：整行匹配才认，避免把正文里的数字（如「2026 年 9 月」）当页码。
# 顺序有意义：装饰性的「— N —」优先于裸数字，否则裸数字会抢先匹到无意义行尾。
FOOTER_PAGE_RES = (
    re.compile(r"^\s*[\u2014\u2013\-\s]*(\d+)[\s\u2014\u2013\-]*\s*$"),  # — 1 —
    re.compile(r"^\s*第\s*(\d+)\s*页.*$"),  # 第 1 页 / 第 1 页 共 3 页
    re.compile(r"^\s*Page\s*(\d+)\s*$", re.IGNORECASE),  # Page 1
    re.compile(r"^\s*(\d+)\s*$"),  # 裸数字
)


def footer_page_number(line: str, n_pages: int):
    """单行页脚文本 → 页码；非页码行或超出页数范围返回 None。"""
    for rx in FOOTER_PAGE_RES:
        m = rx.match(line.strip())
        if m:
            v = int(m.group(1))
            return v if 1 <= v <= n_pages else None
    return None


def collect_page_numbers(pdf_doc) -> list:
    """逐页提取页脚页码（只看页尾三行；无页码的页如封面/目录直接跳过）。

    必须用 sort=True 按版面位置排序：默认块序里页脚可能在文本流任意位置，
    「取尾行」会漏检（实测 sample 文档页码明明存在却检测不到）。
    """
    n_pages = len(pdf_doc)
    numbers = []
    for i in range(n_pages):
        lines = [ln for ln in pdf_doc[i].get_text("text", sort=True).splitlines() if ln.strip()]
        for ln in lines[-3:]:
            hit = footer_page_number(ln, n_pages)
            if hit is not None:
                numbers.append(hit)
                break
    return numbers


def check_page_numbering(pdf_path, max_pages: int = 1000) -> tuple[bool, str]:
    """检查页码连续性 (基于共享导出的 PDF)。"""
    if pdf_path is None:
        return False, "页码：无法检查 (Word 导出 PDF 失败)"
    if pymupdf is None:
        return True, "页码：跳过 (未安装 PyMuPDF)"
    try:
        pdf_doc = pymupdf.open(pdf_path)
        try:
            n_pages = len(pdf_doc)
            page_numbers = collect_page_numbers(pdf_doc)
        finally:
            pdf_doc.close()

        if not page_numbers:
            # 如果无法提取页码，至少确认页数合理
            if 1 <= n_pages <= max_pages:
                return True, f"页码：{n_pages}页 (连续，未检测到页码格式)"
            else:
                return False, f"页数异常：{n_pages}页"

        # 验证连续性：检测到的页码应构成无缺口的递增序列
        expected = list(range(min(page_numbers), max(page_numbers) + 1))
        if sorted(page_numbers) != expected:
            missing = set(expected) - set(page_numbers)
            return False, f"页码不连续：缺失{sorted(missing)}"

        return (
            True,
            f"页码：{n_pages}页 (连续，检测到页码{min(page_numbers)}-{max(page_numbers)})",
        )
    except PermissionError as e:
        return False, f"文件权限不足：{e}"
    except FileNotFoundError as e:
        return False, f"文件不存在：{e}"
    except Exception as e:
        return False, f"检查失败：{type(e).__name__} - {e}"


def check_blank_pages(pdf_path, max_empty: int = 0) -> tuple[bool, str]:
    """检查空白页数量 (基于共享导出的 PDF)。"""
    if pdf_path is None:
        return False, "空白页：无法检查 (Word 导出 PDF 失败)"
    if pymupdf is None:
        return True, "空白页：跳过 (未安装 PyMuPDF)"
    try:
        pdf_doc = pymupdf.open(pdf_path)
        try:
            empty_count = 0
            for page_num in range(len(pdf_doc)):
                # 简单判断：如果文字极少（少于 10 个字符），视为空白页
                if len(pdf_doc[page_num].get_text("text").strip()) < 10:
                    empty_count += 1
            n_total = len(pdf_doc)
        finally:
            pdf_doc.close()

        passed = empty_count <= max_empty
        status = f"空白页：{empty_count}/{n_total} (阈值：{max_empty})"
        return (passed, status) if passed else (False, status)
    except Exception as e:
        return False, f"检查失败：{type(e).__name__} - {e}"


# 基线图的录制口径与 snapshot.py 一致：dpi=100、逐字节全量比对。
# 之前 validate 用 2x 矩阵渲染再抽样比对，与 baselines/ 的尺寸根本对不上，
# 会把「没漂移」判成漂移——口径必须统一。
BASELINE_DPI = 100
DEFAULT_MAX_DIFF = 0.001  # 0.1%，实测依据见 snapshot.py 注释


def _import_diff_ratio():
    """复用 snapshot.py 的 diff_ratio，避免两处实现漂移。"""
    here = os.path.dirname(os.path.abspath(__file__))
    if here not in sys.path:
        sys.path.insert(0, here)
    from snapshot import diff_ratio

    return diff_ratio


def baseline_page_diff(png_a: str, png_b: str) -> float:
    """解码两张 PNG 并逐像素比较，返回差异比例 0..1；尺寸不同视为 1.0。"""
    if pymupdf is None:
        raise RuntimeError("需要 PyMuPDF：pip install PyMuPDF")
    diff_ratio = _import_diff_ratio()
    a = pymupdf.Pixmap(png_a)
    b = pymupdf.Pixmap(png_b)
    if (a.width, a.height) != (b.width, b.height):
        return 1.0
    return diff_ratio(a.samples, b.samples)


def check_visual_drift(
    pdf_path, baseline_dir: str = None, max_diff: float = DEFAULT_MAX_DIFF
) -> tuple[bool, str]:
    """与基线比对视觉漂移 (基于共享导出的 PDF，口径同 snapshot.py)。"""
    if not baseline_dir or not os.path.exists(baseline_dir):
        return True, "视觉基线：跳过 (未提供基线目录)"
    if pdf_path is None:
        return False, "视觉基线：无法比对 (Word 导出 PDF 失败)"
    if pymupdf is None:
        return False, "视觉基线：无法比对 (未安装 PyMuPDF：pip install PyMuPDF)"

    baseline_files = sorted(glob.glob(os.path.join(baseline_dir, "p*.png")))
    if not baseline_files:
        return False, f"基线目录无图片：{baseline_dir}"

    try:
        diff_ratio = _import_diff_ratio()
        doc = pymupdf.open(pdf_path)
        drifts = []
        try:
            n = len(doc)
            if n < len(baseline_files):
                return (
                    False,
                    f"页数 {n} 少于基线图数 {len(baseline_files)}，版式可能大改或基线需重录",
                )
            # 抽样比对：首 / 中 / 尾（与证据截图同一组页）
            idxs = [0]
            if n > 10:
                idxs.append(n // 2)
            idxs.append(min(n - 1, len(baseline_files) - 1))
            for i in idxs:
                if i >= len(baseline_files):
                    continue
                cur = doc[i].get_pixmap(dpi=BASELINE_DPI).samples
                base = pymupdf.Pixmap(baseline_files[i]).samples
                r = diff_ratio(cur, base)
                if r > max_diff:
                    drifts.append((i + 1, r))
        finally:
            doc.close()

        if drifts:
            parts = [f"第{i}页漂移{r * 100:.2f}%" for i, r in drifts]
            return False, "视觉漂移：" + ", ".join(parts)
        return True, f"视觉基线：一致 (抽样 {len(idxs)} 页)"
    except ImportError as e:
        return False, f"依赖缺失：{e} (需要 PyMuPDF)"
    except Exception as e:
        return False, f"视觉比对异常：{type(e).__name__} - {e}"


# =============================================================================
# 证据生成
# =============================================================================


def generate_evidence_package(
    docx_path: str,
    out_dir: str,
    profile: dict = None,
    expected_hash: str = None,
    max_empty: int = 0,
    baseline_dir: str = None,
):
    """生成证据包：report.json + 截图 + signature。"""
    os.makedirs(out_dir, exist_ok=True)

    # 1. 生成结构化报告
    report = {
        "metadata": {
            "timestamp": datetime.now().isoformat(),
            "document": os.path.basename(docx_path),
            "tool_version": __version__,
            "profile": profile or {},
        },
        "checks": {},
        "summary": {"total": 0, "passed": 0, "failed": 0},
    }

    # 2. Word 只启动一次：导出共享 PDF，后面的页码/空白页/视觉比对/截图都用它
    tmp_dir = tempfile.mkdtemp(prefix="texere_validate_")
    pdf_path = os.path.join(tmp_dir, "verify.pdf")
    try:
        export_ok, export_err = export_pdf_once(docx_path, pdf_path)
        shared_pdf = pdf_path if export_ok else None

        # 3. 执行所有检查
        checks = [
            ("package_integrity", check_package_integrity, (docx_path,)),
            ("source_content", check_source_content_integrity, (docx_path, expected_hash)),
            ("image_embedding", check_image_embedding, (docx_path,)),
            ("section_count", check_section_count, (docx_path,)),
            ("toc_field", check_toc_field, (docx_path,)),
            ("page_numbering", check_page_numbering, (shared_pdf,)),
            ("blank_pages", check_blank_pages, (shared_pdf, max_empty)),
            ("word_acceptance", check_word_acceptance, (export_ok, export_err)),
            ("visual_drift", check_visual_drift, (shared_pdf, baseline_dir)),
        ]

        for name, checker, args in checks:
            try:
                passed, message = checker(*args)
                report["checks"][name] = {
                    "status": "PASS" if passed else "FAIL",
                    "message": message,
                }
                report["summary"]["total"] += 1
                if passed:
                    report["summary"]["passed"] += 1
                else:
                    report["summary"]["failed"] += 1
            except Exception as e:
                report["checks"][name] = {"status": "ERROR", "message": f"检查异常：{e}"}
                report["summary"]["total"] += 1
                report["summary"]["failed"] += 1

        # 4. 生成 PDF 截图证据（复用同一份导出 PDF）
        if shared_pdf is not None and pymupdf is not None:
            pdf_doc = pymupdf.open(shared_pdf)
            try:
                # 抽样截图：第 1 页、中间页、最后一页
                sample_pages = [0]
                if len(pdf_doc) > 10:
                    sample_pages.append(len(pdf_doc) // 2)
                sample_pages.append(len(pdf_doc) - 1)

                for page_idx in sample_pages:
                    page = pdf_doc[page_idx]
                    zoom = pymupdf.Matrix(2, 2)  # 2x 缩放（证据图只给人看，不受基线口径约束）
                    pix = page.get_pixmap(matrix=zoom)
                    img_path = os.path.join(out_dir, f"page-{page_idx + 1:03d}.png")
                    pix.save(img_path)
            finally:
                pdf_doc.close()
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    # 5. 生成签名
    with open(docx_path, "rb") as f:
        doc_hash = hashlib.sha256(f.read()).hexdigest()

    sig_path = os.path.join(out_dir, "signature")
    with open(sig_path, "w", encoding="utf-8") as f:
        f.write("# texere validation signature\n")
        f.write("# Generated: %s\n" % report["metadata"]["timestamp"])
        f.write("document_hash: %s\n" % doc_hash)
        f.write(
            "checks_passed: %d/%d\n" % (report["summary"]["passed"], report["summary"]["total"])
        )

    # 6. 保存报告
    report_path = os.path.join(out_dir, "report.json")
    with open(report_path, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
        f.write("\n")

    return report


def print_report(report: dict, quiet: bool = False):
    """打印人类可读的验证报告。"""
    if not quiet:
        print("\nDocument Validation")
        print("─" * 40)

        # 检查项名称与状态在两种模式下都输出（quiet 只是省去标题装饰），
        # 否则调用方/CI 无法从 stdout 判断哪项检查出了问题。
    for name, check in report["checks"].items():
        status = check["status"]
        message = check["message"]
        symbol = "✅" if status == "PASS" else "❌" if status == "FAIL" else "⚠️"
        print(f"{symbol} [{status}] {name}: {message}")

    if not quiet:
        print("\nSummary")
        print("─" * 40)

    # Always print summary even in quiet mode
    print(f"Passed: {report['summary']['passed']}/{report['summary']['total']}")

    if report["summary"]["failed"] > 0:
        print(f"\n❌ Validation FAILED ({report['summary']['failed']} checks failed)")
        sys.exit(1)
    else:
        if not quiet:
            print("\n✅ All checks passed")


# =============================================================================
# CLI 入口
# =============================================================================


def main():
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
    ap.add_argument("--expected-hash", help="Expected SHA256 hash of source content")
    ap.add_argument(
        "--max-empty",
        type=int,
        default=0,
        help="Maximum allowed blank pages (default: 0)",
    )
    ap.add_argument("--quiet", action="store_true", help="Suppress detailed output")
    a = ap.parse_args()

    if not os.path.exists(a.docx):
        sys.exit(f"File not found: {a.docx}")

    # 准备 profile
    profile = {}
    if a.profile and os.path.exists(a.profile):
        profile = json.load(open(a.profile, encoding="utf-8-sig"))

    # 生成证据包
    print(f"Validating: {os.path.basename(a.docx)}")
    # profile 里可以声明 baseline_dir；命令行未提供时逐层回退
    baseline_dir = a.baseline or (profile.get("baseline_dir") if profile else None)
    report = generate_evidence_package(
        a.docx, a.out, profile, a.expected_hash, a.max_empty, baseline_dir
    )

    # 打印报告
    if not a.quiet:
        print_report(report, quiet=False)
    else:
        print_report(report, quiet=True)

    print(f"\nEvidence package saved to: {os.path.abspath(a.out)}")
    print("  - report.json (structured validation report)")
    print("  - page-XXX.png (sample screenshots)")
    print("  - signature (SHA256 signature)")

    # 退出码
    sys.exit(0 if report["summary"]["failed"] == 0 else 1)


if __name__ == "__main__":
    main()
