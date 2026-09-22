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
import glob
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass, field
from datetime import datetime

try:
    import pymupdf  # PyMuPDF：包结构/分节等基础检查不需要它，页码/空白页/视觉比对才需要
except ImportError:
    pymupdf = None
from docx import Document
from docx.oxml.ns import qn

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
# 检查项实现 —— 统一返回 (status, message)，status ∈ PASS / FAIL / SKIP / ERROR
# =============================================================================

PASS, FAIL, SKIP, ERROR = "PASS", "FAIL", "SKIP", "ERROR"


@dataclass
class CheckResult:
    """统一的检查结论对象（Layer 0 integrity 地基）。

    取代各处散落的 `(status, message)` / `(False, ...)` 元组，让 profile enforcement、
    provenance、CLI report 都依赖同一个类型，而不是各自拼字符串。
    evidence 装机器可读的判定依据（如命中/缺失段数、图片 sha 列表），report.json
    一并落盘，便于将来做可解释审计；现在先预留，不强制每个 check 都填。
    """

    name: str
    status: str
    message: str
    evidence: dict = field(default_factory=dict)


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


# ---------------------------------------------------------------- 源内容比对


def _normalize(text: str) -> str:
    """归一化：去掉 Markdown 记号与全部空白，用于「正文是否同源」比对。

    去掉空白是因为 Word 会在中英文交界、断行处插入不可见字符，逐字符等价
    在这里不成立；比对的是「去掉排版噪声后的可见文字序列」。
    """
    # pandoc/Word 会把直引号排成弯引号，正文比对前先统一回直引号
    t = text.translate(str.maketrans({"\u201c": '"', "\u201d": '"', "\u2018": "'", "\u2019": "'"}))
    t = re.sub(r"!\[[^\]]*\]\([^)]*\)", "", t)  # 图片：连 alt 一起丢
    t = re.sub(r"\[([^\]]*)\]\([^)]*\)", r"\1", t)  # 链接：只留文字
    t = re.sub(r"[*_`>#|~]+", "", t)  # 强调 / 标题 / 引用 / 表格竖线
    return re.sub(r"\s+", "", t)


def _source_md_segments(md_text: str, min_len: int = 8) -> list[str]:
    """抽出值得比对正文的 Markdown 片段。

    跳过的都不是正文，留着只会误报（真实样例上这几条曾贡献 7 处假 FAIL）：
      - ``` 代码块整体
      - ::: / :::: pandoc fenced div 的栅栏行
      - |:---|---| 表格分隔行
      - 列表符号 `- ` / `* ` / `1. `（Word 里没有这个字符）
      - 行尾硬换行的 `\\`
      - 归一化后过短的片段（< 8 字符）
    """
    segs, in_fence = [], False
    for line in md_text.splitlines():
        if line.lstrip().startswith("```"):
            in_fence = not in_fence
            continue
        if in_fence:
            continue
        stripped = line.strip()
        if not stripped or re.fullmatch(r"[-=:|\s]+", stripped):
            continue
        if stripped.startswith(":::"):  # pandoc fenced div
            continue
        stripped = re.sub(r"^([-*+]|\d+[.)])\s+", "", stripped)  # 列表符号
        stripped = stripped.rstrip("\\").strip()  # Markdown 硬换行
        norm = _normalize(stripped)
        if len(norm) >= min_len:
            segs.append(norm)
    return segs


def _docx_text(doc) -> str:
    """docx 的可见文字：正文段落 + 表格单元格（页眉页脚不算正文）。"""
    parts = [p.text for p in doc.paragraphs]
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                parts.append(cell.text)
    return "\n".join(parts)


def check_source_content_integrity(
    docx_path: str, expected_hash: str = None, source_md: str = None
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
            doc_text = _normalize(_docx_text(Document(docx_path)))
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
        sha256 = hashlib.sha256()
        with open(docx_path, "rb") as f:
            for chunk in iter(lambda: f.read(8192), b""):
                sha256.update(chunk)

        actual = sha256.hexdigest()
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


# `![alt](path)` —— path 后面可能带 pandoc 属性段 `![](a.png){width=3cm}`
MD_IMAGE_REF = re.compile(r"!\[[^\]]*\]\(([^)]+)\)")


def _md_image_paths(md_text: str, md_dir: str = None) -> list[str | None]:
    """按文档顺序解析 `![](path)` 指向的真实文件；解析不到的位置留 None。

    解析顺序：md 所在目录 → 当前目录 → 原样（绝对路径）。
    解析不到不判失败，只在消息里说明「N 张无法定位」——路径规则属于
    render.py 的 resource_paths，这里不该重复实现一套。
    """
    found = []
    for m in MD_IMAGE_REF.finditer(md_text):
        raw = m.group(1).strip()
        raw = re.split(r"\s+[{]", raw)[0].strip().strip("<>").strip()
        raw = raw.split(" ")[0]  # `path "title"` 形式
        cand = raw
        if not os.path.isabs(cand):
            for base in (md_dir, os.getcwd()):
                if not base:
                    continue
                p = os.path.join(base, cand)
                if os.path.exists(p):
                    cand = p
                    break
        found.append(cand if os.path.exists(cand) else None)
    return found


def _sha256_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


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


def _is_subsequence(needle: list[str], hay: list[str]) -> bool:
    """needle 是否按顺序出现在 hay 中（允许 hay 里夹着额外图片，如模板 logo）。"""
    it = iter(hay)
    return all(any(x == y for y in it) for x in needle)


def check_image_embedding(
    docx_path: str, md_ref_text: str = None, md_path: str = None
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
        doc = Document(docx_path)
        doc_shas = _docx_image_shas(doc)
        n_img = len(doc_shas) or len(doc.inline_shapes)

        if not md_ref_text:
            return CheckResult(
                "image_embedding", SKIP, f"跳过 (未提供 Markdown 引用；docx 共 {n_img} 张图)"
            )

        n_ref = len(re.findall(r"!\[", md_ref_text))
        if n_img < n_ref:
            return CheckResult("image_embedding", FAIL, f"图片缺失：引用{n_ref}张，只嵌入{n_img}张")

        if not md_path:
            return CheckResult(
                "image_embedding",
                PASS,
                f"图片嵌入：{n_img}/{n_ref} ok (仅数量，不校验对应关系)",
            )

        # ---- 逐图身份比对 ----
        md_dir = os.path.dirname(os.path.abspath(md_path))
        paths = _md_image_paths(md_ref_text, md_dir)
        resolved = [p for p in paths if p]
        unresolved = len(paths) - len(resolved)

        expected = [_sha256_file(p) for p in resolved]
        missing = [resolved[i] for i, s in enumerate(expected) if s not in doc_shas]
        if missing:
            sample = " / ".join(os.path.basename(m) for m in missing[:3])
            return CheckResult(
                "image_embedding",
                FAIL,
                f"图片对应错误：{len(missing)}/{len(resolved)} 张引用的图没出现在 docx 里"
                f"（例：{sample}）",
            )

        # 源图必须按顺序出现；docx 里允许夹带模板 logo 等额外图片
        if not _is_subsequence(expected, doc_shas):
            return CheckResult(
                "image_embedding",
                FAIL,
                f"图片顺序不一致：{len(resolved)} 张引用的图与 docx 出现次序不同",
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
            )
        return CheckResult(
            "image_embedding",
            PASS,
            f"图片逐图比对：{len(resolved)}/{len(resolved)} 张身份与顺序一致{extra}",
        )
    except PermissionError as e:
        return CheckResult("image_embedding", FAIL, f"文件权限不足：{e}")
    except FileNotFoundError as e:
        return CheckResult("image_embedding", FAIL, f"文件不存在：{e}")
    except Exception as e:
        return CheckResult("image_embedding", ERROR, f"未知错误：{type(e).__name__} - {e}")


def check_section_count(docx_path: str) -> CheckResult:
    """检查分节数合理性。"""
    try:
        doc = Document(docx_path)
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


def count_toc_fields(doc) -> int:
    """数正文里真实的 TOC 域：w:instrText 文本或 w:fldSimple 的 w:instr。

    python-docx 没有 tables_of_contents 属性（1.2.0 实测 hasattr 为 False），
    旧实现靠 hasattr 短路，于是这一项恒定「跳过」——9 项里有 1 项是死的。
    这里下到 OOXML，与 post.py 注入 TOC 的写法（add_field → w:instrText）对齐。
    """
    from docx.oxml.ns import qn

    body = doc.element.body
    n = 0
    for el in body.iter(qn("w:instrText")):
        if re.search(r"(?i)\bTOC\b", el.text or ""):
            n += 1
    for el in body.iter(qn("w:fldSimple")):
        if re.search(r"(?i)\bTOC\b", el.get(qn("w:instr")) or ""):
            n += 1
    return n


def check_toc_field(docx_path: str) -> CheckResult:
    """检查目录域是否存在。

    文档没有 TOC 域时返回 SKIP 而不是 PASS：不配目录是合法配置（toc:false、
    表单类文档），但那意味着「这项没验证」，不该计进 passed。
    """
    try:
        doc = Document(docx_path)
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


def check_word_acceptance(export_ok: bool, export_err: str) -> CheckResult:
    """真机验收：Word 打开 + 导出 PDF（基于步骤 1 的共享导出结果）。"""
    if export_ok:
        return CheckResult("word_acceptance", PASS, "Word 验收：OK")
    return CheckResult("word_acceptance", FAIL, f"Word 验收失败：{export_err}")


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
            return CheckResult("page_numbering", FAIL, f"页码不连续：缺失{sorted(missing)}")

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
            for page_num in range(len(pdf_doc)):
                # 简单判断：如果文字极少（少于 10 个字符），视为空白页
                if len(pdf_doc[page_num].get_text("text").strip()) < 10:
                    empty_count += 1
            n_total = len(pdf_doc)
        finally:
            pdf_doc.close()

        passed = empty_count <= max_empty
        status = f"空白页：{empty_count}/{n_total} (阈值：{max_empty})"
        return CheckResult("blank_pages", PASS if passed else FAIL, status)
    except Exception as e:
        return CheckResult("blank_pages", ERROR, f"检查失败：{type(e).__name__} - {e}")


# 基线图的录制口径与 snapshot.py 一致：dpi=100、逐字节比对。
# 之前 validate 用 2x 矩阵渲染再抽样比对，与 baselines/ 的尺寸根本对不上，
# 会把「没漂移」判成漂移——dpi 口径必须统一。
#
# 覆盖面同样要统一：snapshot.py 是逐页全量，validate 曾经只比首/中/尾 3 页。
# 272 页的标书第 137 页表格溢出时，抽样的 3 页可能全都干净，于是「视觉漂移」
# 通过了——同一份文档在两套工具里给出两个结论。默认改为全量，
# --sample-visual 才退回抽样（长文档快速预检用）。
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


def sample_page_indices(n: int) -> list[int]:
    """抽样页号 (0 基)：首 / 中 / 尾。只在 --sample-visual 下使用。"""
    idxs = [0]
    if n > 10:
        idxs.append(n // 2)
    if n > 1:
        idxs.append(n - 1)
    return sorted({i for i in idxs if 0 <= i < n})


def check_visual_drift(
    pdf_path,
    baseline_dir: str = None,
    max_diff: float = DEFAULT_MAX_DIFF,
    sample: bool = False,
) -> CheckResult:
    """与基线比对视觉漂移 (基于共享导出的 PDF，口径同 snapshot.py: 逐页全量)。"""
    if not baseline_dir or not os.path.exists(baseline_dir):
        return CheckResult("visual_drift", SKIP, "跳过 (未提供基线目录)")
    if pdf_path is None:
        return CheckResult("visual_drift", SKIP, "跳过 (Word 导出 PDF 失败，无 PDF 可比)")
    if pymupdf is None:
        return CheckResult(
            "visual_drift", FAIL, "视觉基线：无法比对 (未安装 PyMuPDF：pip install PyMuPDF)"
        )

    baseline_files = sorted(glob.glob(os.path.join(baseline_dir, "p*.png")))
    if not baseline_files:
        return CheckResult("visual_drift", FAIL, f"基线目录无图片：{baseline_dir}")

    try:
        diff_ratio = _import_diff_ratio()
        doc = pymupdf.open(pdf_path)
        drifts = []
        try:
            n = len(doc)
            n_base = len(baseline_files)
            # 双向都要卡：变少是大改，变多同样是版式变了（旧实现只对变少报错）
            if n != n_base:
                return CheckResult(
                    "visual_drift",
                    FAIL,
                    f"页数 {n} != 基线 {n_base} 页，版式可能大改或基线需重录",
                )
            idxs = sample_page_indices(n) if sample else range(n)
            for i in idxs:
                cur = doc[i].get_pixmap(dpi=BASELINE_DPI).samples
                base = pymupdf.Pixmap(baseline_files[i]).samples
                r = diff_ratio(cur, base)
                if r > max_diff:
                    drifts.append((i + 1, r))
        finally:
            doc.close()

        if drifts:
            parts = [f"第{i}页漂移{r * 100:.2f}%" for i, r in drifts[:5]]
            more = "" if len(drifts) <= 5 else f" …共 {len(drifts)} 页"
            return CheckResult("visual_drift", FAIL, "视觉漂移：" + ", ".join(parts) + more)
        scope = "抽样 %d 页" % len(sample_page_indices(n)) if sample else "全量 %d 页" % n
        return CheckResult("visual_drift", PASS, f"视觉基线：一致 ({scope})")
    except ImportError as e:
        return CheckResult("visual_drift", ERROR, f"依赖缺失：{e} (需要 PyMuPDF)")
    except Exception as e:
        return CheckResult("visual_drift", ERROR, f"视觉比对异常：{type(e).__name__} - {e}")


# =============================================================================
# Profile enforcement —— profile 的声明字段编译成可执行断言（Level 1: Structural）
# =============================================================================
#
# 这是「Profile = executable document contract」的第一步：profile 不再是只选 baseline
# 的 metadata，而是机器可执行的文档规范。每条断言以 profile.<field> 命名，独立可追溯。
# 目前只覆盖 Structural 级（页面/边距/字体/标题层级/表格边框/目录域），这些都能用
# python-docx 直接读 OOXML 判定，不依赖 Word；Semantic / Rendered 级后续接入。
# 只在显式 --enforce-profile 时才计入门禁；profile 是按需 opt-in 的 contract。

EMU_PER_CM = 360000.0


def _emu_cm(v):
    return (v or 0) / EMU_PER_CM


def _get_style(doc, name):
    """python-docx 的 Styles 没有 .get，这里做缺失安全的取值。"""
    try:
        return doc.styles[name]
    except KeyError:
        return None


def _east_asia_of_style(style):
    """取 style 的 w:rFonts/@w:eastAsia（中文主字体）。"""
    rpr = style.element.find(qn("w:rPr"))
    if rpr is None:
        return None
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        return None
    return fonts.get(qn("w:eastAsia"))


def _check_page(page, doc) -> CheckResult:
    sec = doc.sections[0]
    w_cm, h_cm = _emu_cm(sec.page_width), _emu_cm(sec.page_height)
    pw, ph = float(page.get("width", 0)), float(page.get("height", 0))
    issues = []
    if pw and abs(w_cm - pw) > 0.1:
        issues.append(f"宽 {w_cm:.2f}≠{pw}")
    if ph and abs(h_cm - ph) > 0.1:
        issues.append(f"高 {h_cm:.2f}≠{ph}")
    for pk, sk in (
        ("margin_top", "top_margin"),
        ("margin_bottom", "bottom_margin"),
        ("margin_left", "left_margin"),
        ("margin_right", "right_margin"),
    ):
        if page.get(pk) is not None:
            actual = _emu_cm(getattr(sec, sk))
            if abs(actual - float(page[pk])) > 0.2:
                issues.append(f"{pk} {actual:.2f}≠{page[pk]}")
    if issues:
        return CheckResult("profile.page", FAIL, "页面/边距不符：" + "，".join(issues))
    return CheckResult("profile.page", PASS, f"页面 {w_cm:.2f}×{h_cm:.2f}cm、边距符合规范")


def _check_body_font(body, doc) -> CheckResult:
    st = _get_style(doc, "Normal")
    if st is None:
        return CheckResult("profile.body_font", SKIP, "跳过（无 Normal 样式）")
    font = st.font
    issues = []
    if body.get("font_latin") and font.name and font.name != body["font_latin"]:
        issues.append(f"西文 {font.name}≠{body['font_latin']}")
    ea = _east_asia_of_style(st)
    if body.get("font_eastAsia") and ea and ea != body["font_eastAsia"]:
        issues.append(f"中文 {ea}≠{body['font_eastAsia']}")
    if body.get("size") and font.size is not None and abs(font.size.pt - float(body["size"])) > 0.5:
        issues.append(f"字号 {font.size.pt}≠{body['size']}")
    if issues:
        return CheckResult("profile.body_font", FAIL, "正文样式：" + "，".join(issues))
    return CheckResult("profile.body_font", PASS, "正文样式（字体/字号）符合规范")


def _check_heading_styles(styles, doc) -> CheckResult:
    issues = []
    for lvl, wname in (("h1", "Heading 1"), ("h2", "Heading 2"), ("h3", "Heading 3")):
        spec = styles.get(lvl)
        if not spec:
            continue
        st = _get_style(doc, wname)
        if st is None:
            issues.append(f"{wname} 样式缺失")
            continue
        font = st.font
        if spec.get("font_eastAsia"):
            ea = _east_asia_of_style(st)
            if ea and ea != spec["font_eastAsia"]:
                issues.append(f"{wname} 中文 {ea}≠{spec['font_eastAsia']}")
        if spec.get("font_latin") and font.name and font.name != spec["font_latin"]:
            issues.append(f"{wname} 西文 {font.name}≠{spec['font_latin']}")
        if spec.get("size") and font.size is not None and abs(font.size.pt - spec["size"]) > 0.5:
            issues.append(f"{wname} 字号 {font.size.pt}≠{spec['size']}")
        if spec.get("bold") is not None and bool(font.bold) != bool(spec["bold"]):
            issues.append(f"{wname} 加粗 {font.bold}≠{spec['bold']}")
    if issues:
        return CheckResult("profile.heading", FAIL, "标题样式：" + "，".join(issues))
    return CheckResult("profile.heading", PASS, "标题层级样式符合规范")


def _check_table_borders(doc) -> CheckResult:
    if not doc.tables:
        return CheckResult("profile.table", SKIP, "跳过（文档无表格，border 断言不适用）")
    bordered = 0
    for tbl in doc.tables:
        borders = tbl._tbl.tblPr.find(qn("w:tblBorders"))
        if borders is not None and any(
            (b.get(qn("w:sz")) and int(b.get(qn("w:sz")) or 0) > 0)
            for b in borders.findall(qn("w:border"))
        ):
            bordered += 1
    if bordered:
        return CheckResult(
            "profile.table", PASS, f"表格边框：{bordered}/{len(doc.tables)} 个表含可见边框"
        )
    return CheckResult("profile.table", FAIL, f"表格边框：{len(doc.tables)} 个表均无可见边框")


def _check_profile_toc(doc) -> CheckResult:
    n = count_toc_fields(doc)
    if n:
        return CheckResult("profile.toc", PASS, f"目录域：{n} 个")
    return CheckResult("profile.toc", FAIL, "profile 要求目录，但文档中未发现 TOC 域")


def compile_profile_checks(profile, docx_path) -> list:
    """把 profile 的声明字段编译成可执行断言（profile.<field>）。

    只覆盖 profile 真正声明了的字段；未声明的字段不凭空编造断言。
    返回 [] 当且仅当 profile 为空（未传 --profile）。
    """
    out = []
    if not profile:
        return out
    try:
        doc = Document(docx_path)
    except Exception as e:
        return [CheckResult("profile.load", ERROR, f"无法打开文档以执行 profile 断言：{e}")]
    page = profile.get("page") or {}
    if page:
        out.append(_check_page(page, doc))
    styles = profile.get("styles") or {}
    if styles.get("body"):
        out.append(_check_body_font(styles["body"], doc))
    if any(styles.get(k) for k in ("h1", "h2", "h3")):
        out.append(_check_heading_styles(styles, doc))
    table = profile.get("table") or {}
    if table.get("border") and table["border"] != "none":
        out.append(_check_table_borders(doc))
    if profile.get("toc"):
        out.append(_check_profile_toc(doc))
    return out


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
    source_md: str = None,
    sample_visual: bool = False,
    enforce_profile: bool = False,
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
        # skipped 独立于 passed：「没检查」不许冒充实测通过。
        "summary": {"total": 0, "passed": 0, "failed": 0, "skipped": 0},
    }

    # 2. Word 只启动一次：导出共享 PDF，后面的页码/空白页/视觉比对/截图都用它
    tmp_dir = tempfile.mkdtemp(prefix="texere_validate_")
    pdf_path = os.path.join(tmp_dir, "verify.pdf")
    try:
        export_ok, export_err = export_pdf_once(docx_path, pdf_path)
        shared_pdf = pdf_path if export_ok else None

        # 3. 执行所有检查
        md_ref_text = None
        if source_md and os.path.exists(source_md):
            md_ref_text = open(source_md, encoding="utf-8-sig").read()

        checks = [
            ("package_integrity", check_package_integrity, (docx_path,)),
            (
                "source_content",
                check_source_content_integrity,
                (docx_path, expected_hash, source_md),
            ),
            ("image_embedding", check_image_embedding, (docx_path, md_ref_text, source_md)),
            ("section_count", check_section_count, (docx_path,)),
            ("toc_field", check_toc_field, (docx_path,)),
            ("page_numbering", check_page_numbering, (shared_pdf,)),
            ("blank_pages", check_blank_pages, (shared_pdf, max_empty)),
            ("word_acceptance", check_word_acceptance, (export_ok, export_err)),
            (
                "visual_drift",
                check_visual_drift,
                (shared_pdf, baseline_dir, DEFAULT_MAX_DIFF, sample_visual),
            ),
        ]

        for name, checker, args in checks:
            try:
                res = checker(*args)
                if not isinstance(res, CheckResult) or res.status not in (
                    PASS,
                    FAIL,
                    SKIP,
                    ERROR,
                ):
                    res = CheckResult(name, ERROR, f"检查返回了非预期结果：{res!r}")
            except Exception as e:
                res = CheckResult(name, ERROR, f"检查异常：{type(e).__name__} - {e}")

            report["checks"][name] = {
                "status": res.status,
                "message": res.message,
                "evidence": res.evidence,
            }
            status = res.status
            report["summary"]["total"] += 1
            if status == PASS:
                report["summary"]["passed"] += 1
            elif status == SKIP:
                report["summary"]["skipped"] += 1
            else:
                # FAIL 与 ERROR 都算不合格：检查崩了不等于文档合格
                report["summary"]["failed"] += 1

        # 3b. profile 契约断言（仅在 --enforce-profile 时计入门禁；profile 即文档规范）
        if enforce_profile:
            for res in compile_profile_checks(profile, docx_path):
                report["checks"][res.name] = {
                    "status": res.status,
                    "message": res.message,
                    "evidence": res.evidence,
                }
                report["summary"]["total"] += 1
                if res.status == PASS:
                    report["summary"]["passed"] += 1
                elif res.status == SKIP:
                    report["summary"]["skipped"] += 1
                else:
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

    # 5. 先落盘 report.json —— 它的 hash 要进证据清单，顺序不能反
    report_path = os.path.join(out_dir, "report.json")
    with open(report_path, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
        f.write("\n")

    # 6. 证据清单（文件名沿用 signature，但它是 checksum manifest，不是密码学签名：
    #    没有密钥，任何人都能重算这些 hash。它证明的是「这份证据记录了哪个产物」，
    #    不是「这份证据没被改过」。真签名需要非对称密钥 + 验签方持公钥。）
    with open(docx_path, "rb") as f:
        doc_hash = hashlib.sha256(f.read()).hexdigest()
    with open(report_path, "rb") as f:
        report_hash = hashlib.sha256(f.read()).hexdigest()

    sig_path = os.path.join(out_dir, "signature")
    with open(sig_path, "w", encoding="utf-8") as f:
        f.write("# texere validation manifest (checksums, not a cryptographic signature)\n")
        f.write("# Generated: %s\n" % report["metadata"]["timestamp"])
        f.write("document_hash: %s\n" % doc_hash)
        f.write("report_hash: %s\n" % report_hash)
        f.write(
            "checks_passed: %d/%d\n" % (report["summary"]["passed"], report["summary"]["total"])
        )
        # 单独记 skipped：证据里也要能看出「9 项里有几项其实没查」
        f.write("checks_skipped: %d\n" % report["summary"]["skipped"])

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
        sys.exit(1)
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
        help="把 profile 里声明的页面/字体/标题/表格/目录编译成硬断言（profile 即文档规范；默认只当 baseline 选择器）",
    )
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
        a.docx,
        a.out,
        profile,
        a.expected_hash,
        a.max_empty,
        baseline_dir,
        a.source_md,
        a.sample_visual,
        a.enforce_profile,
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
