"""verify 能力的纯逻辑核（可导入、无 CLI 副作用）。

从 validate.py 外提：统一结论类型 CheckResult / 状态常量、docx 打开原语
_open_doc、TOC 域计数 count_toc_fields，以及 profile enforcement（把 profile
声明字段编译成可执行断言 profile.<field>）。validate.py 作为 thin CLI 壳
re-export 这些符号，保持 `python scripts/validate.py` 行为与 stdout/exit-code
逐字不变。裸名 sibling import，禁止反向 import validate。
"""

import os
import re
from dataclasses import dataclass, field

from _ooxml import style_fonts
from _shared import FOOTER_PAGE_RES
from docx import Document
from docx.oxml.ns import qn

# =============================================================================
# 统一结论类型与状态 —— PASS / FAIL / SKIP / ERROR
# =============================================================================

PASS, FAIL, SKIP, ERROR = "PASS", "FAIL", "SKIP", "ERROR"

# 九项检查的名字与顺序 —— 名单的单一来源。
# validate.py 的 checks 注册表、scripts/make_hero.py 的 CHECKS、README 双语版与
# docs/showcase.md 里的 report 摘录，全部由
# tests/test_docs_sync.py::test_check_names_have_single_source 对拍到这个名字表。
# 为什么要有这条：Hero 图与 Showcase 是对「不摆拍」的承诺，改名没同步过去就是谎报。
CHECK_NAMES = (
    "package_integrity",
    "source_content",
    "image_embedding",
    "section_count",
    "toc_field",
    "page_numbering",
    "blank_pages",
    "renderer_acceptance",
    "visual_drift",
)


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
    # evidence 是「只观测、不裁决」的机器可读判定依据：profile.page / profile.body_font /
    # profile.heading / profile.table / profile.toc / image_embedding 等检查会填实它，
    # 其余检查留空 {}。它绝不参与 status 判定，缺失 evidence 也绝不改变 verdict——
    # 仅用于可解释审计与将来的 Build Manifest 直接消费。


def _open_doc(docx_path: str, doc=None):
    """共用已解析的 Document，没传才自己解。

    调用方（generate_evidence_package）顶层解析一次注入；预解析失败时传 None，
    这里回落自己解 → 调用方自己的 FileNotFoundError / PermissionError 分支
    能照原样报错，不会把「文档打不开」误报成内容不符。
    """
    return Document(docx_path) if doc is None else doc


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


# =============================================================================
# 源内容 / 图片引用 / 页码 —— 零 I/O 纯文本助手
# =============================================================================


def _normalize(text: str) -> str:
    """归一化：去掉 Markdown 记号与全部空白，用于「正文是否同源」比对。

    去掉空白是因为 Word 会在中英文交界、断行处插入不可见字符，逐字符等价
    在这里不成立；比对的是「去掉排版噪声后的可见文字序列」。
    """
    # pandoc/Word 会把直引号排成弯引号，-- 变 en/em-dash，... 变 ellipsis，正文比对前先统一回 ASCII
    t = text.translate(
        str.maketrans(
            {
                "\u201c": '"',
                "\u201d": '"',
                "\u2018": "'",
                "\u2019": "'",
                "\u2013": "-",
                "\u2014": "-",
                "\u2026": "...",
            }
        )
    )
    # pandoc 的 smart 排版会把源里的 -- 收成 en-dash、--- 收成 em-dash，
    # 于是同一句话两侧分别是「2020--2024」和「2020–2024」。破折号长度不是正文
    # 契约（实测源文件里写 -- 还是—纯属作者习惯），统一压成单个 - 再比。
    t = re.sub(r"-{2,}", "-", t)
    t = re.sub(
        r"!\[[^\]]*\]\([^)]*\)\s*(?:\{[^}]*\})?", "", t
    )  # 图片：连 alt 和尾随的 pandoc 属性 {width=...} 一起丢
    t = re.sub(r"\[([^\]]*)\]\([^)]*\)(?:\s*\{[^}]*\})?", r"\1", t)  # 链接：只留文字，属性同样丢掉
    t = re.sub(r"[*_`>#|~]+", "", t)  # 强调 / 标题 / 引用 / 表格竖线
    return re.sub(r"\s+", "", t)


def _source_md_segments(md_text: str, min_len: int = 8) -> list[str]:
    """抽出值得比对正文的 Markdown 片段。

    跳过的都不是正文，留着只会误报（真实样例上这几条曾贡献 7 处假 FAIL）：
      - ``` / ~~~ 代码块整体（围栏判定与 _compile._outside_code_fences 同一条规则：
        记录开栏字符，只有同样的字符才关闭 —— ``` 包 inside 里的 ~~~ 不会误闭）
      - ::: / :::: pandoc fenced div 的栅栏行
      - |:---|---| 表格分隔行
      - 列表符号 `- ` / `* ` / `1. `（Word 里没有这个字符）
      - 行尾硬换行的 `\\`
      - 归一化后过短的片段（< 8 字符）
    """
    segs, fence = [], None
    for line in md_text.splitlines():
        s = line.lstrip()
        if fence is None and (s.startswith("```") or s.startswith("~~~")):
            fence = s[:3]
            continue
        if fence is not None:
            if s.startswith(fence):
                fence = None
            continue
        stripped = line.strip()
        if not stripped or re.fullmatch(r"[-=:\s]+", stripped):
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


# `![alt](path)` —— path 后面可能带 pandoc 属性段 `![](a.png){width=3cm}`
MD_IMAGE_REF = re.compile(r"!\[[^\]]*\]\(([^)]+)\)")


def _md_image_raw_refs(md_text: str) -> list[str]:
    """按文档顺序取回 `![](path)` 里的原始引用文本（已去掉 title / <>/ 尾随属性）。

    与 _md_image_paths 共用同一套解析规则；路径解析失败（None）的位置对应的
    原始引用从这里按序取回，供 evidence 记录「哪几张图定位不到」。
    """
    refs = []
    for m in MD_IMAGE_REF.finditer(md_text):
        raw = m.group(1).strip()
        raw = re.split(r"\s+[{]", raw)[0].strip().strip("<>").strip()
        raw = raw.split(" ")[0]  # `path "title"` 形式
        refs.append(raw)
    return refs


def _md_image_paths(md_text: str, md_dir: str = None) -> list[str | None]:
    """按文档顺序解析 `![](path)` 指向的真实文件；解析不到的位置留 None。

    解析顺序：md 所在目录 → 当前目录 → 原样（绝对路径）。
    解析不到不判失败，只在消息里说明「N 张无法定位」——路径规则属于
    render.py 的 resource_paths，这里不该重复实现一套。
    """
    found = []
    for raw in _md_image_raw_refs(md_text):
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


def _is_subsequence(needle: list[str], hay: list[str]) -> bool:
    """needle 是否按顺序出现在 hay 中（允许 hay 里夹着额外图片，如模板 logo）。"""
    it = iter(hay)
    return all(any(x == y for y in it) for x in needle)


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
    return style_fonts(style)[0]


def _check_page(page, doc) -> CheckResult:
    sec = doc.sections[0]
    w_cm, h_cm = _emu_cm(sec.page_width), _emu_cm(sec.page_height)
    pw = float(page.get("width", 0) or 0)
    ph = float(page.get("height", 0) or 0)
    # evidence：把「实际观测值」与 profile 期望值并排，便于审计 / Manifest 直接消费。
    # 注意：evidence 只是观测记录，下方 issues / status 的判定逻辑一字未改。
    evidence = {
        "width": {
            "expected_cm": round(pw, 2) if pw else None,
            "actual_cm": round(w_cm, 2),
        },
        "height": {
            "expected_cm": round(ph, 2) if ph else None,
            "actual_cm": round(h_cm, 2),
        },
    }
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
        exp = page.get(pk)
        actual = _emu_cm(getattr(sec, sk))
        evidence[pk] = {
            "expected_cm": round(float(exp), 2) if exp is not None else None,
            "actual_cm": round(actual, 2),
        }
        if exp is not None and abs(actual - float(exp)) > 0.2:
            issues.append(f"{pk} {actual:.2f}≠{exp}")
    if issues:
        return CheckResult(
            "profile.page", FAIL, "页面/边距不符：" + "，".join(issues), evidence=evidence
        )
    return CheckResult(
        "profile.page",
        PASS,
        f"页面 {w_cm:.2f}×{h_cm:.2f}cm、边距符合规范",
        evidence=evidence,
    )


def _check_body_font(body, doc) -> CheckResult:
    st = _get_style(doc, "Normal")
    if st is None:
        return CheckResult("profile.body_font", SKIP, "跳过（无 Normal 样式）")
    font = st.font
    issues = []
    # contract semantics：profile 要求某字段时，「实际缺失」与「值不符」都应判 FAIL。
    # evidence 用 fields 数组把每个声明字段的 expected/actual 并排：actual=None 天然
    # 表达「缺失 ≠ 不符」，审计方无需从中文 message 里反解。
    fields = []
    if body.get("font_latin"):
        actual = font.name
        fields.append(
            {
                "field": "styles.body.font_latin",
                "expected": body["font_latin"],
                "actual": actual,
            }
        )
        if actual is None:
            issues.append(f"西文缺失（要求 {body['font_latin']}）")
        elif actual != body["font_latin"]:
            issues.append(f"西文 {actual}≠{body['font_latin']}")
    ea = _east_asia_of_style(st)
    if body.get("font_eastAsia"):
        fields.append(
            {
                "field": "styles.body.font_eastAsia",
                "expected": body["font_eastAsia"],
                "actual": ea,
            }
        )
        if ea is None:
            issues.append(f"中文缺失（要求 {body['font_eastAsia']}）")
        elif ea != body["font_eastAsia"]:
            issues.append(f"中文 {ea}≠{body['font_eastAsia']}")
    if body.get("size"):
        actual = round(font.size.pt, 2) if font.size is not None else None
        fields.append({"field": "styles.body.size", "expected": body["size"], "actual": actual})
        if font.size is None:
            issues.append(f"字号缺失（要求 {body['size']}）")
        elif abs(font.size.pt - float(body["size"])) > 0.5:
            issues.append(f"字号 {font.size.pt}≠{body['size']}")
    evidence = {"fields": fields, "source": "profile"}
    if issues:
        return CheckResult(
            "profile.body_font", FAIL, "正文样式：" + "，".join(issues), evidence=evidence
        )
    return CheckResult(
        "profile.body_font", PASS, "正文样式（字体/字号）符合规范", evidence=evidence
    )


def _check_heading_styles(styles, doc) -> CheckResult:
    issues = []
    fields = []
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
            fields.append(
                {
                    "field": f"styles.{lvl}.font_eastAsia",
                    "expected": spec["font_eastAsia"],
                    "actual": ea,
                }
            )
            if ea is None:
                issues.append(f"{wname} 中文缺失（要求 {spec['font_eastAsia']}）")
            elif ea != spec["font_eastAsia"]:
                issues.append(f"{wname} 中文 {ea}≠{spec['font_eastAsia']}")
        if spec.get("font_latin"):
            actual = font.name
            fields.append(
                {
                    "field": f"styles.{lvl}.font_latin",
                    "expected": spec["font_latin"],
                    "actual": actual,
                }
            )
            if actual is None:
                issues.append(f"{wname} 西文缺失（要求 {spec['font_latin']}）")
            elif actual != spec["font_latin"]:
                issues.append(f"{wname} 西文 {actual}≠{spec['font_latin']}")
        if spec.get("size"):
            actual = round(font.size.pt, 2) if font.size is not None else None
            fields.append(
                {"field": f"styles.{lvl}.size", "expected": spec["size"], "actual": actual}
            )
            if font.size is None:
                issues.append(f"{wname} 字号缺失（要求 {spec['size']}）")
            elif abs(font.size.pt - float(spec["size"])) > 0.5:
                issues.append(f"{wname} 字号 {font.size.pt}≠{spec['size']}")
        if spec.get("bold") is not None:
            actual = bool(font.bold) if font.bold is not None else None
            fields.append(
                {
                    "field": f"styles.{lvl}.bold",
                    "expected": bool(spec["bold"]),
                    "actual": actual,
                }
            )
            if font.bold is None:
                issues.append(f"{wname} 加粗缺失（要求 {spec['bold']}）")
            elif bool(font.bold) != bool(spec["bold"]):
                issues.append(f"{wname} 加粗 {font.bold}≠{spec['bold']}")
    evidence = {"fields": fields, "source": "profile"}
    if issues:
        return CheckResult(
            "profile.heading", FAIL, "标题样式：" + "，".join(issues), evidence=evidence
        )
    return CheckResult("profile.heading", PASS, "标题层级样式符合规范", evidence=evidence)


def _check_table_borders(doc) -> CheckResult:
    if not doc.tables:
        return CheckResult("profile.table", SKIP, "跳过（文档无表格，border 断言不适用）")
    bordered = 0
    for tbl in doc.tables:
        borders = tbl._tbl.tblPr.find(qn("w:tblBorders"))
        if borders is None:
            continue
        # OOXML 里 w:tblBorders 的子元素是 w:top / w:left / w:bottom / w:right /
        # w:insideH / w:insideV（没有名为 w:border 的元素）。早期实现误用
        # findall(qn("w:border"))，恒返回空 → 所有表都被误判为「无边框」，
        # 于是 render 默认加的全框线被错杀成 FAIL。这里直接遍历 tblBorders 的子元素。
        if any((b.get(qn("w:sz")) and int(b.get(qn("w:sz")) or 0) > 0) for b in borders):
            bordered += 1
    # evidence 记录粗粒度语义：「至少一个表存在可见边框」，不逐边校验颜色/粗细——
    # 用 rule 字段把这条边界写死，防止后人误读成「所有表全部符合边框规范」。
    evidence = {
        "bordered": bordered,
        "total": len(doc.tables),
        "rule": "at_least_one_visible_border",
    }
    if bordered:
        return CheckResult(
            "profile.table",
            PASS,
            f"表格边框：{bordered}/{len(doc.tables)} 个表含可见边框",
            evidence=evidence,
        )
    return CheckResult(
        "profile.table",
        FAIL,
        f"表格边框：{len(doc.tables)} 个表均无可见边框",
        evidence=evidence,
    )


def _check_profile_toc(doc) -> CheckResult:
    n = count_toc_fields(doc)
    evidence = {"has_field": bool(n), "count": n}
    if n:
        return CheckResult("profile.toc", PASS, f"目录域：{n} 个", evidence=evidence)
    return CheckResult(
        "profile.toc", FAIL, "profile 要求目录，但文档中未发现 TOC 域", evidence=evidence
    )


def compile_profile_checks(profile, docx_path, doc=None) -> list:
    """把 profile 的声明字段编译成可执行断言（profile.<field>）。

    只覆盖 profile 真正声明了的字段；未声明的字段不凭空编造断言。
    返回 [] 当且仅当 profile 为空（未传 --profile）。
    """
    out = []
    if not profile:
        return out
    try:
        doc = _open_doc(docx_path, doc)
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
