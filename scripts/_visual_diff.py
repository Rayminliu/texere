"""视觉差异能力的纯逻辑核（可导入、无 CLI 副作用）。

从 validate.py / snapshot.py 外提：像素差异比例 diff_ratio、单页比对
baseline_page_diff、抽样页号 sample_page_indices，以及把 check_visual_drift
内核收拢为 compute_visual_diff（返回统一 CheckResult）。

- validate.py 作为 thin CLI 壳 re-export 这些符号（并保留 check_visual_drift
  别名），使 `python scripts/validate.py` 的 stdout / exit-code 逐字不变；
  原 `_import_diff_ratio` 的 sys.path hack 在此一并消灭。
- snapshot.py 改 `from _visual_diff import diff_ratio`，单一实现不再漂移。

裸名 sibling import，禁止反向 import validate / snapshot。
"""

import glob
import operator
import os

try:
    import pymupdf  # 可选：视觉比对才需要，缺时相应检查降级为 SKIP/FAIL
except ImportError:
    pymupdf = None

from _shared import BASELINE_DPI, DEFAULT_MAX_DIFF
from _verify import ERROR, FAIL, PASS, SKIP, CheckResult


def diff_ratio(a, b):
    """逐字节差异比例；长度不同视为全变。

    分块短路：页面像素绝大部分块是相同的（真实漂移 <1%），先按 64K 块
    比较，只有不等的块才逐字节数差异——语义与全量 sum(map(ne)) 一致，
    但整页 PNG 比对快一个数量级。
    """
    if len(a) != len(b):
        return 1.0
    if a == b:
        return 0.0
    n = len(a)
    step = 65536
    diff = 0
    for i in range(0, n, step):
        ca, cb = a[i : i + step], b[i : i + step]
        if ca != cb:
            diff += sum(map(operator.ne, ca, cb))
    return diff / n


def baseline_page_diff(png_a: str, png_b: str) -> float:
    """解码两张 PNG 并逐像素比较，返回差异比例 0..1；尺寸不同视为 1.0。"""
    if pymupdf is None:
        raise RuntimeError("需要 PyMuPDF：pip install PyMuPDF")
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


def compute_visual_diff(
    pdf_path,
    baseline_dir: str = None,
    max_diff: float = DEFAULT_MAX_DIFF,
    sample: bool = False,
) -> CheckResult:
    """与基线比对视觉漂移 (基于共享导出的 PDF，口径同 snapshot.py: 逐页全量)。"""
    if not baseline_dir or not os.path.exists(baseline_dir):
        return CheckResult(
            "visual_drift",
            SKIP,
            "跳过 (未提供基线目录；如需版面漂移防护：先 python scripts/snapshot.py <pdf> --update 录基线，"
            "再用 --baseline <目录> 或 profile 的 baseline_dir 指定)",
        )
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
