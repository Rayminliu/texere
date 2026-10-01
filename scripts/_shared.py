"""texere 脚本共享基础设施：版本、控制台编码、SHA256、子进程环境、显示宽度、快照口径常量、
语言词表（中文/英文正式文档措辞默认值的单一事实源）。

各脚本（render / validate / patch / edit / snapshot / post / …）统一从这里引用，
避免逐字复制导致的漂移。

注：语言词表曾计划单独建 `_locale.py`，但 `_locale` 是 CPython 内置模块名，
永远遮蔽脚本目录里的同名文件，故并入本模块。config.json 的覆盖路径
（style.east_font、caption_words、toc_heading 等）行为不变——这里只拥有默认值。
"""

import hashlib
import os
import re
import sys
import unicodedata

from _version import __version__  # noqa: F401 — re-export

# ---------------------------------------------------------------------------
# 控制台编码
# ---------------------------------------------------------------------------
# Windows 控制台默认 GBK，子进程输出里若出现无法编码的字符会抛 UnicodeEncodeError。
# 统一 reconfigure 为 utf-8 + errors="replace"，让不可编码字符降级为 ? 而非崩溃。


def force_utf8_stdio():
    """在脚本入口处调用一次，确保 stdout/stderr 不会因编码问题崩溃。"""
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


# ---------------------------------------------------------------------------
# 子进程 UTF-8 环境
# ---------------------------------------------------------------------------
# 强制子进程管道输出 UTF-8，整条链编码统一。
UTF8_ENV = {**os.environ, "PYTHONIOENCODING": "utf-8"}


# ---------------------------------------------------------------------------
# SHA256 文件哈希
# ---------------------------------------------------------------------------


def sha256_file(path):
    """流式计算文件 SHA256；path 为空或文件不存在时返回 None。"""
    if not path or not os.path.exists(path):
        return None
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


# ---------------------------------------------------------------------------
# 显示宽度
# ---------------------------------------------------------------------------
# East Asian W/F 记 2 列，其余记 1：align_tables 的重排与 render --doctor 的
# 表格对齐共用这一把尺子，避免两套实现算出不同的列宽。


def display_width(s: str) -> int:
    """显示宽度：East Asian W/F 记 2，其余记 1。"""
    return sum(2 if unicodedata.east_asian_width(c) in ("W", "F") else 1 for c in s)


# ---------------------------------------------------------------------------
# 快照口径（snapshot.py 与 validate.py 的视觉漂移必须同一把尺子）
# ---------------------------------------------------------------------------
# 0.1%。实测：同一文档重复导出 PDF 的差异为 0.00%，而改一个页眉文字会产生 0.16%，
# 所以阈值必须压到 0.1% 才能抓住这种"小但真实"的漂移；0.5% 会直接漏报。
# 两处若各定义一份，同一文档会在两套工具里得到相反结论（曾踩）。
BASELINE_DPI = 100
DEFAULT_MAX_DIFF = 0.001


# ---------------------------------------------------------------------------
# 语言词表（中文/英文正式文档的措辞默认值）
# ---------------------------------------------------------------------------
# _shared.DEFAULT_CAPTION_WORDS 与 scripts/filters/captions.lua 的
# DEFAULT_TABLE_WORDS/DEFAULT_FIGURE_WORDS 必须逐词一致，由
# tests/test_docs_sync.py 的守卫钉死（只解析、不执行 lua）。

# 字体对
DEFAULT_EAST_FONT = "宋体"
DEFAULT_LATIN_FONT = "Times New Roman"

# 页码 / 目录文案
PAGE_NUMBER_TEMPLATE = "— {n} —"
TOC_HEADING = "目　　录"
TOC_PLACEHOLDER = "【目录将在打开文档时自动生成；若未显示请全选后按 F9】"

# 样式别名：兼容中文模板（reference_doc 来自中文 Word 时一级标题样式名为「标题 1」）
H1_STYLE_ALIASES = {"Heading 1", "标题 1"}
TITLE_STYLE_ALIASES = {"Title", "Subtitle", "Author", "Date", "标题", "副标题"}

# 题注词表：覆盖 表 1-1 / 表1.1 / 表１－１ / 图 2-3 / Table 1-1 / Figure 1-2 / Fig. 3
# 「Fig.」与「Fig」两个写法都列：Python 侧按词匹配，lua 侧按前缀匹配，
# 各自需要的形式不同但语义集合一致。
DEFAULT_CAPTION_WORDS = {
    "table": ["表", "表格", "Table"],
    "figure": ["图", "圖", "图片", "圖片", "Figure", "Fig.", "Fig"],
}

# 页脚行识别：整行匹配才认，避免把正文里的数字（如「2026 年 9 月」）当页码。
# 顺序有意义：装饰性的「— N —」优先于裸数字，否则裸数字会抢先匹到无意义行尾。
# 多语言容错：Word 页脚域各种写法（Page N of M、N / M、S. N——北欧语页码）。
FOOTER_PAGE_RES = (
    re.compile(r"^\s*[\u2014\u2013\-\s]*(\d+)[\s\u2014\u2013\-]*\s*$"),  # — 1 —
    re.compile(r"^\s*第\s*(\d+)\s*页.*$"),  # 第 1 页 / 第 1 页 共 3 页
    re.compile(r"^\s*Page\s*(\d+)\s*(?:of\s*\d+)?\s*$", re.IGNORECASE),  # Page 1 / Page 1 of 8
    re.compile(r"^\s*(\d+)\s*[/／]\s*\d+\s*$"),  # 1 / 8（总页码式页脚）
    re.compile(r"^\s*(?:S|Nr)\.\s*(\d+)\s*$", re.IGNORECASE),  # S. 1 / Nr. 1（北欧/德语页码缩写）
    re.compile(r"^\s*(\d+)\s*$"),  # 裸数字
)

# 近空白页的**有意稀疏**豁免：签字 / 盖章 / 无正文声明出现在页面上，说明该页
# 本来就该只有这几行——不计入空页（README「已知边界」的表单尾页误报）。
SPARSE_OK_RE = re.compile(
    r"签字|签章|盖章|签署|公章|以下无正文|本页面故意留白|此页有意留白"
    r"|intentionally\s+\w*\s*left\s+blank|left\s+blank\s+intentionally",
    re.IGNORECASE,
)
