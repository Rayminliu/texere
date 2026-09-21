"""文档同步守卫：脚本的每个 CLI 参数都必须写进 docs/SCRIPT_HELP.md。

治的是"功能先行、文档滞后"：新增 add_argument 而不补文档，本测试即失败，
pre-commit hook 会在 commit 时拦住。反向失配（文档写了代码里没有的参数）
由 test_script_help_options_exist_in_code 兜住——虚构参数同样是文档病。

另外三条守的是文档内部的漂移：同一个配置键在 README 里被定义两次（0.6.1
的治理只拦了 SKILL/BEST_PRACTICES，README 自己漏了）、中英 README 镜像结构
失配（中文版曾整段丢掉验证章节）、以及文档里写死的断言数与实际收集数不符。
"""

import ast
import os
import re
import subprocess
import sys
import unicodedata

import pytest

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS_DIR = os.path.join(KIT, "scripts")
DOC_PATH = os.path.join(KIT, "docs", "SCRIPT_HELP.md")

# 用 argparse 的脚本（finalize/check_pdf/snapshot/post 是手解 sys.argv，另行覆盖）
ARGPARSE_SCRIPTS = [
    "render.py",
    "validate.py",
    "patch.py",
    "edit.py",
    "distill.py",
    "make_ref.py",
]
SKIP_OPTIONS = {"--help"}


def _options_of(py_path):
    """ast 提取脚本里所有 add_argument("--xxx") 的长选项。"""
    with open(py_path, encoding="utf-8") as f:
        tree = ast.parse(f.read())
    opts = set()
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        if getattr(node.func, "attr", "") != "add_argument":
            continue
        for a in node.args:
            if isinstance(a, ast.Constant) and isinstance(a.value, str):
                if a.value.startswith("--"):
                    # "--section body|all|N" 这类带示例值的，取选项名本身
                    opts.add(a.value.split()[0])
    return opts - SKIP_OPTIONS


@pytest.fixture(scope="module")
def doc_text():
    with open(DOC_PATH, encoding="utf-8") as f:
        return f.read()


@pytest.mark.parametrize("script", ARGPARSE_SCRIPTS)
def test_every_cli_option_is_documented(script, doc_text):
    """代码里有的参数，文档必须出现。"""
    opts = _options_of(os.path.join(SCRIPTS_DIR, script))
    assert opts, "%s 一个 argparse 参数都没有？" % script
    missing = sorted(o for o in opts if o not in doc_text)
    assert not missing, "%s 的新参数没写进 docs/SCRIPT_HELP.md: %s" % (script, missing)


def test_script_help_docs_have_no_invented_options(doc_text):
    """文档里出现的 --xxx，必须真在某个脚本的 argparse 里存在。

    曾踩过：SCRIPT_HELP 给 snapshot.py 编了 --baseline/--quiet/--threshold
    三个不存在的参数。虚构参数比漏写更恶劣——它教用户用不存在的东西。
    （只查带反引号的形如 `--opt` 的文档参数，避免误伤正文里的破折号词。）
    """
    real = set()
    for s in ARGPARSE_SCRIPTS:
        real |= _options_of(os.path.join(SCRIPTS_DIR, s))
    # 手解 argv 的脚本，显式登记它们的真实选项
    real |= {"--update", "--dpi", "--max-diff", "--max-empty"}  # snapshot / check_pdf
    documented = set(re.findall(r"`(--[a-z][a-z0-9-]*)`", doc_text))
    fake = sorted(o for o in documented if o not in real)
    assert not fake, "文档写了代码里不存在的参数: %s" % fake


# 字段说明表行：首列是单个反引号标识符，如 "| `cover` | 封面行 ... |"。
# 这种表只允许住在手册（README*）和 CLI 参考（SCRIPT_HELP）里——
# SKILL/BEST_PRACTICES 复述字段表就是漂移源（文档地图约定的单一来源）。
FIELD_TABLE_ROW = re.compile(r"^\|\s*`[A-Za-z_][A-Za-z0-9_]*`\s*\|.*$", re.MULTILINE)


@pytest.mark.parametrize("doc", ["SKILL.md", "BEST_PRACTICES.md"])
def test_no_field_reference_tables_outside_manual(doc):
    """字段/参数说明表不得在 SKILL.md 与 BEST_PRACTICES.md 中复述，只能引用。"""
    with open(os.path.join(KIT, doc), encoding="utf-8") as f:
        text = f.read()
    hits = FIELD_TABLE_ROW.findall(text)
    assert not hits, "%s 复述了字段说明表（应改为引用 README/SCRIPT_HELP）: %s" % (
        doc,
        [h[:40] for h in hits[:3]],
    )


IDENT = re.compile(r"`([A-Za-z_][A-Za-z0-9_]*)`")


def _table_rows(text):
    """产出表格行的单元格列表（跳过代码块内的伪表格行，按未转义的 | 切分）。"""
    rows = []
    in_fence = False
    for line in text.splitlines():
        if line.lstrip().startswith("```"):
            in_fence = not in_fence
            continue
        if in_fence or not line.lstrip().startswith("|"):
            continue
        cells = [c.strip() for c in re.split(r"(?<!\\)\|", line.strip())]
        rows.append([c for c in cells][1:-1])  # 丢掉行首行尾竖线切出的空串
    return rows


@pytest.mark.parametrize("doc", ["README.md", "README.zh-CN.md"])
def test_no_key_defined_twice_in_manual(doc):
    """`键 | 默认 | 说明` 式的定义行，同一个键在一个文件里只能出现一次。

    README 一度在「style 段」和「表格视觉控制」两张表里各写了一份 header_rows /
    table_border 等 6 个键的默认值——单一来源约定在手册内部同样成立，
    否则改一处忘一处，读者拿到的是两份互相竞争的默认值。
    """
    with open(os.path.join(KIT, doc), encoding="utf-8") as f:
        rows = _table_rows(f.read())
    seen = {}
    for cells in rows:
        if len(cells) != 3:
            continue
        keys = IDENT.findall(cells[0]) + IDENT.findall(cells[1])
        if not keys or not cells[2]:
            continue
        for k in keys:
            seen.setdefault(k, 0)
            seen[k] += 1
    dup = sorted(k for k, n in seen.items() if n > 1)
    assert not dup, "%s 里这些配置键被定义了两次，默认值会各说各话: %s" % (doc, dup)


FRONTMATTER_UNQUOTED_KV = re.compile(r"^([A-Za-z_][\w-]*):\s+(?!['\"])(.*\S.*)$")


def test_skill_frontmatter_is_valid_yaml():
    """SKILL.md 的 front matter 必须能被严格 YAML 解析器读。

    踩过：`description: ... with unified validation: 9 automated checks ...`
    —— 未加引号的值里出现 `: `，PyYAML 直接报
    “mapping values are not allowed in this context”，技能加载失败。
    不依赖 PyYAML（它不是项目依赖），用结构检查兼容这个单一错型。
    """
    with open(os.path.join(KIT, "SKILL.md"), encoding="utf-8") as f:
        lines = f.read().splitlines()
    assert lines and lines[0].strip() == "---", "SKILL.md 首行必须是 front matter 分隔符"
    fm = []
    for line in lines[1:]:
        if line.strip() == "---":
            break
        fm.append(line)
    else:
        raise AssertionError("SKILL.md 的 front matter 没有闭合")
    assert fm, "front matter 为空"
    bad = []
    for line in fm:
        m = FRONTMATTER_UNQUOTED_KV.match(line)
        if m and ": " in m.group(2):
            bad.append((m.group(1), m.group(2)[max(0, m.group(2).find(": ") - 20) :][:60]))
    assert not bad, (
        "front matter 里这些值含未引用的 `: `，会被 YAML 当成嵌套映射而解析失败，"
        "请用双引号包裹整个值: %s" % (bad,)
    )


def _headings(path):
    """标题层级序列（跳过代码块），用于比对中英 README 的结构。"""
    with open(path, encoding="utf-8") as f:
        lines = f.read().splitlines()
    out, in_fence = [], False
    for line in lines:
        if line.lstrip().startswith("```"):
            in_fence = not in_fence
            continue
        if in_fence:
            continue
        m = re.match(r"^(#{1,3}) ", line)
        if m:
            out.append(len(m.group(1)))
    return out


def test_readme_mirrors_share_structure():
    """中英 README 必须是同一份内容的两个语言版本，不是两份各写各的文档。

    中文版曾整段没有验证章节（9 项检查、证据包），英文版的三条支柱也已改写，
    两份越差越远——标题层级序列相等这个粗粒度约束足以暴露「一边多了一节」。
    真要结构性分叉时，改这个测试，别只改一份 README。
    """
    en = _headings(os.path.join(KIT, "README.md"))
    zh = _headings(os.path.join(KIT, "README.zh-CN.md"))
    assert en == zh, "中英 README 章节结构已分叉（层级序列不同），把两份对齐: %s / %s" % (en, zh)


def test_readme_mirrors_cover_same_scripts():
    """两份 README 提到的脚本集合要一致——只在一边介绍某脚本就是漏写。"""
    pat = re.compile(r"scripts/([a-z_]+\.py)")
    found = {}
    for doc in ("README.md", "README.zh-CN.md"):
        with open(os.path.join(KIT, doc), encoding="utf-8") as f:
            found[doc] = set(pat.findall(f.read()))
    en, zh = found["README.md"], found["README.zh-CN.md"]
    assert en == zh, "中英 README 提到的脚本不一致: 仅英文 %s / 仅中文 %s" % (
        sorted(en - zh),
        sorted(zh - en),
    )


def _slug(title):
    """按 GitHub 的规则把标题转成锚点（去内联代码/加粗，只留字母数词与连字符）。"""
    title = re.sub(r"`([^`]*)`", r"\1", title)
    title = re.sub(r"\*\*([^*]*)\*\*", r"\1", title)
    out = []
    for ch in title.strip().lower():
        if ch in " -":
            out.append("-")
        elif ch == "_" or unicodedata.category(ch).startswith(("L", "N")):
            out.append(ch)
    return "".join(out)


def _headings_and_slugs(path):
    in_fence, slugs = False, set()
    with open(path, encoding="utf-8") as f:
        for line in f.read().splitlines():
            if line.lstrip().startswith("```"):
                in_fence = not in_fence
                continue
            if in_fence:
                continue
            m = re.match(r"^#{1,6}\s+(.*)$", line)
            if m:
                slugs.add(_slug(m.group(1)))
    return slugs


@pytest.mark.parametrize("doc", ["README.md", "README.zh-CN.md"])
def test_internal_anchors_resolve(doc):
    """README 顶部的锚点导航不能断——改标题就得同步改链接。

    导航是这一版才加上的（23 个锚点），而错一个锚点只会默默跳到页首，
    预览里看不出来，所以只能由测试兜。
    """
    path = os.path.join(KIT, doc)
    slugs = _headings_and_slugs(path)
    with open(path, encoding="utf-8") as f:
        text = f.read()
    # 只查看起来像锚点的（正文里有 `[1](#...)` 这种被当作反例引用的伪锚点）
    anchors = (
        a for a in re.findall(r"\]\(#([^)\s]+)\)", text) if re.fullmatch(r"[\w\u4e00-\u9fff-]+", a)
    )
    dead = sorted(a for a in set(anchors) if a not in slugs)
    assert not dead, "%s 里的锁定链接指向不存在的标题: %s" % (doc, dead)


STATED_COUNT = [
    re.compile(r"(\d{2,4})\s+assertions"),  # README.md / SKILL.md
    re.compile(r"(\d{2,4})\s+pytest\s+assertions"),
    re.compile(r"(\d{2,4})\s*项[^，。\n]{0,8}断言"),  # README.zh-CN.md
]
COUNT_DOCS = ["README.md", "README.zh-CN.md", "SKILL.md"]


def test_stated_test_count_is_current():
    """文档里写死的「N 项断言」必须等于 pytest 实际收集数。

    这个数字在 3 份文档里出现 6 次，加一条用例忘记改文档，就是「文档吹嘘比
    实际多」——恰好是本套件一直在治的那种病，这次轮到治自己。
    """
    result = subprocess.run(
        [sys.executable, "-m", "pytest", "--collect-only", "-q", "--no-header"],
        cwd=KIT,
        capture_output=True,
        text=True,
    )
    m = re.search(r"(\d+) tests? collected", result.stdout)
    assert m, "没能从 pytest 输出里拿到收集数:\n" + result.stdout[-500:]
    real = int(m.group(1))
    stale = []
    for doc in COUNT_DOCS:
        with open(os.path.join(KIT, doc), encoding="utf-8") as f:
            text = f.read()
        for pat in STATED_COUNT:
            for hit in pat.findall(text):
                if int(hit) != real:
                    stale.append((doc, hit))
    assert not stale, "文档里的断言数过期了（实际收集 %d 项）: %s" % (real, stale)
