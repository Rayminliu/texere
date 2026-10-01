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
import sys
import unicodedata

import pytest

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS_DIR = os.path.join(KIT, "scripts")
DOC_PATH = os.path.join(KIT, "docs", "SCRIPT_HELP.md")

# 用 argparse 且有长选项的脚本（post.py 只有位置参数，无选项可对照，不入表）
ARGPARSE_SCRIPTS = [
    "render.py",
    "validate.py",
    "patch.py",
    "edit.py",
    "distill.py",
    "make_ref.py",
    "snapshot.py",
    "check_pdf.py",
    "finalize.py",
    "align_tables.py",
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
    real = set(SKIP_OPTIONS)  # --help 是 argparse 自带的，写进文档不算虚构
    for s in ARGPARSE_SCRIPTS:
        real |= _options_of(os.path.join(SCRIPTS_DIR, s))
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


# 手工目录导航的起头标记（两份 README 不同语言）
NAV_MARKERS = {"README.md": "**Contents**", "README.zh-CN.md": "**目录**"}


def _nav_anchors(path, marker):
    """取手工目录那一段（从 marker 到下一个空行）里的全部锚点。"""
    out, started = [], False
    for line in open(path, encoding="utf-8").read().splitlines():
        if marker in line:
            started = True
        elif started and not line.strip():
            break
        if started:
            out.extend(re.findall(r"\]\(#([^)\s]+)\)", line))
    return out


@pytest.mark.parametrize("doc", ["README.md", "README.zh-CN.md"])
def test_contents_nav_has_no_duplicate_entries(doc):
    """首屏那条手工目录不许有重复项。

    英文版一度把 Design principles 列了两次。镜像守卫只比对标题层级序列——
    「标题没多没少」和「目录列表没写错」是两件事，而目录就在第一屏，
    错一眼就能看见，所以单独守一条。
    """
    anchors = _nav_anchors(os.path.join(KIT, doc), NAV_MARKERS[doc])
    assert anchors, "%s 里没找到目录导航（标记：%s）" % (doc, NAV_MARKERS[doc])
    dup = sorted({a for a in anchors if anchors.count(a) > 1})
    assert not dup, "%s 的目录里有重复项: %s" % (doc, dup)


@pytest.mark.parametrize("doc", ["README.md", "README.zh-CN.md"])
def test_contents_nav_stays_short(doc):
    """目录是首屏导航，不是第二套站点地图——超过 12 项就该往下沉了。"""
    anchors = _nav_anchors(os.path.join(KIT, doc), NAV_MARKERS[doc])
    assert len(anchors) <= 12, "%s 的目录有 %d 项，太重了: %s" % (doc, len(anchors), anchors)


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


# 下沉到 docs/ 的参考文档同样是中英成对的；README 的镜像守卫管不到它们，另立一条
DOCS_MIRROR_PAIRS = [
    ("docs/CONFIG.md", "docs/CONFIG.zh-CN.md"),
    ("docs/TABLES.md", "docs/TABLES.zh-CN.md"),
    ("docs/VALIDATION.md", "docs/VALIDATION.zh-CN.md"),
    ("docs/EDITING.md", "docs/EDITING.zh-CN.md"),
]


@pytest.mark.parametrize("en, zh", DOCS_MIRROR_PAIRS)
def test_docs_mirrors_share_structure(en, zh):
    """从 README 下沉出去的参考文档也要中英同构。

    内容一旦离开 README，就不再被 test_readme_mirrors_share_structure 覆盖；
    不补这条的话，「下沉」会悄悄变成「英文版有一套、中文版有另一套」。
    """
    a = _headings(os.path.join(KIT, en))
    b = _headings(os.path.join(KIT, zh))
    assert a == b, "%s / %s 章节结构已分叉（层级序列不同）: %s / %s" % (en, zh, a, b)


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


def test_readme_repo_map_covers_all_scripts():
    """scripts/*.py 每一个都要出现在两份 README 的仓库地图里。

    新增脚本却不出现在地图 = 用户与 agent 的入口清单失真。基础设施模块
    （_shared / _ooxml 这类）也要列——它们不是秘密，写了反而降低上手成本；
    确有不该公开的才进白名单（当前为空）。
    """
    infra_allowlist: set[str] = set()
    actual = {
        f for f in os.listdir(os.path.join(KIT, "scripts")) if f.endswith(".py")
    } - infra_allowlist
    for doc in ("README.md", "README.zh-CN.md"):
        text = open(os.path.join(KIT, doc), encoding="utf-8").read()
        listed = set(re.findall(r"`scripts/([A-Za-z0-9_]+\.py)`", text))
        missing = actual - listed
        assert not missing, "%s 仓库地图漏列脚本: %s" % (doc, sorted(missing))


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
    s = "".join(out)
    # 与 GitHub 的锚点生成对齐：丢弃 emoji（非 L/N/_ 字符）后，其后的空格会变成首尾
    # 连字符；GitHub 会一并去掉。不处理的话，带 emoji 的标题会被误判为死链。
    return re.sub(r"-{2,}", "-", s).strip("-")


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


def test_docs_do_not_hardcode_test_count():
    """文档不得写死精确断言数——加一条用例不该改三份文档。

    旧守卫要求「文档数字 == pytest 收集数」，每加一条用例就要同步 README×2 +
    SKILL 共 7 处，维护成本大于收益且正是本套件一直在治的漂移病。改为反向约束：
    出现 `N assertions` / `N pytest assertions` / `N 项…断言` 即失败；精确计数下沉到
    实际运行 pytest 时才知道，文档只讲 L1–L4 分层与结构常量（9 检查 / 四态 / 0.00%）。
    """
    stale = []
    for doc in COUNT_DOCS:
        with open(os.path.join(KIT, doc), encoding="utf-8") as f:
            text = f.read()
        for pat in STATED_COUNT:
            for hit in pat.findall(text):
                stale.append((doc, hit))
    assert not stale, (
        "文档不应写死精确断言数（改用 L1–L4 分层叙述，见 docs/VALIDATION.md）: %s" % stale
    )


def test_config_schema_matches_docs_and_render():
    """config.schema.json ↔ CONFIG 双语键表 ↔ render 白名单 三方一致（外部审计 R3 #2）。

    schema 是静态预校验层，CONFIG 键表是唯一来源，render 的 [warn] 白名单是运行时
    执行者——三方漂移任何一方，键就会进入「文档说有/实际没有」的失真状态。
    """
    import importlib.util
    import json as _json

    schema = _json.load(open(os.path.join(KIT, "config.schema.json"), encoding="utf-8"))
    top = set(schema["properties"])
    style = set(schema["properties"]["style"]["properties"])

    sys.path.insert(0, os.path.join(KIT, "scripts"))  # render.py 依赖兄弟模块 renderers
    spec = importlib.util.spec_from_file_location(
        "render_guard", os.path.join(KIT, "scripts", "render.py")
    )
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    assert top == set(mod.KNOWN_CONFIG_KEYS), (
        "schema 顶层键与 render.py KNOWN_CONFIG_KEYS 漂移: %s" % (top ^ set(mod.KNOWN_CONFIG_KEYS))
    )

    def _table_keys(path, start_marker, end_marker, key_cell):
        text = open(path, encoding="utf-8").read()
        i = text.index(start_marker)
        # style 段可能是文件最后一节（无下一个 ## 标题）——取不到终点就到文末
        j = text.find(end_marker, i)
        j = j if j != -1 else len(text)
        keys = set()
        for line in text[i:j].splitlines():
            if not line.lstrip().startswith("|"):
                continue
            cells = [c.strip() for c in line.strip().strip("|").split("|")]
            if len(cells) > key_cell:
                keys.update(re.findall(r"`([A-Za-z_][A-Za-z0-9_]*)`", cells[key_cell]))
        return keys

    zh_top = _table_keys(
        os.path.join(KIT, "docs", "CONFIG.zh-CN.md"), "### config.json 字段", "### `style` 段", 0
    )
    en_top = _table_keys(
        os.path.join(KIT, "docs", "CONFIG.md"),
        "### config.json fields",
        "### The `style` section",
        0,
    )
    assert top == zh_top, "schema 顶层键与 zh 键表漂移: %s" % (top ^ zh_top)
    assert top == en_top, "schema 顶层键与 en 键表漂移: %s" % (top ^ en_top)

    zh_style = _table_keys(
        os.path.join(KIT, "docs", "CONFIG.zh-CN.md"), "### `style` 段", "\n## ", 1
    )
    en_style = _table_keys(
        os.path.join(KIT, "docs", "CONFIG.md"), "### The `style` section", "\n## ", 1
    )
    assert style == zh_style, "schema style 键与 zh 表漂移: %s" % (style ^ zh_style)
    assert style == en_style, "schema style 键与 en 表漂移: %s" % (style ^ en_style)


# ---------------------------------------------------------------------------
# style 键四方同步：post.py 的 S ↔ apply_style_cfg 处理清单 ↔ config.schema.json
# ↔ CONFIG 双语键表。上面那条守卫只钉键名集合；这里补另外两个「写了不生效、
# 静默默认」的漏点：S 的每个键必须真被 apply_style_cfg 消费，文档 Default 列
# 必须等于 S 的实际值。
# ---------------------------------------------------------------------------

# apply_style_cfg 里在 _STYLE_* 四张清单之外显式特判的键
_STYLE_SPECIAL_KEYS = {"header_rows", "caption_keep_with_next", "table_zebra"}

# schema style 段里合法、但不进 S 的键：caption_words 是语义不是样式，
# apply_style_cfg 走兼容路径（顶层 cfg["caption_words"] or style.caption_words）直接消费
_STYLE_SCHEMA_EXTRA_KEYS = {"caption_words"}

# 文档里用文字描述而非直接值的键 → (en 措辞, zh 措辞)
_STYLE_DOC_EXEMPT = {
    "header_rows": ("auto", "自动"),
    "table_header_color": ("unset", "不指定"),
    "toc_placeholder": ("see source", "见提示语"),
}


def _module_literals(path):
    """解析模块顶层的 literal 赋值（含多目标元组拆 zip）→ {名字: 值}。"""
    tree = ast.parse(open(path, encoding="utf-8").read())
    consts = {}
    for node in tree.body:
        if not isinstance(node, ast.Assign):
            continue
        try:
            value = ast.literal_eval(node.value)
        except (ValueError, SyntaxError):
            continue
        if not isinstance(value, (str, int, float, tuple)):
            continue
        for t in node.targets:
            if isinstance(t, ast.Name):
                consts[t.id] = value
            elif isinstance(t, ast.Tuple) and isinstance(value, tuple):
                consts.update((e.id, v) for e, v in zip(t.elts, value) if isinstance(e, ast.Name))
    return consts


def _post_style_surface():
    """ast 解构 post.py（不 import，避免拖入 python-docx）：
    返回 (S 键序列, S 默认值表, _STYLE_* 清单字典)。

    默认值若引用 _shared 的语言词表常量（单一事实源），会顺着
    `from _shared import ...` 把 _shared 的字面值解进 consts。"""
    src = open(os.path.join(SCRIPTS_DIR, "post.py"), encoding="utf-8").read()
    tree = ast.parse(src)
    shared_consts = _module_literals(os.path.join(SCRIPTS_DIR, "_shared.py"))
    consts = {}  # 模块级常量（GRAY 等），S 的字面值可能引用它们
    s_node = None
    style_lists = {}
    for node in tree.body:
        if isinstance(node, ast.ImportFrom) and node.module == "_shared":
            for alias in node.names:
                name = alias.asname or alias.name
                if name in shared_consts:
                    consts[name] = shared_consts[name]
            continue
        if not isinstance(node, ast.Assign):
            continue
        try:
            value = ast.literal_eval(node.value)
        except (ValueError, SyntaxError):
            value = None
        if value is not None and isinstance(value, (str, int, float, tuple)):
            single = [t.id for t in node.targets if isinstance(t, ast.Name)]
            if len(single) == 1 and single[0].startswith("_STYLE_"):
                style_lists[single[0]] = list(value)
                continue
            # 单名常量赋值，或多目标（如旧版 SONG, LATIN = ...）拆 zip
            for t in node.targets:
                if isinstance(t, ast.Name):
                    consts[t.id] = value
                elif isinstance(t, ast.Tuple) and isinstance(value, tuple):
                    consts.update(
                        (e.id, v) for e, v in zip(t.elts, value) if isinstance(e, ast.Name)
                    )
            continue
        names = [t.id for t in node.targets if isinstance(t, ast.Name)]
        if len(names) != 1:
            continue
        if names[0] == "S" and isinstance(node.value, ast.Dict):
            s_node = node.value
    assert s_node is not None, "post.py 里找不到模块级 S 字典"
    s_keys, s_defaults = [], {}
    for k, v in zip(s_node.keys, s_node.values):
        s_keys.append(k.value)
        try:
            s_defaults[k.value] = ast.literal_eval(v)
        except (ValueError, SyntaxError):
            s_defaults[k.value] = consts.get(getattr(v, "id", None), ast.unparse(v))
    return s_keys, s_defaults, style_lists


def _style_default_repr(key, value):
    """S 的值 → 文档单元格里应有的字面形式（None/长文案键不走这里，见豁免表）。"""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, tuple) and len(value) == 3 and all(isinstance(c, int) for c in value):
        return "%02X%02X%02X" % value  # 颜色元组 → 6 位十六进制
    if isinstance(value, float) and value == int(value):
        return "%d" % value  # 9.0 在文档里写 9
    return str(value)


def test_style_keys_consumed_by_apply_style_cfg():
    """S 的键集合 == _STYLE_* 四张清单并集 + 特判键 == schema style 键：三方互为镜像。

    漏一个键的表现正是本项目最怕的「用户在 config 写了不生效、静默吃默认」。
    """
    import json as _json

    s_keys, _, style_lists = _post_style_surface()
    processed = set().union(*style_lists.values()) | _STYLE_SPECIAL_KEYS
    assert set(style_lists) == {"_STYLE_STR", "_STYLE_FLOAT", "_STYLE_INT", "_STYLE_COLOR"}, (
        "post.py 出现了新的 _STYLE_* 清单，本守卫未覆盖: %s" % sorted(style_lists)
    )
    assert set(s_keys) == processed, "S 键与 apply_style_cfg 处理清单漂移: %s" % (
        set(s_keys) ^ processed
    )
    schema = _json.load(open(os.path.join(KIT, "config.schema.json"), encoding="utf-8"))
    schema_style = set(schema["properties"]["style"]["properties"])
    assert schema_style == processed | _STYLE_SCHEMA_EXTRA_KEYS, (
        "schema style 键与 S/处理清单漂移: %s"
        % (schema_style ^ processed ^ _STYLE_SCHEMA_EXTRA_KEYS)
    )


@pytest.mark.parametrize(
    "doc, start, end",
    [
        ("docs/CONFIG.md", "### The `style` section", "\n## "),
        ("docs/CONFIG.zh-CN.md", "### `style` 段", "\n## "),
    ],
)
def test_docs_style_defaults_match_post_constants(doc, start, end):
    """CONFIG 键表的 Default 列 == post.py S 的实际默认值（逐键 containment）。

    逐行不拆列（一行多键时拆分列序易碎），改查「该键所在行的 Default 单元格里
    出现格式化后的默认值」；豁免表里用文字描述的键检查对应措辞存在。
    """
    _, s_defaults, _ = _post_style_surface()
    text = open(os.path.join(KIT, doc), encoding="utf-8").read()
    i = text.index(start)
    j = text.find(end, i)
    j = j if j != -1 else len(text)
    rows = {}
    for line in text[i:j].splitlines():
        if not line.lstrip().startswith("|"):
            continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(cells) != 3:
            continue
        for k in re.findall(r"`([A-Za-z_][A-Za-z0-9_]*)`", cells[1]):
            rows[k] = cells[2]
    marker_idx = 0 if doc.endswith("CONFIG.md") else 1
    bad = []
    for key, value in s_defaults.items():
        assert key in rows, "%s 的 style 表漏了键 %s" % (doc, key)
        cell = rows[key]
        if key in _STYLE_DOC_EXEMPT:
            if _STYLE_DOC_EXEMPT[key][marker_idx] not in cell:
                bad.append((key, cell[:40]))
            continue
        if _style_default_repr(key, value).lower() not in cell.lower():
            bad.append((key, "%r 应在 %s" % (value, cell[:40])))
    assert not bad, "%s 的 Default 列与 post.py S 实际值漂移: %s" % (doc, bad)


def test_lua_default_caption_words_match_shared():
    """captions.lua 默认题注词表与 _shared.DEFAULT_CAPTION_WORDS 逐词一致。

    词表历史上三份互不一致（post.py 正则含 Fig.? 缺 圖片、lua TABLE 侧误收
    圖片、FIGURE 侧缺 Fig）。词表合一后本守卫钉死唯一声明点与 lua 的同步，
    只解析文本、不执行 lua。
    """
    shared_src = open(os.path.join(SCRIPTS_DIR, "_shared.py"), encoding="utf-8").read()
    words = None
    for node in ast.parse(shared_src).body:
        if isinstance(node, ast.Assign) and any(
            isinstance(t, ast.Name) and t.id == "DEFAULT_CAPTION_WORDS" for t in node.targets
        ):
            words = ast.literal_eval(node.value)
    assert isinstance(words, dict), "_shared.py 里找不到 DEFAULT_CAPTION_WORDS"

    lua = open(os.path.join(SCRIPTS_DIR, "filters", "captions.lua"), encoding="utf-8").read()
    for lua_name, key in (("DEFAULT_TABLE_WORDS", "table"), ("DEFAULT_FIGURE_WORDS", "figure")):
        m = re.search(lua_name + r"\s*=\s*\{([^}]*)\}", lua)
        assert m, f"captions.lua 找不到 {lua_name} 默认词表"
        lua_words = re.findall(r'"([^"]+)"', m.group(1))
        assert lua_words == words[key], (
            f"{lua_name} 与 _shared.DEFAULT_CAPTION_WORDS[{key!r}] 漂移：lua={lua_words} py={words[key]}"
        )


# 视觉漂移口径常量（DPI / 差异阈值）的单一来源守卫——轨道 0 条目 5，
# 亦是轨道 1 阈值收敛的回归对象。常量此前在 snapshot.py 与 validate.py
# 各存一份，两套值会让同一产物得到相反的视觉漂移结论（曾踩）。
DRIFT_CONSTANT_SRC = "_shared.py"
DRIFT_CONSTANT_NAMES = ("BASELINE_DPI", "DEFAULT_MAX_DIFF")
# 引用方脚本：不得再把 DPI/阈值赋成独立数字字面量，必须引用 _shared。
DRIFT_CONSUMER_SCRIPTS = ("snapshot.py", "validate.py", "check_pdf.py")


def test_snapshot_thresholds_have_single_source():
    """DPI/差异阈值只在 _shared.py 定义，消费脚本必须引用而非重新硬编码。

    两层断言：(1) 单一来源确实落在 _shared 且值逐字不变（100 / 0.001）；
    (2) 扫描消费脚本模块顶层，凡名字含 DPI 或 MAX_DIFF 的赋值，右值不得是
    数字字面量——必须是引用 _shared 的名字。只查模块顶层赋值，不碰函数
    内逻辑与 argparse 默认值，避免误伤无关的 100。
    """
    # (1) 单一来源：定义在 _shared 且值逐字不变
    shared_src = open(os.path.join(SCRIPTS_DIR, DRIFT_CONSTANT_SRC), encoding="utf-8").read()
    defined = {}
    for node in ast.parse(shared_src).body:
        if isinstance(node, ast.Assign) and len(node.targets) == 1:
            t = node.targets[0]
            if isinstance(t, ast.Name) and t.id in DRIFT_CONSTANT_NAMES:
                defined[t.id] = node.value
    for name in DRIFT_CONSTANT_NAMES:
        assert name in defined, f"_shared.py 缺少阈值单一来源 {name}"
        assert isinstance(defined[name], ast.Constant), f"_shared.{name} 应直接定义为字面量"
    assert defined["BASELINE_DPI"].value == 100, "BASELINE_DPI 值漂移"
    assert defined["DEFAULT_MAX_DIFF"].value == 0.001, "DEFAULT_MAX_DIFF 值漂移"

    # (2) 消费脚本顶层不得把 DPI/阈值再赋成独立数字字面量
    for script in DRIFT_CONSUMER_SCRIPTS:
        src = open(os.path.join(SCRIPTS_DIR, script), encoding="utf-8").read()
        for node in ast.parse(src).body:
            if not isinstance(node, ast.Assign) or len(node.targets) != 1:
                continue
            t = node.targets[0]
            if not (isinstance(t, ast.Name) and ("DPI" in t.id or "MAX_DIFF" in t.id)):
                continue
            assert not isinstance(node.value, ast.Constant), (
                f"{script}:{t.lineno} 把 {t.id} 又硬编码成字面量 {node.value.value!r}；"
                f"请引用 _shared.BASELINE_DPI / _shared.DEFAULT_MAX_DIFF"
            )


# =============================================================================
# 本轮新增的三张面（Hero 图 / Showcase / make_hero）原本完全在守卫之外：
# test_internal_anchors_resolve 只查 `](#锚点)`，断言计数守卫不含 showcase，
# 仓库地图也不管图与表格里的检查名。下面三条把刚开的口子焊住。
# =============================================================================

LINK_DOCS = ["README.md", "README.zh-CN.md", "docs/showcase.md"]
LINK_IN_DOC = re.compile(r"!?\[[^\]]*\]\(([^)]+)")


def _outside_fences(text):
    """去掉 ``` 围栏内容（代码示例里的 `![...]` 不是真链接）。"""
    out, in_fence = [], False
    for line in text.splitlines():
        if line.lstrip().startswith("```"):
            in_fence = not in_fence
            continue
        if not in_fence:
            out.append(line)
    return "\n".join(out)


@pytest.mark.parametrize("doc", LINK_DOCS)
def test_relative_links_in_docs_resolve(doc):
    """文档里的相对链接与图片必须指向真实存在的文件。

    锁定的是 `assets/hero-verified.png`、`docs/showcase.md`、`../evidence/page-00*.png`
    这类跨目录相对路径——错一个只会默默显示碎图，预览里不报锚点错，只能由测试兜。
    目录目标合法（GitHub 能链到目录）；行内代码里的写法示例（如输入要求里的
    `` ![图 1-1](a.jpg){width=13cm} ``）是语法说明不是引用，不计入。
    """
    path = os.path.join(KIT, doc)
    text = _outside_fences(open(path, encoding="utf-8").read())
    text = re.sub(r"`[^`\n]*`", "", text)  # 去掉行内代码：里面的 []() 是语法示例
    base = os.path.dirname(path)
    dead = []
    for raw in set(LINK_IN_DOC.findall(text)):
        target = raw.strip().split("#", 1)[0]
        if not target or target.startswith(("<", "http://", "https://", "mailto:")):
            continue
        if not os.path.exists(os.path.normpath(os.path.join(base, target))):
            dead.append(raw)
    assert not dead, "%s 里这些相对链接指向不存在的文件: %s" % (doc, sorted(dead))


def _module_literal_list(py_path, var):
    """取回模块级 `VAR = [...] / (...)` 的字面量列表（ast，不 import 避开依赖）。"""
    for node in ast.parse(open(py_path, encoding="utf-8").read()).body:
        if isinstance(node, ast.Assign) and any(
            isinstance(t, ast.Name) and t.id == var for t in node.targets
        ):
            return list(ast.literal_eval(node.value))
    raise AssertionError("%s 里找不到模块级 %s" % (py_path, var))


def _validate_registry_names():
    """validate.py 里 checks 注册表的检查名，按运行时顺序。"""
    src = open(os.path.join(SCRIPTS_DIR, "validate.py"), encoding="utf-8").read()
    for node in ast.walk(ast.parse(src)):
        if not isinstance(node, ast.Assign) or not any(
            isinstance(t, ast.Name) and t.id == "checks" for t in node.targets
        ):
            continue
        names = []
        for elt in node.value.elts:
            assert isinstance(elt, ast.Tuple) and isinstance(elt.elts[0], ast.Constant), (
                'validate.py 的 checks 注册表项必须是 ("name", fn, args, kwargs)'
            )
            names.append(elt.elts[0].value)
        return names
    raise AssertionError("validate.py 里找不到 checks 注册表")


REPORT_KEY = re.compile(r'"([a-z_]+)"\s*:\s*\{\s*"status"')
SHOWCASE_TABLE_HEAD = "| Check |"


def _showcase_check_table(text):
    """docs/showcase.md 里 `| Check | … |` 表格首列的反引号检查名。"""
    lines = text.splitlines()
    try:
        start = next(i for i, line in enumerate(lines) if line.startswith(SHOWCASE_TABLE_HEAD))
    except StopIteration:
        raise AssertionError("docs/showcase.md 缺少 `%s` 表头" % SHOWCASE_TABLE_HEAD)
    names = []
    for line in lines[start + 2 :]:
        if not line.startswith("|"):
            break
        m = re.match(r"^\|\s*`([a-z_]+)`\s*\|", line)
        assert m, "showcase 检查表首列不是反引号检查名: %s" % line
        names.append(m.group(1))
    return names


def test_check_names_have_single_source():
    """九项检查名：代码 ↔ Hero 图 ↔ README ↔ Showcase 逐字对齐（含顺序）。

    CHANGELOG 声称它们“逐字对齐”，而这类声明没有守卫就只是善意：validate 改名
    而 hero/showcase 没跟着改，首屏那张“凭据”就成了谎报 —— 正是本项目立身之本
    的反面。不用 import 读取（make_hero 要 pymupdf、validate 要 docx，CI 未必装齐），
    全部走 ast / 文本解析，确定且零依赖。
    """
    declared = _module_literal_list(os.path.join(SCRIPTS_DIR, "_verify.py"), "CHECK_NAMES")
    assert len(declared) == 9, "检查名单应当是 9 项（Hero 图上写的就是 9 checks）: %s" % declared

    assert _validate_registry_names() == declared, (
        "validate.py 的 checks 注册表与 _verify.CHECK_NAMES 漂移"
    )
    assert _module_literal_list(os.path.join(SCRIPTS_DIR, "make_hero.py"), "CHECKS") == declared, (
        "scripts/make_hero.py 的 CHECKS 与 _verify.CHECK_NAMES 漂移：Hero 图会谎报"
    )

    for doc in ("README.md", "README.zh-CN.md"):
        text = open(os.path.join(KIT, doc), encoding="utf-8").read()
        got = REPORT_KEY.findall(text)
        assert got == declared, "%s 的 report 摘录检查名与单一来源不符: %s" % (doc, got)

    sc = open(os.path.join(KIT, "docs", "showcase.md"), encoding="utf-8").read()
    assert _showcase_check_table(sc) == declared, "showcase 检查表与单一来源不符"
    # 证据节里的 json 是省略摘录：必须仍是声明顺序的子序列
    excerpt = REPORT_KEY.findall(sc)
    assert set(excerpt) <= set(declared), "showcase 引用了不存在的检查名: %s" % sorted(
        set(excerpt) - set(declared)
    )
    assert excerpt == [n for n in declared if n in set(excerpt)], (
        "showcase 的 report 摘录检查名顺序与声明顺序不一致: %s" % excerpt
    )


H1_AS_MDNAME = re.compile(r"^#\s+([A-Za-z0-9_.-]+\.md)\s*$")


def _first_h1(path):
    for line in _outside_fences(open(path, encoding="utf-8").read()).splitlines():
        m = re.match(r"^# \S", line)
        if m:
            return line
    return None


@pytest.mark.parametrize(
    "doc",
    ["README.md", "README.zh-CN.md"]
    + sorted("docs/" + f for f in os.listdir(os.path.join(KIT, "docs")) if f.endswith(".md")),
)
def test_doc_h1_does_not_name_another_markdown_file(doc):
    """文档 H1 长得像文件名时，就必须是它自己的名字。

    踩过的坑：docs/TABLES.zh-CN.md 顶着一行 `# CONFIG.zh-CN.md`（复制粘贴没改标题）
    一直活着。散文式标题（SCRIPT_HELP.md 的“Texere 脚本帮助文档”、showcase.md 的
    长标题）不受约束，只拦这一类错型。
    """
    h1 = _first_h1(os.path.join(KIT, doc))
    assert h1 is not None, "%s 没有 H1 标题" % doc
    m = H1_AS_MDNAME.match(h1)
    if m:
        assert m.group(1) == os.path.basename(doc), (
            "%s 的 H1 写的是 `%s`，应为 `# %s`（复制粘贴没改标题）"
            % (doc, m.group(1), os.path.basename(doc))
        )
