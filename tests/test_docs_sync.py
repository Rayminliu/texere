"""文档同步守卫：脚本的每个 CLI 参数都必须写进 docs/SCRIPT_HELP.md。

治的是"功能先行、文档滞后"：新增 add_argument 而不补文档，本测试即失败，
pre-commit hook 会在 commit 时拦住。反向失配（文档写了代码里没有的参数）
由 test_script_help_options_exist_in_code 兜住——虚构参数同样是文档病。
"""

import ast
import os
import re

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
