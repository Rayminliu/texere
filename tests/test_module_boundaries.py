"""依赖方向守卫：core 能力模块（scripts/_*.py）绝不 import CLI 壳。

这次重构把纯逻辑从巨型 CLI 脚本外提为可导入的 `_*.py` 能力模块，脚本收敛为
thin 壳并 re-export 同名符号。边界现在是干净的：

    CLI facade  ->  capability  ->  primitive / OOXML
   (validate.py)     (_verify.py)     (_docx_edit.py / _ooxml.py)

这类倒退通常不是一次大事故，而是将来有人为快速修 bug 顺手「就 import 一个函数
嘛」，把依赖拉回 core -> CLI 的反向耦合，三个月后重新缠死。所以趁边界最干净的
现在把它钉住：任何 `scripts/_*.py` 里出现对 CLI 壳的 import 即失败。

只锁「禁止方向」这一件事，不做通用架构检查框架——正向（壳 import core）与同级
（core import core/_shared）一律放行。
"""

import os
import re

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(KIT, "scripts")

# CLI 薄壳：带 argparse / main() 的入口脚本，处在依赖方向最上层，不该被下层反向引用。
# 刻意不含 renderers.py（渲染器后端是 peer capability，不是 CLI 壳）。
CLI_FACADES = {
    "align_tables",
    "check_pdf",
    "distill",
    "edit",
    "finalize",
    "make_ref",
    "patch",
    "post",
    "render",
    "snapshot",
    "validate",
}

# 顶层或函数内缩进的 import / from-import，取被导入模块的根名。
_IMPORT_LINE = re.compile(r"^[ \t]*(?:import|from)[ \t]+([A-Za-z_][A-Za-z0-9_.]*)", re.MULTILINE)


def _root_module(dotted: str) -> str:
    return dotted.split(".", 1)[0]


def test_core_modules_do_not_import_cli_facades():
    core_files = sorted(f for f in os.listdir(SCRIPTS) if f.startswith("_") and f.endswith(".py"))
    assert core_files, "scripts/ 下没找到任何 _*.py 核心模块？路径或命名可能变了"

    violations = []
    for fname in core_files:
        path = os.path.join(SCRIPTS, fname)
        with open(path, encoding="utf-8") as fh:
            text = fh.read()
        for token in _IMPORT_LINE.findall(text):
            root = _root_module(token)
            if root in CLI_FACADES:
                violations.append((fname, root))

    if violations:
        lines = ["Dependency boundary violation — core (_*.py) must not import CLI facades:", ""]
        for fname, mod in sorted(set(violations)):
            lines.append(f"  scripts/{fname} imports {mod}")
        lines += [
            "",
            "Direction is: CLI facade -> capability -> primitive / OOXML.",
            "Move the shared logic into the core module (or an existing _*.py) instead",
            "of reaching back up into a CLI script — that re-couples what this refactor",
            "just decoupled.",
        ]
        raise AssertionError("\n".join(lines))
