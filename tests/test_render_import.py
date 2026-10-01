"""CLI 壳的 import 副作用守卫（0.6.x 起头于 render.py，本批铺到全部壳）。

历史上这些脚本在模块顶层直接调用 force_utf8_stdio() 与 cleanup_old_temp()：
任何 `import render`（测试复用、工具脚本、未来模块化）都会悄悄改宿主 stdio
编码、扫描并删除系统临时目录。0.6.x 把 render.py 收敛为「副作用只在 main()
执行」，本批把同一条既定约定铺到其余 8 个壳，守卫随之参数化。

为什么值得机器守住：`force_utf8_stdio()` 用的是 `errors="replace"`，单向有损——
宿主进程（pytest、被它 exec 的 4 个测试文件）的 stdio 被改掉后，潜伏的编码
问题会被静默替换掉而不是报错，等于把缺陷藏进噪音里。

make_ref.py 刻意不在名单内：它没有 main()，整个模块体就是脚本本体（顶层
parse_args → 末尾 doc.save），只被 `python scripts/make_ref.py` 直接运行、
从不被 import，所以不存在 import 副作用；把它改造成 main() 是结构重构。
"""

import ast
import os
import subprocess
import sys

import pytest

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(KIT, "scripts")

# 有 main() 入口、可被安全 import 的 CLI 壳
CLI_SHIELDS = [
    "render.py",
    "validate.py",
    "patch.py",
    "edit.py",
    "distill.py",
    "snapshot.py",
    "check_pdf.py",
    "finalize.py",
    "align_tables.py",
]

SIDE_EFFECT_CALLS = ("force_utf8_stdio", "cleanup_old_temp")


def _module(py_name):
    path = os.path.join(SCRIPTS, py_name)
    with open(path, encoding="utf-8") as f:
        return ast.parse(f.read(), filename=path)


def _call_names(node):
    """收集节点内所有调用的被调名（只取 func 位，参数引用不算调用）。"""
    names = set()
    for n in ast.walk(node):
        if isinstance(n, ast.Call):
            f = n.func
            if isinstance(f, ast.Name):
                names.add(f.id)
            elif isinstance(f, ast.Attribute):
                names.add(f.attr)
    return names


def _imported_side_effects(tree):
    """本模块 import 进来的副作用函数名。没 import 就谈不上调用，不该强求。"""
    names = set()
    for node in tree.body:
        if isinstance(node, ast.ImportFrom):
            names |= {a.asname or a.name for a in node.names}
    return names & set(SIDE_EFFECT_CALLS)


@pytest.mark.parametrize("py_name", CLI_SHIELDS)
def test_no_side_effect_calls_at_module_top_level(py_name):
    """顶层不得执行 stdio 重配置或临时目录清理。"""
    top_calls = set()
    for stmt in _module(py_name).body:
        if isinstance(stmt, ast.Expr) and isinstance(stmt.value, ast.Call):
            top_calls |= _call_names(stmt.value)
    assert not top_calls & set(SIDE_EFFECT_CALLS), (
        "%s 顶层存在 import 副作用调用：%s——移到 main()"
        % (py_name, sorted(top_calls & set(SIDE_EFFECT_CALLS)))
    )


@pytest.mark.parametrize("py_name", CLI_SHIELDS)
def test_side_effects_run_in_main(py_name):
    """副作用必须在 main() 入口执行——CLI 行为不能因收敛而丢失。"""
    tree = _module(py_name)
    fn = next(n for n in ast.walk(tree) if isinstance(n, ast.FunctionDef) and n.name == "main")
    called = _call_names(fn)
    for name in sorted(_imported_side_effects(tree)):
        assert name in called, "%s 的 main() 未调用 %s()：CLI 行为回归" % (py_name, name)


@pytest.mark.parametrize("py_name", CLI_SHIELDS)
def test_importing_script_is_silent(py_name, tmp_path):
    """真实 import 一次：退出码 0、零输出（不打印、不告警、不动宿主 stdio）。"""
    stem = os.path.splitext(py_name)[0]
    code = "import sys; sys.path.insert(0, %r); import %s" % (SCRIPTS, stem)
    r = subprocess.run(
        [sys.executable, "-c", code],
        capture_output=True,
        text=True,
        encoding="utf-8",
        cwd=str(tmp_path),
    )
    assert r.returncode == 0, r.stderr
    assert r.stdout == "", "import %s 产生了输出（副作用回归？）：%r" % (py_name, r.stdout)
