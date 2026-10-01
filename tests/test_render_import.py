"""render.py 的 import 副作用守卫。

历史上 render.py 在模块顶层直接调用 force_utf8_stdio() 与 cleanup_old_temp()：
任何 `import render`（测试复用、工具脚本、未来模块化）都会悄悄改宿主 stdio
编码、扫描并删除系统临时目录。0.6.x 收敛为副作用只在 main() 入口执行，
本守卫把这条边界钉死，防止回归。
"""

import ast
import os
import subprocess
import sys

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RENDER_PATH = os.path.join(KIT, "scripts", "render.py")

SIDE_EFFECT_CALLS = ("force_utf8_stdio", "cleanup_old_temp")


def _module():
    with open(RENDER_PATH, encoding="utf-8") as f:
        return ast.parse(f.read(), filename=RENDER_PATH)


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


def test_no_side_effect_calls_at_module_top_level():
    """顶层不得执行 stdio 重配置或临时目录清理。"""
    top_calls = set()
    for stmt in _module().body:
        if isinstance(stmt, ast.Expr) and isinstance(stmt.value, ast.Call):
            top_calls |= _call_names(stmt.value)
    assert not top_calls & set(SIDE_EFFECT_CALLS), (
        "render.py 顶层存在 import 副作用调用：%s——移到 main()"
        % sorted(top_calls & set(SIDE_EFFECT_CALLS))
    )


def test_side_effects_run_in_main():
    """两个副作用必须在 main() 入口执行——CLI 行为不能因收敛而丢失。"""
    fn = next(n for n in ast.walk(_module()) if isinstance(n, ast.FunctionDef) and n.name == "main")
    called = _call_names(fn)
    for name in SIDE_EFFECT_CALLS:
        assert name in called, "main() 未调用 %s()：CLI 行为回归" % name


def test_importing_render_is_silent(tmp_path):
    """真实 import 一次：退出码 0、零输出（不打印、不告警、不抛异常）。"""
    code = "import sys; sys.path.insert(0, %r); import render" % os.path.join(KIT, "scripts")
    r = subprocess.run(
        [sys.executable, "-c", code],
        capture_output=True,
        text=True,
        encoding="utf-8",
        cwd=str(tmp_path),
    )
    assert r.returncode == 0, r.stderr
    assert r.stdout == "", "import render 产生了输出（副作用回归？）：%r" % r.stdout
