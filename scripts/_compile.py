"""compile 能力的纯逻辑核（可导入、无 CLI 副作用）。

从 render.py 外提把 Markdown 编成 pandoc 命令、判断是否需要人工提示的纯助手：
_should_prompt（是否当场问验收）、_outside_code_fences（剥围栏代码块）、
_build_pandoc_cmd（拼 pandoc 命令行，返回 list 可断言）。render.py 作为 thin
CLI 壳 re-export 同名符号，`run` / `preflight` / `doctor` / `_export_and_check`
等含 sys.exit / 子进程 / 版式的编排仍留壳。裸名 sibling import，禁止反向 import render。
"""

import os

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _should_prompt(can_read, can_write):
    """该不该当场问验收？只有「人在终端前」才问。

    两个流都得是 TTY：脚本 / CI / pytest 里 capture_output=True 时 stdout 是管道，
    那时子进程的 stdin 可能仍然继承了终端——只查 stdin 会挂在那里等输入。
    """
    return bool(can_read and can_write)


def _outside_code_fences(md_text: str) -> str:
    """去掉 ``` / ~~~ 围栏代码块，只留围栏外的正文。"""
    out, fence = [], None
    for line in md_text.splitlines():
        s = line.strip()
        if fence is None and (s.startswith("```") or s.startswith("~~~")):
            fence = s[:3]
            continue
        if fence is not None:
            if s.startswith(fence):
                fence = None
            continue
        out.append(line)
    return "\n".join(out)


def _build_pandoc_cmd(all_md, body, ref, uniq, cfg):
    """拼 pandoc 命令行；captions.lua 存在时把题注关键字一并喂给它。"""
    lua_filter = os.path.join(KIT, "scripts", "filters", "captions.lua")
    cmd = [
        "pandoc",
        all_md,
        "-o",
        body,
        "--reference-doc=" + ref,
        "--resource-path=" + os.pathsep.join(uniq),
        "-f",
        "markdown+pipe_tables+raw_html",
        "--wrap=none",
    ]
    if os.path.exists(lua_filter):
        # AST 层标记表题/图注，post.py 就不用再靠正则猜
        cmd.append("--lua-filter=" + lua_filter)
        # 题注关键字可配：同一份配置同时喂给 lua filter 与 post.py
        cw = cfg.get("caption_words") or {}
        for key, meta_name in (
            ("table", "dk-table-words"),
            ("figure", "dk-figure-words"),
        ):
            if cw.get(key):
                cmd += ["-M", "%s=%s" % (meta_name, ",".join(cw[key]))]
    else:
        print("[warn] 缺少 filters/captions.lua，题注退回文本正则判定")
    return cmd
