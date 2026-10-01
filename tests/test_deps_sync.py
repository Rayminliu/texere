"""依赖声明双轨一致性守卫：pyproject.toml ↔ requirements.txt 不许漂移。

背景（0.7.3 全面审查发现的最高风险项）：pyproject 声明 `python-docx>=1.1,<2.0`
而 requirements 写着 `>=1.2,<3`——CI 按 requirements 安装、uv 用户按 pyproject 解析，
两边可以装出不同大版本；而 post.py 大量依赖 OOXML 底层 API，主版本漂移是静默版式风险。
本守卫钉死两轨的区间一致，新增依赖必须两边同批声明。

纯文本解析（不依赖 tomllib / packaging），Python 3.10 也能跑。
"""

import os
import re

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# requirements.txt 里允许**不**出现在 pyproject 核心 deps 的运行时包 → 理由。
# 守卫之外的新白名单必须写清为什么。
RUNTIME_WHITELIST = {
    # lxml 的传递依赖；显式钉下限只为防降级报错，故意不进 pyproject（避免收紧 uv 解析）。
    "typing-extensions": "transitive pin only",
    # ruff 在 requirements 是 == 精确钉版（与 .pre-commit-config.yaml 的 rev 同步），
    # pyproject dev 组只给区间下限；守卫按「精确版必须落在区间内」校验。
    "ruff": "exact pin vs range",
    # 覆盖率度量为本机可选工具，pyproject dev 组未声明（CI 不装也能全绿）。
    "pytest-cov": "local-only tool",
}

# pyproject extras（可选依赖）也纳入可匹配范围：requirements 里的 pymupdf 对应 check extra。
_EXTRA_RE = re.compile(r"^(\w[\w-]*)\s*=\s*\[([^\]]*)\]", re.MULTILINE)
_DEPS_RE = re.compile(r"^dependencies\s*=\s*\[([^\]]*)\]", re.MULTILINE)
_ITEM_RE = re.compile(r'"([^"]+)"')
_REQ_LINE_RE = re.compile(r"^([A-Za-z0-9_.-]+)\s*(.*)$")
_CLAUSE_RE = re.compile(r"(>=|<=|==|~=|!=|>|<)\s*([^,]+)")


def _norm_name(name):
    return re.sub(r"[-_.]+", "-", name).lower()


def _norm_ver(ver):
    """PEP440 尾零等价：2 == 2.0 == 2.0.0。非纯数字版本原样返回。"""
    parts = ver.strip().split(".")
    try:
        ints = [int(p) for p in parts]
    except ValueError:
        return ver.strip()
    while len(ints) > 1 and ints[-1] == 0:
        ints.pop()
    return tuple(ints)


def _clauses(spec):
    """'>=1.2,<2.0' → frozenset({('>=', (1,2)), ('<', (2,))})"""
    return frozenset((op, _norm_ver(v)) for op, v in _CLAUSE_RE.findall(spec.replace(" ", "")))


def _split_req(text):
    """'python-docx>=1.2,<2.0' → (name, spec)；文本已去空格。"""
    text = text.replace(" ", "")
    m = re.match(r"([A-Za-z0-9_.-]+)(.*)", text)
    return _norm_name(m.group(1)), m.group(2)


def _pyproject_index():
    src = open(os.path.join(KIT, "pyproject.toml"), encoding="utf-8").read()
    index = {}
    core = _DEPS_RE.search(src)
    assert core, "pyproject.toml 里找不到 [project].dependencies"
    for item in _ITEM_RE.findall(core.group(1)):
        name, spec = _split_req(item)
        index[name] = _clauses(spec)
    # extras 与 dev 组：按小节切块，避免把 [tool.*] 的表误当依赖
    for section in ("[project.optional-dependencies]", "[dependency-groups]"):
        start = src.find(section)
        if start < 0:
            continue
        end = src.find("\n[", start + len(section))
        block = src[start + len(section) : end if end > 0 else len(src)]
        for m in _EXTRA_RE.finditer(block):
            for item in _ITEM_RE.findall(m.group(2)):
                name, spec = _split_req(item)
                index.setdefault(name, _clauses(spec))
    return index


def _requirements_index():
    index = {}
    path = os.path.join(KIT, "requirements.txt")
    for raw in open(path, encoding="utf-8").read().splitlines():
        line = raw.split("#", 1)[0].strip()
        if not line or line.startswith("-"):
            continue
        m = _REQ_LINE_RE.match(line)
        assert m, f"requirements.txt 出现无法解析的行: {raw!r}"
        index[_norm_name(m.group(1))] = _clauses(m.group(2))
    return index


def test_requirements_specs_match_pyproject():
    """两轨共同声明的包，区间必须逐子句相等；requirements 不得引入未白名单的野生包。"""
    pp = _pyproject_index()
    req = _requirements_index()
    for name, clauses in req.items():
        if name in RUNTIME_WHITELIST:
            continue
        assert name in pp, f"requirements.txt 的 {name} 未在 pyproject.toml 声明（双轨之一漏写）"
        assert clauses == pp[name], (
            f"{name} 区间漂移：requirements={sorted(map(str, clauses))} vs "
            f"pyproject={sorted(map(str, pp[name]))}；两处必须同批修改"
        )


def test_pyproject_core_deps_all_in_requirements():
    """pyproject 核心运行时 deps 必须都出现在 requirements（CI 只装 requirements）。"""
    src = open(os.path.join(KIT, "pyproject.toml"), encoding="utf-8").read()
    core = _DEPS_RE.search(src)
    req = _requirements_index()
    for item in _ITEM_RE.findall(core.group(1)):
        name, _ = _split_req(item)
        assert name in req, f"pyproject 核心依赖 {name} 在 requirements.txt 缺失，CI 会装不到"


def test_exact_pins_satisfy_pyproject_range():
    """白名单里 == 精确钉版的工具（ruff）必须落在 pyproject 声明的区间内。"""
    pp = _pyproject_index()
    req = _requirements_index()
    for name in RUNTIME_WHITELIST:
        if name not in req or name not in pp:
            continue
        exact = [c for c in req[name] if c[0] == "=="]
        if not exact:
            continue
        ((_, pinned),) = exact
        bounds = [c for c in pp[name] if c[0] in (">=", "<", ">", "~=", "!=")]
        for op, ver in bounds:
            if op == ">=":
                assert pinned >= ver, f"{name} 钉版 {pinned} 低于下限 {ver}"
            elif op == "<":
                assert pinned < ver, f"{name} 钉版 {pinned} 越过上限 {ver}"
            elif op == ">":
                assert pinned > ver, f"{name} 钉版 {pinned} 不高于下限 {ver}"


def test_whitelist_entries_are_justified():
    """白名单只收「两轨之一有意缺席/区间不同」的包；全部收进 pyproject 后应删条目。"""
    pp = _pyproject_index()
    req = _requirements_index()
    for name in RUNTIME_WHITELIST:
        assert name in req or name in pp, f"白名单 {name} 两边都不存在，是死条目，删掉"
