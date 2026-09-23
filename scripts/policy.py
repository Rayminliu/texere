"""Texere 验收策略模型（Flexible input, explicit policy, deterministic verification）。

核心思想：验收的「严格程度」不是二元的（全严格 / 全不严格），而是**按属性（property）**
选择四级严重度：

    ignore   不观测、不计入（跳过）
    observe  观测并如实报告，但绝不门禁（纯信息，不影响结论）
    warn     观测；若不符合则降级为 WARN（亮黄灯，不阻断交付）
    enforce  观测；若不符合则硬失败（门禁）

三种使用模式只是这套属性的预设：

    free      只生成：默认一切 observe，不强制任何 profile
    advisory  机器发现、用户选择：选定的属性设为 warn
    strict    完整验收：选定的属性设为 enforce（配合 profile + baseline + renderer）

策略可分层叠加（后者覆盖前者），对应「模板检测 spec → 用户已批准约束 → 本次任务临时覆盖」：

    base (template spec)  →  approved overrides  →  task overrides
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Iterable


class Severity(str, Enum):
    IGNORE = "ignore"
    OBSERVE = "observe"
    WARN = "warn"
    ENFORCE = "enforce"

    @property
    def gates(self) -> bool:
        """该级别是否会改变最终门禁结论（FAIL / WARN）。"""
        return self is Severity.WARN or self is Severity.ENFORCE


# 语义属性 → 受管辖的底层检查名（与 validate.py 的 CheckResult.name 对齐；
# profile.* 系列来自 --enforce-profile 编译出的可执行断言）。Phase 2 接入时据此反查。
PROPERTY_CHECKS: dict[str, list[str]] = {
    "package_integrity": ["package_integrity"],
    "source_content": ["source_content"],
    "image_embedding": ["image_embedding"],
    "section_count": ["section_count"],
    "toc_field": ["toc_field"],
    "caption": ["caption_recognition"],
    "page_numbering": ["page_numbering"],
    "blank_pages": ["blank_pages"],
    "visual_baseline": ["visual_drift"],
    "renderer_acceptance": ["renderer_acceptance"],
    "page": ["profile.page"],
    "body_font": ["profile.body_font"],
    "heading_font": ["profile.heading"],
    "table_border": ["profile.table"],
    "toc": ["profile.toc", "toc_field"],
}

# 反向：check name → property
CHECK_PROPERTY: dict[str, str] = {c: p for p, cs in PROPERTY_CHECKS.items() for c in cs}


def severity_for_check(
    check_name: str, policy: dict[str, str], default: Severity = Severity.OBSERVE
) -> Severity:
    """查某个底层检查在给定策略下的严重度；未登记的检查回落到 default（默认 observe）。"""
    prop = CHECK_PROPERTY.get(check_name)
    if prop is None:
        return default
    return Severity(policy.get(prop, default.value))


# 三种模式的预设（只声明关心的属性；未声明的属性回落到 default=observe）。
FREE_MODE: dict[str, str] = {}  # 空策略：一切走默认 observe，不强制
ADVISORY_MODE: dict[str, str] = {
    "page": Severity.WARN.value,
    "body_font": Severity.WARN.value,
    "heading_font": Severity.WARN.value,
    "toc": Severity.WARN.value,
    "caption": Severity.WARN.value,
    "page_numbering": Severity.WARN.value,
    "blank_pages": Severity.WARN.value,
    "renderer_acceptance": Severity.WARN.value,
    "visual_baseline": Severity.WARN.value,
}
STRICT_MODE: dict[str, str] = {
    "package_integrity": Severity.ENFORCE.value,
    "source_content": Severity.ENFORCE.value,
    "page": Severity.ENFORCE.value,
    "body_font": Severity.ENFORCE.value,
    "heading_font": Severity.ENFORCE.value,
    "table_border": Severity.ENFORCE.value,
    "toc": Severity.ENFORCE.value,
    "caption": Severity.ENFORCE.value,
    "page_numbering": Severity.ENFORCE.value,
    "blank_pages": Severity.ENFORCE.value,
    "renderer_acceptance": Severity.ENFORCE.value,
    "visual_baseline": Severity.ENFORCE.value,
}

MODE_PRESETS: dict[str, dict[str, str]] = {
    "free": FREE_MODE,
    "advisory": ADVISORY_MODE,
    "strict": STRICT_MODE,
}


def resolve_policy(*layers: dict[str, str]) -> dict[str, str]:
    """分层叠加策略；靠后的层覆盖靠前的层（override layering）。"""
    out: dict[str, str] = {}
    for layer in layers:
        if layer:
            out.update({k: str(v) for k, v in layer.items()})
    return out


def required_renderers(acceptance: dict | None) -> list[str]:
    """从 acceptance 配置里取「必须通过的所有 renderer」（renderer 矩阵验收）。

    例：{"renderers": ["word", "wps"]} → 两个环境都要通过。
    """
    if not acceptance:
        return []
    rs = acceptance.get("renderers") or []
    return [str(r) for r in rs]


@dataclass
class AcceptanceSummary:
    mode: str
    enforced: int = 0
    warnings: int = 0
    skipped: int = 0
    renderer: str | None = None
    visual_baseline: str | None = None
    renderers_required: list[str] = field(default_factory=list)


def evaluate(
    checks: Iterable,
    policy: dict[str, str],
    *,
    mode: str = "advisory",
    renderer: str | None = None,
    visual_baseline: str | None = None,
    renderers_required: Iterable[str] | None = None,
) -> tuple[AcceptanceSummary, str]:
    """把策略套到一组检查上，算出摘要与最终结论。

    检查项需暴露 `.name` 与 `.status`（CheckResult / 简单对象皆可）。

    严重度语义：
      - IGNORE  → 该检查强制 SKIP（不观测、不门禁）
      - OBSERVE → 如实保留结论，但绝不改变最终结论（纯信息）
      - WARN    → FAIL 降级为 WARN（亮灯不阻断）
      - ENFORCE → FAIL 仍是 FAIL（门禁）

    最终结论优先级：FAIL > WARN > PASS。
    """
    summary = AcceptanceSummary(
        mode=mode,
        renderer=renderer,
        visual_baseline=visual_baseline,
        renderers_required=list(renderers_required or []),
    )
    verdict = "PASS"
    for chk in checks:
        name = getattr(chk, "name", None)
        raw = getattr(chk, "status", None)
        sev = severity_for_check(name, policy)
        if sev is Severity.IGNORE:
            summary.skipped += 1
            continue
        if sev is Severity.WARN and raw == "FAIL":
            summary.warnings += 1
            if verdict != "FAIL":
                verdict = "WARN"
        elif sev is Severity.ENFORCE and raw == "FAIL":
            verdict = "FAIL"
        if sev is Severity.ENFORCE and raw != "SKIP":
            summary.enforced += 1
    return summary, verdict
