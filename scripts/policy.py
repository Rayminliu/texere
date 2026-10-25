"""Texere 验收策略模型（Flexible input, explicit policy, deterministic verification）。

0.9.3 起接入 validate.py 退出码：`--acceptance-mode advisory|strict` 时，profile
契约层（`profile.*` 断言）按本模型裁决门禁；核心九项不变量与 renderer 矩阵
不归策略管，FAIL/ERROR 永远硬门禁（见 evaluate 的 enforced_check_results）。

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
    "caption": [],  # caption_recognition 检查尚未实现，此映射留空（原幽灵值已清除）
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


def required_renderers(acceptance: dict | None) -> list[str]:
    """从 acceptance 配置里取「必须通过的所有 renderer」（renderer 矩阵验收）。

    例：{"renderers": ["word", "wps"]} → 两个环境都要通过。
    """
    if not acceptance:
        return []
    rs = acceptance.get("renderers") or []
    return [str(r) for r in rs]


# ---- validate.py 接线（0.9.3）------------------------------------------------
# 属性→检查名映射在这里消费；接线端只需要 mode 解析与核心不变量名单两个小函数，
# 放回 policy.py 本体（它是纯函数库、无 CLI 副作用，不是 CLI 壳——不违反模块边界）。


def mode_policy(mode: str | None) -> dict[str, str]:
    """按模式名取策略层；None / 未知名回落空层（一切 observe，不门禁）。"""
    return MODE_PRESETS.get(mode or "", FREE_MODE)


def core_invariants() -> dict[str, str]:
    """核心验收链的名单：九项检查 + profile.load。

    这些名字进 evaluate 的 enforced_check_results —— 它们是文档的物理不变量，
    任何模式下 FAIL/ERROR 都硬门禁，severity 调不松。profile.* 断言不在此列：
    那是契约层，严格程度归策略管。renderer 矩阵（acceptance.renderers）不在这里：
    单次运行只跑一个渲染器，「多个渲染器都通过」只能由编排层跨运行聚合，
    validate 如实把 renderers_required 记进 metadata.acceptance 供其消费。
    """
    out = {n: "ENFORCE" for n in CHECK_PROPERTY if not n.startswith("profile.")}
    out["profile.load"] = "ENFORCE"
    return out


@dataclass
class AcceptanceSummary:
    mode: str
    enforced: int = 0
    warnings: int = 0
    skipped: int = 0
    renderer: str | None = None
    visual_baseline: str | None = None
    renderers_required: list[str] = field(default_factory=list)
    # 被降级（observe）但实际 FAIL/ERROR 的检查名——不改变 verdict（那是 severity
    # 的语义），但必须可观测：否则「严格度被调松 + 核心项挂了」会在绿字里溜过去。
    observed_failures: list[str] = field(default_factory=list)


def evaluate(
    checks: Iterable,
    policy: dict[str, str],
    *,
    mode: str = "advisory",
    renderer: str | None = None,
    visual_baseline: str | None = None,
    renderers_required: Iterable[str] | None = None,
    enforced_check_results: dict[str, str] | None = None,
) -> tuple[AcceptanceSummary, str]:
    """把策略套到一组检查上，算出摘要与最终结论。

    检查项需暴露 `.name` 与 `.status`（CheckResult / 简单对象皆可）。

    严重度语义：
      - IGNORE  → 该检查强制 SKIP（不观测、不门禁）
      - OBSERVE → 如实保留结论，但绝不改变最终结论（纯信息）
      - WARN    → FAIL 降级为 WARN（亮灯不阻断）
      - ENFORCE → FAIL 仍是 FAIL（门禁）

    ERROR 不归 severity 裁决：FAIL 是「属性不符」（策略可按严格度降级），ERROR 是
    「无法判定」（检查自身崩了）——任何策略都无权把「不知道」降级成「合格」，
    除 IGNORE（明确不观测）外一律硬门禁。这是 validate 的既有铁律（查不了当 PASS
    是最大失真）在本模型里的对应约束。

    enforced_check_results：{check_name: status}——**在这些键里，severity 不作数**。
    0.9.3 接线语义：核心验收链（九项检查 + profile.load）是文档的物理不变量，
    它们的 FAIL/ERROR 永远硬门禁（FAIL 优先级高于 ERROR）；策略只裁决 profile
    契约层（`profile.*` 断言）——那里才是「严格程度按属性选」的适用范围。
    未列出的名字走正常 severity 裁决，向后兼容。

    最终结论优先级：FAIL > WARN > PASS。
    """
    enforced_map = dict(enforced_check_results or {})
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
        elif sev is Severity.ENFORCE and raw in ("FAIL", "ERROR"):
            verdict = "FAIL"
        if sev is Severity.ENFORCE and raw != "SKIP":
            summary.enforced += 1
        # observe 层的失真必须可观测：策略说「只看不拦」，但 FAIL/ERROR 静音成
        # 绿字就是撒谎（与 validate「查不了当 PASS 是失真」同一立场）。
        if sev is Severity.OBSERVE and raw in ("FAIL", "ERROR"):
            summary.observed_failures.append(name)
        # ERROR 硬门禁（fail-close）：「无法判定」不可被任何 severity 降级为合格，
        # 只有 IGNORE（明确不观测）例外。FAIL 优先级仍高于 ERROR 场景的组合判定。
        if raw == "ERROR" and sev is not Severity.IGNORE:
            verdict = "FAIL"
        # 核心不变量硬门禁：severity 降级不了物理事实（0.9.3 接线语义）。
        # ERROR 算不合格（与 validate 口径一致：查不了≠合格），但 FAIL 优先级更高。
        if enforced_map and name in enforced_map and raw in ("FAIL", "ERROR"):
            if raw == "FAIL" or enforced_map[name] != "FAIL":
                verdict = "FAIL"
    return summary, verdict
