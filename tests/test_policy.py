"""policy 模型单测（纯函数，无需 Word / 渲染器）。

守住：四级严重度语义、属性→检查反查、模式预设、renderer 矩阵解析。
"""

from collections import namedtuple

from scripts import policy

_C = namedtuple("_C", "name status message")


def _checks(*pairs):
    return [_C(n, s, "") for n, s in pairs]


def test_severity_values_and_gates():
    assert policy.Severity.IGNORE.value == "ignore"
    assert policy.Severity.OBSERVE.value == "observe"
    assert policy.Severity.WARN.value == "warn"
    assert policy.Severity.ENFORCE.value == "enforce"
    assert policy.Severity.IGNORE.gates is False
    assert policy.Severity.OBSERVE.gates is False
    assert policy.Severity.WARN.gates is True
    assert policy.Severity.ENFORCE.gates is True


def test_severity_for_check_uses_property_registry():
    pol = {"body_font": "enforce"}
    assert policy.severity_for_check("profile.body_font", pol) is policy.Severity.ENFORCE
    # toc 属性同时管辖 profile.toc 与 toc_field
    assert policy.severity_for_check("toc_field", {"toc": "warn"}) is policy.Severity.WARN
    # 未登记的检查回落到默认 observe
    assert policy.severity_for_check("mystery_check", {}) is policy.Severity.OBSERVE


def test_free_mode_never_gates():
    pol = policy.MODE_PRESETS["free"]
    summary, verdict = policy.evaluate(
        _checks(("profile.body_font", "FAIL"), ("blank_pages", "FAIL")),
        pol,
        mode="free",
    )
    assert verdict == "PASS"  # 自由模式：不门禁
    assert summary.enforced == 0


def test_advisory_mode_downgrades_fail_to_warn():
    pol = policy.MODE_PRESETS["advisory"]
    summary, verdict = policy.evaluate(
        _checks(("profile.body_font", "FAIL"), ("blank_pages", "PASS")),
        pol,
        mode="advisory",
    )
    assert verdict == "WARN"  # 建议模式：亮黄灯不阻断
    assert summary.warnings == 1


def test_strict_mode_fails_on_enforced_fail():
    pol = policy.MODE_PRESETS["strict"]
    summary, verdict = policy.evaluate(
        _checks(("profile.body_font", "FAIL"), ("blank_pages", "PASS")),
        pol,
        mode="strict",
    )
    assert verdict == "FAIL"  # 严格模式：硬失败
    assert summary.enforced >= 1


def test_strict_mode_fails_on_enforced_error():
    """ERROR 与 FAIL 同等对待，不允许『查不了=合格』。"""
    pol = policy.MODE_PRESETS["strict"]
    summary, verdict = policy.evaluate(
        _checks(("profile.body_font", "ERROR"), ("blank_pages", "PASS")),
        pol,
        mode="strict",
    )
    assert verdict == "FAIL"


def test_ignore_skips_and_ignores_fail():
    pol = {"blank_pages": "ignore"}
    summary, verdict = policy.evaluate(
        _checks(("blank_pages", "FAIL"), ("package_integrity", "PASS")),
        pol,
        mode="advisory",
    )
    assert summary.skipped == 1
    assert verdict == "PASS"  # 被 ignore 的 FAIL 不影响结论


def test_observe_reports_but_never_gates():
    pol = {"blank_pages": "observe"}
    summary, verdict = policy.evaluate(
        _checks(("blank_pages", "FAIL")),
        pol,
        mode="advisory",
    )
    assert verdict == "PASS"  # observe：如实报告但不门禁
    assert summary.warnings == 0


def test_required_renderers_matrix():
    assert policy.required_renderers({"renderers": ["word", "wps"]}) == ["word", "wps"]
    assert policy.required_renderers(None) == []
    assert policy.required_renderers({}) == []


# ------------------------------- 0.9.3 接线语义：enforced_check_results / core_invariants


def test_core_invariants_exclude_contract_layer():
    """核心名单 = 九项 + profile.load；profile.* 契约层不进去（severity 归策略管）。"""
    inv = policy.core_invariants()
    for n in (
        "package_integrity",
        "source_content",
        "image_embedding",
        "section_count",
        "toc_field",
        "page_numbering",
        "blank_pages",
        "renderer_acceptance",
        "visual_drift",
        "profile.load",
    ):
        assert n in inv, f"核心不变量缺 {n}"
    for contract in (
        "profile.page",
        "profile.body_font",
        "profile.heading",
        "profile.table",
        "profile.toc",
    ):
        assert contract not in inv, f"契约层 {contract} 不该进核心名单"


def test_enforced_results_override_observe_downgrade():
    """策略把核心项 observe 掉也白搭：enforced_check_results 里 FAIL 照样 FAIL（接线语义）。"""
    pol = {"package_integrity": "observe"}
    summary, verdict = policy.evaluate(
        _checks(("package_integrity", "FAIL")),
        pol,
        mode="free",
        enforced_check_results=policy.core_invariants(),
    )
    assert verdict == "FAIL"  # severity 降级不了物理事实
    assert "package_integrity" in summary.observed_failures  # 但降级必须可观测


def test_enforced_results_error_also_gates_but_fail_wins():
    """ERROR 同样硬门禁（查不了≠合格）；同批同时出现 FAIL 与 ERROR 时 FAIL 定级。"""
    _s, v1 = policy.evaluate(
        _checks(("package_integrity", "ERROR")),
        {},
        enforced_check_results={"package_integrity": "ENFORCE"},
    )
    assert v1 == "FAIL"
    _s2, v2 = policy.evaluate(
        _checks(("package_integrity", "ERROR"), ("toc_field", "FAIL")),
        {},
        enforced_check_results=policy.core_invariants(),
    )
    assert v2 == "FAIL"


def test_enforced_map_ignores_pass_skip_and_absent():
    """enforced 只对 FAIL/ERROR 起作用；PASS/SKIP 与不在名单里的检查走正常裁决。"""
    _s, v = policy.evaluate(
        _checks(
            ("package_integrity", "PASS"), ("blank_pages", "SKIP"), ("profile.body_font", "FAIL")
        ),
        {"body_font": "warn"},
        mode="advisory",
        enforced_check_results=policy.core_invariants(),
    )
    assert v == "WARN"  # 契约层 warn 生效；PASS/SKIP 不触发 enforced 分支


def test_contract_layer_still_governed_by_severity():
    """接线不越界：profile.body_font FAIL 在 advisory 下仍降 WARN（策略管契约层）。"""
    summary, verdict = policy.evaluate(
        _checks(("profile.body_font", "FAIL")),
        policy.MODE_PRESETS["advisory"],
        mode="advisory",
        enforced_check_results=policy.core_invariants(),
    )
    assert verdict == "WARN"
    assert summary.warnings == 1


def test_error_is_fail_close_under_any_severity():
    """ERROR（无法判定）不归 severity 裁决：advisory/observe 都拦，只有 ignore 放行。

    FAIL 是「属性不符」（策略可按严格度降级），ERROR 是「检查崩了不知道结果」——
    把不知道降级成合格，正是 validate 立身铁律禁止的失真。
    """
    for pol in (policy.MODE_PRESETS["advisory"], {"body_font": "observe"}, {}):
        _s, v = policy.evaluate(_checks(("profile.body_font", "ERROR")), pol, mode="advisory")
        assert v == "FAIL", f"policy={pol} 下 ERROR 未硬门禁"
    _s2, v2 = policy.evaluate(
        _checks(("profile.body_font", "ERROR")),
        {"body_font": "ignore"},
        mode="advisory",
    )
    assert v2 == "PASS"  # ignore 是明确声明「不观测」，不是失真


def test_error_from_core_invariants_gates_even_in_free_mode():
    """free 模式 + 核心项 ERROR：双保险（fail-close 与 enforced_map）都必须拦。"""
    _s, v = policy.evaluate(
        _checks(("package_integrity", "ERROR")),
        {},
        mode="free",
        enforced_check_results=policy.core_invariants(),
    )
    assert v == "FAIL"


def test_mode_policy_resolution():
    assert policy.mode_policy("strict") is policy.MODE_PRESETS["strict"]
    assert policy.mode_policy("advisory") is policy.MODE_PRESETS["advisory"]
    assert policy.mode_policy("free") == {}
    assert policy.mode_policy(None) == {}
    assert policy.mode_policy(" nonsense ") == {}
