"""policy 模型单测（纯函数，无需 Word / 渲染器）。

守住：四级严重度语义、属性→检查反查、模式预设、覆盖分层、renderer 矩阵解析。
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


def test_resolve_policy_layers_override():
    base = {"page": "enforce", "body_font": "enforce"}
    override = {"body_font": "warn"}  # 任务临时把正文字体降级为 warn
    merged = policy.resolve_policy(base, override)
    assert merged["page"] == "enforce"
    assert merged["body_font"] == "warn"


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
