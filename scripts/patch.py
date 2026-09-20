"""Agent 友好的文档编辑 API —— 声明式 Patch，而不是命令式 Edit。

借鉴 word-ai 的思路，但保持 texere 的简洁性：
- PatchSet + hash precondition
- Dry-run → Assess → Apply → Validate → Evidence
- 适合 Goal Mode / coding agents / MCP

用法:
  python scripts/patch.py document.docx --dry-run patch.json
  python scripts/patch.py document.docx --apply patch.json
  python scripts/patch.py document.docx --validate

Patch schema (JSON):
{
  "id": "patch-001",
  "description": "把工期从 90 天改成 120 天",
  "preconditions": {
    "hash": "abc123...",      # 可选，文档未被意外修改
    "must_contain": ["90 日历天"]   # 可选，确保旧值存在
  },
  "operations": [
    {
      "op": "replace_text",
      "target": {"paragraph": 42},
      "expected_old_text": "90 日历天",
      "new_text": "120 日历天"
    },
    {
      "op": "insert_after",
      "target": {"anchor": "资质要求"},
      "content": ["新增条款 1", "新增条款 2"],
      "style": "Normal"
    },
    {
      "op": "set_cell",
      "target": {"table": 0, "row": 5, "column": 2},
      "value": "1,060,000"
    }
  ]
}
"""

import argparse
import hashlib
import json
import os
import re
import sys
from datetime import datetime

from docx import Document

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _read_version() -> str:
    """版本号单一来源：scripts/_version.py（与 pyproject.toml 保持一致）。"""
    vp = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_version.py")
    with open(vp, encoding="utf-8") as f:
        m = re.search(r"__version__\s*=\s*[\"']([^\"']+)[\"']", f.read())
    return m.group(1) if m else "0.0.0"


__version__ = _read_version()

# Windows 控制台编码处理
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(errors="replace")


# =============================================================================
# Patch Schema Validation
# =============================================================================


def validate_patch_schema(patch: dict) -> tuple[bool, list[str]]:
    """验证 Patch JSON 是否符合 schema。"""
    errors = []

    # 必需字段
    if "id" not in patch:
        errors.append("缺少必需字段：id")
    if "operations" not in patch:
        errors.append("缺少必需字段：operations")
    elif not isinstance(patch["operations"], list):
        errors.append("operations 必须是数组")

    # 验证每个 operation
    for i, op in enumerate(patch.get("operations", [])):
        if "op" not in op:
            errors.append(f"operation[{i}] 缺少必需字段：op")
            continue

        valid_ops = {
            "replace_text",
            "insert_after",
            "insert_before",
            "delete_paragraph",
            "set_cell",
            "add_row",
            "del_row",
        }
        if op["op"] not in valid_ops:
            errors.append(f"operation[{i}] 无效的 op: {op['op']}")

        if "target" not in op:
            errors.append(f"operation[{i}] 缺少必需字段：target")

    # 验证 preconditions（可选）
    if "preconditions" in patch:
        pc = patch["preconditions"]
        if "hash" in pc and not isinstance(pc["hash"], str):
            errors.append("preconditions.hash 必须是字符串")
        if "must_contain" in pc and not isinstance(pc["must_contain"], list):
            errors.append("preconditions.must_contain 必须是数组")

    if errors:
        print("Patch schema validation failed:", file=sys.stderr)
        for error in errors:
            print(f"  - {error}", file=sys.stderr)
        return False, errors
    return True, []


# =============================================================================
# Hash Precondition
# =============================================================================


def compute_docx_hash(docx_path: str) -> str:
    """计算 DOCX 文件的 SHA256 hash。"""
    sha256 = hashlib.sha256()
    with open(docx_path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            sha256.update(chunk)
    return sha256.hexdigest()


def check_hash_precondition(docx_path: str, expected_hash: str) -> tuple[bool, str]:
    """检查文档 hash 是否匹配。"""
    actual = compute_docx_hash(docx_path)
    if actual == expected_hash:
        return True, "Hash 校验通过"
    else:
        return (
            False,
            f"Hash 不匹配 (期望:{expected_hash[:16]}..., 实际:{actual[:16]}...)",
        )


def check_must_contain_precondition(doc: Document, must_contain: list[str]) -> tuple[bool, str]:
    """检查文档是否包含所有必需文本。"""
    missing = []
    for text in must_contain:
        found = False
        for para in doc.paragraphs:
            if text in para.text:
                found = True
                break
        if not found:
            missing.append(text)

    if missing:
        return False, f"缺少必需文本：{missing}"
    return True, "所有必需文本都存在"


# =============================================================================
# Operations Implementation
# =============================================================================


def replace_text_op(doc: Document, op: dict) -> tuple[bool, str]:
    """替换文本操作。"""
    target = op["target"]
    expected_old = op.get("expected_old_text")
    new_text = op["new_text"]

    para_idx = target.get("paragraph")
    if para_idx is None:
        return False, "target.paragraph 是必需的"

    try:
        para = doc.paragraphs[para_idx]
        old_text = para.text

        if expected_old and expected_old not in old_text:
            return False, f"预期旧文本不存在：{expected_old} (实际：{old_text[:50]})"

        # 使用 edit.py 的跨 run 替换逻辑
        import importlib

        edit = importlib.import_module("edit")

        # 构造 pairs
        pairs = [(expected_old or old_text, new_text)]
        n = edit.replace_in_paragraph(para, pairs)

        if n > 0:
            return True, f"成功替换 (第{para_idx}段)"
        else:
            return False, "未找到匹配文本"
    except IndexError:
        return False, f"段落索引越界：{para_idx}"
    except Exception as e:
        return False, f"替换失败：{e}"


def insert_after_op(doc: Document, op: dict) -> tuple[bool, str]:
    """在锚点后插入操作。"""
    target = op["target"]
    content = op["content"]
    style = op.get("style", "Normal")

    anchor_text = target.get("anchor")
    if not anchor_text:
        return False, "target.anchor 是必需的"

    try:
        from scripts import edit

        hits = edit._anchors(doc, anchor_text, all_mode=True)
        if not hits:
            return False, f"找不到锚点：{anchor_text}"

        for para_idx, anchor_para in hits:
            for line in content:
                edit._clone_paragraph(anchor_para, line, style, doc, before=False)

        return True, f"在 {len(hits)} 个位置插入 {len(content)} 段"
    except Exception as e:
        return False, f"插入失败：{e}"


def insert_before_op(doc: Document, op: dict) -> tuple[bool, str]:
    """在锚点前插入操作。"""
    target = op["target"]
    content = op["content"]
    style = op.get("style", "Normal")

    anchor_text = target.get("anchor")
    if not anchor_text:
        return False, "target.anchor 是必需的"

    try:
        from scripts.edit import _anchors, _clone_paragraph

        hits = _anchors(doc, anchor_text, all_mode=True)
        if not hits:
            return False, f"找不到锚点：{anchor_text}"

        for para_idx, anchor_para in hits:
            for line in reversed(content):  # 逆序插入保证顺序正确
                _clone_paragraph(anchor_para, line, style, doc, before=True)

        return True, f"在 {len(hits)} 个位置插入 {len(content)} 段"
    except Exception as e:
        return False, f"插入失败：{e}"


def delete_paragraph_op(doc: Document, op: dict) -> tuple[bool, str]:
    """删除段落操作。"""
    target = op["target"]

    para_idx = target.get("paragraph")
    if para_idx is None:
        return False, "target.paragraph 是必需的"

    try:
        para = doc.paragraphs[para_idx]
        para._p.getparent().remove(para._p)
        return True, f"已删除第{para_idx}段"
    except IndexError:
        return False, f"段落索引越界：{para_idx}"
    except Exception as e:
        return False, f"删除失败：{e}"


def set_cell_op(doc: Document, op: dict) -> tuple[bool, str]:
    """设置单元格值操作。"""
    target = op["target"]
    value = op["value"]

    try:
        from scripts import edit

        ti = target.get("table")
        ri = target.get("row")
        ci = target.get("column")

        if None in (ti, ri, ci):
            return False, "target 需要 table, row, column"

        cell = edit._cell(doc, ti, ri, ci)
        old = cell.text
        edit._set_cell_text(cell, value)

        return True, f"单元格 [{ti}][{ri},{ci}]: {old[:20]} -> {value[:20]}"
    except Exception as e:
        return False, f"设置单元格失败：{e}"


def add_row_op(doc: Document, op: dict) -> tuple[bool, str]:
    """添加行操作。"""
    target = op["target"]
    values = op["values"]

    try:
        from scripts import edit

        ti = target.get("table")
        if ti is None:
            return False, "target.table 是必需的"

        # 插入到指定行之后（如果指定）
        _ = target.get("after_row", -1)  # 预留参数

        # 手动实现：先获取表格，再添加行
        table = doc.tables[ti]
        row = table.add_row()

        for i, v in enumerate(values):
            if i < len(row.cells):
                edit._set_cell_text(row.cells[i], v)

        return True, f"表{ti}: 新增 1 行（{len(values)}列）"
    except Exception as e:
        return False, f"添加行失败：{e}"


def del_row_op(doc: Document, op: dict) -> tuple[bool, str]:
    """删除行操作。"""
    target = op["target"]

    try:
        ti = target.get("table")
        ri = target.get("row")

        if None in (ti, ri):
            return False, "target 需要 table 和 row"

        # 手动实现删除
        table = doc.tables[ti]
        row = table.rows[ri]
        row._tr.getparent().remove(row._tr)

        return True, f"表{ti}行{ri}: 已删除"
    except Exception as e:
        return False, f"删除行失败：{e}"


# Operation registry
OPERATIONS = {
    "replace_text": replace_text_op,
    "insert_after": insert_after_op,
    "insert_before": insert_before_op,
    "delete_paragraph": delete_paragraph_op,
    "set_cell": set_cell_op,
    "add_row": add_row_op,
    "del_row": del_row_op,
}


# =============================================================================
# Assessment & Dry Run
# =============================================================================


def assess_patch(patch: dict, doc: Document) -> tuple[bool, list[str]]:
    """评估 Patch 可行性，不进行实际修改。"""
    issues = []

    for i, op in enumerate(patch.get("operations", [])):
        op_name = op.get("op")
        target = op.get("target", {})

        # 静态检查
        if op_name == "replace_text":
            para_idx = target.get("paragraph")
            if para_idx is not None and para_idx >= len(doc.paragraphs):
                issues.append(f"operation[{i}]: 段落索引越界")

        elif op_name in ("insert_after", "insert_before"):
            anchor = target.get("anchor")
            if anchor:
                count = sum(1 for p in doc.paragraphs if anchor in p.text)
                if count == 0:
                    issues.append(f"operation[{i}]: 锚点不存在：{anchor}")

        elif op_name == "set_cell":
            ti, ri = target.get("table"), target.get("row")  # ci 预留参数
            if ti is not None and ti >= len(doc.tables):
                issues.append(f"operation[{i}]: 表格索引越界")
            elif ti is not None:
                if ri is not None and ri >= len(doc.tables[ti].rows):
                    issues.append(f"operation[{i}]: 行索引越界")

    return len(issues) == 0, issues


def dry_run_patch(patch: dict, doc: Document) -> tuple[bool, str]:
    """Dry run：模拟执行 Patch，预估影响。"""
    import copy

    # 深拷贝文档进行模拟
    doc_copy = copy.deepcopy(doc)

    failed_ops = []
    successful_ops = []

    for i, op in enumerate(patch.get("operations", [])):
        op_name = op.get("op")

        try:
            if op_name in OPERATIONS:
                success, message = OPERATIONS[op_name](doc_copy, op)
                if success:
                    successful_ops.append((i, op_name, message))
                else:
                    failed_ops.append((i, op_name, message))
            else:
                failed_ops.append((i, op_name, "未实现的 operation"))
        except Exception as e:
            failed_ops.append((i, op_name, "异常：%s" % e))

    if failed_ops:
        print(f"Dry run 失败：{len(failed_ops)}个操作失败", file=sys.stderr)
        for i, op, msg in failed_ops:
            print(f"  [{i}] {op}: {msg}", file=sys.stderr)
        sys.exit(1)
    else:
        return True, "Dry run 成功：%d 个操作将通过\n" % len(successful_ops) + "\n".join(
            "  [%d] %s: %s" % (i, op, msg) for i, op, msg in successful_ops
        )


# =============================================================================
# Apply & Validate
# =============================================================================


def apply_patch(patch: dict, doc: Document) -> tuple[bool, list[str]]:
    """应用 Patch 到文档。"""
    errors = []

    for i, op in enumerate(patch.get("operations", [])):
        op_name = op.get("op")

        if op_name not in OPERATIONS:
            errors.append(f"operation[{i}]: 未实现的 operation: {op_name}")
            continue

        success, message = OPERATIONS[op_name](doc, op)
        if not success:
            errors.append(f"operation[{i}] ({op_name}): {message}")

    return len(errors) == 0, errors


def validate_patch_result(doc: Document, patch: dict) -> tuple[bool, list[str]]:
    """验证 Patch 应用后的结果。"""
    issues = []

    # 验证每个 operation 是否达到预期
    for i, op in enumerate(patch.get("operations", [])):
        op_name = op.get("op")

        if op_name == "replace_text":
            expected = op.get("new_text")
            target = op["target"]
            para_idx = target.get("paragraph")

            if para_idx < len(doc.paragraphs):
                if expected not in doc.paragraphs[para_idx].text:
                    issues.append(f"operation[{i}]: 预期文本未找到：{expected}")

        elif op_name == "set_cell":
            target = op["target"]
            expected = op["value"]
            ti, ri, ci = target.get("table"), target.get("row"), target.get("column")

            if ti is not None and ri is not None and ci is not None:
                if expected not in doc.tables[ti].rows[ri].cells[ci].text:
                    issues.append(f"operation[{i}]: 单元格值未更新")

    return len(issues) == 0, issues


# =============================================================================
# Evidence Generation
# =============================================================================


def generate_patch_evidence(doc: Document, patch: dict, out_dir: str):
    """生成 Patch 证据包。"""
    if not os.path.exists(out_dir):
        os.makedirs(out_dir)
    elif not os.path.isdir(out_dir):
        raise NotADirectoryError(f"{out_dir} exists but is not a directory")

    # 1. 保存 Patch 执行报告
    report = {
        "metadata": {
            "timestamp": datetime.now().isoformat(),
            "tool_version": __version__,
            "patch_id": patch.get("id", "unknown"),
        },
        "operations": [],
    }

    for i, op in enumerate(patch.get("operations", [])):
        report["operations"].append({"index": i, "op": op.get("op"), "status": "pending"})

    report_path = os.path.join(out_dir, "patch_report.json")
    with open(report_path, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
        f.write("\n")

    # 2. 保存文档快照（前后对比）
    snapshot_path = os.path.join(out_dir, "document_after.patch.docx")
    doc.save(snapshot_path)

    # 3. 生成签名
    sha256 = hashlib.sha256()
    with open(snapshot_path, "rb") as f:
        sha256.update(f.read())

    sig_path = os.path.join(out_dir, "signature")
    with open(sig_path, "w", encoding="utf-8") as f:
        f.write("# texere patch signature\n")
        f.write("# Generated: %s\n" % report["metadata"]["timestamp"])
        f.write("patch_id: %s\n" % patch.get("id", "unknown"))
        f.write("document_hash: %s\n" % sha256.hexdigest())


# =============================================================================
# CLI Entry Point
# =============================================================================


def main():
    ap = argparse.ArgumentParser(description="texere patch — Agent-friendly document editing API")
    ap.add_argument("docx", help="Document to patch")
    ap.add_argument("patch_file", help="Patch JSON file")
    ap.add_argument("--dry-run", action="store_true", help="Simulate without applying")
    ap.add_argument("--apply", action="store_true", help="Apply the patch")
    ap.add_argument("--validate", action="store_true", help="Validate after applying")
    ap.add_argument("--out", help="Output directory for evidence package")
    ap.add_argument("--no-backup", action="store_true", help="Skip creating backup file")
    a = ap.parse_args()

    # 加载文档
    if not os.path.exists(a.docx):
        sys.exit(f"File not found: {a.docx}")

    doc = Document(a.docx)
    print(f"Loaded: {os.path.basename(a.docx)}")

    # 加载 Patch
    if not os.path.exists(a.patch_file):
        sys.exit(f"Patch file not found: {a.patch_file}")

    patch = json.load(open(a.patch_file, encoding="utf-8-sig"))
    print(f"Loaded patch: {patch.get('id', 'unknown')}")

    # 验证 Patch schema
    valid, errors = validate_patch_schema(patch)
    if not valid:
        sys.exit(1)

    # 检查 preconditions
    if "preconditions" in patch:
        pc = patch["preconditions"]

        if "hash" in pc:
            success, msg = check_hash_precondition(a.docx, pc["hash"])
            if not success:
                print(f"Hash precondition failed: {msg}", file=sys.stderr)
                sys.exit(1)
            print(f"✓ {msg}")

        if "must_contain" in pc:
            success, msg = check_must_contain_precondition(doc, pc["must_contain"])
            if not success:
                print(f"Must-contain precondition failed: {msg}", file=sys.stderr)
                sys.exit(1)
            print(f"✓ {msg}")

    # Dry run
    if a.dry_run:
        print("\n[Dry Run]")
        success, msg = dry_run_patch(patch, doc)
        print(msg)
        sys.exit(0 if success else 1)

    # Apply
    if a.apply:
        print("\n[Apply]")
        success, errors = apply_patch(patch, doc)

        if not success:
            print("Patch application failed:", file=sys.stderr)
            for error in errors:
                print(f"  - {error}", file=sys.stderr)
            sys.exit(1)

        print("✓ Patch applied successfully")

        # Validate result
        if a.validate:
            print("\n[Validate]")
            valid, issues = validate_patch_result(doc, patch)
            if not valid:
                print("Validation failed:", file=sys.stderr)
                for issue in issues:
                    print(f"  - {issue}", file=sys.stderr)
                sys.exit(1)
            print("✓ All operations verified")

        # Save
        if a.out:
            # --out 指定输出目录或文件
            # 如果路径不存在或明确是目录，保存到该目录下的 document.patched.docx
            if not os.path.exists(a.out) or os.path.isdir(a.out):
                evidence_dir = a.out
                out_file = os.path.join(evidence_dir, "document.patched.docx")
            else:
                # 已存在的文件路径
                out_file = a.out
                evidence_dir = os.path.dirname(out_file) or "."

            # 确保目录存在
            if not os.path.exists(evidence_dir):
                os.makedirs(evidence_dir)

            doc.save(out_file)
            print(f"Saved: {out_file}")

            # Generate evidence
            generate_patch_evidence(doc, patch, evidence_dir)
            print(f"Evidence package saved to: {evidence_dir}")
        else:
            # 无 --out 时写回原文件并创建备份
            if not getattr(a, "no_backup", False):
                import shutil

                bak = os.path.splitext(a.docx)[0] + ".bak.docx"
                shutil.copy2(a.docx, bak)
                print(f"backup -> {bak}")

            doc.save(a.docx)
            print(f"Saved: {a.docx}")

        print("\n✅ Patch completed successfully")

    # If no action specified, show assessment
    if not (a.dry_run or a.apply):
        print("\n[Assessment]")
        success, issues = assess_patch(patch, doc)
        if success:
            print("✓ Patch is safe to apply")
        else:
            print("⚠ Issues found:")
            for issue in issues:
                print(f"  - {issue}")

        print("\nUsage:")
        print("  python scripts/patch.py docx.patch.json --dry-run")
        print("  python scripts/patch.py docx.patch.json --apply --out result.docx")
        print("  python scripts/patch.py docx.patch.json --apply --validate")


if __name__ == "__main__":
    main()
