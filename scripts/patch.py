"""Agent 友好的文档编辑 API —— 声明式 Patch，而不是命令式 Edit。

借鉴 word-ai 的思路，但保持 texere 的简洁性：
- PatchSet + hash precondition
- Dry-run → Assess → Apply → Validate → Evidence
- 适合 Goal Mode / coding agents / MCP

用法:
  python scripts/patch.py document.docx patch.json --dry-run
  python scripts/patch.py document.docx patch.json --apply --out result.docx
  python scripts/patch.py document.docx patch.json --apply --validate

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
    },
    {
      "op": "add_row",
      "target": {"table": 0, "after_row": 3},
      "values": ["接入层", "设备", "320,000"]
    }
  ]
}

`insert_after` / `insert_before` 的 `content` 顺序就是它们在文档里的出现顺序：
`"content": ["A", "B"]` 得到 A 在上、B 在下。此前多行插入会被写成倒序（已修正），
按旧行为写的补丁需要把 content 重排回正常阅读顺序。

`add_row` 的 `target.after_row` 是 0 基行号，新行插在该行之后；
省略或取 -1 表示追加到表尾。这个字段此前只是「预留参数」——schema 收下、
实现忽略，Agent 按文档写了会被静默追加到表尾，现在按声明语义落地。

实现分层：声明式操作的执行内核（schema 校验 / 前置条件 / 7 个 op 原语 /
assess·dry_run·apply·validate 编排）已下沉到 `_mutate.py`，统一返回
`MutationResult`。本文件是 thin CLI 壳——re-export 同名符号、把 `MutationResult`
降级回历史一致的 `tuple[bool, str]`，并保留 argparse / sys.exit / stdout 措辞 /
证据写盘等 CLI 语义。
"""

import argparse
import json
import os
import sys
from datetime import datetime

from _mutate import (
    _edit_module,  # noqa: F401  re-export：test_patch 哨兵直访 p._edit_module
    apply_patch,
    assess_patch,
    check_hash_precondition,
    check_must_contain_precondition,
    compute_docx_hash,
    dry_run_patch,
    validate_patch_result,
    validate_patch_schema,
)
from _mutate import add_row_op as _add_row_op
from _mutate import del_row_op as _del_row_op
from _mutate import delete_paragraph_op as _delete_paragraph_op
from _mutate import insert_after_op as _insert_after_op
from _mutate import insert_before_op as _insert_before_op
from _mutate import replace_text_op as _replace_text_op
from _mutate import set_cell_op as _set_cell_op
from _shared import __version__, force_utf8_stdio
from docx import Document

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


# =============================================================================
# Operations —— 壳层适配：_mutate 的 MutationResult 降级回 tuple[bool, str]
# （签名与措辞与历史逐字节一致，test_patch.py 的进程内用例不改一行即绿）
# =============================================================================


def replace_text_op(doc: Document, op: dict) -> tuple[bool, str]:
    """替换文本操作。"""
    res = _replace_text_op(doc, op)
    return res.ok, res.message


def insert_after_op(doc: Document, op: dict) -> tuple[bool, str]:
    """在锚点后插入操作。"""
    res = _insert_after_op(doc, op)
    return res.ok, res.message


def insert_before_op(doc: Document, op: dict) -> tuple[bool, str]:
    """在锚点前插入操作。"""
    res = _insert_before_op(doc, op)
    return res.ok, res.message


def delete_paragraph_op(doc: Document, op: dict) -> tuple[bool, str]:
    """删除段落操作。"""
    res = _delete_paragraph_op(doc, op)
    return res.ok, res.message


def set_cell_op(doc: Document, op: dict) -> tuple[bool, str]:
    """设置单元格值操作。"""
    res = _set_cell_op(doc, op)
    return res.ok, res.message


def add_row_op(doc: Document, op: dict) -> tuple[bool, str]:
    """添加行操作：插在 target.after_row（0 基）之后，缺省追加到表尾。"""
    res = _add_row_op(doc, op)
    return res.ok, res.message


def del_row_op(doc: Document, op: dict) -> tuple[bool, str]:
    """删除行操作。"""
    res = _del_row_op(doc, op)
    return res.ok, res.message


# =============================================================================
# Evidence Generation
# =============================================================================


def generate_patch_evidence(
    doc: Document, patch: dict, out_dir: str, before_sha256: str = None, verify_result: str = None
):
    """生成 Patch 证据包。

    编辑链的契约是「只改指定处、其余字节不动」——证据必须同时记录
    before_sha256 与事后 document_hash，才能自证「其余字节确实没动」
    （外部审计 R3 #1：渲染链有 report+signature，编辑链不能只靠 stdout）。
    """
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
            "before_sha256": before_sha256,
            "verify": verify_result,
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
    document_hash = compute_docx_hash(snapshot_path)

    sig_path = os.path.join(out_dir, "signature")
    with open(sig_path, "w", encoding="utf-8") as f:
        f.write("# texere patch signature\n")
        f.write("# Generated: %s\n" % report["metadata"]["timestamp"])
        f.write("patch_id: %s\n" % patch.get("id", "unknown"))
        if before_sha256:
            f.write("before_sha256: %s\n" % before_sha256)
        f.write("document_hash: %s\n" % document_hash)
        if verify_result:
            f.write("verify: %s\n" % verify_result)


# =============================================================================
# CLI Entry Point
# =============================================================================


def main():
    # 副作用只在入口执行：被 import（工具复用/测试）时不碰宿主 stdio
    force_utf8_stdio()

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

    with open(a.patch_file, encoding="utf-8-sig") as f:
        patch = json.load(f)
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
        print(msg, file=sys.stdout if success else sys.stderr)
        sys.exit(0 if success else 1)

    # Apply
    if a.apply:
        print("\n[Apply]")
        before_sha = compute_docx_hash(a.docx)  # 变更前指纹进证据（外部审计 R3 #1）
        success, errors = apply_patch(patch, doc)

        if not success:
            print("Patch application failed:", file=sys.stderr)
            for error in errors:
                print(f"  - {error}", file=sys.stderr)
            sys.exit(1)

        print("✓ Patch applied successfully")

        # Validate result
        verify_result = None
        if a.validate:
            print("\n[Validate]")
            valid, issues = validate_patch_result(doc, patch)
            if not valid:
                print("Validation failed:", file=sys.stderr)
                for issue in issues:
                    print(f"  - {issue}", file=sys.stderr)
                sys.exit(1)
            print("✓ All operations verified")
            verify_result = "PASS"

        # Save
        if a.out:
            # --out 指定输出目录或文件
            # 有扩展名时视为文件路径；不存在且无扩展名时视为目录
            out_ext = os.path.splitext(a.out)[1].lower()
            if out_ext in (".docx", ".doc"):
                out_file = a.out
                evidence_dir = os.path.dirname(os.path.abspath(out_file)) or "."
            elif not os.path.exists(a.out) or os.path.isdir(a.out):
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
            generate_patch_evidence(
                doc, patch, evidence_dir, before_sha256=before_sha, verify_result=verify_result
            )
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
