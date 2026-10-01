"""Evidence / Manifest 能力的纯逻辑核（横切，不反向控制其它能力）。

从 validate.py 外提：报告骨架构建 _new_report（含 pandoc provenance memo
_pandoc_version）、把单条 CheckResult 记进报告并累计 summary 的 _tally、落盘
report.json 的 _write_report，以及显式证据清单 build_manifest 与写盘壳
_write_signature。validate.py 作为 thin CLI 壳 re-export 这些符号，保持
report.json / signature 输出逐字节不变（build_manifest 返回有序行，
_write_signature 用 writelines 落盘，与原逐行 f.write 结果逐字节等价）。

裸名 sibling import，禁止反向 import validate。
"""

import json
import os
import sys
from datetime import datetime

from _shared import __version__
from _shared import sha256_file as _sha256_file
from _verify import PASS, SKIP

_PANDOC_VERSION_CACHE = None


def _pandoc_version() -> str:
    """pandoc 版本进证据：provenance 缺「转换器是谁」就少一环（外部审计 R2）。

    memo：spawn 一次 pandoc 约百毫秒，同一进程内多次写证据不该重复探；
    encoding 显式 UTF-8：GBK locale 下 text=True 解码子进程输出会炸。
    """
    global _PANDOC_VERSION_CACHE
    if _PANDOC_VERSION_CACHE is not None:
        return _PANDOC_VERSION_CACHE
    try:
        import subprocess

        out = subprocess.run(
            ["pandoc", "--version"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            # 探测必须有时限：这条在 report.json 的写入关键路径上（metadata.pandoc_version），
            # pandoc 挂起会让整个证据包永不产出。超时走下面的 except → "unknown"，
            # 不改变任何检查结论，只是不再永挂（与 renderers._soffice_version 同档）。
            timeout=30,
        ).stdout
        _PANDOC_VERSION_CACHE = (out.splitlines() or ["unknown"])[0].strip()
    except Exception:
        _PANDOC_VERSION_CACHE = "unknown"
    return _PANDOC_VERSION_CACHE


def _scrub_path(arg: str) -> str:
    """剥掉路径的目录部分只留文件名；剥空了回落到原值（不许把参数弄丢）。"""
    return os.path.basename(arg.rstrip("/\\")) or arg


def _cli_line() -> str:
    """命令行进证据，但不带目录：证据是对外交付物。

    以前直接 `" ".join(sys.argv)`，真跑一次的 report.json 里就是
    `...\\scripts\\validate.py C:\\Users\\<user>\\AppData\\Local\\Temp\\...`——
    把操作者的用户名、工作目录与临时路径一并交给客户/审计方。选项名与
    文件名保留（它们才回答「用哪些开关跑的哪份文档」）。
    """
    parts = []
    for i, a in enumerate(sys.argv):
        if i == 0:
            parts.append(_scrub_path(a))
        elif a.startswith("-") and "=" in a:
            flag, _, val = a.partition("=")
            parts.append(f"{flag}={_scrub_path(val)}")
        elif a.startswith("-"):
            parts.append(a)
        else:
            parts.append(_scrub_path(a))
    return " ".join(parts)


def _new_report(docx_path: str, profile: dict, config_path: str, reference_doc: str) -> dict:
    """报告骨架：metadata + 空 checks + 归零的 summary。"""
    return {
        "metadata": {
            "timestamp": datetime.now().isoformat(),
            "document": os.path.basename(docx_path),
            "tool_version": __version__,
            "profile": profile or {},
            # provenance：证据要能回答「哪个产物 + 哪份契约 + 哪个转换器」（外部审计 R2）
            "cli": _cli_line(),
            "pandoc_version": _pandoc_version(),
            "config_sha256": _sha256_file(config_path),
            "reference_sha256": _sha256_file(reference_doc),
        },
        "checks": {},
        # skipped 独立于 passed：「没检查」不许冒充实测通过。
        "summary": {"total": 0, "passed": 0, "failed": 0, "skipped": 0},
    }


def _tally(report: dict, res, name: str = None):
    """把一条 CheckResult 记进 report 并累计 summary。

    FAIL 与 ERROR 都算不合格（检查崩了不等于文档合格）；SKIP 单独计数。
    """
    key = name or res.name
    report["checks"][key] = {
        "status": res.status,
        "message": res.message,
        "evidence": res.evidence,
    }
    report["summary"]["total"] += 1
    if res.status == PASS:
        report["summary"]["passed"] += 1
    elif res.status == SKIP:
        report["summary"]["skipped"] += 1
    else:
        report["summary"]["failed"] += 1


def build_manifest(report: dict, out_dir: str, docx_path: str, report_path: str) -> list:
    """显式证据清单：返回 signature 文件的有序文本行（每行自带结尾换行）。

    这是 checksum manifest，不是密码学签名：没有密钥，任何人都能重算这些 hash。
    它证明的是「这份证据记录了哪个产物」，不是「这份证据没被改过」。
    _write_report 必须先落盘（它的 hash 要进清单），顺序不能反。
    全量文件清单：report.json + 全部截图的 sha256——证据目录里任何一个文件
    被事后改动都可检出（此前只盖 docx 与 report，截图是漏项）。
    signature 自身排除（自引用不可行）；排序保证清单确定性。
    """
    evidence_files = {}
    for name in sorted(os.listdir(out_dir)):
        p = os.path.join(out_dir, name)
        if name == "signature" or not os.path.isfile(p):
            continue
        evidence_files[name] = _sha256_file(p)

    lines = [
        "# texere validation manifest (checksums, not a cryptographic signature)\n",
        "# Generated: %s\n" % report["metadata"]["timestamp"],
        "document_hash: %s\n" % _sha256_file(docx_path),
        "report_hash: %s\n" % _sha256_file(report_path),
        "checks_passed: %d/%d\n" % (report["summary"]["passed"], report["summary"]["total"]),
        # 单独记 skipped：证据里也要能看出「9 项里有几项其实没查」
        "checks_skipped: %d\n" % report["summary"]["skipped"],
        "generated_at: %s\n" % report["metadata"]["timestamp"],
        "tool_version: %s\n" % __version__,
        "pandoc_version: %s\n" % report["metadata"].get("pandoc_version", "unknown"),
    ]
    # 契约指纹：config 与 ref.docx 决定版式——证据缺了它们就缺「哪个契约」这一环
    if report["metadata"].get("config_sha256"):
        lines.append("config_sha256: %s\n" % report["metadata"]["config_sha256"])
    if report["metadata"].get("reference_sha256"):
        lines.append("reference_sha256: %s\n" % report["metadata"]["reference_sha256"])
    for name, digest in evidence_files.items():
        lines.append("file[%s]: %s\n" % (name, digest))
    return lines


def _write_report(report: dict, out_dir: str) -> str:
    """先落盘 report.json —— 它的 hash 要进证据清单，顺序不能反。"""
    report_path = os.path.join(out_dir, "report.json")
    with open(report_path, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
        f.write("\n")
    return report_path


def _write_signature(report: dict, out_dir: str, docx_path: str, report_path: str):
    """证据清单写盘壳：把 build_manifest 的有序行落盘（逐字节等价原逐行写）。"""
    sig_path = os.path.join(out_dir, "signature")
    with open(sig_path, "w", encoding="utf-8") as f:
        f.writelines(build_manifest(report, out_dir, docx_path, report_path))
