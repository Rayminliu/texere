"""texere 脚本共享基础设施：版本、控制台编码、SHA256、子进程环境。

各脚本（render / validate / patch / edit / snapshot / …）统一从这里引用，
避免逐字复制导致的漂移。
"""

import hashlib
import os
import sys

from _version import __version__  # noqa: F401 — re-export

# ---------------------------------------------------------------------------
# 控制台编码
# ---------------------------------------------------------------------------
# Windows 控制台默认 GBK，子进程输出里若出现无法编码的字符会抛 UnicodeEncodeError。
# 统一 reconfigure 为 utf-8 + errors="replace"，让不可编码字符降级为 ? 而非崩溃。


def force_utf8_stdio():
    """在脚本入口处调用一次，确保 stdout/stderr 不会因编码问题崩溃。"""
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


# ---------------------------------------------------------------------------
# 子进程 UTF-8 环境
# ---------------------------------------------------------------------------
# 强制子进程管道输出 UTF-8，整条链编码统一。
UTF8_ENV = {**os.environ, "PYTHONIOENCODING": "utf-8"}


# ---------------------------------------------------------------------------
# SHA256 文件哈希
# ---------------------------------------------------------------------------


def sha256_file(path):
    """流式计算文件 SHA256；path 为空或文件不存在时返回 None。"""
    if not path or not os.path.exists(path):
        return None
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()
