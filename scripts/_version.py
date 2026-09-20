"""版本号的单一来源。

发版时改这里，并与 pyproject.toml 的 [project] version 保持一致
（tests/test_version.py 会校验两处一致，防止漏改）。

render.py / validate.py / patch.py 都从这里读取版本，不再各自硬编码。
"""

__version__ = "0.4.0"
