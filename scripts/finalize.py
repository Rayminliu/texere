"""Word 验收 + 导出 PDF 的 CLI 壳：python scripts/finalize.py <in.docx> [out.pdf] [--save-updated-fields]

实现已抽到 scripts/renderers.py 的 WordRenderer（RendererAdapter 的一个实现）。
本文件只负责参数解析与「人读」输出，行为与原 finalize.py 完全一致：
- 验收信号：Word 打不开 = OOXML 结构有问题（立即失败）；
- 刷新目录域并重排页码后导出 PDF；
- 默认只读验收（绝不写回输入 docx），只有 --save-updated-fields 才写回。
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from renderers import WordRenderer

# 解析参数：位置参数 <in.docx> [out.pdf]，开关 --save-updated-fields
save_updated = "--save-updated-fields" in sys.argv
positional = [a for a in sys.argv[1:] if not a.startswith("--")]
if len(positional) < 1:
    sys.exit("用法: python scripts/finalize.py <in.docx> [out.pdf] [--save-updated-fields]")
SRC = positional[0]
PDF = positional[1] if len(positional) > 1 else os.path.splitext(SRC)[0] + ".pdf"

res = WordRenderer().render(SRC, PDF, save_updated_fields=save_updated)

# 保留与原 finalize.py 完全一致的人读输出，避免破坏任何下游解析。
print("Engine    :", res.renderer_name, res.renderer_version, "|", res.engine_path)
s = res.stats
print("Pages      :", s.get("pages"))
print("Words      :", s.get("words"))
print("Tables     :", s.get("tables"))
print("InlineShapes:", s.get("inline_shapes"))
print("Sections   :", s.get("sections"))
print("PDF written:", os.path.exists(PDF))
if res.errors:
    print("Errors     :", "; ".join(res.errors))
    sys.exit(1)
print("OK")
