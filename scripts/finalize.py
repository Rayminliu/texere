"""Word COM 验收 + 导出 PDF：python scripts/finalize.py <in.docx> [out.pdf] [--save-updated-fields]

验收信号：Word 打不开 = OOXML 结构有问题（立即失败）；同时刷新目录域并重排页码后导出 PDF。

默认**只读验收**：在临时副本上刷新域 / 重排 / 导 PDF，绝不写回输入 docx——
验证 / 编辑校验场景下，被验的 artifact 必须保持字节不可变（见 Layer 0 integrity 设计）。
只有显式传 --save-updated-fields 才把刷新后的域写回原文件（render --pdf 的交付物需要）。

out.pdf 缺省取输入同名 .pdf（长中文名手打容易错，这里自动推导）。
输出：页数 / 字数 / 表数 / 图数 / 节数。
"""

import os
import shutil
import sys
import tempfile

import pythoncom
import win32com.client as win32

# 解析参数：位置参数 <in.docx> [out.pdf]，开关 --save-updated-fields
save_updated = "--save-updated-fields" in sys.argv
positional = [a for a in sys.argv[1:] if not a.startswith("--")]
if len(positional) < 1:
    sys.exit("用法: python scripts/finalize.py <in.docx> [out.pdf] [--save-updated-fields]")
SRC = positional[0]
PDF = positional[1] if len(positional) > 1 else os.path.splitext(SRC)[0] + ".pdf"

WD_PAGES, WD_WORDS = 2, 0
WD_PDF = 17

pythoncom.CoInitialize()
word = win32.DispatchEx("Word.Application")
word.Visible = False
word.DisplayAlerts = 0
# 记录验收引擎身份：换机器/装了 WPS 时，"Word 能打开"这句话的分量不一样
print("Engine    :", word.Name, word.Version, "|", word.Path)

# 只读验收：在临时副本上操作，原 docx 字节绝不被动；--save-updated-fields 才写回。
tmp_dir = tempfile.mkdtemp(prefix="texere_finalize_")
tmp_src = os.path.join(tmp_dir, os.path.basename(SRC))
shutil.copy2(SRC, tmp_src)
try:
    doc = word.Documents.Open(os.path.abspath(tmp_src), False, False, False)
    for i in range(1, doc.TablesOfContents.Count + 1):
        doc.TablesOfContents(i).Update()
    doc.Fields.Update()
    doc.Repaginate()
    print("Pages      :", doc.ComputeStatistics(WD_PAGES))
    print("Words      :", doc.ComputeStatistics(WD_WORDS))
    print("Tables     :", doc.Tables.Count)
    print("InlineShapes:", doc.InlineShapes.Count)
    print("Sections   :", doc.Sections.Count)
    doc.ExportAsFixedFormat(os.path.abspath(PDF), WD_PDF)
    print("PDF written:", os.path.exists(PDF))
    if save_updated:
        # 显式把刷新后的域写回原文件（默认不写，验证阶段保持 artifact 不可变）
        doc.SaveAs(os.path.abspath(SRC))
    doc.Close(0)
    print("OK")
finally:
    word.Quit()
    pythoncom.CoUninitialize()
    shutil.rmtree(tmp_dir, ignore_errors=True)
