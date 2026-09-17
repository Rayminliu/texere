# -*- coding: utf-8 -*-
"""Word COM 验收 + 导出 PDF：python scripts/finalize.py <in.docx> <out.pdf>

验收信号：Word 打不开 = OOXML 结构有问题（立即失败）；同时刷新目录域并重排页码后存回。
输出：页数 / 字数 / 表数 / 图数 / 节数。
"""
import os
import sys

import pythoncom
import win32com.client as win32

if len(sys.argv) < 3:
    sys.exit("用法: python scripts/finalize.py <in.docx> <out.pdf>")
SRC, PDF = sys.argv[1], sys.argv[2]

WD_PAGES, WD_WORDS = 2, 0
WD_PDF = 17

pythoncom.CoInitialize()
word = win32.DispatchEx("Word.Application")
word.Visible = False
word.DisplayAlerts = 0
# 记录验收引擎身份：换机器/装了 WPS 时，"Word 能打开"这句话的分量不一样
print("Engine    :", word.Name, word.Version, "|", word.Path)
try:
    doc = word.Documents.Open(os.path.abspath(SRC), False, False, False)
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
    doc.Save()
    doc.Close(0)
    print("OK")
finally:
    word.Quit()
    pythoncom.CoUninitialize()
