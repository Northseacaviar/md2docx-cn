# -*- coding: utf-8 -*-
"""用本机 Word 把 docx 导出为 PDF（只用于排版预览）。用法：python docx_to_pdf.py <docx> <pdf>"""
import sys
from pathlib import Path

import win32com.client

src, dst = (str(Path(a).resolve()) for a in sys.argv[1:3])
word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
word.DisplayAlerts = 0  # wdAlertsNone：不弹“更新域”确认框
try:
    doc = word.Documents.Open(src, ReadOnly=True)
    doc.Fields.Update()
    doc.ExportAsFixedFormat(dst, ExportFormat=17)  # wdExportFormatPDF
    print("pages:", doc.ComputeStatistics(2))  # wdStatisticPages
    doc.Close(False)
finally:
    word.Quit()
