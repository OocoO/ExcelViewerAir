# -*- coding: utf-8 -*-
"""造一个"改过几处"的 xlsx 副本，用来验证 diff 功能。"""
import os
import shutil
import sys

import openpyxl

src = sys.argv[1] if len(sys.argv) > 1 else r"test\Item.xlsx"
dst = sys.argv[2] if len(sys.argv) > 2 else os.path.join(
    os.environ.get("LOCALAPPDATA", "."), "ExcelViewer", "difftest", "Item_edited.xlsx")

os.makedirs(os.path.dirname(dst), exist_ok=True)
shutil.copy2(src, dst)

wb = openpyxl.load_workbook(dst)
ws = wb.worksheets[0]
print(f"sheet={ws.title} dims={ws.dimensions}")

# 1) 改两个已知单元格（row5/col2 与 row6/col2：铜矿/铁矿 名称列）
edits = [(5, 3, "铜矿(改)"), (6, 3, "铁矿(改)"), (5, 4, 99999)]
for row, col, value in edits:
    old = ws.cell(row=row, column=col).value
    ws.cell(row=row, column=col).value = value
    print(f"  edit r{row}c{col}: {old!r} -> {value!r}")

# 2) 把一个原本空的格子填上内容，验证"从空到有值"
ws.cell(row=8, column=13).value = "新增内容"
print("  edit r8c13: None -> '新增内容'")

wb.save(dst)
wb.close()
print(f"saved: {dst}")
