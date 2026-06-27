#!/usr/bin/env python3
"""
Generate data.js from the Excel source file.
Usage: python3 scripts/gen_data.py
Requires: pip install openpyxl
"""
import openpyxl
import json
import pathlib
import glob
from datetime import date

ROOT = pathlib.Path(__file__).parent.parent
OUTPUT_PATH = ROOT / "data.js"
SHEET_NAME = "货品基础资料P"

# Auto-detect xlsx file: prefer 货品基础资料.xlsx, fall back to first .xlsx found
preferred = ROOT / "货品基础资料.xlsx"
if preferred.exists():
    EXCEL_PATH = preferred
else:
    candidates = sorted(ROOT.glob("*.xlsx"))
    if not candidates:
        raise FileNotFoundError("根目录下未找到任何 .xlsx 文件")
    EXCEL_PATH = candidates[0]

print(f"Reading: {EXCEL_PATH.name}")
wb = openpyxl.load_workbook(EXCEL_PATH)
ws = wb[SHEET_NAME]

rows = []
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0:  # skip header
        continue
    sku   = str(row[5]).strip() if row[5] is not None else ""
    name  = str(row[0]).strip() if row[0] is not None else ""
    spec  = str(row[2]).strip() if row[2] is not None else ""
    aux   = str(row[3]).strip() if row[3] is not None else ""
    brand = str(row[1]).strip() if row[1] is not None else ""
    price = row[4]

    if not sku:
        continue

    # Validate price: must be a positive number
    if price is not None:
        try:
            p = float(price)
            price = str(price).strip() if p > 0 else None
        except (ValueError, TypeError):
            price = None

    rows.append([sku, name, spec, aux, brand, price])

js = (
    f"// Generated from {EXCEL_PATH.name} on {date.today()} — DO NOT EDIT MANUALLY\n"
    f"// Total products: {len(rows)}\n"
    "const PRODUCTS=" +
    json.dumps(rows, ensure_ascii=False, separators=(',', ':')) +
    ";\n"
)

OUTPUT_PATH.write_text(js, encoding="utf-8")
print(f"Written {len(rows)} products to {OUTPUT_PATH}")
print(f"File size: {OUTPUT_PATH.stat().st_size / 1024:.1f} KB")
