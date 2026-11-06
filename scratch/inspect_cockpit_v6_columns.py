"""
scratch/inspect_cockpit_v6_columns.py
Inspect all columns of COCKPIT_V6 sheet.
"""
import sys, openpyxl
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')
wb = openpyxl.load_workbook('AI_BTB_Dashboard.xlsm', read_only=True, data_only=True)
ws = wb['COCKPIT_V6']

# Find header row
for r_idx, row in enumerate(ws.iter_rows(values_only=True)):
    if r_idx < 10:
        non_null = [(c, v) for c, v in enumerate(row) if v is not None and str(v).strip() != '']
        print(f"Row {r_idx+1}: {[f'Col{c+1}: {v}' for c, v in non_null[:15]]}")
        if any('Ticker' in str(v) for _, v in non_null):
            header_row = [v for v in row]
            print(f"\nFound Header at row {r_idx+1}:")
            for c_idx, h in enumerate(header_row):
                if h is not None:
                    print(f"  Col {c_idx+1}: {h}")
            break
