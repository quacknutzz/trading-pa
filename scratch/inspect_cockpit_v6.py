"""
scratch/inspect_cockpit_v6.py
Inspect the COCKPIT_V6 sheet in AI_BTB_Dashboard.xlsm.
"""
import sys, openpyxl

sys.stdout.reconfigure(encoding='utf-8')
wb = openpyxl.load_workbook('AI_BTB_Dashboard.xlsm', read_only=True, data_only=True)
ws = wb['COCKPIT_V6']

print("=== COCKPIT_V6 ROWS (Sample of first 40 rows) ===")
count = 0
for i, row in enumerate(ws.iter_rows(values_only=True)):
    non_null = [(c, v) for c, v in enumerate(row) if v is not None and str(v).strip() != '']
    if non_null:
        count += 1
        print(f"Row {i+1}: {[(c, str(v)[:25]) for c, v in non_null[:10]]}")
        if count >= 35:
            break
