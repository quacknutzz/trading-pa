"""
scratch/scan_all_excel_data.py
Scan all sheets in AI_BTB_Dashboard.xlsm to find actual trade logs and P&L entries.
"""
import sys, openpyxl

sys.stdout.reconfigure(encoding='utf-8')
wb = openpyxl.load_workbook('AI_BTB_Dashboard.xlsm', data_only=True)

print("Scanning all sheets for non-empty rows...")

for sname in wb.sheetnames:
    ws = wb[sname]
    non_empty_rows = 0
    sample_rows = []
    
    # Check rows
    for r in range(1, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, min(30, ws.max_column + 1))]
        non_null = [v for v in vals if v is not None and str(v).strip() != '']
        if len(non_null) > 0:
            non_empty_rows += 1
            if len(sample_rows) < 5:
                sample_rows.append((r, non_null[:8]))
            elif non_empty_rows % 50 == 0:
                sample_rows.append((r, non_null[:8]))
                
    if non_empty_rows > 0:
        print(f"\nSheet '{sname}': {non_empty_rows} non-empty rows out of {ws.max_row}")
        for r_num, s_vals in sample_rows[:6]:
            print(f"  Row {r_num}: {s_vals}")
