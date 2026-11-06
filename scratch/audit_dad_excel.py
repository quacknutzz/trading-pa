"""
scratch/audit_dad_excel.py
Deep audit of Dad's AI_BTB_Dashboard.xlsm.
Check:
1. Trade log sheets: 'Trades', 'New_Trades', 'Paper_Trades', 'COCKPITV6_Log', 'Live', 'Live_New'
2. Columns, number of rows, date ranges, symbols traded, gross P&L, fees/brokerage/charges, net P&L
3. Summary sheets: 'COCKPIT_V6', 'NL_ID_Cumulative'
"""
import sys, os
import pandas as pd
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('AI_BTB_Dashboard.xlsm', data_only=True)
sheets = wb.sheetnames
print(f"Total sheets: {len(sheets)}")

sheets_to_inspect = ['COCKPITV6_Log', 'Trades', 'New_Trades', 'Paper_Trades', 'Live', 'Live_New', 'NL_ID_Cumulative', 'COCKPIT_V6']

for sname in sheets_to_inspect:
    if sname not in wb.sheetnames:
        continue
    ws = wb[sname]
    print(f"\n==================== SHEET: {sname} ====================")
    print(f"Dimensions: max_row={ws.max_row}, max_column={ws.max_column}")
    
    # Read first 5 rows to see headers and structure
    rows = []
    for r in range(1, min(10, ws.max_row + 1)):
        row_vals = [ws.cell(r, c).value for c in range(1, min(25, ws.max_column + 1))]
        if any(v is not None for v in row_vals):
            rows.append(row_vals)
            
    if rows:
        print(f"Header / Sample rows (first {len(rows)}):")
        for i, r in enumerate(rows):
            # Print non-empty items
            non_empty = [f"Col{idx+1}: {val}" for idx, val in enumerate(r) if val is not None]
            print(f"  Row {i+1}: {', '.join(non_empty[:8])}")
    else:
        print("  [Sheet appears empty]")
