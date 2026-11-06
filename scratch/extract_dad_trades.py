"""
scratch/extract_dad_trades.py
Extract and analyze trade logs from Dad's Excel workbook:
- 'Trades' sheet
- 'New_Trades' sheet
- 'Paper_Trades' sheet
- 'COCKPITV6_Log' sheet
"""
import sys, os
import pandas as pd
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('AI_BTB_Dashboard.xlsm', data_only=True)

for sheet_name in ['Trades', 'New_Trades', 'Paper_Trades']:
    ws = wb[sheet_name]
    print(f"\n==================== SHEET: {sheet_name} ====================")
    data = []
    headers = None
    for r in range(1, ws.max_row + 1):
        row_vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        # Check if this row looks like header
        if headers is None:
            # Look for row containing 'Ticker' or 'Action' or 'Entry'
            str_vals = [str(v) if v is not None else '' for v in row_vals]
            if any('Ticker' in v or 'Action' in v for v in str_vals):
                headers = [v if v is not None else f'Col_{i+1}' for i, v in enumerate(row_vals)]
                print(f"Found headers at row {r}: {[h for h in headers if not str(h).startswith('Col_')][:15]}")
                continue
        elif any(v is not None for v in row_vals):
            data.append(row_vals[:len(headers)])
            
    if headers and data:
        df = pd.DataFrame(data, columns=headers)
        # Drop completely empty columns
        df = df.dropna(how='all', axis=1)
        print(f"Total non-empty rows: {len(df)}")
        print("Columns:", list(df.columns[:20]))
        print("\nFirst 3 rows:")
        print(df.iloc[:3, :10])
        print("\nLast 3 rows:")
        print(df.iloc[-3:, :10])
        df.to_csv(f'scratch/dad_{sheet_name}.csv', index=False)
        print(f"Saved to scratch/dad_{sheet_name}.csv")
    else:
        print(f"Could not parse tabular data from {sheet_name}")
