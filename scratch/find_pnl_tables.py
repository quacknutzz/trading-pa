"""
scratch/find_pnl_tables.py
Search across all sheets in AI_BTB_Dashboard.xlsm for P&L tables, realized returns, charges, or trade records.
"""
import sys, openpyxl

sys.stdout.reconfigure(encoding='utf-8')
wb = openpyxl.load_workbook('AI_BTB_Dashboard.xlsm', read_only=True, data_only=True)

keywords = ['p&l', 'p/l', 'profit', 'loss', 'gross', 'net', 'brokerage', 'stt', 'charges', 'realized', 'win', 'closed']

print("Searching all sheets for P&L or trade record keywords...")

for sname in wb.sheetnames:
    ws = wb[sname]
    matches = []
    for r_idx, row in enumerate(ws.iter_rows(values_only=True)):
        for c_idx, val in enumerate(row):
            if val is not None:
                s_val = str(val).lower()
                if any(k in s_val for k in keywords):
                    matches.append((r_idx + 1, c_idx + 1, str(val)[:50]))
                    if len(matches) > 15:
                        break
        if len(matches) > 15:
            break
    if matches:
        print(f"\nSheet '{sname}' has {len(matches)} keyword matches:")
        for r, c, val in matches[:8]:
            print(f"  Row {r}, Col {c}: {val}")
