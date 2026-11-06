"""
scratch/extract_cockpit_v6_trades.py
Extract and analyze all rows from COCKPIT_V6 in AI_BTB_Dashboard.xlsm.
"""
import sys, openpyxl
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')
wb = openpyxl.load_workbook('AI_BTB_Dashboard.xlsm', read_only=True, data_only=True)
ws = wb['COCKPIT_V6']

headers = [
    'Ticker', 'Action', 'LTP', 'Sector', 'S3', 'S2', 'S1', 'PP', 'R1', 'R2', 'R3',
    'RSI_5m', 'RSI_15m', 'RSI_1h', 'RSI_1d', 'Supertrend', 'EMA_10_30', 'EMA_10_50',
    'Ent', 'Tgt', 'SL', 'CP', 'Open_Time', 'Close_Time', 'Outcome', 'Bottomline',
    'Comments', 'No_Of_Bars', 'Apply_To_FYERS', 'Peak_Price', 'Best_MFE_Pct', 'Last_MFE_Bar',
    'MFE_Bar1', 'Is_Runner', 'Last_Bar_TS', 'Exec_State', 'FYERS_Order_ID', 'Filled_Qty',
    'Requested_Qty', 'Request_ID', 'Exit_Fill_Px', 'Request_Type', 'Trade_Origin',
    'FYERS_Has_SL', 'FYERS_Has_TP'
]

records = []
for r_idx, row in enumerate(ws.iter_rows(values_only=True)):
    if r_idx < 4:
        continue # skip top rows and header
    ticker = row[0]
    if ticker is not None and str(ticker).strip() != '':
        rec = {}
        for c_idx, h in enumerate(headers):
            val = row[c_idx] if c_idx < len(row) else None
            rec[h] = val
        records.append(rec)

df = pd.DataFrame(records)
print(f"Total rows in COCKPIT_V6: {len(df)}")
df.to_csv('scratch/dad_cockpit_v6_raw.csv', index=False)
print("Saved to scratch/dad_cockpit_v6_raw.csv")

# Filter rows that have Open_Time or Outcome
trades = df[df['Outcome'].notna() | df['Open_Time'].notna()]
print(f"Rows with Trade Activity (Outcome or Open_Time): {len(trades)}")

if len(trades) > 0:
    print("\nOutcome Value Counts:")
    print(trades['Outcome'].value_counts())
    
    print("\nExec State Counts:")
    print(trades['Exec_State'].value_counts())
    
    print("\nSample 10 trades:")
    cols_to_show = ['Ticker', 'Action', 'Ent', 'CP', 'Open_Time', 'Close_Time', 'Outcome', 'Bottomline', 'Filled_Qty', 'Exec_State']
    print(trades[cols_to_show].head(10))
