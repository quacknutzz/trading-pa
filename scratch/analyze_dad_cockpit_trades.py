"""
scratch/analyze_dad_cockpit_trades.py
Deep audit of Dad's 272 trades from COCKPIT_V6.
Analyze:
- Dates and times of trades
- Capital per trade / Quantities
- Gross P&L in Rupees (Qty * Price Diff)
- Taxes and charges paid
- Net realized P&L
- Broker execution details (Fyers Order IDs, Rejections, Cancellations)
"""
import sys
import pandas as pd
import numpy as np

sys.stdout.reconfigure(encoding='utf-8')

df = pd.read_csv('scratch/dad_cockpit_v6_raw.csv')

print("=== COLUMNS IN DAD'S TRADE LOG ===")
print(df.columns.tolist())

# Inspect Open_Time and Close_Time formats
print("\nSample Open_Time:")
print(df['Open_Time'].dropna().head(10))

print("\nSample Close_Time:")
print(df['Close_Time'].dropna().head(10))

# Quantities
print("\nFilled_Qty stats:")
print(df['Filled_Qty'].describe())

print("\nRequested_Qty stats:")
print(df['Requested_Qty'].describe())

# Exec_State
print("\nExec_State value counts:")
print(df['Exec_State'].value_counts(dropna=False))

# Calculate P&L
# If Filled_Qty is NaN, what quantity was Dad using?
# Let's inspect the Comments, FYERS_Order_ID, etc.
sample_rows = df[df['Exec_State'] == 'CLOSED'].head(5)
for i, r in sample_rows.iterrows():
    print(f"\nTrade {r['Ticker']} {r['Action']}: Ent={r['Ent']}, Exit={r['CP']}, Bottomline={r['Bottomline']}, Open={r['Open_Time']}, Close={r['Close_Time']}, FilledQty={r['Filled_Qty']}, OrderID={r['FYERS_Order_ID']}")
