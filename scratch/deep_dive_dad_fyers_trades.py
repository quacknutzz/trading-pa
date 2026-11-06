"""
scratch/deep_dive_dad_fyers_trades.py
Parse all dates and trades from COCKPIT_V6.
Group by Date and calculate:
- Number of trades (total, filled, cancelled/rejected)
- Wins vs Losses
- Gross P&L in Rupees (Filled_Qty * Price Diff)
- Approximate Brokerage and Taxes
- Net P&L in Rupees
"""
import sys
import pandas as pd
import numpy as np

sys.stdout.reconfigure(encoding='utf-8')

df = pd.read_csv('scratch/dad_cockpit_v6_raw.csv')

# Parse date from Open_Time
df['Date'] = pd.to_datetime(df['Open_Time']).dt.date

print(f"Total rows: {len(df)}")
print("Dates present in Dad's workbook:")
date_counts = df['Date'].value_counts().sort_index()
print(date_counts)

# Distinguish between LIVE trades with FYERS Order IDs vs Paper/Sim trades in sheet
df['Has_Fyers_ID'] = df['FYERS_Order_ID'].notna()
print("\nTrades with FYERS Order ID by Date:")
print(df[df['Has_Fyers_ID']]['Date'].value_counts().sort_index())

# Group by Date and compute metrics
daily_summary = []
for d, group in df.groupby('Date'):
    total_rows = len(group)
    closed = group[group['Exec_State'] == 'CLOSED']
    filled = group[group['Filled_Qty'] > 0]
    
    # Calculate rupees for filled trades
    rupees_pnl = 0.0
    for _, r in filled.iterrows():
        qty = r['Filled_Qty']
        diff = r['Bottomline']
        rupees_pnl += qty * diff
        
    wins = len(group[group['Outcome'] == 'WIN'])
    losses = len(group[group['Outcome'] == 'LOSS'])
    wr = (wins / total_rows * 100) if total_rows else 0
    
    daily_summary.append({
        'Date': str(d),
        'Total_Signals': total_rows,
        'Exec_Closed': len(closed),
        'Fyers_Filled': len(filled),
        'Wins': wins,
        'Losses': losses,
        'WinRate_Pct': round(wr, 1),
        'Filled_Rupees_P&L': round(rupees_pnl, 2)
    })

df_sum = pd.DataFrame(daily_summary)
print("\n=== DAD'S WORKBOOK DAILY SUMMARY ===")
print(df_sum.to_string(index=False))

# Show all trades with Fyers Order IDs
fyers_trades = df[df['Has_Fyers_ID']]
print(f"\nTotal Live Fyers Trades: {len(fyers_trades)}")
cols_fyers = ['Date', 'Open_Time', 'Ticker', 'Action', 'Ent', 'CP', 'Bottomline', 'Filled_Qty', 'FYERS_Order_ID', 'Outcome', 'Exec_State']
print(fyers_trades[cols_fyers].to_string(index=False))
