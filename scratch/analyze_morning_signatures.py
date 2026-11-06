"""
scratch/analyze_morning_signatures.py
Analyze the 54 morning trades: 37 winners vs 17 losers.
Extract candle properties at entry bar:
1. RVOL_TOD
2. Body-to-Range ratio
3. Wick ratios (upper vs lower)
4. Distance to VWAP (%)
5. NIFTY 5m return & 10m EMA state at that minute
6. Sector trend
7. ADX
"""
import sys, os, pickle, math
from datetime import datetime
import pandas as pd
import numpy as np

sys.stdout.reconfigure(encoding='utf-8')

df_morning = pd.read_csv('scratch/morning_trades_raw.csv')
print(f"Loaded {len(df_morning)} morning trades.")

with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
    payload = pickle.load(f)

raw = payload['raw']
idx = payload['idx']
universe = payload['universe']
ndf = idx['NIFTY 50']

features = []

for _, row in df_morning.iterrows():
    sym = row['symbol']
    y_sym = sym + '.NS'
    if y_sym not in raw:
        continue
    df = raw[y_sym]
    entry_dt = pd.to_datetime(row['entry_time'])
    
    # Get bars up to entry
    sub = df[df.index <= entry_dt]
    if len(sub) < 3:
        continue
        
    entry_bar = sub.iloc[-1]
    prev_bar = sub.iloc[-2]
    
    o, h, l, c = entry_bar['open'], entry_bar['high'], entry_bar['low'], entry_bar['close']
    rng = h - l if h > l else 0.01
    body = abs(c - o)
    body_ratio = body / rng
    
    upper_wick = (h - max(o, c)) / rng
    lower_wick = (min(o, c) - l) / rng
    
    side = row['side']
    # For BUY: adverse wick is upper wick (selling pressure). For SELL: adverse wick is lower wick.
    adverse_wick = upper_wick if side == 'BUY' else lower_wick
    support_wick = lower_wick if side == 'BUY' else upper_wick
    
    # Volume vs 20-day TOD average
    tod = entry_dt.time()
    same_tod_bars = df[df.index.time == tod]
    past_same_tod = same_tod_bars[same_tod_bars.index < entry_dt]
    rvol_tod = (entry_bar['volume'] / past_same_tod['volume'].tail(20).mean()) if len(past_same_tod) >= 5 and past_same_tod['volume'].tail(20).mean() > 0 else 1.0
    
    # NIFTY state
    n_sub = ndf[ndf.index <= entry_dt]
    n_ret_morning = 0.0
    if len(n_sub) >= 2:
        # Day open to now
        day_n = n_sub[n_sub.index.date == entry_dt.date()]
        if len(day_n) >= 1:
            n_ret_morning = (day_n['close'].iloc[-1] - day_n['open'].iloc[0]) / day_n['open'].iloc[0] * 100
            
    # NIFTY alignment: BUY wants NIFTY > 0 or not heavily red; SELL wants NIFTY < 0
    nifty_aligned = (side == 'BUY' and n_ret_morning >= -0.10) or (side == 'SELL' and n_ret_morning <= 0.10)
    
    features.append({
        'symbol': sym,
        'date': row['date'],
        'entry_tod': row['entry_tod'],
        'side': side,
        'win': row['win'],
        'gross': row['gross'],
        'net': row['net'],
        'ret_pct': row['ret_pct'],
        'body_ratio': round(body_ratio, 2),
        'adverse_wick': round(adverse_wick, 2),
        'support_wick': round(support_wick, 2),
        'rvol_tod': round(rvol_tod, 2),
        'nifty_ret_morning': round(n_ret_morning, 2),
        'nifty_aligned': 1 if nifty_aligned else 0
    })

df_feat = pd.DataFrame(features)
print(f"Computed features for {len(df_feat)} trades.")

winners = df_feat[df_feat['win'] == 1]
losers = df_feat[df_feat['win'] == 0]

print("\n=== METRIC COMPARISON: WINNERS (True Momentum) vs LOSERS (Fake Trap) ===")
print(f"Feature                   | Winners (N={len(winners)}) | Losers (N={len(losers)})")
print("-" * 65)
print(f"RVOL vs Time-of-Day       | {winners['rvol_tod'].mean():.2f}x               | {losers['rvol_tod'].mean():.2f}x")
print(f"Candle Body / Range Ratio | {winners['body_ratio'].mean():.2f}                | {losers['body_ratio'].mean():.2f}")
print(f"Adverse Wick Ratio        | {winners['adverse_wick'].mean():.2f}                | {losers['adverse_wick'].mean():.2f}")
print(f"Support Wick Ratio        | {winners['support_wick'].mean():.2f}                | {losers['support_wick'].mean():.2f}")
print(f"NIFTY Morning Return (%)  | {winners['nifty_ret_morning'].mean():+.2f}%              | {losers['nifty_ret_morning'].mean():+.2f}%")
print(f"NIFTY Direction Aligned   | {winners['nifty_aligned'].mean()*100:.1f}%              | {losers['nifty_aligned'].mean()*100:.1f}%")

df_feat.to_csv('scratch/morning_trade_features.csv', index=False)
