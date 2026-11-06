"""
scratch/audit_today_trades.py
Deep forensic audit of today's (2026-11-03) trades.
Inspect:
- Exact entry time & exit time
- Setup engine (Engine 1 Cross vs Engine 2 VWAP Cont)
- Nifty 50 intraday movement during trade
- MFE, MAE, and bar action
"""
import sys, os, pickle, json
from datetime import datetime
import pandas as pd
import numpy as np

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, 'sim')
import config as C
from simulator import Simulator

with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
    payload = pickle.load(f)

raw = payload['raw']
idx = payload['idx']
universe = payload['universe']

target_day = sorted(list(set(raw['RELIANCE.NS'].index.date)))[-1]
print(f"Auditing Date: {target_day}")

sim = Simulator('replay', datetime.combine(target_day, datetime.min.time()), strategy='v6_profit_max', starting_capital=50000.0)
sim.load(raw, universe, idx)

for t in sim.session_bar_times():
    sim.step(t)

snap = sim.snapshot()
closed = snap['closed']

print("\n" + "="*95)
print(f"DETAILED FORENSIC TRADE LOG FOR {target_day}")
print("="*95)

ndf = idx['NIFTY 50']
ndf_day = ndf[ndf.index.date == target_day]

for c in closed:
    sym = c['symbol']
    e_time = c['entry_time']
    x_time = c['exit_time']
    setup = c.get('setup', 'N/A')
    
    # Check Nifty at entry
    n_entry_cl = ndf_day.loc[ndf_day.index <= pd.Timestamp(e_time)]['close'].iloc[-1] if len(ndf_day.loc[ndf_day.index <= pd.Timestamp(e_time)]) else 0
    n_exit_cl = ndf_day.loc[ndf_day.index <= pd.Timestamp(x_time)]['close'].iloc[-1] if len(ndf_day.loc[ndf_day.index <= pd.Timestamp(x_time)]) else 0
    n_ret_during = (n_exit_cl - n_entry_cl) / n_entry_cl * 100 if n_entry_cl else 0
    
    print(f"\nSymbol: {sym} ({c['side']}) | Setup: {setup}")
    print(f"  Entry: Rs {c['entry']} @ {e_time} | Exit: Rs {c['exit']} @ {x_time}")
    print(f"  Bars Held: {c['bars_held']} ({(c['bars_held'])*5} mins) | MFE: {c['mfe']:+.2f}% | MAE: {c['mae']:+.2f}%")
    print(f"  Gross: Rs {c['gross']:+,.2f} | Fees: Rs {c['costs']:,.2f} | Net: Rs {c['net']:+,.2f}")
    print(f"  Exit Reason: {c['reason']}")
    print(f"  Nifty Move During Trade: {n_ret_during:+.2f}% (Nifty {n_entry_cl:.0f} -> {n_exit_cl:.0f})")

print("\n" + "="*95)
print("NIFTY 50 INTRADAY TIMELINE TODAY")
print("="*95)
for i in range(0, len(ndf_day), 6):
    row = ndf_day.iloc[i]
    ts = ndf_day.index[i].strftime('%H:%M')
    o, h, l, cl = row['open'], row['high'], row['low'], row['close']
    ret = (cl - ndf_day.iloc[0]['open']) / ndf_day.iloc[0]['open'] * 100
    print(f"  {ts} IST: Nifty {cl:,.1f} ({ret:+.2f}% from Open)")
