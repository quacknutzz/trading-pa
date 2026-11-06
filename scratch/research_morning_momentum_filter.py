"""
scratch/research_morning_momentum_filter.py
Investigate all morning trades taken between 09:15 and 09:45 across 62 days.
Isolate:
1. Winning morning trades (True Momentum) vs Losing morning trades (Fake Momentum / Trap)
2. Quantitative signatures:
   - RVOL (Relative Volume vs Time-of-Day)
   - Candle Body / Range Ratio (Marubozu vs Wick Rejection)
   - Distance from VWAP (% difference from VWAP)
   - NIFTY 10m EMA & NIFTY Return at 09:25-09:45
   - ADX strength at entry
   - Structural Target Distance (nearest Camarilla/Pivot)
"""
import sys, os, pickle, json
from datetime import datetime, time as dtime
import numpy as np
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, '.')
sys.path.insert(0, 'sim')
import config as C
from simulator import Simulator, BAR

with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
    payload = pickle.load(f)

raw = payload['raw']
idx = payload['idx']
universe = payload['universe']
sym0 = list(raw.keys())[0]
all_dates = sorted(list(set(raw[sym0].index.date)))

C.LEVERAGE = 2.0
C.ENTRY_START = dtime(9, 25)

# Run full 62-day simulation and record every trade with timestamp and indicators
dt0 = datetime.combine(all_dates[0], datetime.min.time())
master = Simulator('replay', dt0, strategy='v6_profit_max', starting_capital=50000.0)
master.load(raw, universe, idx)

all_closed_trades = []

for d in all_dates:
    dt = datetime.combine(d, datetime.min.time())
    sim = Simulator('replay', dt, strategy='v6_profit_max', starting_capital=50000.0)
    sim.syms = master.syms
    sim.indices = master.indices
    sim.nifty_pos = master.nifty_pos
    sim.nifty_e10 = master.nifty_e10
    sim.nifty_e21 = master.nifty_e21
    sim.nifty_e30 = master.nifty_e30
    
    for t in sim.session_bar_times():
        sim.step(t)
        
    for tr in sim.closed:
        entry_dt = datetime.fromisoformat(tr['entry_time'])
        all_closed_trades.append({
            'date': d.strftime('%Y-%m-%d'),
            'entry_time': tr['entry_time'],
            'entry_tod': entry_dt.time().strftime('%H:%M'),
            'exit_time': tr['exit_time'],
            'symbol': tr['symbol'],
            'side': tr['side'],
            'qty': tr['qty'],
            'entry': tr['entry'],
            'exit': tr['exit'],
            'gross': tr['gross'],
            'costs': tr['costs'],
            'net': tr['net'],
            'ret_pct': tr['ret_pct'],
            'reason': tr['reason'],
            'win': 1 if tr['gross'] > 0 else 0
        })

df_trades = pd.DataFrame(all_closed_trades)
print(f"Total trades recorded: {len(df_trades)}")

morning_trades = df_trades[df_trades['entry_tod'] < '09:45']
rest_trades = df_trades[df_trades['entry_tod'] >= '09:45']

print(f"\nMorning Trades (09:15 - 09:45): {len(morning_trades)}")
print(f"  Wins: {morning_trades['win'].sum()} ({morning_trades['win'].mean()*100:.1f}%)")
print(f"  Gross: Rs {morning_trades['gross'].sum():+,.2f}")
print(f"  Costs: Rs {morning_trades['costs'].sum():,.2f}")
print(f"  Net:   Rs {morning_trades['net'].sum():+,.2f}")

print(f"\nRest of Day Trades (09:45+): {len(rest_trades)}")
print(f"  Wins: {rest_trades['win'].sum()} ({rest_trades['win'].mean()*100:.1f}%)")
print(f"  Gross: Rs {rest_trades['gross'].sum():+,.2f}")
print(f"  Costs: Rs {rest_trades['costs'].sum():,.2f}")
print(f"  Net:   Rs {rest_trades['net'].sum():+,.2f}")

print("\n--- Morning Trades by Month ---")
for m in ['2026-08', '2026-09', '2026-10', '2026-11']:
    sub = morning_trades[morning_trades['date'].str.startswith(m)]
    if len(sub) > 0:
        print(f"  Month {m}: {len(sub)} trades, {sub['win'].mean()*100:.1f}% WR, Gross: Rs {sub['gross'].sum():+,.2f}, Costs: Rs {sub['costs'].sum():,.2f}, Net: Rs {sub['net'].sum():+,.2f}")

# Save morning trades for indicator extraction
morning_trades.to_csv('scratch/morning_trades_raw.csv', index=False)
