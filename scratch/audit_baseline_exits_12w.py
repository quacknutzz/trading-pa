import sys, pickle, math
from datetime import datetime
import pandas as pd
sys.path.insert(0, 'sim')
import datafeed
from simulator import Simulator

u = datafeed.load_universe()
with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
    payload = pickle.load(f)
raw = payload['raw']
idx = payload['idx']
sym0 = list(raw.keys())[0]
dates = sorted(list(set(raw[sym0].index.date)))

master = Simulator('replay', datetime.combine(dates[0], datetime.min.time()), strategy='v6_baseline', starting_capital=500000.0)
master.load(raw, u, idx)

all_closed = []
for d in dates:
    dt = datetime.combine(d, datetime.min.time())
    sim = Simulator('replay', dt, strategy='v6_baseline', starting_capital=500000.0)
    sim.syms = master.syms
    sim.indices = master.indices
    sim.nifty_pos = master.nifty_pos
    sim.nifty_e10 = master.nifty_e10
    sim.nifty_e21 = master.nifty_e21
    sim.nifty_e30 = master.nifty_e30
    for t in sim.session_bar_times():
        sim.step(t)
    all_closed.extend(sim.closed)

df = pd.DataFrame(all_closed)
print(f"Total baseline closed trades: {len(df)}")

def clean_lbl(r):
    r = str(r)
    if 'ADX' in r: return 'ADX < 18'
    if '10M' in r: return '10M_CROSS_CONTRA'
    if '5M' in r: return '5M_CROSS_REVERSE'
    if '3_PM' in r: return '3_PM_EXIT'
    if 'TARGET' in r: return 'TARGET (1.2%)'
    if 'STOP_LOSS' in r: return 'STOP_LOSS (1.0%)'
    return 'OTHER'

df['Exit_Type'] = df['reason'].apply(clean_lbl)
summary = df.groupby('Exit_Type').agg(
    trades=('net', 'count'),
    wins=('net', lambda s: (s > 0).sum()),
    win_rate=('net', lambda s: round((s > 0).mean() * 100, 1)),
    gross=('gross', 'sum'),
    costs=('costs', 'sum'),
    net=('net', 'sum'),
    avg_mfe=('mfe', 'mean'),
    avg_mae=('mae', 'mean'),
    bars_held=('bars_held', 'mean')
)
print("\n--- 12-WEEK BASELINE EXIT BREAKDOWN (681 TRADES) ---")
print(summary.to_string())
