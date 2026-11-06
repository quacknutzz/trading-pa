import sys, pickle, json
from datetime import datetime
import pandas as pd
import numpy as np

sys.path.insert(0, '.')
sys.path.insert(0, 'sim')
import config as C
from simulator import Simulator

with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
    payload = pickle.load(f)
raw, idx, universe = payload['raw'], payload['idx'], payload['universe']
sym0 = list(raw.keys())[0]
all_dates = sorted(list(set(raw[sym0].index.date)))

print(f"Total dates: {len(all_dates)} ({all_dates[0]} to {all_dates[-1]})")

dt0 = datetime.combine(all_dates[0], datetime.min.time())
master = Simulator('replay', dt0, strategy='v6_profit_max', starting_capital=50000.0)
master.load(raw, universe, idx)

# Check all trades on Rs 50k
all_closed = []
for d in all_dates:
    dt = datetime.combine(d, datetime.min.time())
    sim = Simulator('replay', dt, strategy='v6_profit_max', starting_capital=50000.0)
    sim.syms = master.syms; sim.indices = master.indices
    sim.nifty_pos = master.nifty_pos; sim.nifty_e10 = master.nifty_e10
    sim.nifty_e21 = master.nifty_e21; sim.nifty_e30 = master.nifty_e30
    for t in sim.session_bar_times():
        sim.step(t)
    all_closed.extend(sim.closed)

df = pd.DataFrame(all_closed)
print(f"Total Trades: {len(df)}")
print("Exit Reasons breakdown:")
print(df['reason'].value_counts())
print("\nTrade Net P&L Summary:")
print(df['net'].describe())
wins = (df['net'] > 0).sum()
wr = wins / len(df) * 100
print(f"Trades with Net > 0: {wins} / {len(df)} ({wr:.2f}%)")
print(f"Total Gross: Rs {df['gross'].sum():,.2f}")
print(f"Total Fees:  Rs {df['costs'].sum():,.2f}")
print(f"Total Net:   Rs {df['net'].sum():,.2f}")

# Look at worst 10 losing trades
worst = df.sort_values('net').head(10)
print("\nWorst 10 Trades:")
for idx, r in worst.iterrows():
    print(f"  {r['symbol']:<10} {r['side']} | Entry Rs {r['entry']} -> Exit Rs {r['exit']} ({r['ret_pct']:+.2f}%) | Net Rs {r['net']} (Cost Rs {r['costs']}) | {r['reason']}")
