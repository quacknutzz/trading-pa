import sys, pickle
from datetime import datetime
sys.path.insert(0, '.')
sys.path.insert(0, 'sim')
from scratch.fast_decoupling_backtest import DecoupledSimulator, init_worker

init_worker()
with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
    payload = pickle.load(f)
raw = payload['raw']
u = payload['universe']
idx = payload['idx']

losing_dates = ['2026-09-04', '2026-09-16', '2026-09-18', '2026-10-22']
for d_str in losing_dates:
    dt = datetime.strptime(d_str, '%Y-%m-%d')
    sim = DecoupledSimulator('replay', dt, decoupling_mode='decoupled_engine', starting_capital=50000.0)
    sim.load(raw, u, idx)
    for t in sim.session_bar_times():
        sim.step(t)
    print(f'=== Date: {d_str} ===')
    for c in sim.closed:
        print(f"  {c['symbol']} {c['side']} entry {c['entry']} exit {c['exit']} net Rs {c['net']:+.2f} | Reason: {c['reason']} | Bars: {c['bars_held']}")
