import pickle, pandas as pd
from datetime import datetime
import sys
sys.path.insert(0, '.')
sys.path.insert(0, 'sim')
from scratch.fast_decoupling_backtest import DecoupledSimulator, init_worker

init_worker()
with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
    payload = pickle.load(f)
raw = payload['raw']
u = payload['universe']
idx = payload['idx']

flat_dates = ['2026-08-10', '2026-09-01', '2026-09-04', '2026-09-10', '2026-09-16', '2026-09-18', '2026-09-29', '2026-10-22', '2026-10-28']
print("Date       | Symbol     | Side | Net PnL   | MFE    | MAE    | Ret%   | Reason")
print("-" * 80)
for d_str in flat_dates:
    dt = datetime.strptime(d_str, '%Y-%m-%d')
    sim = DecoupledSimulator('replay', dt, decoupling_mode='decoupled_engine', starting_capital=50000.0)
    sim.load(raw, u, idx)
    for t in sim.session_bar_times():
        sim.step(t)
    for c in sim.closed:
        sym = c['symbol']
        side = c['side']
        net = c['net']
        mfe = c.get('mfe', 0.0)
        mae = c.get('mae', 0.0)
        ret = c.get('ret_pct', 0.0)
        reason = c['reason']
        print(f"{d_str} | {sym:10s} | {side:4s} | Rs {net:>7.2f} | {mfe:>5.2f}% | {mae:>5.2f}% | {ret:>5.2f}% | {reason}")
