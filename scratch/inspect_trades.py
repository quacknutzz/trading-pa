import sys, pickle
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
ndf = idx['NIFTY 50']

days = sorted({ts.normalize() for ts in ndf.index})[3:]

all_trades = []

for d in days:
    day_df = ndf[ndf.index.normalize() == d]
    if len(day_df) < 5:
        continue
    o = day_df['open'].iloc[0]
    c = day_df['close'].iloc[-1]
    ret = (c - o) / o * 100
    day_cat = 'GREEN' if ret > 0.25 else ('RED' if ret < -0.25 else 'FLAT')
    
    dt = datetime.strptime(str(d.date()), '%Y-%m-%d')
    sim = Simulator('replay', dt, strategy='v6_profit_max', starting_capital=50000.0)
    sim.load(raw, u, idx)
    for t in sim.session_bar_times():
        sim.step(t)
    
    for tr in sim.closed:
        tr_copy = dict(tr)
        tr_copy['date'] = str(d.date())
        tr_copy['day_cat'] = day_cat
        tr_copy['nifty_day_ret'] = ret
        all_trades.append(tr_copy)

df_tr = pd.DataFrame(all_trades)
print(f"Total closed trades recorded: {len(df_tr)}")
df_tr.to_csv('scratch/baseline_closed_trades.csv', index=False)

print("\n=== SAMPLE TRADES ON FLAT DAYS ===")
flat_trades = df_tr[df_tr['day_cat'] == 'FLAT']
cols_show = ['date', 'symbol', 'side', 'entry', 'exit', 'gross', 'fees', 'net', 'reason']
print(flat_trades[cols_show].to_string(index=False))

print("\n=== WIN RATE & NET PROFIT BY RVOL ON FLAT/RED DAYS ===")
df_tr['rvol_high'] = df_tr['rvol'] >= 1.5
print(df_tr.groupby(['day_cat', 'rvol_high']).agg(
    trades=('net', 'count'),
    wins=('net', lambda x: (x > 0).sum()),
    win_rate=('net', lambda x: (x > 0).sum() / len(x) * 100),
    net_pnl=('net', 'sum'),
    gross_pnl=('gross', 'sum'),
    fees=('fees', 'sum')
))
