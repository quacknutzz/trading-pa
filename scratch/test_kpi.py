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

days = sorted({ts.normalize() for ts in ndf.index})[3:] # skip first 3 warmups
print(f'Evaluating {len(days)} trading days...')

# Test one day to see kpi keys
dt_test = datetime.strptime(str(days[0].date()), '%Y-%m-%d')
sim_test = Simulator('replay', dt_test, strategy='v6_profit_max', starting_capital=50000.0)
sim_test.load(raw, u, idx)
for t in sim_test.session_bar_times():
    sim_test.step(t)
k_test = sim_test.snapshot()['kpi']
print('KPI keys:', list(k_test.keys()))
print('Sample KPI:', k_test)
