import sys, pickle
import numpy as np
import pandas as pd
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, 'sim')

with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
    payload = pickle.load(f)
raw = payload['raw']
idx = payload['idx']

nifty_df = idx.get('NIFTY 50')
if nifty_df is None:
    print('NIFTY not found in idx')
    sys.exit(0)

# Calculate daily performance of NIFTY
nifty_daily = nifty_df['close'].resample('1D').agg({'open': 'first', 'close': 'last'}).dropna()
nifty_daily['ret'] = (nifty_daily['close'] / nifty_daily['open'] - 1) * 100

print(f"Total days analyzed: {len(nifty_daily)}")
flat_days = nifty_daily[(nifty_daily['ret'] >= -0.3) & (nifty_daily['ret'] <= 0.3)]
red_days = nifty_daily[nifty_daily['ret'] < -0.3]
green_days = nifty_daily[nifty_daily['ret'] > 0.3]

print(f"  Green days (> +0.3%):  {len(green_days)}")
print(f"  Flat / Chop days:      {len(flat_days)}")
print(f"  Red days (< -0.3%):    {len(red_days)}")
