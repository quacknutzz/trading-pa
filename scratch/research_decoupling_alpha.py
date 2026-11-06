import sys, pickle
from datetime import datetime
import pandas as pd
import numpy as np

print("Loading dataset...")
with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
    payload = pickle.load(f)

raw = payload['raw']
idx = payload['idx']
universe = {item['symbol']: item for item in payload['universe']}
ndf = idx['NIFTY 50'].copy()

# Compute Nifty indicators
ndf['nifty_ret_5m'] = ndf['close'].pct_change()
ndf['nifty_ret_15m'] = ndf['close'].pct_change(3)
ndf['nifty_ret_30m'] = ndf['close'].pct_change(6)
ndf['nifty_ema10'] = ndf['close'].ewm(span=10, adjust=False).mean()
ndf['nifty_ema30'] = ndf['close'].ewm(span=30, adjust=False).mean()

# Map days to regime
daily_nifty = []
days = sorted({ts.normalize() for ts in ndf.index})
for d in days:
    day_df = ndf[ndf.index.normalize() == d]
    if len(day_df) < 5:
        continue
    o = day_df['open'].iloc[0]
    c = day_df['close'].iloc[-1]
    ret = (c - o) / o * 100
    reg = 'GREEN' if ret > 0.25 else ('RED' if ret < -0.25 else 'FLAT')
    daily_nifty.append({'date': d, 'nifty_day_ret': ret, 'nifty_regime': reg})

df_daily_nifty = pd.DataFrame(daily_nifty).set_index('date')

# Precompute Time-of-Day (TOD) mean volume for each stock across all days
# For each stock and each 5m bar time of day (e.g., 09:15, 09:20... 15:25),
# compute the rolling 20-day mean volume at that exact time.
print("Precomputing Time-of-Day (TOD) 20-day baseline volumes...")

tod_rvol_dict = {}
for y_sym, df in raw.items():
    sym = y_sym.replace('.NS', '')
    if len(df) < 200:
        continue
    df_copy = df.copy()
    df_copy['tod'] = df_copy.index.time
    # Group by TOD and compute rolling 20 mean of volume shifted by 1 day
    # To do this correctly: pivot table (date x tod)
    df_copy['date'] = df_copy.index.normalize()
    pv = df_copy.pivot(index='date', columns='tod', values='volume')
    pv_mean20 = pv.shift(1).rolling(20, min_periods=5).mean()
    # Unpivot back
    unpivoted = pv_mean20.unstack().reset_index()
    unpivoted.columns = ['tod', 'date', 'tod_mean_vol']
    # Merge back to df_copy
    df_merged = pd.merge(df_copy.reset_index(), unpivoted, on=['date', 'tod'], how='left')
    df_merged = df_merged.set_index('timestamp')
    df_copy['rvol_tod'] = df_copy['volume'] / df_merged['tod_mean_vol']
    tod_rvol_dict[sym] = df_copy['rvol_tod']

print("TOD RVOL precomputation complete.")
