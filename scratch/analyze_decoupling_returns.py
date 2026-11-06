import sys, pickle
from datetime import datetime, time
import pandas as pd
import numpy as np

print("Loading dataset...")
with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
    payload = pickle.load(f)

raw = payload['raw']
idx = payload['idx']
universe = {item['symbol']: item for item in payload['universe']}
ndf = idx['NIFTY 50'].copy()

# Precompute Nifty metrics
ndf['nifty_ret_5m'] = ndf['close'].pct_change()
ndf['nifty_ret_15m'] = ndf['close'].pct_change(3)
ndf['nifty_ret_30m'] = ndf['close'].pct_change(6)
ndf['nifty_e10'] = ndf['close'].ewm(span=10, adjust=False).mean()
ndf['nifty_e30'] = ndf['close'].ewm(span=30, adjust=False).mean()
ndf['nifty_rs_ratio'] = ndf['close'] / ndf['nifty_e10']

# Precompute daily return of Nifty
daily_nifty = {}
for d, df_day in ndf.groupby(ndf.index.normalize()):
    o = df_day['open'].iloc[0]
    c = df_day['close'].iloc[-1]
    ret = (c - o) / o * 100
    daily_nifty[d] = ret

sector_map = {item['symbol']: item.get('sector', 'UNKNOWN') for item in payload['universe']}

print("Calculating features across 139 stocks...")
stock_features = {}

for y_sym, df in raw.items():
    sym = y_sym.replace('.NS', '')
    if len(df) < 300:
        continue
    df = df.copy()
    df['symbol'] = sym
    df['sector'] = sector_map.get(sym, 'UNKNOWN')
    
    df['date'] = df.index.normalize()
    df['tod'] = df.index.time
    
    # Vectorized Intraday VWAP
    typ = (df['high'] + df['low'] + df['close']) / 3.0
    cum_pv = (typ * df['volume']).groupby(df['date']).cumsum()
    cum_v = df['volume'].groupby(df['date']).cumsum()
    df['vwap'] = cum_pv / cum_v.replace(0, np.nan)
    
    # Rolling 20-bar volume RVOL
    df['rvol_rolling'] = df['volume'] / df['volume'].shift(1).rolling(20, min_periods=5).mean()
    
    # TOD 20-day RVOL
    pv = df.pivot(index='date', columns='tod', values='volume')
    pv_mean20 = pv.shift(1).rolling(20, min_periods=5).mean()
    unpivoted = pv_mean20.unstack().reset_index()
    unpivoted.columns = ['tod', 'date', 'tod_mean_vol']
    
    df_reset = df.reset_index()
    dt_col = df_reset.columns[0] # 'Datetime'
    df_merged = pd.merge(df_reset, unpivoted, on=['date', 'tod'], how='left').set_index(dt_col)
    df['rvol_tod'] = df['volume'] / df_merged['tod_mean_vol']
    
    # 5m EMAs
    df['e10'] = df['close'].ewm(span=10, adjust=False).mean()
    df['e30'] = df['close'].ewm(span=30, adjust=False).mean()
    df['stock_rs_ratio'] = df['close'] / df['e10']
    
    # Forward returns for analysis: 15m (3 bars), 30m (6 bars), 60m (12 bars)
    df['fwd_ret_15m'] = (df['close'].shift(-3) - df['close']) / df['close'] * 100
    df['fwd_ret_30m'] = (df['close'].shift(-6) - df['close']) / df['close'] * 100
    df['fwd_ret_60m'] = (df['close'].shift(-12) - df['close']) / df['close'] * 100
    
    # Forward MFE and MAE over next 6 bars (30m)
    fwd_high_6 = df['high'].iloc[::-1].rolling(6).max().iloc[::-1].shift(-1)
    fwd_low_6 = df['low'].iloc[::-1].rolling(6).min().iloc[::-1].shift(-1)
    df['fwd_mfe_30m'] = (fwd_high_6 - df['close']) / df['close'] * 100
    df['fwd_mae_30m'] = (fwd_low_6 - df['close']) / df['close'] * 100
    
    stock_features[sym] = df

print(f"Features computed for {len(stock_features)} stocks.")

records = []
t_start = time(9, 45)
t_end = time(14, 15)

for sym, df in stock_features.items():
    common_idx = df.index.intersection(ndf.index)
    if len(common_idx) == 0:
        continue
    
    sub_df = df.loc[common_idx]
    sub_nifty = ndf.loc[common_idx]
    
    mask_time = [t_start <= ts.time() <= t_end for ts in sub_df.index]
    sub_df = sub_df.loc[mask_time]
    sub_nifty = sub_nifty.loc[mask_time]
    
    rs_formula = (sub_df['stock_rs_ratio'] / sub_nifty['nifty_rs_ratio'] - 1.0) * 100.0
    stock_ret_15m = sub_df['close'].pct_change(3) * 100.0
    nifty_ret_15m = sub_nifty['nifty_ret_15m'] * 100.0
    ret_spread_15m = stock_ret_15m - nifty_ret_15m
    nifty_day_rets = [daily_nifty.get(ts.normalize(), 0.0) for ts in sub_df.index]
    
    # Absorption flag: Nifty 15m ret < -0.10% (pulling back), while Stock 15m ret >= +0.05% and Stock > VWAP and RVOL >= 1.5
    absorption = (nifty_ret_15m < -0.10) & (stock_ret_15m >= 0.05) & (sub_df['close'] > sub_df['vwap']) & (sub_df['rvol_tod'] >= 1.5)
    
    for i in range(len(sub_df)):
        ts = sub_df.index[i]
        n_day_ret = nifty_day_rets[i]
        
        if n_day_ret > 0.25:
            day_cat = 'GREEN'
        elif n_day_ret < -0.25:
            day_cat = 'RED'
        else:
            day_cat = 'FLAT'
            
        rvol_t = sub_df['rvol_tod'].iloc[i]
        rvol_r = sub_df['rvol_rolling'].iloc[i]
        rs = rs_formula.iloc[i]
        spread = ret_spread_15m.iloc[i]
        cl = sub_df['close'].iloc[i]
        vwap_val = sub_df['vwap'].iloc[i]
        above_vwap = cl > vwap_val if not np.isnan(vwap_val) else False
        fwd_15 = sub_df['fwd_ret_15m'].iloc[i]
        fwd_30 = sub_df['fwd_ret_30m'].iloc[i]
        fwd_60 = sub_df['fwd_ret_60m'].iloc[i]
        mfe_30 = sub_df['fwd_mfe_30m'].iloc[i]
        mae_30 = sub_df['fwd_mae_30m'].iloc[i]
        abs_flag = absorption.iloc[i]
        
        if np.isnan(fwd_30) or np.isnan(rvol_t) or np.isnan(rs):
            continue
            
        records.append({
            'timestamp': ts,
            'symbol': sym,
            'day_cat': day_cat,
            'nifty_day_ret': n_day_ret,
            'nifty_15m_ret': nifty_ret_15m.iloc[i],
            'rvol_tod': rvol_t,
            'rvol_rolling': rvol_r,
            'rs_formula': rs,
            'ret_spread_15m': spread,
            'above_vwap': above_vwap,
            'absorption': abs_flag,
            'fwd_ret_15m': fwd_15,
            'fwd_ret_30m': fwd_30,
            'fwd_ret_60m': fwd_60,
            'mfe_30m': mfe_30,
            'mae_30m': mae_30
        })

df_all = pd.DataFrame(records)
print(f"Total bar observations collected: {len(df_all)}")
df_all.to_parquet('scratch/decoupling_bar_observations.parquet', index=False)
print("Saved bar observations to scratch/decoupling_bar_observations.parquet")
