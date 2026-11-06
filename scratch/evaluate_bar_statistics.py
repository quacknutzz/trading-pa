import pandas as pd
import numpy as np

print("Reading bar observations parquet...")
df = pd.read_parquet('scratch/decoupling_bar_observations.parquet')
print(f"Total rows: {len(df):,}")

# Clean and bin variables
df['rvol_bin'] = pd.cut(df['rvol_tod'], bins=[-np.inf, 1.0, 1.5, 2.5, np.inf], labels=['<1.0x', '1.0-1.5x', '1.5-2.5x', '>2.5x'])
df['rs_bin'] = pd.cut(df['rs_formula'], bins=[-np.inf, 0.0, 0.2, 0.5, np.inf], labels=['RS < 0 (Lagging)', 'RS 0-0.2%', 'RS 0.2-0.5%', 'RS > 0.5%'])

# Success defined as forward 30m return >= +0.30% (solid win over 0.113% tax hurdle)
df['win_30m'] = df['fwd_ret_30m'] >= 0.30
df['loss_30m'] = df['fwd_ret_30m'] <= -0.20

print("\n" + "="*80)
print("1. OVERALL STATS ON FLAT & RED DAYS (NIFTY PATHETIC / ADVERSE)")
print("="*80)

flat_red = df[df['day_cat'].isin(['FLAT', 'RED'])].copy()

# Group by (Above VWAP, RVOL bin, RS bin)
agg_stats = flat_red.groupby(['above_vwap', 'rvol_bin', 'rs_bin'], observed=False).agg(
    n_bars=('fwd_ret_30m', 'count'),
    mean_fwd_30m=('fwd_ret_30m', 'mean'),
    median_fwd_30m=('fwd_ret_30m', 'median'),
    win_rate_30m=('win_30m', lambda x: (x.sum() / len(x) * 100)),
    mfe_30m=('mfe_30m', 'mean'),
    mae_30m=('mae_30m', 'mean')
).reset_index()

# Filter for meaningful buckets
agg_stats['reward_to_risk'] = (agg_stats['mfe_30m'] / agg_stats['mae_30m'].abs()).round(2)
agg_stats['mean_fwd_30m'] = agg_stats['mean_fwd_30m'].round(3)
agg_stats['win_rate_30m'] = agg_stats['win_rate_30m'].round(1)

print("\n=== PERFORMANCE MATRIX ON FLAT/RED DAYS ===")
# Show only Above VWAP == True
vwap_true = agg_stats[agg_stats['above_vwap'] == True]
print(vwap_true.to_string(index=False))

print("\n" + "="*80)
print("2. WYCKOFF PRICE-VOLUME ABSORPTION ON NIFTY PULLBACKS")
print("="*80)
# Nifty is pulling back on 15m window (ret < -0.10%)
nifty_pullbacks = df[df['nifty_15m_ret'] < -0.10].copy()
print(f"Total 5m bars where Nifty was actively falling (15m ret < -0.10%): {len(nifty_pullbacks):,}")

abs_comp = nifty_pullbacks.groupby(['above_vwap', nifty_pullbacks['rvol_tod'] >= 1.5, nifty_pullbacks['rs_formula'] > 0.2]).agg(
    count=('fwd_ret_30m', 'count'),
    mean_fwd_30m=('fwd_ret_30m', 'mean'),
    win_rate_30m=('win_30m', lambda x: (x.sum() / len(x) * 100)),
    mean_mfe=('mfe_30m', 'mean'),
    mean_mae=('mae_30m', 'mean')
).reset_index()
abs_comp.columns = ['Above_VWAP', 'RVOL_ge_1.5', 'RS_gt_0.2%', 'Count', 'Mean_Fwd_30m', 'Win_Rate_%', 'Mean_MFE', 'Mean_MAE']
abs_comp['R:R'] = (abs_comp['Mean_MFE'] / abs_comp['Mean_MAE'].abs()).round(2)
print(abs_comp.to_string(index=False))

print("\n" + "="*80)
print("3. FLASH DUMP RESILIENCE: WHAT HAPPENS WHEN NIFTY DUMPS HARD (15m < -0.30%)?")
print("="*80)
nifty_dumps = df[df['nifty_15m_ret'] < -0.30].copy()
print(f"Total bars during steep Nifty flash drops (< -0.30% in 15m): {len(nifty_dumps):,}")

dump_comp = nifty_dumps.groupby([nifty_dumps['rvol_tod'] >= 1.5, nifty_dumps['rs_formula'] > 0.2, 'above_vwap']).agg(
    count=('fwd_ret_30m', 'count'),
    mean_fwd_30m=('fwd_ret_30m', 'mean'),
    mean_mae=('mae_30m', 'mean'),
    break_vwap_pct=('fwd_mae_30m', lambda x: (x < -0.5).sum() / len(x) * 100)
).reset_index()
dump_comp.columns = ['RVOL_ge_1.5', 'RS_gt_0.2%', 'Above_VWAP', 'Count', 'Mean_Fwd_30m', 'Mean_MAE', 'Severe_Drawdown_%']
print(dump_comp.to_string(index=False))
