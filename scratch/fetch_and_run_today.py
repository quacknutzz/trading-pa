"""
scratch/fetch_and_run_today.py
Fetch today's (2026-11-03) completed 5m market session data, append to historical cache,
and run both:
1. Latest Strategy (Dual-Engine: Crossover + VWAP Continuation)
2. Original Baseline Strategy
Across starting capitals: Rs. 30k, Rs. 50k, Rs. 100k
"""
import sys, os, pickle, time, json
from datetime import datetime, date as ddate
import pandas as pd
import numpy as np

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, 'sim')
import config as C
import datafeed
from simulator import Simulator, BAR

CACHE_FILE = 'sim/cache/nse_5m_60d.pkl'
with open(CACHE_FILE, 'rb') as f:
    payload = pickle.load(f)

raw = payload['raw']
idx = payload['idx']
universe = payload['universe']

# Check latest date
sym0 = list(raw.keys())[0]
latest_cached_date = raw[sym0].index[-1].date()
today_date = datetime.now().date()
print(f"Latest date in cache: {latest_cached_date} | Today's date: {today_date}")

# Fetch today's 1d data if not in cache
if latest_cached_date < today_date:
    print(f"Fetching today's ({today_date}) completed session for 139 stocks + indices...")
    all_tickers = [u['yahoo'] for u in universe] + list(C.INDEX_TICKERS.values())
    fresh_data = datafeed.fetch(all_tickers, period='5d') # fetch last 5d to ensure full today bars
    
    # Update indices
    for idx_name, tick in C.INDEX_TICKERS.items():
        if tick in fresh_data and len(fresh_data[tick]) > 0:
            df_new = fresh_data[tick]
            df_old = idx.get(idx_name)
            idx[idx_name] = df_new if df_old is None else pd.concat([df_old, df_new]).pipe(lambda x: x[~x.index.duplicated(keep='last')]).sort_index()

    # Update stocks
    for u in universe:
        ysym = u['yahoo']
        if ysym in fresh_data and len(fresh_data[ysym]) > 0:
            df_new = fresh_data[ysym]
            df_old = raw.get(ysym)
            raw[ysym] = df_new if df_old is None else pd.concat([df_old, df_new]).pipe(lambda x: x[~x.index.duplicated(keep='last')]).sort_index()

    # Update cache on disk
    with open(CACHE_FILE, 'wb') as f:
        pickle.dump({'raw': raw, 'idx': idx, 'universe': universe}, f)
    print("Updated sim/cache/nse_5m_60d.pkl with today's market session!")

# Verify today's date is available
today_sym0 = raw[sym0]
today_dates = sorted(list(set(today_sym0.index.date)))
target_day = today_dates[-1]
print(f"Running simulation on session: {target_day} ({len(today_sym0[today_sym0.index.date == target_day])} 5m bars)")

# Run comparison on target_day across capital levels
print("\n" + "="*85)
print(f"TODAY'S MARKET SESSION AUDIT ({target_day}): LATEST STRATEGY VS ORIGINAL BASELINE")
print("="*85)

capital_levels = [30000.0, 50000.0, 100000.0]

for cap in capital_levels:
    print(f"\n-------------------------------------------------------------")
    print(f"CAPITAL LEVEL: Rs. {cap:,.0f}")
    print(f"-------------------------------------------------------------")
    
    # 1. Original Baseline
    sim_base = Simulator('replay', datetime.combine(target_day, datetime.min.time()), strategy='v6_baseline', starting_capital=cap)
    sim_base.load(raw, universe, idx)
    for t in sim_base.session_bar_times():
        sim_base.step(t)
    snap_base = sim_base.snapshot()
    k_base = snap_base['kpi']
    
    # 2. Latest Strategy (Dual-Engine)
    sim_opt = Simulator('replay', datetime.combine(target_day, datetime.min.time()), strategy='v6_profit_max', starting_capital=cap)
    sim_opt.load(raw, universe, idx)
    for t in sim_opt.session_bar_times():
        sim_opt.step(t)
    snap_opt = sim_opt.snapshot()
    k_opt = snap_opt['kpi']

    wr_base = f"{k_base['win_rate']:.1f}%" if k_base['trades'] else "N/A"
    wr_opt = f"{k_opt['win_rate']:.1f}%" if k_opt['trades'] else "N/A"
    
    print(f"Original Baseline:   Trades: {k_base['trades']:>2} | Wins: {k_base['wins']:>2} ({wr_base:<5}) | Gross: Rs {k_base['gross']:>+9,.2f} | Fees: Rs {k_base['fees']:>7,.2f} | Net: Rs {k_base['realized']:>+9,.2f} ({k_base['realized']/cap*100:+.2f}%)")
    print(f"Latest Strategy:     Trades: {k_opt['trades']:>2} | Wins: {k_opt['wins']:>2} ({wr_opt:<5}) | Gross: Rs {k_opt['gross']:>+9,.2f} | Fees: Rs {k_opt['fees']:>7,.2f} | Net: Rs {k_opt['realized']:>+9,.2f} ({k_opt['realized']/cap*100:+.2f}%)")
    
    if cap == 50000.0:
        print("\nDetailed Trades Taken Today by Latest Strategy (Rs. 50k):")
        for tr in snap_opt['closed']:
            p_ret = tr['ret_pct']
            print(f"  * {tr['side']:<4} {tr['symbol']:<10} {tr['qty']:>4} sh @ Rs {tr['entry']:<7.2f} -> Exit Rs {tr['exit']:<7.2f} ({p_ret:+.2f}%) | Net: Rs {tr['net']:>+7.2f} | {tr['reason']}")
