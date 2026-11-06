"""
scratch/run_grand_comprehensive_backtest.py
Rigorous 62-day multi-month backtest comparing:
1. Original Baseline (Dad's V6, 1.0x cash, fixed 0.75% target/SL, 09:25 start)
2. Original Baseline (2.0x margin, fixed 0.75% target/SL, 09:25 start)
3. Latest Strategy (Dual-Engine, 2.0x margin, 09:25 start)
4. Latest Strategy (Dual-Engine, 2.0x margin, 09:45 start - Skipping First 30 Mins)

Across all 62 available trading sessions (Aug 5, 2026 to Nov 3, 2026) in sim/cache/nse_5m_60d.pkl.
Includes Capital Scaling analysis (Rs 30k, Rs 50k, Rs 1L).
"""
import sys, os, pickle, json, time
from datetime import datetime, time as dtime
import pandas as pd
import numpy as np

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, '.')
sys.path.insert(0, 'sim')
import config as C
from simulator import Simulator

# 1. Load data
cache_path = 'sim/cache/nse_5m_60d.pkl'
print(f"Loading data from {cache_path}...")
t0 = time.time()
with open(cache_path, 'rb') as f:
    payload = pickle.load(f)

raw = payload['raw']
idx = payload['idx']
universe = payload['universe']
print(f"Loaded in {time.time() - t0:.2f}s.")

sym0 = list(raw.keys())[0]
all_dates = sorted(list(set(raw[sym0].index.date)))
print(f"Total available trading days: {len(all_dates)} ({all_dates[0]} to {all_dates[-1]})")

ndf = idx['NIFTY 50']

def run_suite(strat_name, entry_start, leverage=2.0, starting_cap=50000.0):
    results = []
    dt0 = datetime.combine(all_dates[0], datetime.min.time())
    
    # Configure global parameters
    C.LEVERAGE = leverage
    C.ENTRY_START = entry_start
    
    master = Simulator('replay', dt0, strategy=strat_name, starting_capital=starting_cap)
    master.load(raw, universe, idx)
    
    for d in all_dates:
        day_df = ndf[ndf.index.date == d]
        n_ret = 0.0
        day_type = "FLAT"
        if len(day_df) >= 2:
            o = day_df['open'].iloc[0]
            c = day_df['close'].iloc[-1]
            n_ret = (c - o) / o * 100
            day_type = "GREEN" if n_ret > 0.25 else ("RED" if n_ret < -0.25 else "FLAT")
            
        dt = datetime.combine(d, datetime.min.time())
        C.LEVERAGE = leverage
        C.ENTRY_START = entry_start
        sim = Simulator('replay', dt, strategy=strat_name, starting_capital=starting_cap)
        sim.syms = master.syms
        sim.indices = master.indices
        sim.nifty_pos = master.nifty_pos
        sim.nifty_e10 = master.nifty_e10
        sim.nifty_e21 = master.nifty_e21
        sim.nifty_e30 = master.nifty_e30
        
        for t in sim.session_bar_times():
            sim.step(t)
            
        k = sim.snapshot()['kpi']
        results.append({
            "date": d.strftime('%Y-%m-%d'),
            "month": d.strftime('%Y-%m'),
            "weekday": d.strftime('%a'),
            "day_type": day_type,
            "nifty_ret": round(n_ret, 2),
            "trades": k['trades'],
            "wins": k['wins'],
            "losses": k['trades'] - k['wins'],
            "wr": round(k['win_rate'], 1) if k['trades'] else 0.0,
            "gross": round(k['gross'], 2),
            "fees": round(k['fees'], 2),
            "net": round(k['realized'], 2)
        })
    return results

def compute_metrics(res_list, initial_cap=50000.0):
    total_trades = sum(r['trades'] for r in res_list)
    total_wins = sum(r['wins'] for r in res_list)
    total_losses = sum(r['losses'] for r in res_list)
    total_gross = sum(r['gross'] for r in res_list)
    total_fees = sum(r['fees'] for r in res_list)
    total_net = sum(r['net'] for r in res_list)
    
    wr = round(total_wins / total_trades * 100, 1) if total_trades else 0.0
    
    # Calculate drawdown
    cum_net = 0.0
    peak = 0.0
    max_dd = 0.0
    daily_nets = [r['net'] for r in res_list]
    for n in daily_nets:
        cum_net += n
        if cum_net > peak:
            peak = cum_net
        dd = peak - cum_net
        if dd > max_dd:
            max_dd = dd
            
    # Profit factor: sum of winning days net vs sum of losing days net (or trade-level gross wins vs gross losses)
    # Let's compute gross win sum and gross loss sum across positive and negative daily grosses
    gross_wins = sum(r['gross'] for r in res_list if r['gross'] > 0)
    gross_losses = abs(sum(r['gross'] for r in res_list if r['gross'] < 0))
    profit_factor = round(gross_wins / gross_losses, 2) if gross_losses > 0 else 99.0
    
    green_days = sum(1 for r in res_list if r['net'] > 0)
    red_days = sum(1 for r in res_list if r['net'] < 0)
    flat_days = sum(1 for r in res_list if r['net'] == 0)
    
    ret_pct = round((total_net / initial_cap) * 100, 2)
    max_dd_pct = round((max_dd / initial_cap) * 100, 2)
    avg_daily_net = round(total_net / len(res_list), 2) if res_list else 0.0
    
    return {
        "trades": total_trades,
        "wins": total_wins,
        "losses": total_losses,
        "win_rate": wr,
        "gross": round(total_gross, 2),
        "fees": round(total_fees, 2),
        "net": round(total_net, 2),
        "return_pct": ret_pct,
        "max_dd": round(max_dd, 2),
        "max_dd_pct": max_dd_pct,
        "profit_factor": profit_factor,
        "green_days": green_days,
        "red_days": red_days,
        "flat_days": flat_days,
        "avg_daily_net": avg_daily_net
    }

print("\n1. Running Strategy 1: Original Baseline (1.0x Margin, 09:25 Start)...")
t_start = time.time()
res_baseline_1x = run_suite('v6_baseline', dtime(9, 25), leverage=1.0, starting_cap=50000.0)
print(f"Done in {time.time() - t_start:.2f}s")

print("\n2. Running Strategy 2: Latest Strategy (09:25 AM Start, 2.0x Margin)...")
t_start = time.time()
res_latest_0925 = run_suite('v6_profit_max', dtime(9, 25), leverage=2.0, starting_cap=50000.0)
print(f"Done in {time.time() - t_start:.2f}s")

print("\n3. Running Strategy 3: Latest Strategy (09:45 AM Start, 2.0x Margin - Skipping 30m)...")
t_start = time.time()
res_latest_0945 = run_suite('v6_profit_max', dtime(9, 45), leverage=2.0, starting_cap=50000.0)
print(f"Done in {time.time() - t_start:.2f}s")

# Monthly Breakdown
months = sorted(list(set(r['month'] for r in res_baseline_1x)))

monthly_summary = {}
for m in months:
    b1_m = [r for r in res_baseline_1x if r['month'] == m]
    l25_m = [r for r in res_latest_0925 if r['month'] == m]
    l45_m = [r for r in res_latest_0945 if r['month'] == m]
    
    monthly_summary[m] = {
        "baseline": compute_metrics(b1_m),
        "latest_0925": compute_metrics(l25_m),
        "latest_0945": compute_metrics(l45_m)
    }

# Last 7 Days Breakdown
last7_dates = all_dates[-7:]
last7_b1 = [r for r in res_baseline_1x if r['date'] in [d.strftime('%Y-%m-%d') for d in last7_dates]]
last7_l25 = [r for r in res_latest_0925 if r['date'] in [d.strftime('%Y-%m-%d') for d in last7_dates]]
last7_l45 = [r for r in res_latest_0945 if r['date'] in [d.strftime('%Y-%m-%d') for d in last7_dates]]

overall_summary = {
    "baseline_1x": compute_metrics(res_baseline_1x),
    "latest_0925": compute_metrics(res_latest_0925),
    "latest_0945": compute_metrics(res_latest_0945),
    "last7_baseline": compute_metrics(last7_b1),
    "last7_latest_0925": compute_metrics(last7_l25),
    "last7_latest_0945": compute_metrics(last7_l45),
    "monthly": monthly_summary,
    "daily_records": {
        "baseline_1x": res_baseline_1x,
        "latest_0925": res_latest_0925,
        "latest_0945": res_latest_0945
    }
}

with open('scratch/grand_comprehensive_backtest.json', 'w') as f:
    json.dump(overall_summary, f, indent=2)

print("\n=== OVERALL 62-DAY RESULTS (₹50,000 Capital) ===")
print("Strategy 1 (Baseline 1.0x):", overall_summary["baseline_1x"])
print("Strategy 2 (Latest 09:25):", overall_summary["latest_0925"])
print("Strategy 3 (Latest 09:45):", overall_summary["latest_0945"])

print("\n=== LAST 7-DAY RESULTS (₹50,000 Capital) ===")
print("Strategy 1 (Baseline 1.0x):", overall_summary["last7_baseline"])
print("Strategy 2 (Latest 09:25):", overall_summary["last7_latest_0925"])
print("Strategy 3 (Latest 09:45):", overall_summary["last7_latest_0945"])
