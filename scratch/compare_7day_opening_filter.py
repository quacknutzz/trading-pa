"""
scratch/compare_7day_opening_filter.py
Compare Variant 1 (ENTRY_START = 09:25) vs Variant 2 (ENTRY_START = 09:45)
Across the last 7 trading days (Oct 26 - Nov 03) and the full 62-day dataset.
Starting Capital: Rs. 50,000 (2.0x MIS Intraday Margin).
"""
import sys, os, pickle, json
from datetime import datetime, time as dtime
import pandas as pd
import numpy as np

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, 'sim')
import config as C
from simulator import Simulator

with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
    payload = pickle.load(f)

raw = payload['raw']
idx = payload['idx']
universe = payload['universe']

sym0 = list(raw.keys())[0]
all_dates = sorted(list(set(raw[sym0].index.date)))
last_7_dates = all_dates[-7:]

ndf = idx['NIFTY 50']

def run_simulation(dates_list, entry_start_time, cap=50000.0):
    results = []
    # Preload master once for speed
    dt0 = datetime.combine(dates_list[0], datetime.min.time())
    master = Simulator('replay', dt0, strategy='v6_profit_max', starting_capital=cap)
    master.load(raw, universe, idx)
    
    for d in dates_list:
        day_df = ndf[ndf.index.date == d]
        n_ret = 0.0
        day_type = "FLAT"
        if len(day_df) >= 2:
            o = day_df['open'].iloc[0]
            c = day_df['close'].iloc[-1]
            n_ret = (c - o) / o * 100
            day_type = "GREEN" if n_ret > 0.25 else ("RED" if n_ret < -0.25 else "FLAT")
            
        dt = datetime.combine(d, datetime.min.time())
        C.ENTRY_START = entry_start_time
        sim = Simulator('replay', dt, strategy='v6_profit_max', starting_capital=cap)
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

print("Running 7-Day Comparison: 09:25 Start vs 09:45 Start...")
res_7d_0925 = run_simulation(last_7_dates, dtime(9, 25))
res_7d_0945 = run_simulation(last_7_dates, dtime(9, 45))

print("Running Full 62-Day Dataset Comparison...")
res_all_0925 = run_simulation(all_dates, dtime(9, 25))
res_all_0945 = run_simulation(all_dates, dtime(9, 45))

summary_output = {
    "seven_day": {
        "v1_0925": res_7d_0925,
        "v2_0945": res_7d_0945
    },
    "full_62d_totals": {
        "v1_0925": {
            "trades": sum(r['trades'] for r in res_all_0925),
            "wins": sum(r['wins'] for r in res_all_0925),
            "losses": sum(r['losses'] for r in res_all_0925),
            "gross": round(sum(r['gross'] for r in res_all_0925), 2),
            "fees": round(sum(r['fees'] for r in res_all_0925), 2),
            "net": round(sum(r['net'] for r in res_all_0925), 2),
        },
        "v2_0945": {
            "trades": sum(r['trades'] for r in res_all_0945),
            "wins": sum(r['wins'] for r in res_all_0945),
            "losses": sum(r['losses'] for r in res_all_0945),
            "gross": round(sum(r['gross'] for r in res_all_0945), 2),
            "fees": round(sum(r['fees'] for r in res_all_0945), 2),
            "net": round(sum(r['net'] for r in res_all_0945), 2),
        }
    }
}

with open('scratch/7day_opening_comparison.json', 'w') as f:
    json.dump(summary_output, f, indent=2)

print("Comparison complete! Results saved to scratch/7day_opening_comparison.json")
