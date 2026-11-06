import sys, os, pickle, json
from datetime import datetime
import pandas as pd
import numpy as np

sys.path.insert(0, '.')
sys.path.insert(0, 'sim')
import config as C
from simulator import Simulator

CACHE_FILE = 'sim/cache/nse_5m_60d.pkl'
with open(CACHE_FILE, 'rb') as f:
    payload = pickle.load(f)

raw = payload['raw']
idx = payload['idx']
universe = payload['universe']
sym0 = list(raw.keys())[0]
all_dates = sorted(list(set(raw[sym0].index.date)))

print(f"Generating comprehensive 64-day report for {len(all_dates)} trading days ({all_dates[0]} to {all_dates[-1]})...")

dt0 = datetime.combine(all_dates[0], datetime.min.time())
master = Simulator('replay', dt0, strategy='v6_profit_max', starting_capital=500000.0)
master.load(raw, universe, idx)

# Group dates into calendar weeks (Monday to Friday)
date_objs = [pd.Timestamp(d) for d in all_dates]
df_dates = pd.DataFrame({'date': all_dates, 'ts': date_objs})
df_dates['year'] = df_dates['ts'].dt.isocalendar().year
df_dates['week_num'] = df_dates['ts'].dt.isocalendar().week

weeks_grouped = []
for (yr, wnum), grp in df_dates.groupby(['year', 'week_num'], sort=False):
    w_dates = list(grp['date'])
    weeks_grouped.append(w_dates)

print(f"Total weeks: {len(weeks_grouped)}")

def run_strategy_on_dates(strat_name, cap=500000.0):
    all_closed = []
    week_stats = []
    
    for w_idx, w_dates in enumerate(weeks_grouped):
        w_closed = []
        for d in w_dates:
            dt = datetime.combine(d, datetime.min.time())
            sim = Simulator('replay', dt, strategy=strat_name, starting_capital=cap)
            sim.syms = master.syms; sim.indices = master.indices
            sim.nifty_pos = master.nifty_pos; sim.nifty_e10 = master.nifty_e10
            sim.nifty_e21 = master.nifty_e21; sim.nifty_e30 = master.nifty_e30
            for t in sim.session_bar_times():
                sim.step(t)
            w_closed.extend(sim.closed)
            all_closed.extend(sim.closed)
        
        n = len(w_closed)
        w_wins = sum(1 for c in w_closed if c['net'] > 0)
        wr = round(w_wins / n * 100, 1) if n else 0.0
        gross = round(sum(c['gross'] for c in w_closed), 2)
        fees = round(sum(c['costs'] for c in w_closed), 2)
        net = round(sum(c['net'] for c in w_closed), 2)
        
        d_start_str = w_dates[0].strftime("%b %d")
        d_end_str = w_dates[-1].strftime("%b %d")
        lbl = f"W{w_idx+1:02d} ({d_start_str} - {d_end_str})"
        
        week_stats.append({
            'week': lbl,
            'days': len(w_dates),
            'trades': n,
            'wins': w_wins,
            'wr': wr,
            'gross': gross,
            'fees': fees,
            'net': net
        })
        
    n_tot = len(all_closed)
    tot_wins = sum(1 for c in all_closed if c['net'] > 0)
    tot_wr = round(tot_wins / n_tot * 100, 1) if n_tot else 0.0
    tot_gross = round(sum(c['gross'] for c in all_closed), 2)
    tot_fees = round(sum(c['costs'] for c in all_closed), 2)
    tot_net = round(sum(c['net'] for c in all_closed), 2)
    
    return {
        'totals': {
            'trades': n_tot,
            'wins': tot_wins,
            'win_rate': tot_wr,
            'gross': tot_gross,
            'fees': tot_fees,
            'net': tot_net
        },
        'weeks': week_stats
    }

print("Running v6_baseline...")
res_base = run_strategy_on_dates('v6_baseline', cap=500000.0)
print("Running v6_profit_max...")
res_opt = run_strategy_on_dates('v6_profit_max', cap=500000.0)

# Merge into output json format matching web UI
merged_weeks = []
for i in range(len(weeks_grouped)):
    wb = res_base['weeks'][i]
    wo = res_opt['weeks'][i]
    merged_weeks.append({
        'week': wb['week'],
        'days': wb['days'],
        'baseline': wb,
        'profit_max': wo
    })

report_data = {
    'generated_at': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    'period': f"{all_dates[0]} to {all_dates[-1]} ({len(all_dates)} Trading Days / {len(weeks_grouped)} Weeks)",
    'capital': 500000,
    'leverage': 2.0,
    'totals': {
        'v6_baseline': res_base['totals'],
        'v6_profit_max': res_opt['totals']
    },
    'weeks': merged_weeks
}

out_path = 'sim/journal/backtest_12week_report.json'
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(report_data, f, indent=2)

print(f"Updated {out_path} successfully!")
print("Baseline Totals:  ", res_base['totals'])
print("ProfitMax Totals: ", res_opt['totals'])
