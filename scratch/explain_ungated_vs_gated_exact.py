"""
scratch/explain_ungated_vs_gated_exact.py
Run an exact side-by-side test of:
1. 5 Slots: Ungated vs Gated (RVOL >= 1.25)
2. 3 Slots: Ungated vs Gated (RVOL >= 1.25)

On Rs 50,000 capital across all 62 days.
Isolate:
- Trades taken
- Win rate
- Gross profit
- Fees paid
- Net profit
- Why slot concentration flipped Gated to BEAT Ungated!
"""
import sys, os, pickle, json
from datetime import datetime, time as dtime
import pandas as pd
import numpy as np

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, '.')
sys.path.insert(0, 'sim')
import config as C
from simulator import Simulator, BAR

with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
    payload = pickle.load(f)

raw = payload['raw']
idx = payload['idx']
universe = payload['universe']
sym0 = list(raw.keys())[0]
all_dates = sorted(list(set(raw[sym0].index.date)))
ndf = idx['NIFTY 50']

def run_experiment(max_slots, gate_enabled, cap=50000.0):
    C.LEVERAGE = 2.0
    C.MAX_CONCURRENT_POSITIONS = max_slots
    C.ENTRY_START = dtime(9, 25)
    C.TRUE_MOMENTUM_GATE_ENABLED = gate_enabled
    C.MORNING_RVOL_MIN = 1.25
    
    dt0 = datetime.combine(all_dates[0], datetime.min.time())
    master = Simulator('replay', dt0, strategy='v6_profit_max', starting_capital=cap)
    master.load(raw, universe, idx)
    
    tot_trades, tot_wins = 0, 0
    tot_gross, tot_fees, tot_net = 0.0, 0.0, 0.0
    daily_results = []
    
    for d in all_dates:
        dt = datetime.combine(d, datetime.min.time())
        C.LEVERAGE = 2.0
        C.MAX_CONCURRENT_POSITIONS = max_slots
        C.ENTRY_START = dtime(9, 25)
        C.TRUE_MOMENTUM_GATE_ENABLED = gate_enabled
        C.MORNING_RVOL_MIN = 1.25
        
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
        tot_trades += k['trades']
        tot_wins += k['wins']
        tot_gross += k['gross']
        tot_fees += k['fees']
        tot_net += k['realized']
        daily_results.append((d.strftime('%Y-%m-%d'), k['realized']))
        
    wr = round(tot_wins / tot_trades * 100, 1) if tot_trades else 0.0
    return {
        'slots': max_slots,
        'gated': gate_enabled,
        'trades': tot_trades,
        'wins': tot_wins,
        'wr': wr,
        'gross': round(tot_gross, 2),
        'fees': round(tot_fees, 2),
        'net': round(tot_net, 2),
        'daily': daily_results
    }

print("Running 4 experiments to explain the exact mechanics...")
print("1. 5 Slots - Ungated...")
r_5_ungated = run_experiment(5, False)

print("2. 5 Slots - Gated...")
r_5_gated = run_experiment(5, True)

print("3. 3 Slots - Ungated...")
r_3_ungated = run_experiment(3, False)

print("4. 3 Slots - Gated...")
r_3_gated = run_experiment(3, True)

results = {
    '5_slots_ungated': r_5_ungated,
    '5_slots_gated': r_5_gated,
    '3_slots_ungated': r_3_ungated,
    '3_slots_gated': r_3_gated
}

with open('scratch/slots_and_gate_comparison.json', 'w') as f:
    json.dump(results, f, indent=2)

print("\n=== SUMMARY TABLE ===")
for k, v in results.items():
    print(f"{k:18} | Trades: {v['trades']} | WR: {v['wr']}% | Gross: Rs {v['gross']:+,.2f} | Fees: Rs {v['fees']:,.2f} | Net: Rs {v['net']:+,.2f}")
