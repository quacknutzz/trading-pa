"""
scratch/optimize_momentum_gate_threshold.py
Test different MORNING_RVOL_MIN thresholds (1.05, 1.10, 1.15, 1.20, 1.25)
to see if we can capture the high-momentum September runners while still
blocking the November 03 low-volume traps (< 0.6x).
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

def test_rvol_threshold(rvol_thresh, cap=50000.0):
    C.LEVERAGE = 2.0
    C.ENTRY_START = dtime(9, 25)
    C.TRUE_MOMENTUM_GATE_ENABLED = True
    C.MORNING_RVOL_MIN = rvol_thresh
    
    dt0 = datetime.combine(all_dates[0], datetime.min.time())
    master = Simulator('replay', dt0, strategy='v6_profit_max', starting_capital=cap)
    master.load(raw, universe, idx)
    
    tot_trades, tot_wins = 0, 0
    tot_gross, tot_fees, tot_net = 0.0, 0.0, 0.0
    daily_nets = []
    
    for d in all_dates:
        dt = datetime.combine(d, datetime.min.time())
        C.LEVERAGE = 2.0
        C.ENTRY_START = dtime(9, 25)
        C.TRUE_MOMENTUM_GATE_ENABLED = True
        C.MORNING_RVOL_MIN = rvol_thresh
        
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
        daily_nets.append((d.strftime('%Y-%m-%d'), k['realized']))
        
    wr = round(tot_wins / tot_trades * 100, 1) if tot_trades else 0.0
    nov3_net = [n for dt_str, n in daily_nets if dt_str == '2026-11-03'][0]
    nov2_net = [n for dt_str, n in daily_nets if dt_str == '2026-11-02'][0]
    
    return {
        'threshold': rvol_thresh,
        'trades': tot_trades,
        'wins': tot_wins,
        'wr': wr,
        'gross': round(tot_gross, 2),
        'fees': round(tot_fees, 2),
        'net': round(tot_net, 2),
        'nov2_net': round(nov2_net, 2),
        'nov3_net': round(nov3_net, 2)
    }

print("Testing MORNING_RVOL_MIN thresholds on Rs 50,000 capital...")
for thresh in [0.0, 1.05, 1.10, 1.15, 1.20, 1.25]:
    # 0.0 represents Ungated
    if thresh == 0.0:
        C.TRUE_MOMENTUM_GATE_ENABLED = False
    else:
        C.TRUE_MOMENTUM_GATE_ENABLED = True
    res = test_rvol_threshold(thresh)
    mode_name = "Ungated (09:25 All)" if thresh == 0.0 else f"Gated (RVOL >= {thresh:.2f})"
    print(f"{mode_name:24} | Trades: {res['trades']} | WR: {res['wr']}% | Gross: Rs {res['gross']:+,.2f} | Fees: Rs {res['fees']:,.2f} | Net: Rs {res['net']:+,.2f} | Nov 2: Rs {res['nov2_net']:+,.2f} | Nov 3: Rs {res['nov3_net']:+,.2f}")
