"""
scratch/compare_all_three_week_and_quarter.py
Compare:
1. Strategy 1: Initial Dad Baseline (Sim & Actual Live from Excel)
2. Strategy 2: Latest Strategy without 30m skip (09:25 All)
3. Strategy 3: Latest Strategy with True Momentum Gate (RVOL >= 1.25, Clean Wick, NIFTY aligned)

For both:
A) The exact week Dad live traded (Oct 26 - Oct 30)
B) The full 62-day (~60-day) historical period
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

def run_strategy(mode, cap=50000.0):
    """
    mode:
      - 'baseline': Dad's V6 code (1.0x margin, 09:25 start)
      - 'latest_0925_all': Latest dual engine (2.0x margin, 09:25 all)
      - 'latest_true_momentum_gate': Latest dual engine with True Momentum Gate
    """
    if mode == 'baseline':
        strat_key = 'v6_baseline'
        leverage = 1.0
        gate = False
    elif mode == 'latest_0925_all':
        strat_key = 'v6_profit_max'
        leverage = 2.0
        gate = False
    else:
        strat_key = 'v6_profit_max'
        leverage = 2.0
        gate = True
        
    C.LEVERAGE = leverage
    C.ENTRY_START = dtime(9, 25)
    
    dt0 = datetime.combine(all_dates[0], datetime.min.time())
    master = Simulator('replay', dt0, strategy=strat_key, starting_capital=cap)
    master.load(raw, universe, idx)
    
    daily_res = []
    
    for d in all_dates:
        dt = datetime.combine(d, datetime.min.time())
        C.LEVERAGE = leverage
        C.ENTRY_START = dtime(9, 25)
        
        sim = Simulator('replay', dt, strategy=strat_key, starting_capital=cap)
        sim.syms = master.syms
        sim.indices = master.indices
        sim.nifty_pos = master.nifty_pos
        sim.nifty_e10 = master.nifty_e10
        sim.nifty_e21 = master.nifty_e21
        sim.nifty_e30 = master.nifty_e30
        
        for t in sim.session_bar_times():
            decide = (t + BAR).time()
            sim.clock = t + BAR
            
            if strat_key == 'v6_baseline':
                sim._manage_pending_baseline(t)
                sim._manage_open_baseline(t, decide)
                if C.ENTRY_START <= decide < C.ENTRY_CUTOFF:
                    sim._scan_baseline(t, decide)
            else:
                sim._manage_open_profit_max(t, decide)
                sim._manage_pending_profit_max(t)
                if C.ENTRY_START <= decide < C.OPT_ENTRY_CUTOFF:
                    sim._scan_profit_max(t, decide)
                    
                    if gate and decide < dtime(9, 45):
                        # True Momentum Gate filter
                        for sym in list(sim.pending.keys()):
                            p = sim.pending[sym]
                            if p.get('signal_bar') == t:
                                rvol = p.get('rvol', 1.0)
                                if rvol < 1.25:
                                    del sim.pending[sym]
                                    continue
                                sd = sim.syms.get(sym)
                                if sd and t in sd.pos:
                                    i = sd.pos[t]
                                    o_px = float(sd.df['open'].iloc[i])
                                    h_px = float(sd.df['high'].iloc[i])
                                    l_px = float(sd.df['low'].iloc[i])
                                    c_px = float(sd.df['close'].iloc[i])
                                    rng = h_px - l_px if h_px > l_px else 0.01
                                    adv_wick = (h_px - max(o_px, c_px)) / rng if p['side'] == 'BUY' else (min(o_px, c_px) - l_px) / rng
                                    if adv_wick > 0.35:
                                        del sim.pending[sym]
                                        continue
                                day_n = ndf[(ndf.index.date == d) & (ndf.index <= t + BAR)]
                                if len(day_n) >= 2:
                                    n_open = day_n['open'].iloc[0]
                                    n_now = day_n['close'].iloc[-1]
                                    n_ret = (n_now - n_open) / n_open * 100
                                    if p['side'] == 'BUY' and n_ret < -0.10:
                                        del sim.pending[sym]
                                        continue
                                    elif p['side'] == 'SELL' and n_ret > 0.10:
                                        del sim.pending[sym]
                                        continue
                                        
        k = sim.snapshot()['kpi']
        daily_res.append({
            'date': d.strftime('%Y-%m-%d'),
            'trades': k['trades'],
            'wins': k['wins'],
            'losses': k['trades'] - k['wins'],
            'wr': round(k['win_rate'], 1) if k['trades'] else 0.0,
            'gross': round(k['gross'], 2),
            'fees': round(k['fees'], 2),
            'net': round(k['realized'], 2)
        })
    return daily_res

print("Running 3-Way Comparative Replay across all 62 days...")
res_base = run_strategy('baseline')
res_0925_all = run_strategy('latest_0925_all')
res_gate = run_strategy('latest_true_momentum_gate')

target_dates = ['2026-10-26', '2026-10-27', '2026-10-28', '2026-10-29', '2026-10-30']

def aggregate(res_list, dates=None):
    if dates:
        sub = [r for r in res_list if r['date'] in dates]
    else:
        sub = res_list
    tr = sum(r['trades'] for r in sub)
    w = sum(r['wins'] for r in sub)
    wr = round(w / tr * 100, 1) if tr else 0.0
    g = sum(r['gross'] for r in sub)
    f = sum(r['fees'] for r in sub)
    n = sum(r['net'] for r in sub)
    return {'trades': tr, 'wins': w, 'wr': wr, 'gross': round(g, 2), 'fees': round(f, 2), 'net': round(n, 2)}

summary = {
    'dad_week': {
        'baseline': aggregate(res_base, target_dates),
        'latest_0925_all': aggregate(res_0925_all, target_dates),
        'latest_gated': aggregate(res_gate, target_dates)
    },
    'full_62d': {
        'baseline': aggregate(res_base),
        'latest_0925_all': aggregate(res_0925_all),
        'latest_gated': aggregate(res_gate)
    },
    'daily_dad_week': {
        'baseline': [r for r in res_base if r['date'] in target_dates],
        'latest_0925_all': [r for r in res_0925_all if r['date'] in target_dates],
        'latest_gated': [r for r in res_gate if r['date'] in target_dates]
    }
}

with open('scratch/grand_dad_vs_gate_comparison.json', 'w') as f:
    json.dump(summary, f, indent=2)

print("\n=== RESULTS FOR DAD'S LIVE TRADING WEEK (OCT 26 - OCT 30) ===")
print("1. Baseline:            ", summary['dad_week']['baseline'])
print("2. Latest 09:25 All:    ", summary['dad_week']['latest_0925_all'])
print("3. Latest Gated:        ", summary['dad_week']['latest_gated'])

print("\n=== FULL 62-DAY RESULTS ===")
print("1. Baseline:            ", summary['full_62d']['baseline'])
print("2. Latest 09:25 All:    ", summary['full_62d']['latest_0925_all'])
print("3. Latest Gated:        ", summary['full_62d']['latest_gated'])
