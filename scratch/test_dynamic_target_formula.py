"""
scratch/test_dynamic_target_formula.py
Test Dynamic Target Threshold based on:
1. Fee-to-Profit Ratio: Cost must be <= X% of Target (testing X = 15%, 20%, 25%)
2. Stock Volatility: Target >= 1.5 * ATR_5m
3. Capital & Slot-based position sizing

Compare against:
- Current fixed >= 0.75%
- Hardcoded >= 1.00%
- Hardcoded >= 1.25%
- Dynamic Mathematical Target
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

def run_target_test(mode, max_fee_ratio=0.20, fixed_val=0.0075, cap=50000.0, max_slots=3):
    C.LEVERAGE = 2.0
    C.MAX_CONCURRENT_POSITIONS = max_slots
    C.ENTRY_START = dtime(9, 25)
    C.TRUE_MOMENTUM_GATE_ENABLED = True
    C.MORNING_RVOL_MIN = 1.25
    C.MARKET_SLIPPAGE = 0.00015 # SL-Limit
    
    pos_size = (cap * C.LEVERAGE) / max_slots
    est_fee = 55.0 # approx round trip
    
    # Calculate required min target %
    if mode == 'FIXED':
        req_min_tgt = fixed_val
    elif mode == 'DYNAMIC_FEE':
        req_min_tgt = est_fee / (pos_size * max_fee_ratio)
    elif mode == 'DYNAMIC_HYBRID':
        # Fee floor, but capped reasonably
        fee_floor = est_fee / (pos_size * max_fee_ratio)
        req_min_tgt = max(0.008, fee_floor)
        
    dt0 = datetime.combine(all_dates[0], datetime.min.time())
    master = Simulator('replay', dt0, strategy='v6_profit_max', starting_capital=cap)
    master.load(raw, universe, idx)
    
    tr_count, w_count, tot_gross, tot_fees, tot_net = 0, 0, 0.0, 0.0, 0.0
    
    for d in all_dates:
        dt = datetime.combine(d, datetime.min.time())
        sim = Simulator('replay', dt, strategy='v6_profit_max', starting_capital=cap)
        sim.syms = master.syms; sim.indices = master.indices; sim.nifty_pos = master.nifty_pos
        sim.nifty_e10 = master.nifty_e10; sim.nifty_e21 = master.nifty_e21; sim.nifty_e30 = master.nifty_e30
        
        for t in sim.session_bar_times():
            decide = (t + BAR).time()
            sim.clock = t + BAR
            
            sim._manage_open_profit_max(t, decide)
            sim._manage_pending_profit_max(t)
            
            if C.ENTRY_START <= decide < C.OPT_ENTRY_CUTOFF:
                sim._scan_profit_max(t, decide)
                
                # Check pending orders against the target threshold
                for sym in list(sim.pending.keys()):
                    p = sim.pending[sym]
                    if p.get('signal_bar') == t:
                        sd = sim.syms.get(sym)
                        if sd and t in sd.pos:
                            i = sd.pos[t]
                            # Check distance to target
                            tgt_px = p.get('target', 0.0)
                            limit_px = p.get('limit', 0.0)
                            if tgt_px > 0 and limit_px > 0:
                                dist_pct = abs(tgt_px - limit_px) / limit_px
                                
                                # If dynamic hybrid, also consider stock ATR
                                atr_val = float(sd.atr[i]) if not np.isnan(sd.atr[i]) else 0.0
                                atr_pct = (atr_val / limit_px) if limit_px > 0 else 0.002
                                dynamic_threshold = max(req_min_tgt, 1.2 * atr_pct) if mode == 'DYNAMIC_HYBRID' else req_min_tgt
                                
                                if dist_pct < dynamic_threshold:
                                    del sim.pending[sym]
                                    
        k = sim.snapshot()['kpi']
        tr_count += k['trades']
        w_count += k['wins']
        tot_gross += k['gross']
        tot_fees += k['fees']
        tot_net += k['realized']
        
    wr = round(w_count / tr_count * 100, 1) if tr_count else 0
    ret_pct = round(tot_net / cap * 100, 2)
    fee_pct_gross = round(tot_fees / tot_gross * 100, 1) if tot_gross else 0
    return {
        'mode': mode,
        'req_min_tgt_pct': round(req_min_tgt * 100, 2),
        'trades': tr_count,
        'wins': w_count,
        'wr': wr,
        'gross': round(tot_gross, 2),
        'fees': round(tot_fees, 2),
        'net': round(tot_net, 2),
        'fee_pct_gross': fee_pct_gross,
        'roi': ret_pct
    }

print("Running Target Optimization tests across all 62 days (Rs 50k, 3 Slots)...")
tests = [
    ("Fixed 0.75% (Current)", "FIXED", 0.20, 0.0075),
    ("Fixed 1.00%", "FIXED", 0.20, 0.0100),
    ("Fixed 1.25%", "FIXED", 0.20, 0.0125),
    ("Dynamic Fee-Based (Fee <= 20% of Target)", "DYNAMIC_FEE", 0.20, 0.0075),
    ("Dynamic Fee-Based (Fee <= 15% of Target)", "DYNAMIC_FEE", 0.15, 0.0075),
    ("Dynamic Hybrid (Fee + ATR Volatility)", "DYNAMIC_HYBRID", 0.20, 0.0075),
]

for label, mode, ratio, fixed_val in tests:
    res = run_target_test(mode, max_fee_ratio=ratio, fixed_val=fixed_val)
    print(f"{label:38} | MinTgt: {res['req_min_tgt_pct']:4.2f}% | Trades: {res['trades']:3d} | WR: {res['wr']:4.1f}% | Gross: Rs {res['gross']:+9.2f} | Fees: Rs {res['fees']:8.2f} ({res['fee_pct_gross']:4.1f}%) | Net: Rs {res['net']:+9.2f} ({res['roi']:+5.2f}%)")
