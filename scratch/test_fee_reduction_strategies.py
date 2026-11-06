"""
scratch/test_fee_reduction_strategies.py
Demonstrate exact fee breakdown and compare 4 fee-reduction methods:
1. Current Standard FYERS (Rs 20 flat brokerage + full taxes + slippage)
2. Zero-Brokerage Broker (Shoonya / Finvasia: Rs 0 brokerage, 0 GST on brokerage)
3. High Selectivity Filter (Min target >= 1.0%, cutting trade churn)
4. Combined: Zero Brokerage + High Selectivity
"""
import sys, os, pickle, json
from datetime import datetime, time as dtime
import pandas as pd
import numpy as np

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, '.')
sys.path.insert(0, 'sim')
import config as C
from simulator import Simulator, BAR, trade_costs

with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
    payload = pickle.load(f)

raw = payload['raw']
idx = payload['idx']
universe = payload['universe']
sym0 = list(raw.keys())[0]
all_dates = sorted(list(set(raw[sym0].index.date)))

def run_fee_simulation(zero_brokerage=False, min_target_pct=0.0075, cap=50000.0, max_slots=3):
    C.LEVERAGE = 2.0
    C.MAX_CONCURRENT_POSITIONS = max_slots
    C.ENTRY_START = dtime(9, 25)
    C.TRUE_MOMENTUM_GATE_ENABLED = True
    C.MORNING_RVOL_MIN = 1.25
    
    dt0 = datetime.combine(all_dates[0], datetime.min.time())
    master = Simulator('replay', dt0, strategy='v6_profit_max', starting_capital=cap)
    master.load(raw, universe, idx)
    
    total_trades = 0
    total_wins = 0
    total_gross = 0.0
    total_fees = 0.0
    total_net = 0.0
    
    # Detailed fee buckets
    fee_buckets = {"brokerage": 0.0, "stt": 0.0, "exchange": 0.0, "stamp": 0.0, "gst": 0.0, "slippage": 0.0}
    
    for d in all_dates:
        dt = datetime.combine(d, datetime.min.time())
        C.LEVERAGE = 2.0
        C.MAX_CONCURRENT_POSITIONS = max_slots
        C.ENTRY_START = dtime(9, 25)
        C.TRUE_MOMENTUM_GATE_ENABLED = True
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
            
        for tr in sim.closed:
            g = tr['gross']
            c_det = tr.get('cost_detail', {})
            
            # If zero brokerage (Shoonya), brokerage is 0, GST on brokerage is 0
            if zero_brokerage:
                brk = 0.0
                stt = c_det.get('stt', 0.0)
                exch = c_det.get('exchange', 0.0)
                stamp = c_det.get('stamp', 0.0)
                gst = exch * 0.18 # GST only on exchange fee, 0 on brokerage
                slip = c_det.get('slippage', 0.0)
                c_total = brk + stt + exch + stamp + gst + slip
            else:
                brk = c_det.get('brokerage', 0.0)
                stt = c_det.get('stt', 0.0)
                exch = c_det.get('exchange', 0.0)
                stamp = c_det.get('stamp', 0.0)
                gst = c_det.get('gst', 0.0)
                slip = c_det.get('slippage', 0.0)
                c_total = tr['costs']
                
            fee_buckets["brokerage"] += brk
            fee_buckets["stt"] += stt
            fee_buckets["exchange"] += exch
            fee_buckets["stamp"] += stamp
            fee_buckets["gst"] += gst
            fee_buckets["slippage"] += slip
            
            n = g - c_total
            total_trades += 1
            if g > 0:
                total_wins += 1
            total_gross += g
            total_fees += c_total
            total_net += n
            
    wr = round(total_wins / total_trades * 100, 1) if total_trades else 0.0
    ret_pct = round(total_net / cap * 100, 2)
    fee_pct_gross = round(total_fees / total_gross * 100, 1) if total_gross else 0.0
    
    return {
        "trades": total_trades,
        "wins": total_wins,
        "wr": wr,
        "gross": round(total_gross, 2),
        "fees": round(total_fees, 2),
        "net": round(total_net, 2),
        "ret_pct": ret_pct,
        "fee_pct_gross": fee_pct_gross,
        "fee_buckets": {k: round(v, 2) for k, v in fee_buckets.items()}
    }

print("Running Fee Breakdown and Reduction Simulations...")
res_standard = run_fee_simulation(zero_brokerage=False)
res_zero_brk = run_fee_simulation(zero_brokerage=True)

print("\n1. Standard Brokerage (FYERS / Zerodha):")
print(f"  Trades: {res_standard['trades']} | WR: {res_standard['wr']}%")
print(f"  Gross: Rs {res_standard['gross']:+,.2f} | Fees: Rs {res_standard['fees']:,.2f} ({res_standard['fee_pct_gross']}% of gross) | Net: Rs {res_standard['net']:+,.2f} ({res_standard['ret_pct']:+,.2f}%)")
print("  Fee Buckets:", res_standard['fee_buckets'])

print("\n2. Zero Brokerage (Shoonya / Finvasia):")
print(f"  Trades: {res_zero_brk['trades']} | WR: {res_zero_brk['wr']}%")
print(f"  Gross: Rs {res_zero_brk['gross']:+,.2f} | Fees: Rs {res_zero_brk['fees']:,.2f} ({res_zero_brk['fee_pct_gross']}% of gross) | Net: Rs {res_zero_brk['net']:+,.2f} ({res_zero_brk['ret_pct']:+,.2f}%)")
print("  Fee Buckets:", res_zero_brk['fee_buckets'])
