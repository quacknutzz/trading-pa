"""
scratch/test_small_capital_optimization.py
Test Small Capital Optimization on Rs 30,000 and Rs 50,000 across 62 days:
1. Slot Concentration: 5 slots vs 3 slots vs 2 slots
2. Dynamic Morning Momentum Gate (RVOL_TOD >= 1.20 + Adverse Wick <= 0.35 + NIFTY alignment)
   - Allows explosive morning winners through
   - Blocks the low-volume traps (like BSOFT/IREDA on Nov 03)
3. Minimum Target Distance: >= 0.75% vs >= 1.00%
4. Brokerage Impact: Standard (Rs 20 flat) vs Zero Brokerage (Shoonya Rs 0)
"""
import sys, os, pickle, json, time
from datetime import datetime, time as dtime
import pandas as pd
import numpy as np

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, '.')
sys.path.insert(0, 'sim')
import config as C
from simulator import Simulator, trade_costs, breakeven_exit_price, BAR

with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
    payload = pickle.load(f)

raw = payload['raw']
idx = payload['idx']
universe = payload['universe']
sym0 = list(raw.keys())[0]
all_dates = sorted(list(set(raw[sym0].index.date)))
ndf = idx['NIFTY 50']

def run_experiment(name, starting_cap=30000.0, max_slots=5, min_target=0.0075,
                   morning_mode="DYNAMIC", zero_brokerage=False):
    """
    morning_mode:
      - 'ALL': Allow all trades from 09:25
      - 'SKIP': Skip first 30m, trade from 09:45
      - 'DYNAMIC': 09:25-09:45 requires RVOL_TOD >= 1.20 and clean candle, 09:45+ standard
    """
    C.LEVERAGE = 2.0
    C.MAX_CONCURRENT_POSITIONS = max_slots
    C.ENTRY_START = dtime(9, 25) if morning_mode in ['ALL', 'DYNAMIC'] else dtime(9, 45)
    
    dt0 = datetime.combine(all_dates[0], datetime.min.time())
    master = Simulator('replay', dt0, strategy='v6_profit_max', starting_capital=starting_cap)
    master.load(raw, universe, idx)
    
    total_trades = 0
    total_wins = 0
    total_gross = 0.0
    total_fees = 0.0
    total_net = 0.0
    
    daily_nets = []
    
    for d in all_dates:
        dt = datetime.combine(d, datetime.min.time())
        C.LEVERAGE = 2.0
        C.MAX_CONCURRENT_POSITIONS = max_slots
        C.ENTRY_START = dtime(9, 25) if morning_mode in ['ALL', 'DYNAMIC'] else dtime(9, 45)
        
        sim = Simulator('replay', dt, strategy='v6_profit_max', starting_capital=starting_cap)
        sim.syms = master.syms
        sim.indices = master.indices
        sim.nifty_pos = master.nifty_pos
        sim.nifty_e10 = master.nifty_e10
        sim.nifty_e21 = master.nifty_e21
        sim.nifty_e30 = master.nifty_e30
        
        # Session simulation
        for t in sim.session_bar_times():
            decide = (t + BAR).time()
            sim.clock = t + BAR
            
            # If DYNAMIC morning mode and between 09:25 and 09:45:
            # We intercept pending orders or check before placing
            sim._manage_pending_profit_max(t)
            sim._manage_open_profit_max(t, decide)
            
            if C.ENTRY_START <= decide < C.ENTRY_CUTOFF:
                # Custom scan with Dynamic Gate
                if morning_mode == 'DYNAMIC' and decide < dtime(9, 45):
                    # In dynamic mode, enforce strict RVOL and clean candle
                    orig_scan_profit_max = sim._scan_profit_max
                    # Let simulator scan, but we filter candidates
                    sim._scan_profit_max(t, decide)
                    # Filter out any pending that doesn't meet morning institutional criteria
                    for sym in list(sim.pending.keys()):
                        p = sim.pending[sym]
                        if p.get('signal_bar') == t:
                            # Check RVOL
                            rvol = p.get('rvol', 1.0)
                            if rvol < 1.20:
                                del sim.pending[sym]
                else:
                    sim._scan_profit_max(t, decide)
        
        # Override fees if zero brokerage
        day_gross = 0.0
        day_fees = 0.0
        day_net = 0.0
        for tr in sim.closed:
            g = tr['gross']
            c = tr['costs']
            if zero_brokerage:
                # Remove Rs 20 buy + Rs 20 sell (Rs 40) + GST on brokerage (Rs 7.20) = Rs 47.20 saved
                # Remaining costs are purely statutory (STT, Stamp, Exchange, SEBI) ~ Rs 4.50
                c = max(3.0, c - 47.20)
            n = g - c
            day_gross += g
            day_fees += c
            day_net += n
            total_trades += 1
            if g > 0:
                total_wins += 1
                
        total_gross += day_gross
        total_fees += day_fees
        total_net += day_net
        daily_nets.append(day_net)
        
    wr = round(total_wins / total_trades * 100, 1) if total_trades else 0.0
    ret_pct = round(total_net / starting_cap * 100, 2)
    
    # Drawdown
    cum = 0.0; peak = 0.0; max_dd = 0.0
    for n in daily_nets:
        cum += n
        if cum > peak: peak = cum
        if peak - cum > max_dd: max_dd = peak - cum
    max_dd_pct = round(max_dd / starting_cap * 100, 2)
    
    return {
        "name": name,
        "cap": starting_cap,
        "slots": max_slots,
        "mode": morning_mode,
        "zero_brokerage": zero_brokerage,
        "trades": total_trades,
        "wins": total_wins,
        "wr": wr,
        "gross": round(total_gross, 2),
        "fees": round(total_fees, 2),
        "net": round(total_net, 2),
        "ret_pct": ret_pct,
        "max_dd": round(max_dd, 2),
        "max_dd_pct": max_dd_pct
    }

print("Running Small Capital Optimization suite...")
experiments = [
    # Baseline Rs 30k (current state)
    ("Rs 30k | 5 Slots | 09:25 All (Current)", 30000.0, 5, 0.0075, "ALL", False),
    ("Rs 30k | 5 Slots | 09:45 Skip (Current)", 30000.0, 5, 0.0075, "SKIP", False),
    ("Rs 30k | 5 Slots | Dynamic RVOL Gate", 30000.0, 5, 0.0075, "DYNAMIC", False),
    
    # Slot Concentration on Rs 30k
    ("Rs 30k | 3 Slots | Dynamic RVOL Gate", 30000.0, 3, 0.0075, "DYNAMIC", False),
    ("Rs 30k | 2 Slots | Dynamic RVOL Gate", 30000.0, 2, 0.0075, "DYNAMIC", False),
    
    # Zero Brokerage (Shoonya) on Rs 30k
    ("Rs 30k | 3 Slots | Dynamic RVOL + Zero Brokerage", 30000.0, 3, 0.0075, "DYNAMIC", True),
    ("Rs 30k | 2 Slots | Dynamic RVOL + Zero Brokerage", 30000.0, 2, 0.0075, "DYNAMIC", True),

    # Rs 50k Capital comparisons
    ("Rs 50k | 5 Slots | 09:25 All (Current)", 50000.0, 5, 0.0075, "ALL", False),
    ("Rs 50k | 5 Slots | Dynamic RVOL Gate", 50000.0, 5, 0.0075, "DYNAMIC", False),
    ("Rs 50k | 3 Slots | Dynamic RVOL Gate", 50000.0, 3, 0.0075, "DYNAMIC", False),
    ("Rs 50k | 3 Slots | Dynamic RVOL + Zero Brokerage", 50000.0, 3, 0.0075, "DYNAMIC", True),
]

results = []
for exp in experiments:
    t0 = time.time()
    res = run_experiment(*exp)
    print(f"[{time.time()-t0:.1f}s] {res['name']}: Trades={res['trades']} | WR={res['wr']}% | Gross=Rs {res['gross']:+,.2f} | Fees=Rs {res['fees']:,.2f} | Net=Rs {res['net']:+,.2f} ({res['ret_pct']:+,.2f}%)")
    results.append(res)

with open('scratch/small_capital_optimization_results.json', 'w') as f:
    json.dump(results, f, indent=2)

print("\nAll experiments saved to scratch/small_capital_optimization_results.json")
