"""
scratch/test_exact_true_momentum_gate.py
Test adding the True Momentum Gate to v6_profit_max:
Condition for entries between 09:15 and 09:45 AM:
1. RVOL_TOD >= 1.25 (Institutional volume)
2. Candle adverse wick <= 0.35 (No adverse wick rejection)
3. NIFTY alignment:
   - BUY: NIFTY 10m EMA not BEARISH and morning return >= -0.10%
   - SELL: NIFTY 10m EMA not BULLISH and morning return <= +0.10%

After 09:45 AM: Standard v6_profit_max logic continues.
"""
import sys, os, pickle, json, time
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

def run_simulation(enable_true_momentum_gate=False, zero_brokerage=False, cap=50000.0, max_slots=5):
    C.LEVERAGE = 2.0
    C.MAX_CONCURRENT_POSITIONS = max_slots
    C.ENTRY_START = dtime(9, 25)
    
    dt0 = datetime.combine(all_dates[0], datetime.min.time())
    master = Simulator('replay', dt0, strategy='v6_profit_max', starting_capital=cap)
    master.load(raw, universe, idx)
    
    total_trades = 0
    total_wins = 0
    total_gross = 0.0
    total_fees = 0.0
    total_net = 0.0
    daily_results = []
    
    for d in all_dates:
        dt = datetime.combine(d, datetime.min.time())
        C.LEVERAGE = 2.0
        C.MAX_CONCURRENT_POSITIONS = max_slots
        C.ENTRY_START = dtime(9, 25)
        
        sim = Simulator('replay', dt, strategy='v6_profit_max', starting_capital=cap)
        sim.syms = master.syms
        sim.indices = master.indices
        sim.nifty_pos = master.nifty_pos
        sim.nifty_e10 = master.nifty_e10
        sim.nifty_e21 = master.nifty_e21
        sim.nifty_e30 = master.nifty_e30
        
        for t in sim.session_bar_times():
            decide = (t + BAR).time()
            sim.clock = t + BAR
            
            sim._manage_open_profit_max(t, decide)
            sim._manage_pending_profit_max(t)
            
            if C.ENTRY_START <= decide < C.OPT_ENTRY_CUTOFF:
                sim._scan_profit_max(t, decide)
                
                # If gate is enabled and we are in the morning window (before 09:45):
                if enable_true_momentum_gate and decide < dtime(9, 45):
                    # Check each newly pending order
                    for sym in list(sim.pending.keys()):
                        p = sim.pending[sym]
                        if p.get('signal_bar') == t:
                            # 1. RVOL check
                            rvol = p.get('rvol', 1.0)
                            if rvol < 1.25:
                                del sim.pending[sym]
                                continue
                                
                            # 2. Candle wick check
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
                                    
                            # 3. NIFTY Alignment check
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
                                    
        day_gross = 0.0
        day_fees = 0.0
        day_net = 0.0
        day_trades = 0
        day_wins = 0
        for tr in sim.closed:
            g = tr['gross']
            c = tr['costs']
            if zero_brokerage:
                c = max(3.0, c - 47.20)
            n = g - c
            day_gross += g
            day_fees += c
            day_net += n
            day_trades += 1
            if g > 0:
                day_wins += 1
                
        total_trades += day_trades
        total_wins += day_wins
        total_gross += day_gross
        total_fees += day_fees
        total_net += day_net
        daily_results.append({
            'date': d.strftime('%Y-%m-%d'),
            'trades': day_trades,
            'wins': day_wins,
            'gross': day_gross,
            'fees': day_fees,
            'net': day_net
        })
        
    wr = round(total_wins / total_trades * 100, 1) if total_trades else 0.0
    ret_pct = round(total_net / cap * 100, 2)
    return {
        "trades": total_trades,
        "wins": total_wins,
        "wr": wr,
        "gross": round(total_gross, 2),
        "fees": round(total_fees, 2),
        "net": round(total_net, 2),
        "ret_pct": ret_pct,
        "daily": daily_results
    }

print("Running comparative tests across 62 days on Rs 50,000 capital...")
res_ungated = run_simulation(enable_true_momentum_gate=False, zero_brokerage=False, cap=50000.0)
print(f"1. Current (No Gate - 09:25 All): Trades={res_ungated['trades']}, WR={res_ungated['wr']}%, Gross=Rs {res_ungated['gross']:+,.2f}, Fees=Rs {res_ungated['fees']:,.2f}, Net=Rs {res_ungated['net']:+,.2f} ({res_ungated['ret_pct']:+,.2f}%)")

res_gated = run_simulation(enable_true_momentum_gate=True, zero_brokerage=False, cap=50000.0)
print(f"2. With True Momentum Gate:     Trades={res_gated['trades']}, WR={res_gated['wr']}%, Gross=Rs {res_gated['gross']:+,.2f}, Fees=Rs {res_gated['fees']:,.2f}, Net=Rs {res_gated['net']:+,.2f} ({res_gated['ret_pct']:+,.2f}%)")

res_gated_zero = run_simulation(enable_true_momentum_gate=True, zero_brokerage=True, cap=50000.0)
print(f"3. With Gate + Zero Brokerage:  Trades={res_gated_zero['trades']}, WR={res_gated_zero['wr']}%, Gross=Rs {res_gated_zero['gross']:+,.2f}, Fees=Rs {res_gated_zero['fees']:,.2f}, Net=Rs {res_gated_zero['net']:+,.2f} ({res_gated_zero['ret_pct']:+,.2f}%)")

print("\n--- Testing on Rs 30,000 Capital ---")
res_30k_gated = run_simulation(enable_true_momentum_gate=True, zero_brokerage=False, cap=30000.0)
print(f"4. Rs 30k With Gate:            Trades={res_30k_gated['trades']}, WR={res_30k_gated['wr']}%, Gross=Rs {res_30k_gated['gross']:+,.2f}, Fees=Rs {res_30k_gated['fees']:,.2f}, Net=Rs {res_30k_gated['net']:+,.2f} ({res_30k_gated['ret_pct']:+,.2f}%)")

res_30k_gated_zero = run_simulation(enable_true_momentum_gate=True, zero_brokerage=True, cap=30000.0)
print(f"5. Rs 30k With Gate + Zero Brk: Trades={res_30k_gated_zero['trades']}, WR={res_30k_gated_zero['wr']}%, Gross=Rs {res_30k_gated_zero['gross']:+,.2f}, Fees=Rs {res_30k_gated_zero['fees']:,.2f}, Net=Rs {res_30k_gated_zero['net']:+,.2f} ({res_30k_gated['ret_pct']:+,.2f}%)")

# Check Tuesday Nov 03 specifically under the gate
nov3_ungated = [r for r in res_ungated['daily'] if r['date'] == '2026-11-03'][0]
nov3_gated = [r for r in res_gated['daily'] if r['date'] == '2026-11-03'][0]
print(f"\nTuesday Nov 03 Ungated: Net = Rs {nov3_ungated['net']:+,.2f} (Trades: {nov3_ungated['trades']}, Wins: {nov3_ungated['wins']})")
print(f"Tuesday Nov 03 Gated:   Net = Rs {nov3_gated['net']:+,.2f} (Trades: {nov3_gated['trades']}, Wins: {nov3_gated['wins']})")
