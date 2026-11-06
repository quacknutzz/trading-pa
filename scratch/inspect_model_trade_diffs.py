"""
scratch/inspect_model_trade_diffs.py
Detailed trade-by-trade comparison between Baseline and Model 2 (Camarilla) / Model 1 (Floor Pivots).
"""
import sys, os, pickle
import pandas as pd
import numpy as np

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, 'sim')
sys.path.insert(0, 'scratch')
import datafeed
from test_structural_exit_models import master, dates, StructuralSimulator

# Run Baseline
sim_base_trades = []
for d in dates:
    dt = pd.Timestamp(d)
    sim = StructuralSimulator(dt, model_type='baseline', starting_capital=500000.0)
    sim.syms, sim.indices = master.syms, master.indices
    sim.nifty_pos, sim.nifty_e10, sim.nifty_e21, sim.nifty_e30 = master.nifty_pos, master.nifty_e10, master.nifty_e21, master.nifty_e30
    for t in sim.session_bar_times():
        sim.step(t)
    sim_base_trades.extend(sim.closed)

# Run Camarilla (Model 2)
sim_cam_trades = []
for d in dates:
    dt = pd.Timestamp(d)
    sim = StructuralSimulator(dt, model_type='camarilla', starting_capital=500000.0)
    sim.syms, sim.indices = master.syms, master.indices
    sim.nifty_pos, sim.nifty_e10, sim.nifty_e21, sim.nifty_e30 = master.nifty_pos, master.nifty_e10, master.nifty_e21, master.nifty_e30
    for t in sim.session_bar_times():
        sim.step(t)
    sim_cam_trades.extend(sim.closed)

print(f"Baseline trades: {len(sim_base_trades)} | Camarilla trades: {len(sim_cam_trades)}")

# Match trades by symbol and date
base_map = {(t['symbol'], t['entry_time'][:10]): t for t in sim_base_trades}
cam_map = {(t['symbol'], t['entry_time'][:10]): t for t in sim_cam_trades}

print("\n--- MAJOR WINNERS UNLOCKED BY STRUCTURAL TARGETS ---")
diffs = []
for k in base_map:
    if k in cam_map:
        b = base_map[k]
        c = cam_map[k]
        diff_net = c['net'] - b['net']
        diffs.append({
            'symbol': k[0], 'date': k[1], 'side': b['side'],
            'base_net': b['net'], 'cam_net': c['net'], 'diff_net': diff_net,
            'base_exit': b['reason'], 'cam_exit': c['reason'],
            'base_ret': b['ret_pct'], 'cam_ret': c['ret_pct']
        })

df_diffs = pd.DataFrame(diffs).sort_values(by='diff_net', ascending=False)
print(f"Top 10 trades where Structural Model outperformed Baseline:")
for idx, r in df_diffs.head(10).iterrows():
    print(f"{r['date']} {r['symbol']} ({r['side']}): Baseline Net: Rs. {r['base_net']:+8.2f} ({r['base_ret']:+.2f}%) -> Camarilla Net: Rs. {r['cam_net']:+8.2f} ({r['cam_ret']:+.2f}%) | Delta: Rs. {r['diff_net']:+8.2f}")
    print(f"   Baseline Exit: {r['base_exit']}")
    print(f"   Camarilla Exit: {r['cam_exit']}\n")

print("\n--- Trades where Baseline outperformed Camarilla ---")
for idx, r in df_diffs.tail(5).iterrows():
    print(f"{r['date']} {r['symbol']} ({r['side']}): Baseline Net: Rs. {r['base_net']:+8.2f} ({r['base_ret']:+.2f}%) -> Camarilla Net: Rs. {r['cam_net']:+8.2f} ({r['cam_ret']:+.2f}%) | Delta: Rs. {r['diff_net']:+8.2f}")
    print(f"   Baseline Exit: {r['base_exit']}")
    print(f"   Camarilla Exit: {r['cam_exit']}\n")
