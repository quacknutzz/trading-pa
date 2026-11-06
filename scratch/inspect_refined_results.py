"""
scratch/inspect_refined_results.py
Detailed inspection of the 0.75% threshold Refined Structural Confluence Engine.
"""
import sys, os, pickle
import pandas as pd
import numpy as np

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, 'sim')
sys.path.insert(0, 'scratch')
from test_refined_confluence_engine import master, dates, RefinedStructuralSimulator

all_trades = []
for d in dates:
    dt = pd.Timestamp(d)
    sim = RefinedStructuralSimulator(dt, min_target_dist=0.0075, starting_capital=500000.0)
    sim.syms, sim.indices = master.syms, master.indices
    sim.nifty_pos, sim.nifty_e10, sim.nifty_e21, sim.nifty_e30 = master.nifty_pos, master.nifty_e10, master.nifty_e21, master.nifty_e30
    for t in sim.session_bar_times():
        sim.step(t)
    all_trades.extend(sim.closed)

print(f"Total trades: {len(all_trades)}")
reasons = {}
for c in all_trades:
    r = c['reason'].split('(')[0].strip()
    reasons[r] = reasons.get(r, 0) + 1

print("\n--- Exit Reason Distribution ---")
for r, cnt in sorted(reasons.items(), key=lambda x: -x[1]):
    print(f"  {r:<25}: {cnt:>2} ({cnt/len(all_trades)*100:.1f}%)")

print("\n--- Structural Target Name Distribution (When Target was hit) ---")
targets = [c for c in all_trades if 'TARGET' in c['reason']]
tgt_names = {}
for c in targets:
    # reason format: TARGET (NAME +PCT%)
    name = c['reason'].split('(')[1].split()[0]
    tgt_names[name] = tgt_names.get(name, 0) + 1

for name, cnt in sorted(tgt_names.items(), key=lambda x: -x[1]):
    print(f"  {name:<15}: {cnt:>2} trades")

print("\n--- Sample of Closed Trades with Structural Exits ---")
for c in all_trades[:15]:
    print(f"{c['entry_time'][:10]} {c['symbol']:<10} {c['side']:<4} | Net: Rs. {c['net']:+8.2f} ({c['ret_pct']:+.2f}%) | MFE: {c['mfe']:.2f}% | Exit: {c['reason']}")
