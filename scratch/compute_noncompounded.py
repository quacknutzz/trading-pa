"""
scratch/compute_noncompounded.py -- Verifies Capital-Agnostic Non-Compounded metrics
(where every single trade across all 5 years is normalized to a constant 1x slot size,
zero compounding effect) vs Compounded metrics across all 8 permutations.
"""
import os
import pickle
import pandas as pd

ROOT_DIR = r"c:\Users\rishi\Downloads\Trading Pa"
with open(os.path.join(ROOT_DIR, "scratch", "ablation_matrix_results.pkl"), "rb") as f:
    res = pickle.load(f)

print(f"{'Code':4s} | {'Permutation':30s} | {'Trades':6s} | {'WR%':5s} | {'GrossPct':8s} | {'NetPct':8s} | {'Const-Slot SumNetPct':20s} | {'NetPF':5s} | {'Friction%':9s} | {'Compounded Net':15s} | {'MaxDD':6s}")
print("-" * 145)
for r in res:
    # Constant-slot total return = Trades * Mean NetPct per trade (100% capital agnostic!)
    const_slot_sum_net_pct = r["trades"] * r["mean_net_pct"]
    # On 3 slots with 2x leverage, each slot is (2/3) of account equity, so 1% on 1 slot = 0.6667% of account equity
    uncompounded_acct_roi = const_slot_sum_net_pct * (2.0 / 3.0)
    uncompounded_rupees_40k = 40000.0 * (uncompounded_acct_roi / 100.0)
    print(
        f"{r['code']:4s} | {r['label']:30s} | {r['trades']:6d} | {r['win_rate']:4.1f}% | "
        f"{r['mean_gross_pct']:+7.3f}% | {r['mean_net_pct']:+7.3f}% | "
        f"{uncompounded_acct_roi:+7.1f}% (Rs.{uncompounded_rupees_40k:+7,.0f}) | "
        f"{r['net_pf']:5.2f} | {r['friction_share']:8.1f}% | "
        f"Rs.{r['net_pnl']:+9,.0f} ({r['net_roi']:+5.1f}%) | {r['max_dd_pct']:4.1f}%"
    )
