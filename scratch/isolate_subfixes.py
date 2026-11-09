import os
import pickle

ROOT_DIR = r"c:\Users\rishi\Downloads\Trading Pa"
with open(os.path.join(ROOT_DIR, "scratch", "ablation_matrix_results.pkl"), "rb") as f:
    ablation_res = pickle.load(f)

print("\n--- YEAR-BY-YEAR BREAKDOWN OF THE 8 PERMUTATIONS ---")
for r in ablation_res:
    yr_str = " | ".join(f"{y}: Rs.{d['net']:+8,.0f} ({d['trades']:3d}t, {d['wr']:4.1f}%)" for y, d in sorted(r["yearly"].items()))
    print(f"[{r['code']}] {r['label']:30s} | Net: Rs.{r['net_pnl']:+9,.0f} ({r['net_roi']:+6.1f}%) | DD:{r['max_dd_pct']:4.1f}% => {yr_str}")
