import pickle
import pandas as pd

with open("scratch/ablation_matrix_results.pkl", "rb") as f:
    res = pickle.load(f)

# Let's inspect the trades of ABC (permutation index 7)
abc_trades = res[7]
print(f"ABC Net P&L: Rs. {abc_trades['net_pnl']:+,.2f}")
print(f"Total costs: Rs. {abc_trades['tot_costs']:,.2f}")
