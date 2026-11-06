"""
scratch/compare_dad_live_vs_sim.py
Direct comparison between Dad's actual live FYERS trades and simulation.
"""
import sys, json
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

with open('scratch/grand_comprehensive_backtest.json', 'r') as f:
    data = json.load(f)

baseline_daily = data['daily_records']['baseline_1x']
target_dates = ['2026-10-26', '2026-10-27', '2026-10-28', '2026-10-29', '2026-10-30']
sim_week = [r for r in baseline_daily if r['date'] in target_dates]

print("=== SIMULATION BASELINE (Oct 26 - Oct 30) ===")
tot_gross, tot_fees, tot_net = 0, 0, 0
for r in sim_week:
    tot_gross += r['gross']
    tot_fees += r['fees']
    tot_net += r['net']
    print(f"{r['date']}: Trades={r['trades']}, Wins={r['wins']}, Gross=Rs {r['gross']:+,.2f}, Fees=Rs {r['fees']:,.2f}, Net=Rs {r['net']:+,.2f}")

print(f"\nTotal Sim Baseline: Gross=Rs {tot_gross:+,.2f}, Fees=Rs {tot_fees:,.2f}, Net=Rs {tot_net:+,.2f}")
