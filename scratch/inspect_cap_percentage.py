import sys, json, os, pickle
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, 'sim')
import datafeed
from simulator import Simulator

# Let's inspect the 12-week report to see what the numbers were on 500k
with open('sim/journal/backtest_12week_report.json', 'r', encoding='utf-8') as f:
    rpt = json.load(f)

print('=== 12-WEEK TOTAL ON Rs. 500,000 CAPITAL ===')
tot_500 = rpt['totals']['v6_profit_max']
net_500 = tot_500['net']
cap_500 = 500000.0
pct_500 = (net_500 / cap_500) * 100
fees_500 = tot_500['fees']
fee_pct_500 = (fees_500 / cap_500) * 100

print(f"Trades: {tot_500['trades']} | WR: {tot_500['wr']}% | Gross: Rs. {tot_500['gross']:+,.2f} | Fees: Rs. {fees_500:,.2f} ({fee_pct_500:.2f}% of cap)")
print(f"Net Realized: Rs. {net_500:+,.2f} | Percentage Return: {pct_500:+.2f}%\n")
