import sys, pickle, os
from datetime import datetime

sys.stdout.reconfigure(encoding='utf-8')
SIM_DIR = 'sim'
sys.path.insert(0, SIM_DIR)
import config as C
import datafeed
from simulator import Simulator

u = datafeed.load_universe()
with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
    payload = pickle.load(f)
raw = payload['raw']
idx = payload['idx']

dt = datetime(2026, 11, 2)

print('=== TESTING TODAY (2026-11-02) WITH CAPITAL Rs. 50,000 vs Rs. 500,000 ===\n')

for cap in [50000.0, 500000.0]:
    print(f'>>> INITIAL CAPITAL: Rs. {cap:,.2f}')
    for strat in ['v6_profit_max', 'v6_fixes_1_to_5', 'v6_baseline']:
        s = Simulator('replay', dt, strategy=strat, starting_capital=cap)
        s.load(raw, u, idx)
        times = s.session_bar_times()
        for t in times:
            s.step(t)
        
        snap = s.snapshot()
        kpi = snap['kpi']
        strat_name = 'Latest Strategy' if strat == 'v6_profit_max' else ('Fixes #1-5' if strat == 'v6_fixes_1_to_5' else "Dad's Original V6")
        print(f"  [{strat_name:18s}] Trades: {kpi['trades']} | Wins: {kpi['wins']} | WR: {kpi['win_rate']}% | Gross: Rs. {kpi['gross']:>8.2f} | Fees: Rs. {kpi['fees']:>7.2f} | Net Realized: Rs. {kpi['realized']:>8.2f}")
        for t in snap['closed']:
            c_cost = t.get('costs', 0)
            print(f"      {t['symbol']:10s} {t['side']:4s} qty={t['qty']:3d} entry={t['entry']:7.2f} exit={t['exit']:7.2f} gross=Rs. {t['gross']:>7.2f} fees=Rs. {c_cost:>6.2f} net=Rs. {t['net']:>7.2f} ({t['reason']})")
    print()
