import sys, pickle, math
from datetime import datetime, time as dtime
import pandas as pd
import numpy as np

sys.path.insert(0, 'sim')
import datafeed
import config as C
from simulator import Simulator, trade_costs, BAR

u = datafeed.load_universe()
with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
    payload = pickle.load(f)
raw = payload['raw']
idx = payload['idx']
sym0 = list(raw.keys())[0]
dates = sorted(list(set(raw[sym0].index.date)))

master = Simulator('replay', datetime.combine(dates[0], datetime.min.time()), strategy='v6_baseline', starting_capital=500000.0)
master.load(raw, u, idx)

# We want to test 5 Exit Rules on the EXACT SAME baseline entries:
# Rule 0: Baseline (ADX < 18, 5m reverse, 10m contra, 1.0% SL, 1.2% Target, 3 PM exit)
# Rule 1: Remove ADX < 18 (Keep 5m reverse, 10m contra, SL, Target, 3 PM)
# Rule 2: Remove ADX < 18 AND Remove 5m reverse (Only 10m contra, SL, Target, 3 PM)
# Rule 3: Fix #6 Original Proposal (ALL 3 must break: 10m contra + Close < VWAP + Close < EMA30, 3-bar hold min)
# Rule 4: Opus Proposal (2-of-3 Vote: exit if any 2 of [10m contra, Close < VWAP, Close < EMA30] break, 3-bar hold min)
# Rule 5: Pure Structural: 10m contra is hard exit; VWAP + EMA30 replaces 5m reverse & ADX; 3-bar hold min

class ExitTestSimulator(Simulator):
    def __init__(self, session_date, exit_mode='baseline', starting_capital=500000.0):
        super().__init__('replay', session_date, strategy='v6_baseline', starting_capital=starting_capital)
        self.exit_mode = exit_mode

    def _manage_open_baseline(self, t, decide: dtime):
        for sym in list(self.open):
            p = self.open[sym]
            sd, i = self._bar(sym, t)
            if sd is None:
                if decide >= C.FORCE_LIQUIDATION:
                    self._close(sym, p["ltp"], t, "3_PM_EXIT", market=True)
                continue
            hi, lo, cl, op = float(sd.h[i]), float(sd.l[i]), float(sd.c[i]), float(sd.o[i])
            p["ltp"], p["bars_held"] = cl, p["bars_held"] + 1
            sgn = 1 if p["side"] == "BUY" else -1
            p["mfe"] = max(p["mfe"], sgn * ((hi if sgn > 0 else lo) / p["entry"] - 1) * 100)
            p["mae"] = min(p["mae"], sgn * ((lo if sgn > 0 else hi) / p["entry"] - 1) * 100)

            # 1. 3 PM session liquidation
            if decide >= C.FORCE_LIQUIDATION:
                self._close(sym, cl, t, "3_PM_EXIT", market=True); continue

            # 2. Hard Stop Loss (1.0%) & Target (1.2%)
            if p["side"] == "BUY":
                if lo <= p["sl"]:
                    self._close(sym, min(p["sl"], op), t, "STOP_LOSS (1.0%)", market=True); continue
                if hi >= p["tgt"]:
                    self._close(sym, max(p["tgt"], op), t, "TARGET (1.2%)", market=False); continue
            else:
                if hi >= p["sl"]:
                    self._close(sym, max(p["sl"], op), t, "STOP_LOSS (1.0%)", market=True); continue
                if lo <= p["tgt"]:
                    self._close(sym, min(p["tgt"], op), t, "TARGET (1.2%)", market=False); continue

            # Indicator Exits depending on exit_mode
            c10, _ = sd.tenmin_state(i)
            c5 = sd.cross5m(i)
            adx = float(sd.adx[i])
            vwap = float(sd.vwap[i])
            e30 = float(sd.ema[30][i])

            is_10m_contra = (p["side"] == "BUY" and c10 == "BEARISH") or (p["side"] == "SELL" and c10 == "BULLISH")
            is_5m_reverse = (p["side"] == "BUY" and c5 == "BEARISH") or (p["side"] == "SELL" and c5 == "BULLISH")
            lost_vwap = (p["side"] == "BUY" and cl < vwap) or (p["side"] == "SELL" and cl > vwap)
            lost_e30 = (p["side"] == "BUY" and cl < e30) or (p["side"] == "SELL" and cl > e30)

            if self.exit_mode == 'baseline':
                if not math.isnan(adx) and adx < C.ADX_EXIT_BELOW:
                    self._close(sym, cl, t, "ADX_5M_BELOW_18", market=True); continue
                if is_10m_contra:
                    self._close(sym, cl, t, "EMA_CROSS_10M_CONTRA", market=True); continue
                if is_5m_reverse:
                    self._close(sym, cl, t, "EMA_CROSS_5M", market=True); continue

            elif self.exit_mode == 'no_adx':
                # Remove ADX exit only
                if is_10m_contra:
                    self._close(sym, cl, t, "EMA_CROSS_10M_CONTRA", market=True); continue
                if is_5m_reverse:
                    self._close(sym, cl, t, "EMA_CROSS_5M", market=True); continue

            elif self.exit_mode == 'no_adx_no_5m':
                # Only 10m contra
                if is_10m_contra:
                    self._close(sym, cl, t, "EMA_CROSS_10M_CONTRA", market=True); continue

            elif self.exit_mode == 'fix6_all_three':
                # ALL 3 break + 3-bar hold min
                if p["bars_held"] >= 3:
                    if is_10m_contra and lost_vwap and lost_e30:
                        self._close(sym, cl, t, "ALL_THREE_BREAK", market=True); continue

            elif self.exit_mode == 'opus_two_of_three':
                # 2-of-3 Vote + 3-bar hold min
                if p["bars_held"] >= 3:
                    votes = sum([is_10m_contra, lost_vwap, lost_e30])
                    if votes >= 2:
                        self._close(sym, cl, t, f"VOTE_2_OF_3 ({votes}/3)", market=True); continue

            elif self.exit_mode == 'structural_tiered':
                # Tiered: 10m contra is hard exit after bar 2;
                # Pullback exit only if BOTH VWAP and EMA30 lost after bar 3
                if p["bars_held"] >= 2 and is_10m_contra:
                    self._close(sym, cl, t, "10M_CONTRA_HARD", market=True); continue
                if p["bars_held"] >= 3 and (lost_vwap and lost_e30):
                    self._close(sym, cl, t, "LOST_VWAP_AND_EMA30", market=True); continue

modes = [
    ('baseline', 'Rule 0: Current Baseline (ADX<18 + 5m reverse + 10m contra)'),
    ('no_adx', 'Rule 1: Drop ADX Exit (Keep 5m reverse + 10m contra)'),
    ('no_adx_no_5m', 'Rule 2: Drop ADX & 5m Reverse (Only 10m contra)'),
    ('fix6_all_three', 'Rule 3: Fix #6 Original (Require ALL 3 to break)'),
    ('opus_two_of_three', 'Rule 4: Opus 2-of-3 Vote (Any 2 of 3 break)'),
    ('structural_tiered', 'Rule 5: Tiered (10m Contra Hard + [VWAP & EMA30] Vote)')
]

results = []
for m_key, m_name in modes:
    all_closed = []
    for d in dates:
        dt = datetime.combine(d, datetime.min.time())
        sim = ExitTestSimulator(dt, exit_mode=m_key, starting_capital=500000.0)
        sim.syms = master.syms
        sim.indices = master.indices
        sim.nifty_pos = master.nifty_pos
        sim.nifty_e10 = master.nifty_e10
        sim.nifty_e21 = master.nifty_e21
        sim.nifty_e30 = master.nifty_e30
        for t in sim.session_bar_times():
            sim.step(t)
        all_closed.extend(sim.closed)
    
    df = pd.DataFrame(all_closed)
    wins = (df['net'] > 0).sum()
    n = len(df)
    results.append({
        'Mode': m_name,
        'Trades': n,
        'Wins': wins,
        'Win Rate (%)': round(wins / n * 100, 1) if n else 0,
        'Gross (Rs)': round(df['gross'].sum(), 2),
        'Costs (Rs)': round(df['costs'].sum(), 2),
        'Net (Rs)': round(df['net'].sum(), 2),
        'Avg Net/Trd (Rs)': round(df['net'].mean(), 2) if n else 0
    })

res_df = pd.DataFrame(results)
print("\n================== EMPIRICAL EXIT MODEL COMPARISON (60 TRADING DAYS) ==================")
print(res_df.to_string(index=False))
