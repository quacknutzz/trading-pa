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

class CooldownExitSimulator(Simulator):
    def __init__(self, session_date, exit_mode='baseline', starting_capital=500000.0):
        super().__init__('replay', session_date, strategy='v6_baseline', starting_capital=starting_capital)
        self.exit_mode = exit_mode
        self.traded_today = set()

    def _scan_baseline(self, t, decide):
        orig_syms = self.syms
        self.syms = {k: v for k, v in orig_syms.items() if k not in self.traded_today}
        try:
            super()._scan_baseline(t, decide)
            for sym in self.pending:
                self.traded_today.add(sym)
            for sym in self.open:
                self.traded_today.add(sym)
        finally:
            self.syms = orig_syms

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

            if decide >= C.FORCE_LIQUIDATION:
                self._close(sym, cl, t, "3_PM_EXIT", market=True); continue

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

            elif self.exit_mode == 'fix6_all_three':
                if p["bars_held"] >= 3:
                    if is_10m_contra and lost_vwap and lost_e30:
                        self._close(sym, cl, t, "ALL_THREE_BREAK", market=True); continue

            elif self.exit_mode == 'opus_two_of_three':
                if p["bars_held"] >= 3:
                    votes = sum([is_10m_contra, lost_vwap, lost_e30])
                    if votes >= 2:
                        self._close(sym, cl, t, f"VOTE_2_OF_3 ({votes}/3)", market=True); continue

modes = [
    ('baseline', 'Baseline (ADX<18 + 5m rev + 10m contra) + Cooldown'),
    ('fix6_all_three', 'Fix #6 ALL-3 Break + Cooldown'),
    ('opus_two_of_three', 'Opus 2-of-3 Vote + Cooldown')
]

results = []
for m_key, m_name in modes:
    all_closed = []
    for d in dates:
        dt = datetime.combine(d, datetime.min.time())
        sim = CooldownExitSimulator(dt, exit_mode=m_key, starting_capital=500000.0)
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
    sl_hits = df[df['reason'].str.contains('STOP_LOSS')]
    tgt_hits = df[df['reason'].str.contains('TARGET')]
    pm_hits = df[df['reason'].str.contains('3_PM')]
    results.append({
        'Mode': m_name,
        'Trades': n,
        'Wins': wins,
        'Win Rate (%)': round(wins / n * 100, 1) if n else 0,
        'Gross (Rs)': round(df['gross'].sum(), 2),
        'Costs (Rs)': round(df['costs'].sum(), 2),
        'Net (Rs)': round(df['net'].sum(), 2),
        'Targets': len(tgt_hits),
        'SLs': len(sl_hits),
        '3PM': len(pm_hits)
    })

res_df = pd.DataFrame(results)
print("\n================== WITH PER-STOCK COOLDOWN (FIX #4B) ==================")
print(res_df.to_string(index=False))
