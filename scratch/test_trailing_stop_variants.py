"""
scratch/test_trailing_stop_variants.py
Compare different structural trailing stop algorithms:
1. 2-Bar Swing Low
2. 3-Bar Swing Low (Current Champion)
3. 4-Bar Swing Low
4. Fractal Swing Low (pivot low with 1 bar before and after)
5. Chandelier ATR Structural Trail (Highest High - 2.0 * ATR)
"""
import sys, os, pickle, math
from datetime import datetime, time as dtime
import numpy as np
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, 'sim')
sys.path.insert(0, 'scratch')
from test_refined_confluence_engine import master, dates, pivot_cache, vwap_std_cache, BAR, trade_costs, breakeven_exit_price, C
from simulator import Simulator

class TrailingVariantSimulator(Simulator):
    def __init__(self, session_date, trail_type='3bar', starting_capital=500000.0):
        super().__init__('replay', session_date, strategy='v6_profit_max', starting_capital=starting_capital)
        self.trail_type = trail_type

    def _manage_pending_profit_max(self, t):
        for sym in list(self.pending):
            o = self.pending[sym]
            sd, i = self._bar(sym, t)
            if sd is None or t <= o["signal_bar"]:
                continue
            lo, hi, op, cl = float(sd.l[i]), float(sd.h[i]), float(sd.o[i]), float(sd.c[i])
            filled = None
            if o["side"] == "BUY" and lo <= o["limit"]:
                filled = min(o["limit"], op)
            elif o["side"] == "SELL" and hi >= o["limit"]:
                filled = max(o["limit"], op)

            if filled is not None:
                c10, _ = sd.tenmin_state(i)
                c5 = sd.cross5m(i)
                if (o["side"] == "BUY" and (c10 == "BEARISH" or c5 == "BEARISH" or float(sd.c10_diff[i]) <= 0)) or \
                   (o["side"] == "SELL" and (c10 == "BULLISH" or c5 == "BULLISH" or float(sd.c10_diff[i]) >= 0)):
                    del self.pending[sym]
                    continue

                vwap_val = float(sd.vwap[i])
                e30 = float(sd.ema[30][i])
                if o["side"] == "BUY":
                    if cl < vwap_val or cl < op or cl < filled * 0.9985 or cl < e30:
                        del self.pending[sym]
                        continue
                else:
                    if cl > vwap_val or cl > op or cl > filled * 1.0015 or cl > e30:
                        del self.pending[sym]
                        continue

                del self.pending[sym]
                curr_d = t.date()
                pivots = pivot_cache.get(sym, {}).get(curr_d)

                sl_pct = 0.0085
                sl_px = round(filled * (1 - sl_pct) if o["side"] == "BUY" else filled * (1 + sl_pct), 2)

                # Target selection (confluence >= 0.75%)
                tgt_name = "R2/H4"
                if pivots is not None:
                    if o["side"] == "BUY":
                        candidates = [("H3", pivots['H3']), ("R1", pivots['R1']), ("PDH", pivots['pdh']), ("H4", pivots['H4']), ("R2", pivots['R2'])]
                        valid = [(name, px) for name, px in sorted(candidates, key=lambda x: x[1]) if (px - filled) / filled >= 0.0075]
                        tgt_name, tgt_px = (valid[0][0], round(valid[0][1], 2)) if valid else ("R2_EXP", round(pivots['R2'], 2))
                    else:
                        candidates = [("L3", pivots['L3']), ("S1", pivots['S1']), ("PDL", pivots['pdl']), ("L4", pivots['L4']), ("S2", pivots['S2'])]
                        valid = [(name, px) for name, px in sorted(candidates, key=lambda x: -x[1]) if (filled - px) / filled >= 0.0075]
                        tgt_name, tgt_px = (valid[0][0], round(valid[0][1], 2)) if valid else ("S2_EXP", round(pivots['S2'], 2))
                else:
                    tgt_px = round(filled * 1.015 if o["side"] == "BUY" else filled * 0.985, 2)
                    tgt_name = "DEFAULT_1.5%"

                tgt_pct = round(abs(tgt_px / filled - 1.0) * 100, 2)

                pos = dict(o, entry=round(filled, 2), entry_time=str(t + BAR), bars_held=0,
                           sl=sl_px, initial_sl=sl_px, tgt=tgt_px, tgt_pct=tgt_pct, tgt_name=tgt_name,
                           ltp=cl, mfe=0.0, mae=0.0, sl_state="INITIAL_0.85PCT", highest_px=filled, lowest_px=filled)
                self.open[sym] = pos
                self.traded_symbols_today.add(sym)
                if (pos["side"] == "BUY" and lo <= pos["sl"]) or (pos["side"] == "SELL" and hi >= pos["sl"]):
                    self._close(sym, pos["sl"], t, "STOP_LOSS (same bar as fill)", market=True)
                continue

            o["bars_waited"] += 1
            if o["bars_waited"] >= C.ENTRY_LIMIT_WAIT_BARS:
                del self.pending[sym]

    def _manage_open_profit_max(self, t, decide: dtime):
        for sym in list(self.open):
            p = self.open[sym]
            sd, i = self._bar(sym, t)
            if sd is None:
                if decide >= C.FORCE_LIQUIDATION:
                    self._close(sym, p["ltp"], t, "3_PM_EXIT (Session Liquidation)", market=True)
                continue
            hi, lo, cl, op = float(sd.h[i]), float(sd.l[i]), float(sd.c[i]), float(sd.o[i])
            p["ltp"], p["bars_held"] = cl, p["bars_held"] + 1
            p["highest_px"] = max(p["highest_px"], hi)
            p["lowest_px"] = min(p["lowest_px"], lo)
            
            sgn = 1 if p["side"] == "BUY" else -1
            bar_mfe = sgn * ((hi if sgn > 0 else lo) / p["entry"] - 1) * 100
            bar_mae = sgn * ((lo if sgn > 0 else hi) / p["entry"] - 1) * 100

            if decide >= C.FORCE_LIQUIDATION:
                p["mfe"] = max(p["mfe"], bar_mfe); p["mae"] = min(p["mae"], bar_mae)
                self._close(sym, cl, t, "3_PM_EXIT (Session Liquidation)", market=True); continue

            # Check SL & Target hits
            if p["side"] == "BUY":
                if lo <= p["sl"]:
                    p["mfe"] = max(p["mfe"], bar_mfe); p["mae"] = min(p["mae"], bar_mae)
                    lbl = f"PROFIT_LOCK ({p['sl_state']})" if not p["sl_state"].startswith("INITIAL") else "STOP_LOSS"
                    self._close(sym, min(p["sl"], op), t, lbl, market=True); continue
                if hi >= p["tgt"]:
                    p["mfe"] = max(p["mfe"], bar_mfe); p["mae"] = min(p["mae"], bar_mae)
                    self._close(sym, max(p["tgt"], op), t, f"TARGET ({p.get('tgt_name', 'TGT')} +{p.get('tgt_pct', 0.8)}%)", market=False); continue
            else:
                if hi >= p["sl"]:
                    p["mfe"] = max(p["mfe"], bar_mfe); p["mae"] = min(p["mae"], bar_mae)
                    lbl = f"PROFIT_LOCK ({p['sl_state']})" if not p["sl_state"].startswith("INITIAL") else "STOP_LOSS"
                    self._close(sym, max(p["sl"], op), t, lbl, market=True); continue
                if lo <= p["tgt"]:
                    p["mfe"] = max(p["mfe"], bar_mfe); p["mae"] = min(p["mae"], bar_mae)
                    self._close(sym, min(p["tgt"], op), t, f"TARGET ({p.get('tgt_name', 'TGT')} -{p.get('tgt_pct', 0.8)}%)", market=False); continue

            p["mfe"] = max(p["mfe"], bar_mfe)
            p["mae"] = min(p["mae"], bar_mae)

            # 1. Statistical Exhaustion Band Check (VWAP + 2.0 Sigma)
            vwap_val = float(sd.vwap[i])
            v_std = float(vwap_std_cache[sym][i])
            upper_2s = vwap_val + 2.0 * v_std
            lower_2s = vwap_val - 2.0 * v_std
            
            if p["side"] == "BUY" and hi >= upper_2s and p["mfe"] >= 0.70:
                if cl < op or (hi - cl) >= (cl - lo):
                    self._close(sym, cl, t, "EXHAUSTION (VWAP +2.0 Sigma Rejection)", market=True); continue
            elif p["side"] == "SELL" and lo <= lower_2s and p["mfe"] >= 0.70:
                if cl > op or (cl - lo) >= (hi - cl):
                    self._close(sym, cl, t, "EXHAUSTION (VWAP -2.0 Sigma Rejection)", market=True); continue

            # 2. Structural Trailing Stop
            if p["bars_held"] >= 3 and p["mfe"] >= 0.40:
                be_px = breakeven_exit_price(p["side"], p["qty"], p["entry"], market_exit=True, min_net_rupees=160.0)
                if (sgn > 0 and be_px > p["sl"]) or (sgn < 0 and be_px < p["sl"]):
                    p["sl"] = be_px

                if self.trail_type == '2bar':
                    k = 2
                    if p["side"] == "BUY":
                        sw_px = round(float(np.min(sd.l[max(0, i-k):i])) * 0.9995, 2)
                        if sw_px > p["sl"]: p["sl"] = sw_px
                    else:
                        sw_px = round(float(np.max(sd.h[max(0, i-k):i])) * 1.0005, 2)
                        if sw_px < p["sl"]: p["sl"] = sw_px

                elif self.trail_type == '3bar':
                    k = 3
                    if p["side"] == "BUY":
                        sw_px = round(float(np.min(sd.l[max(0, i-k):i])) * 0.9995, 2)
                        if sw_px > p["sl"]: p["sl"] = sw_px
                    else:
                        sw_px = round(float(np.max(sd.h[max(0, i-k):i])) * 1.0005, 2)
                        if sw_px < p["sl"]: p["sl"] = sw_px

                elif self.trail_type == '4bar':
                    k = 4
                    if p["side"] == "BUY":
                        sw_px = round(float(np.min(sd.l[max(0, i-k):i])) * 0.9995, 2)
                        if sw_px > p["sl"]: p["sl"] = sw_px
                    else:
                        sw_px = round(float(np.max(sd.h[max(0, i-k):i])) * 1.0005, 2)
                        if sw_px < p["sl"]: p["sl"] = sw_px

                elif self.trail_type == 'chandelier_atr':
                    atr_val = float(sd.atr[i]) if not math.isnan(sd.atr[i]) else (p["entry"] * 0.0025)
                    if p["side"] == "BUY":
                        chan_px = round(p["highest_px"] - 1.8 * atr_val, 2)
                        if chan_px > p["sl"]: p["sl"] = chan_px
                    else:
                        chan_px = round(p["lowest_px"] + 1.8 * atr_val, 2)
                        if chan_px < p["sl"]: p["sl"] = chan_px

            # 3. Universal Structural Reversal Exit
            c10, _ = sd.tenmin_state(i)
            e30 = float(sd.ema[30][i])
            if p["bars_held"] >= 3:
                if p["side"] == "BUY" and c10 == "BEARISH" and cl < vwap_val and cl < e30:
                    self._close(sym, cl, t, "STRUCTURAL_REVERSAL (10m+VWAP+EMA30)", market=True); continue
                if p["side"] == "SELL" and c10 == "BULLISH" and cl > vwap_val and cl > e30:
                    self._close(sym, cl, t, "STRUCTURAL_REVERSAL (10m+VWAP+EMA30)", market=True); continue

print("Testing Trailing Stop Variants on Refined Confluence Target...")
variants = [
    ('2bar', '2-Bar Swing Low/High'),
    ('3bar', '3-Bar Swing Low/High (Standard)'),
    ('4bar', '4-Bar Swing Low/High'),
    ('chandelier_atr', 'Chandelier Exit (Highest High - 1.8*ATR)')
]

for v_code, v_name in variants:
    tot_trades, tot_wins = 0, 0
    tot_gross, tot_costs, tot_net = 0.0, 0.0, 0.0
    all_closed = []
    for d in dates:
        dt = datetime.combine(d, datetime.min.time())
        sim = TrailingVariantSimulator(dt, trail_type=v_code, starting_capital=500000.0)
        sim.syms, sim.indices = master.syms, master.indices
        sim.nifty_pos, sim.nifty_e10, sim.nifty_e21, sim.nifty_e30 = master.nifty_pos, master.nifty_e10, master.nifty_e21, master.nifty_e30
        for t in sim.session_bar_times():
            sim.step(t)
        for c in sim.closed:
            all_closed.append(c)
            tot_trades += 1
            if c['net'] > 0:
                tot_wins += 1
            tot_gross += c['gross']
            tot_costs += c['costs']
            tot_net += c['net']
            
    wr = round(tot_wins / tot_trades * 100, 1) if tot_trades else 0.0
    r_multiples = [c['net'] / (c['entry'] * 0.0085 * c['qty']) for c in all_closed]
    avg_r = np.mean(r_multiples) if r_multiples else 0.0
    print(f"{v_name:<40}: Trades={tot_trades} | WR={wr}% | Gross=Rs. {tot_gross:+10,.2f} | Costs=Rs. {tot_costs:8,.2f} | Net=Rs. {tot_net:+10,.2f} | Avg R={avg_r:+.2f}R")
