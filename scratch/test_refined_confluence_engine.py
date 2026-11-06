"""
scratch/test_refined_confluence_engine.py
Refine the Structural Confluence Engine with:
1. Minimum viable distance threshold (>= 0.75% or 1.0R) to prevent micro-targets
2. Dynamic VWAP + 2.0 Sigma exhaustion harvesting
3. 3-bar Swing Low / Swing High trailing stop loss
4. Benchmark across all 61 days on 139 stocks
"""
import sys, os, pickle, math
from datetime import datetime, time as dtime
import numpy as np
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, 'sim')
import config as C
import datafeed
from simulator import Simulator, trade_costs, breakeven_exit_price, BAR

# Load data and master simulator
u = datafeed.load_universe()
with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
    payload = pickle.load(f)
raw = payload['raw']
idx = payload['idx']

sym0 = list(raw.keys())[0]
dates = sorted(list(set(raw[sym0].index.date)))

master = Simulator('replay', datetime.combine(dates[0], datetime.min.time()), starting_capital=500000.0)
master.load(raw, u, idx)

# Precompute daily pivots
pivot_cache = {}
for s, sd in master.syms.items():
    d_df = sd.df.groupby(sd.df.index.date).agg({
        'open': 'first', 'high': 'max', 'low': 'min', 'close': 'last'
    })
    pivot_cache[s] = {}
    d_list = [d for d in dates if d in d_df.index]
    for i in range(1, len(d_list)):
        prev_d = d_list[i-1]
        curr_d = d_list[i]
        p_row = d_df.loc[prev_d]
        pdh, pdl, pdc = float(p_row['high']), float(p_row['low']), float(p_row['close'])
        rng = pdh - pdl
        P = (pdh + pdl + pdc) / 3.0
        R1 = 2.0 * P - pdl
        S1 = 2.0 * P - pdh
        R2 = P + rng
        S2 = P - rng
        H4 = pdc + 0.55 * rng
        H3 = pdc + 0.275 * rng
        L3 = pdc - 0.275 * rng
        L4 = pdc - 0.55 * rng
        pivot_cache[s][curr_d] = {
            'pdh': pdh, 'pdl': pdl, 'pdc': pdc,
            'P': P, 'R1': R1, 'S1': S1, 'R2': R2, 'S2': S2,
            'H4': H4, 'H3': H3, 'L3': L3, 'L4': L4
        }

# Precompute intraday VWAP Standard Deviation
vwap_std_cache = {}
for s, sd in master.syms.items():
    n = len(sd.df)
    dates_arr = sd.df.index.normalize()
    tp = (sd.h + sd.l + sd.c) / 3.0
    vwap_std = np.zeros(n, dtype=np.float64)
    j = 0
    while j < n:
        d = dates_arr[j]
        k = j
        while k < n and dates_arr[k] == d:
            k += 1
        sub_tp = tp[j:k]
        sub_v = sd.v[j:k]
        cum_v = np.cumsum(sub_v)
        cum_pv = np.cumsum(sub_tp * sub_v)
        cum_pv2 = np.cumsum(sub_tp * sub_tp * sub_v)
        valid = cum_v > 0
        v_vwap = np.zeros(k - j, dtype=np.float64)
        v_vwap[valid] = cum_pv[valid] / cum_v[valid]
        v_var = np.maximum(0.0, (cum_pv2[valid] / cum_v[valid]) - (v_vwap[valid] ** 2))
        v_std = np.zeros(k - j, dtype=np.float64)
        v_std[valid] = np.sqrt(v_var)
        vwap_std[j:k] = v_std
        j = k
    vwap_std_cache[s] = vwap_std

class RefinedStructuralSimulator(Simulator):
    def __init__(self, session_date, min_target_dist=0.0075, starting_capital=500000.0):
        super().__init__('replay', session_date, strategy='v6_profit_max', starting_capital=starting_capital)
        self.min_target_dist = min_target_dist

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

                # STRUCTURAL TARGET SELECTION:
                # Nearest logical resistance offering at least min_target_dist (0.75% or ~0.9R)
                tgt_name = "R2/H4"
                if pivots is not None:
                    if o["side"] == "BUY":
                        candidates = [
                            ("H3", pivots['H3']),
                            ("R1", pivots['R1']),
                            ("PDH", pivots['pdh']),
                            ("H4", pivots['H4']),
                            ("R2", pivots['R2'])
                        ]
                        # Filter for candidates offering at least min_target_dist
                        valid = [(name, px) for name, px in sorted(candidates, key=lambda x: x[1]) if (px - filled) / filled >= self.min_target_dist]
                        if valid:
                            tgt_name, tgt_px = valid[0][0], round(valid[0][1], 2)
                        else:
                            tgt_name, tgt_px = "R2_EXP", round(pivots['R2'], 2)
                    else:
                        candidates = [
                            ("L3", pivots['L3']),
                            ("S1", pivots['S1']),
                            ("PDL", pivots['pdl']),
                            ("L4", pivots['L4']),
                            ("S2", pivots['S2'])
                        ]
                        valid = [(name, px) for name, px in sorted(candidates, key=lambda x: -x[1]) if (filled - px) / filled >= self.min_target_dist]
                        if valid:
                            tgt_name, tgt_px = valid[0][0], round(valid[0][1], 2)
                        else:
                            tgt_name, tgt_px = "S2_EXP", round(pivots['S2'], 2)
                else:
                    tgt_px = round(filled * 1.015 if o["side"] == "BUY" else filled * 0.985, 2)
                    tgt_name = "DEFAULT_1.5%"

                tgt_pct = round(abs(tgt_px / filled - 1.0) * 100, 2)

                pos = dict(o, entry=round(filled, 2), entry_time=str(t + BAR), bars_held=0,
                           sl=sl_px, initial_sl=sl_px, tgt=tgt_px, tgt_pct=tgt_pct, tgt_name=tgt_name,
                           ltp=cl, mfe=0.0, mae=0.0, sl_state="INITIAL_0.85PCT")
                self.open[sym] = pos
                self.traded_symbols_today.add(sym)
                if (pos["side"] == "BUY" and lo <= pos["sl"]) or (pos["side"] == "SELL" and hi >= pos["sl"]):
                    self._close(sym, pos["sl"], t, "STOP_LOSS (same bar as fill)", market=True)
                continue

            o["bars_waited"] += 1
            if o["bars_waited"] >= C.ENTRY_LIMIT_WAIT_BARS:
                del self.pending[sym]

    def _manage_open_profit_max(self, t, decide: dtime):
        mins_left = max(1, ((15 * 60 + 5) - (decide.hour * 60 + decide.minute)))
        for sym in list(self.open):
            p = self.open[sym]
            sd, i = self._bar(sym, t)
            if sd is None:
                if decide >= C.FORCE_LIQUIDATION:
                    self._close(sym, p["ltp"], t, "3_PM_EXIT (Session Liquidation)", market=True)
                continue
            hi, lo, cl, op = float(sd.h[i]), float(sd.l[i]), float(sd.c[i]), float(sd.o[i])
            p["ltp"], p["bars_held"] = cl, p["bars_held"] + 1
            sgn = 1 if p["side"] == "BUY" else -1
            bar_mfe = sgn * ((hi if sgn > 0 else lo) / p["entry"] - 1) * 100
            bar_mae = sgn * ((lo if sgn > 0 else hi) / p["entry"] - 1) * 100

            if decide >= C.FORCE_LIQUIDATION:
                p["mfe"] = max(p["mfe"], bar_mfe); p["mae"] = min(p["mae"], bar_mae)
                self._close(sym, cl, t, "3_PM_EXIT (Session Liquidation)", market=True); continue

            # Check SL & Limit Target hits
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

            # -----------------------------------------------------------------
            # 1. Statistical Exhaustion Band Check (VWAP + 2.0 Sigma)
            # -----------------------------------------------------------------
            vwap_val = float(sd.vwap[i])
            v_std = float(vwap_std_cache[sym][i])
            upper_2s = vwap_val + 2.0 * v_std
            lower_2s = vwap_val - 2.0 * v_std
            
            # If price reaches exhaustion band and shows reversal/wick, harvest!
            if p["side"] == "BUY" and hi >= upper_2s and p["mfe"] >= 0.70:
                if cl < op or (hi - cl) >= (cl - lo):
                    self._close(sym, cl, t, "EXHAUSTION (VWAP +2.0 Sigma Rejection)", market=True)
                    continue
            elif p["side"] == "SELL" and lo <= lower_2s and p["mfe"] >= 0.70:
                if cl > op or (cl - lo) >= (hi - cl):
                    self._close(sym, cl, t, "EXHAUSTION (VWAP -2.0 Sigma Rejection)", market=True)
                    continue

            # -----------------------------------------------------------------
            # 2. Structural Trailing Stop (Swing Low / Swing High)
            # -----------------------------------------------------------------
            if p["bars_held"] >= 3 and p["mfe"] >= 0.40:
                # Step A: Structural Breakeven at MFE >= 0.40%
                be_px = breakeven_exit_price(p["side"], p["qty"], p["entry"], market_exit=True, min_net_rupees=160.0)
                if (sgn > 0 and be_px > p["sl"]) or (sgn < 0 and be_px < p["sl"]):
                    p["sl"] = be_px
                    p["sl_state"] = "STRUCTURAL_BREAKEVEN"

                # Step B: 3-Bar Swing Low Trailing
                start_k = max(0, i - 3)
                if p["side"] == "BUY":
                    swing_low = float(np.min(sd.l[start_k:i]))
                    trail_px = round(swing_low * 0.9995, 2)
                    if trail_px > p["sl"]:
                        p["sl"] = trail_px
                        p["sl_state"] = f"SWING_LOW ({trail_px})"
                else:
                    swing_high = float(np.max(sd.h[start_k:i]))
                    trail_px = round(swing_high * 1.0005, 2)
                    if trail_px < p["sl"]:
                        p["sl"] = trail_px
                        p["sl_state"] = f"SWING_HIGH ({trail_px})"

            # 3. Universal Structural Reversal Exit
            c10, _ = sd.tenmin_state(i)
            e30 = float(sd.ema[30][i])
            if p["bars_held"] >= 3:
                if p["side"] == "BUY" and c10 == "BEARISH" and cl < vwap_val and cl < e30:
                    self._close(sym, cl, t, "STRUCTURAL_REVERSAL (10m+VWAP+EMA30)", market=True); continue
                if p["side"] == "SELL" and c10 == "BULLISH" and cl > vwap_val and cl > e30:
                    self._close(sym, cl, t, "STRUCTURAL_REVERSAL (10m+VWAP+EMA30)", market=True); continue


# Test across multiple distance thresholds: 0.60%, 0.70%, 0.75%, 0.85%
print("Testing Refined Structural Confluence Engine across various minimum target distances...")
for dist_thr in [0.0060, 0.0070, 0.0075, 0.0085]:
    tot_trades, tot_wins = 0, 0
    tot_gross, tot_costs, tot_net = 0.0, 0.0, 0.0
    all_closed = []
    for d in dates:
        dt = datetime.combine(d, datetime.min.time())
        sim = RefinedStructuralSimulator(dt, min_target_dist=dist_thr, starting_capital=500000.0)
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
    print(f"Threshold {dist_thr*100:.2f}%: Trades={tot_trades} | WR={wr}% | Gross=Rs. {tot_gross:+10,.2f} | Costs=Rs. {tot_costs:8,.2f} | Net=Rs. {tot_net:+10,.2f} | Avg R={avg_r:+.2f}R")
