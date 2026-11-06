"""
scratch/deep_research_v6_profit_max.py
Thorough ablation & optimization research on v6_profit_max across all 62 trading days
and multiple capital levels (Rs 50k, Rs 100k, Rs 500k).
Tests every single fix variant discussed with the user and flagged by Mr. Opus.
"""
import sys, os, pickle, math
from datetime import datetime, time as dtime
import numpy as np
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, '.')
sys.path.insert(0, 'sim')
import config as C
from simulator import Simulator, trade_costs, breakeven_exit_price, BAR

with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
    payload = pickle.load(f)

raw = payload['raw']
idx = payload['idx']
universe = payload['universe']
sym0 = list(raw.keys())[0]
all_dates = sorted(list(set(raw[sym0].index.date)))

print(f"Loaded {len(universe)} symbols across {len(all_dates)} trading days ({all_dates[0]} to {all_dates[-1]}).")

# Pre-load master simulator once to reuse precomputed indicator arrays
dt0 = datetime.combine(all_dates[0], datetime.min.time())
master = Simulator('replay', dt0, strategy='v6_profit_max', starting_capital=500000.0)
master.load(raw, universe, idx)


class ResearchProfitMaxSimulator(Simulator):
    def __init__(self, session_date, starting_capital=500000.0, opts=None):
        super().__init__('replay', session_date, strategy='v6_profit_max', starting_capital=starting_capital)
        self.opts = opts or {}

    def step(self, t: pd.Timestamp):
        with self.lock:
            if t in self.processed:
                return
            self.processed.add(t)
            decide = (t + BAR).time()
            self.clock = t + BAR

            cutoff = self.opts.get("entry_cutoff", dtime(13, 15))
            entry_start = self.opts.get("entry_start", dtime(9, 25))

            self._manage_open_custom(t, decide)
            self._manage_pending_custom(t)
            if entry_start <= decide < cutoff:
                self._scan_custom(t, decide)

            self._mark_equity(t)

    def _manage_pending_custom(self, t):
        early_contra_cancel = self.opts.get("early_contra_cancel", True)
        sl_pct = self.opts.get("sl_pct", 0.0085)
        target_mode = self.opts.get("target_mode", "structural")  # "structural", "atr", "hybrid"

        for sym in list(self.pending):
            o = self.pending[sym]
            sd, i = self._bar(sym, t)
            if sd is None or t <= o["signal_bar"]:
                continue
            lo, hi, op, cl = float(sd.l[i]), float(sd.h[i]), float(sd.o[i]), float(sd.c[i])

            c10, _ = sd.tenmin_state(i)
            c5 = sd.cross5m(i)
            c10_diff = float(sd.c10_diff[i]) if not math.isnan(sd.c10_diff[i]) else 0.0

            # Fix #1 Enhancement: Cancel immediately if 10m/5m cross flips contra even BEFORE fill!
            if early_contra_cancel:
                if (o["side"] == "BUY" and (c10 == "BEARISH" or c5 == "BEARISH" or c10_diff <= 0)) or \
                   (o["side"] == "SELL" and (c10 == "BULLISH" or c5 == "BULLISH" or c10_diff >= 0)):
                    del self.pending[sym]
                    continue

            filled = None
            if o["side"] == "BUY" and lo <= o["limit"]:
                filled = min(o["limit"], op)
            elif o["side"] == "SELL" and hi >= o["limit"]:
                filled = max(o["limit"], op)

            if filled is not None:
                if (o["side"] == "BUY" and (c10 == "BEARISH" or c5 == "BEARISH" or c10_diff <= 0)) or \
                   (o["side"] == "SELL" and (c10 == "BULLISH" or c5 == "BULLISH" or c10_diff >= 0)):
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
                curr_d = t.normalize()
                pivots = sd.pivots.get(curr_d)

                sl_px = round(filled * (1 - sl_pct) if o["side"] == "BUY" else filled * (1 + sl_pct), 2)

                tgt_name = "R2/H4"
                tgt_px = round(filled * 1.012 if o["side"] == "BUY" else filled * 0.988, 2)

                if target_mode == "atr":
                    atr_val = float(sd.atr[i]) if not math.isnan(sd.atr[i]) else (filled * 0.0025)
                    atr_pct = atr_val / filled if filled > 0 else 0.0025
                    t_pct = min(max(atr_pct * 2.8, 0.0065), 0.0110)
                    tgt_name = "ATR_LIM"
                    tgt_px = round(filled * (1 + t_pct) if o["side"] == "BUY" else filled * (1 - t_pct), 2)
                elif pivots is not None:
                    min_dist = self.opts.get("min_target_dist", 0.0075)
                    if o["side"] == "BUY":
                        candidates = [("H3", pivots['H3']), ("R1", pivots['R1']), ("PDH", pivots['pdh']), ("H4", pivots['H4']), ("R2", pivots['R2'])]
                        valid = [(name, px) for name, px in sorted(candidates, key=lambda x: x[1]) if (px - filled) / filled >= min_dist]
                        if valid:
                            tgt_name, tgt_px = valid[0][0], round(valid[0][1], 2)
                        else:
                            tgt_name, tgt_px = "EXT_R2", round(max(pivots['R2'], filled * 1.012), 2)
                    else:
                        candidates = [("L3", pivots['L3']), ("S1", pivots['S1']), ("PDL", pivots['pdl']), ("L4", pivots['L4']), ("S2", pivots['S2'])]
                        valid = [(name, px) for name, px in sorted(candidates, key=lambda x: -x[1]) if (filled - px) / filled >= min_dist]
                        if valid:
                            tgt_name, tgt_px = valid[0][0], round(valid[0][1], 2)
                        else:
                            tgt_name, tgt_px = "EXT_S2", round(min(pivots['S2'], filled * 0.988), 2)

                tgt_pct = abs(tgt_px - filled) / filled
                pos = dict(o, entry=round(filled, 2), entry_time=str(t + BAR), bars_held=0,
                           sl=sl_px, initial_sl=sl_px, tgt=tgt_px, tgt_pct=round(tgt_pct * 100, 2),
                           tgt_name=tgt_name, ltp=cl, mfe=0.0, mae=0.0, sl_state="INITIAL_0.85PCT",
                           highest_px=filled, lowest_px=filled)
                self.open[sym] = pos
                self.traded_symbols_today.add(sym)
                if (pos["side"] == "BUY" and lo <= pos["sl"]) or (pos["side"] == "SELL" and hi >= pos["sl"]):
                    self._close(sym, pos["sl"], t, "STOP_LOSS (same bar as fill)", market=True)
                continue

            o["bars_waited"] += 1
            if o["bars_waited"] >= C.ENTRY_LIMIT_WAIT_BARS:
                del self.pending[sym]

    def _manage_open_custom(self, t, decide: dtime):
        n_reg = self.nifty_regime(t)
        is_flat_market = (n_reg == "NEUTRAL")
        mins_left = max(1, ((15 * 60 + 5) - (decide.hour * 60 + decide.minute)))
        bars_left = mins_left // 5

        be_trigger_mfe = self.opts.get("be_trigger_mfe", 0.35)
        exit_mode = self.opts.get("exit_mode", "current")
        # exit_mode options:
        #   "current"          : c10 == "BEARISH" and cl < vwap and cl < e30
        #   "state_all3"       : c10_diff < 0 and cl < vwap and cl < e30
        #   "opus_2of3_state"  : sum([c10_diff < 0, cl < vwap, cl < e30]) >= 2
        #   "opus_2of3_cross"  : (c10 == "BEARISH") or (cl < vwap and cl < e30)
        #   "hybrid_5m_vwap"   : (c10 == "BEARISH" and (cl < vwap or cl < e30)) or (c5 == "BEARISH" and cl < vwap and cl < e30)

        for sym in list(self.open):
            p = self.open[sym]
            sd, i = self._bar(sym, t)
            if sd is None:
                if decide >= C.FORCE_LIQUIDATION:
                    self._close(sym, p["ltp"], t, "3_PM_EXIT (Session Liquidation)", market=True)
                continue
            hi, lo, cl, op = float(sd.h[i]), float(sd.l[i]), float(sd.c[i]), float(sd.o[i])
            p["ltp"], p["bars_held"] = cl, p["bars_held"] + 1
            p["highest_px"] = max(p.get("highest_px", p["entry"]), hi)
            p["lowest_px"] = min(p.get("lowest_px", p["entry"]), lo)
            sgn = 1 if p["side"] == "BUY" else -1
            bar_mfe = sgn * ((hi if sgn > 0 else lo) / p["entry"] - 1) * 100
            bar_mae = sgn * ((lo if sgn > 0 else hi) / p["entry"] - 1) * 100

            if decide >= C.FORCE_LIQUIDATION:
                p["mfe"] = max(p["mfe"], bar_mfe); p["mae"] = min(p["mae"], bar_mae)
                self._close(sym, cl, t, "3_PM_EXIT (Session Liquidation)", market=True); continue

            # Late-Session Zero-Slippage Limit Target Harvest (after 14:15, compress limit target to +0.55%)
            if bars_left <= 10 and p["mfe"] >= 0.35:
                harvest_tgt = round(p["entry"] * (1.0 + sgn * 0.0055), 2)
                if (sgn > 0 and harvest_tgt < p["tgt"] and harvest_tgt > op) or \
                   (sgn < 0 and harvest_tgt > p["tgt"] and harvest_tgt < op):
                    p["tgt"] = harvest_tgt

            if p["side"] == "BUY":
                if lo <= p["sl"]:
                    p["mfe"] = max(p["mfe"], bar_mfe); p["mae"] = min(p["mae"], bar_mae)
                    lbl = f"PROFIT_LOCK ({p['sl_state']})" if not p["sl_state"].startswith("INITIAL") else "STOP_LOSS (0.85%)"
                    self._close(sym, min(p["sl"], op), t, lbl, market=True); continue
                if hi >= p["tgt"]:
                    p["mfe"] = max(p["mfe"], bar_mfe); p["mae"] = min(p["mae"], bar_mae)
                    self._close(sym, max(p["tgt"], op), t, f"TARGET ({p.get('tgt_name', 'R1/H4')} +{p.get('tgt_pct', 0.8)}%)", market=False); continue
            else:
                if hi >= p["sl"]:
                    p["mfe"] = max(p["mfe"], bar_mfe); p["mae"] = min(p["mae"], bar_mae)
                    lbl = f"PROFIT_LOCK ({p['sl_state']})" if not p["sl_state"].startswith("INITIAL") else "STOP_LOSS (0.85%)"
                    self._close(sym, max(p["sl"], op), t, lbl, market=True); continue
                if lo <= p["tgt"]:
                    p["mfe"] = max(p["mfe"], bar_mfe); p["mae"] = min(p["mae"], bar_mae)
                    self._close(sym, min(p["tgt"], op), t, f"TARGET ({p.get('tgt_name', 'S1/L4')} +{p.get('tgt_pct', 0.8)}%)", market=False); continue

            p["mfe"] = max(p["mfe"], bar_mfe)
            p["mae"] = min(p["mae"], bar_mae)

            # 1. Statistical Exhaustion Harvest (VWAP +/- 2.0 Sigma Rejection)
            vwap_val = float(sd.vwap[i])
            vwap_std_val = float(sd.vwap_std[i]) if hasattr(sd, "vwap_std") else 0.0
            vwap_top = vwap_val + 2.0 * vwap_std_val
            vwap_bot = vwap_val - 2.0 * vwap_std_val

            if p["bars_held"] >= 2 and p["mfe"] >= 0.70:
                if p["side"] == "BUY" and hi >= vwap_top and cl < (hi + lo) / 2.0:
                    self._close(sym, cl, t, "EXHAUSTION (VWAP +2.0 Sigma Rejection)", market=True); continue
                elif p["side"] == "SELL" and lo <= vwap_bot and cl > (hi + lo) / 2.0:
                    self._close(sym, cl, t, "EXHAUSTION (VWAP -2.0 Sigma Rejection)", market=True); continue

            # 2. Structural Trailing Stop (2-Bar Swing Low/High) + Adaptive Breakeven
            if p["bars_held"] >= 3 and p["mfe"] >= be_trigger_mfe:
                be_px = breakeven_exit_price(p["side"], p["qty"], p["entry"], market_exit=True, min_net_rupees=50.0 if is_flat_market else 160.0)
                if (sgn > 0 and be_px > p["sl"]) or (sgn < 0 and be_px < p["sl"]):
                    p["sl"] = be_px
                    p["sl_state"] = "TAX_SHIELD_BE"

                k = 2
                if p["side"] == "BUY":
                    sw_px = round(float(np.min(sd.l[max(0, i-k):i])) * 0.9995, 2)
                    if sw_px > p["sl"]:
                        p["sl"] = sw_px
                        p["sl_state"] = "SWING_LOW_TRAIL"
                else:
                    sw_px = round(float(np.max(sd.h[max(0, i-k):i])) * 1.0005, 2)
                    if sw_px < p["sl"]:
                        p["sl"] = sw_px
                        p["sl_state"] = "SWING_HIGH_TRAIL"

            # 3. Structural Reversal Exit
            c10, _ = sd.tenmin_state(i)
            c5 = sd.cross5m(i)
            c10_diff = float(sd.c10_diff[i]) if not math.isnan(sd.c10_diff[i]) else 0.0
            e30 = float(sd.ema[30][i])

            if p["bars_held"] >= 3:
                if exit_mode == "current":
                    if p["side"] == "BUY" and c10 == "BEARISH" and cl < vwap_val and cl < e30:
                        self._close(sym, cl, t, "STRUCTURAL_REVERSAL (10m+VWAP+EMA30)", market=True); continue
                    if p["side"] == "SELL" and c10 == "BULLISH" and cl > vwap_val and cl > e30:
                        self._close(sym, cl, t, "STRUCTURAL_REVERSAL (10m+VWAP+EMA30)", market=True); continue
                elif exit_mode == "state_all3":
                    if p["side"] == "BUY" and c10_diff < 0 and cl < vwap_val and cl < e30:
                        self._close(sym, cl, t, "STRUCTURAL_REVERSAL (10mState+VWAP+EMA30)", market=True); continue
                    if p["side"] == "SELL" and c10_diff > 0 and cl > vwap_val and cl > e30:
                        self._close(sym, cl, t, "STRUCTURAL_REVERSAL (10mState+VWAP+EMA30)", market=True); continue
                elif exit_mode == "opus_2of3_state":
                    if p["side"] == "BUY" and sum([c10_diff < 0, cl < vwap_val, cl < e30]) >= 2:
                        self._close(sym, cl, t, "STRUCTURAL_REVERSAL (2of3 Vote)", market=True); continue
                    if p["side"] == "SELL" and sum([c10_diff > 0, cl > vwap_val, cl > e30]) >= 2:
                        self._close(sym, cl, t, "STRUCTURAL_REVERSAL (2of3 Vote)", market=True); continue
                elif exit_mode == "hybrid_5m_vwap":
                    if p["side"] == "BUY" and ((c10_diff < 0 and cl < vwap_val and cl < e30) or (c5 == "BEARISH" and cl < vwap_val and cl < e30)):
                        self._close(sym, cl, t, "STRUCTURAL_REVERSAL (Hybrid)", market=True); continue
                    if p["side"] == "SELL" and ((c10_diff > 0 and cl > vwap_val and cl > e30) or (c5 == "BULLISH" and cl > vwap_val and cl > e30)):
                        self._close(sym, cl, t, "STRUCTURAL_REVERSAL (Hybrid)", market=True); continue

    def _scan_custom(self, t, decide: dtime):
        adx_mode = self.opts.get("adx_mode", "bar1") # "bar1", "smoothed3", "sma3", "relaxed_dip"
        sizing_mode = self.opts.get("sizing_mode", "conviction") # "conviction", "equal"
        min_atr = self.opts.get("min_atr", 0.14)

        ndf = self.indices.get("NIFTY 50")
        is_nifty_choppy = True
        n_rs_ratio = 1.0
        n_15m_ret = 0.0

        if ndf is not None and t in self.nifty_pos:
            ni = self.nifty_pos[t]
            n_cl = float(ndf["close"].iloc[ni])
            n_e10 = float(self.nifty_e10[ni])
            n_e30 = float(self.nifty_e30[ni])
            n_rs_ratio = (n_cl / n_e10) if n_e10 > 0 else 1.0
            if ni >= 3:
                n_prev3 = float(ndf["close"].iloc[ni - 3])
                n_15m_ret = (n_cl - n_prev3) / n_prev3 * 100.0

            d_norm = t.normalize()
            d_bars = ndf.loc[ndf.index.normalize() == d_norm]
            d_bars_up_to_t = d_bars.loc[d_bars.index <= t]
            if len(d_bars_up_to_t) >= 5:
                n_open = float(d_bars_up_to_t["open"].iloc[0])
                n_high = float(d_bars_up_to_t["high"].max())
                n_low = float(d_bars_up_to_t["low"].min())
                n_ret = (n_cl - n_open) / n_open * 100.0
                n_rng = (n_high - n_low) / n_open * 100.0
                is_nifty_choppy = (abs(n_ret) < 0.35 and n_rng < 0.55) or (abs(n_e10 - n_e30) / n_e30 < 0.0005)

        n_reg = self.nifty_regime(t)
        cands = []
        for sym, sd in self.syms.items():
            if t not in sd.pos:
                continue
            i = sd.pos[t]
            if i < 3:
                continue
            c10, rsi10 = sd.tenmin_state(i)
            ltp = float(sd.c[i])
            adx = float(sd.adx[i])
            prev_adx = float(sd.adx[i - 1])
            vwap_val = float(sd.vwap[i])
            pdi, mdi = float(sd.plus_di[i]), float(sd.minus_di[i])
            e10, e30 = float(sd.ema[10][i]), float(sd.ema[30][i])
            atr_val = float(sd.atr[i]) if not math.isnan(sd.atr[i]) else 0.0
            atr_pct = (atr_val / ltp * 100.0) if ltp > 0 else 0.0

            cur_vol = float(sd.v[i])
            rolling_rvol = float(sd.rvol[i]) if not math.isnan(sd.rvol[i]) else 0.0
            tod_mean = getattr(sd, "tod_mean", {}).get((t.normalize(), t.time()), 0.0)
            eff_rvol = max(rolling_rvol, cur_vol / tod_mean if tod_mean > 0 else rolling_rvol)

            stock_rs_ratio = (ltp / e10) if e10 > 0 else 1.0
            rs_formula_pct = ((stock_rs_ratio / n_rs_ratio) - 1.0) * 100.0

            if sym in self.open or sym in self.pending:
                continue
            elif sym in self.traded_symbols_today:
                continue
            elif math.isnan(rsi10) or math.isnan(adx) or math.isnan(prev_adx) or adx < 20.0 or atr_pct < min_atr:
                continue

            # ADX Rising check based on adx_mode
            if adx_mode == "bar1":
                if adx < prev_adx:
                    continue
            elif adx_mode == "smoothed3":
                prev3_adx = float(sd.adx[i - 3])
                if math.isnan(prev3_adx) or adx < prev3_adx:
                    continue
            elif adx_mode == "sma3":
                sma3 = (float(sd.adx[i - 1]) + float(sd.adx[i - 2]) + float(sd.adx[i - 3])) / 3.0
                if math.isnan(sma3) or adx < sma3:
                    continue
            elif adx_mode == "bar1_or_3":
                prev2_adx = float(sd.adx[i - 2])
                if adx < prev_adx and (math.isnan(prev2_adx) or adx < prev2_adx):
                    continue

            # Check Setup A: Classical 10m Crossover Pullback (Engine 1)
            setup_type = ""
            side = ""
            limit = 0.0
            period = 0

            if c10 in ("BULLISH", "BEARISH"):
                side = "BUY" if c10 == "BULLISH" else "SELL"
                if side == "BUY" and (50.0 <= rsi10 <= 65.0) and ltp > vwap_val and pdi > mdi and e10 >= e30:
                    limit, period = sd.pullback_limit(i, side, ltp)
                    if limit > 0:
                        setup_type = "ENGINE1_CROSS"
                elif side == "SELL" and (35.0 <= rsi10 <= 50.0) and ltp < vwap_val and mdi > pdi and e10 <= e30:
                    limit, period = sd.pullback_limit(i, side, ltp)
                    if limit > 0:
                        setup_type = "ENGINE1_CROSS"

            # Check Setup B: Intraday VWAP Continuation & Momentum Expansion (Engine 2)
            if not setup_type:
                if ltp > vwap_val and e10 > e30 and pdi > mdi and (48.0 <= rsi10 <= 66.0):
                    lo_prev = float(sd.l[i - 1])
                    vwap_prev = float(sd.vwap[i - 1])
                    if lo_prev <= vwap_prev * 1.002 or lo_prev <= float(sd.ema[30][i - 1]):
                        if float(sd.c[i]) > float(sd.o[i]) and float(sd.c[i]) > float(sd.c[i - 1]):
                            side = "BUY"
                            limit = float(sd.c[i])
                            period = 10
                            setup_type = "ENGINE2_VWAP_CONT"
                elif ltp < vwap_val and e10 < e30 and mdi > pdi and (34.0 <= rsi10 <= 52.0):
                    hi_prev = float(sd.h[i - 1])
                    vwap_prev = float(sd.vwap[i - 1])
                    if hi_prev >= vwap_prev * 0.998 or hi_prev >= float(sd.ema[30][i - 1]):
                        if float(sd.c[i]) < float(sd.o[i]) and float(sd.c[i]) < float(sd.c[i - 1]):
                            side = "SELL"
                            limit = float(sd.c[i])
                            period = 10
                            setup_type = "ENGINE2_VWAP_CONT"

            if not setup_type or limit <= 0:
                continue

            # True Momentum Gate: Dynamic Morning Noise Filter (09:15 - 09:45 AM)
            if decide < dtime(9, 45):
                if eff_rvol < 1.25:
                    continue
                o_px = float(sd.o[i]); h_px = float(sd.h[i]); l_px = float(sd.l[i]); c_px = float(sd.c[i])
                rng_px = h_px - l_px if h_px > l_px else 0.01
                adv_wick = (h_px - max(o_px, c_px)) / rng_px if side == "BUY" else (min(o_px, c_px) - l_px) / rng_px
                if adv_wick > 0.35:
                    continue
                if ndf is not None and len(d_bars_up_to_t) >= 2:
                    n_open_m = float(d_bars_up_to_t["open"].iloc[0])
                    n_ret_m = (n_cl - n_open_m) / n_open_m * 100.0
                    if side == "BUY" and (n_ret_m < -0.10 or n_e10 < n_e30):
                        continue
                    elif side == "SELL" and (n_ret_m > 0.10 or n_e10 > n_e30):
                        continue

            must_decouple = is_nifty_choppy or (side == "BUY" and n_reg != "BULL") or (side == "SELL" and n_reg != "BEAR")
            if must_decouple:
                if eff_rvol < 1.30:
                    continue
                if side == "BUY" and rs_formula_pct < 0.12:
                    continue
                if side == "SELL" and rs_formula_pct > -0.12:
                    continue
                if side == "BUY" and n_15m_ret < -0.25 and eff_rvol < 1.8:
                    continue
            else:
                if eff_rvol < 0.8:
                    continue

            conv_score = (eff_rvol * 1.5) + ((adx - 20.0) * 0.1) + (atr_pct * 4.0) + (abs(rs_formula_pct) * 2.0)
            if setup_type == "ENGINE2_VWAP_CONT":
                conv_score += 0.5

            cands.append(dict(symbol=sym, side=side, ltp=ltp, rsi=rsi10, adx=adx, rvol=eff_rvol,
                              atr_pct=atr_pct, conv_score=conv_score, setup=setup_type,
                              limit=limit, ema_period=period, sector=sd.sector))

        # Ranking priority: when sizing_mode == "equal", rank by RVOL + ADX + RS (without 4x ATR volatility bias!)
        if sizing_mode == "equal":
            cands.sort(key=lambda c: (-(c["rvol"] * 1.5 + (c["adx"] - 20.0) * 0.1 + abs(c.get("rs", 0)) * 2.0), -c["rvol"], c["symbol"]))
        else:
            cands.sort(key=lambda c: (-c["conv_score"], -c["rvol"], c["symbol"]))

        free = C.MAX_CONCURRENT_POSITIONS - len(self.open) - len(self.pending)
        total_equity = self.starting_capital + self.cash_realized
        committed_cash = sum(p["entry"] * p["qty"] for p in self.open.values()) + \
                         sum(o["limit"] * o["qty"] for o in self.pending.values())
        avail_cash = max(0.0, total_equity * C.LEVERAGE - committed_cash)
        n_take = min(len(cands), free)

        min_cash_buffer = max(1000.0, total_equity * 0.05)
        for k, c in enumerate(cands):
            if k >= free or avail_cash < min_cash_buffer:
                continue
            if sizing_mode == "equal":
                # Strict Equal Slot Sizing (Fix #9 shelved per Opus & user agreement)
                slot_budget = (total_equity * C.LEVERAGE) / C.MAX_CONCURRENT_POSITIONS
                equal_rem = avail_cash / max(1, (free - k))
                budget = min(avail_cash, slot_budget, equal_rem)
            else:
                equal_share = avail_cash / max(1, (n_take - k))
                target_alloc = total_equity * C.LEVERAGE * 0.48
                budget = min(avail_cash, max(equal_share, target_alloc) if n_take == 1 else equal_share)

            qty = int(budget // c["limit"])
            if qty <= 0:
                continue
            avail_cash -= qty * c["limit"]
            self.seq += 1
            self.pending[c["symbol"]] = dict(
                id=self.seq, symbol=c["symbol"], sector=c["sector"], side=c["side"], qty=qty,
                limit=c["limit"], ema_period=c["ema_period"], signal_bar=t, signal_time=str(t + BAR),
                signal_ltp=c["ltp"], rsi=c["rsi"], adx=c["adx"], rvol=c["rvol"], bars_waited=0)


def eval_variant(name, opts, cap=500000.0):
    all_closed = []
    daily_nets = []
    for d in all_dates:
        dt = datetime.combine(d, datetime.min.time())
        sim = ResearchProfitMaxSimulator(dt, starting_capital=cap, opts=opts)
        sim.syms = master.syms
        sim.indices = master.indices
        sim.nifty_pos = master.nifty_pos
        sim.nifty_e10 = master.nifty_e10
        sim.nifty_e21 = master.nifty_e21
        sim.nifty_e30 = master.nifty_e30
        for t in sim.session_bar_times():
            sim.step(t)
        all_closed.extend(sim.closed)
        daily_nets.append(sim.cash_realized)

    df = pd.DataFrame(all_closed)
    n = len(df)
    wins = (df['net'] > 0).sum() if n else 0
    gross = df['gross'].sum() if n else 0.0
    costs = df['costs'].sum() if n else 0.0
    net = df['net'].sum() if n else 0.0
    sl_cnt = df['reason'].str.contains('STOP_LOSS').sum() if n else 0
    tgt_cnt = df['reason'].str.contains('TARGET').sum() if n else 0
    max_loss = df['net'].min() if n else 0.0
    cum = np.cumsum(daily_nets)
    peak = np.maximum.accumulate(np.maximum(0.0, cum))
    max_dd = float(np.max(peak - cum)) if len(cum) else 0.0
    return {
        "Variant": name,
        "Trades": n,
        "Wins": wins,
        "WR%": round(wins / n * 100, 1) if n else 0.0,
        "Gross": round(gross, 2),
        "Costs": round(costs, 2),
        "Net": round(net, 2),
        "Targets": int(tgt_cnt),
        "SLs": int(sl_cnt),
        "WorstTrd": round(max_loss, 2),
        "MaxDD": round(max_dd, 2),
    }


variants = [
    ("0. Current v6_profit_max", {
        "early_contra_cancel": False, "sizing_mode": "conviction", "exit_mode": "current",
        "adx_mode": "bar1", "entry_cutoff": dtime(13, 15), "be_trigger_mfe": 0.35
    }),
    ("1. Equal Slot Sizing (Fix #9 Shelved)", {
        "early_contra_cancel": False, "sizing_mode": "equal", "exit_mode": "current",
        "adx_mode": "bar1", "entry_cutoff": dtime(13, 15), "be_trigger_mfe": 0.35
    }),
    ("2. Equal Sizing + Pre-Fill Early Contra Cancel (Fix #1)", {
        "early_contra_cancel": True, "sizing_mode": "equal", "exit_mode": "current",
        "adx_mode": "bar1", "entry_cutoff": dtime(13, 15), "be_trigger_mfe": 0.35
    }),
    ("3. + State-Aware All-3 Exit (c10_diff<0 + VWAP + EMA30)", {
        "early_contra_cancel": True, "sizing_mode": "equal", "exit_mode": "state_all3",
        "adx_mode": "bar1", "entry_cutoff": dtime(13, 15), "be_trigger_mfe": 0.35
    }),
    ("4. + Opus 2-of-3 Vote Exit", {
        "early_contra_cancel": True, "sizing_mode": "equal", "exit_mode": "opus_2of3_state",
        "adx_mode": "bar1", "entry_cutoff": dtime(13, 15), "be_trigger_mfe": 0.35
    }),
    ("5. + Hybrid (State-All3 OR 5mRev+VWAP+EMA30)", {
        "early_contra_cancel": True, "sizing_mode": "equal", "exit_mode": "hybrid_5m_vwap",
        "adx_mode": "bar1", "entry_cutoff": dtime(13, 15), "be_trigger_mfe": 0.35
    }),
    ("6. + Cutoff 13:45 (instead of 13:15)", {
        "early_contra_cancel": True, "sizing_mode": "equal", "exit_mode": "state_all3",
        "adx_mode": "bar1", "entry_cutoff": dtime(13, 45), "be_trigger_mfe": 0.35
    }),
    ("7. + Cutoff 14:00 (Opus suggestion)", {
        "early_contra_cancel": True, "sizing_mode": "equal", "exit_mode": "state_all3",
        "adx_mode": "bar1", "entry_cutoff": dtime(14, 0), "be_trigger_mfe": 0.35
    }),
    ("8. + Smoothed ADX (bar1_or_3)", {
        "early_contra_cancel": True, "sizing_mode": "equal", "exit_mode": "state_all3",
        "adx_mode": "bar1_or_3", "entry_cutoff": dtime(13, 45), "be_trigger_mfe": 0.35
    }),
    ("9. + Smoothed ADX (3-bar trend)", {
        "early_contra_cancel": True, "sizing_mode": "equal", "exit_mode": "state_all3",
        "adx_mode": "smoothed3", "entry_cutoff": dtime(13, 45), "be_trigger_mfe": 0.35
    }),
    ("10. + Breakeven Trigger MFE 0.45%", {
        "early_contra_cancel": True, "sizing_mode": "equal", "exit_mode": "state_all3",
        "adx_mode": "bar1", "entry_cutoff": dtime(13, 45), "be_trigger_mfe": 0.45
    }),
]

for cap in [50000.0, 500000.0]:
    print(f"\n==================== CAPITAL: Rs {cap:,.0f} (62 DAYS) ====================")
    rows = [eval_variant(name, opts, cap=cap) for name, opts in variants]
    print(pd.DataFrame(rows).to_string(index=False))
