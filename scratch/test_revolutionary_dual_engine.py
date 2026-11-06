"""
scratch/test_revolutionary_dual_engine.py
Institutional Dual-Engine Intraday Architecture:
Engine 1: Structural 10m Cross + Decoupled Leader Pullback
Engine 2: Intraday VWAP Continuation & Momentum Expansion (Captures established runners without waiting for fresh 10m cross)

Testing across:
- Capital sizes: Rs. 30,000 | Rs. 50,000 | Rs. 1,00,000
- Margin cases: Case A (1.0x Cash) vs Case B (2.0x MIS Intraday Margin)
- Across 61 Trading Days (139 NSE Stocks)
"""
import sys, os, pickle, math, time
from datetime import datetime, time as dtime
import numpy as np
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, 'sim')
import config as C
import datafeed
from simulator import Simulator, trade_costs, breakeven_exit_price, BAR, _cross, ema_fast

with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
    payload = pickle.load(f)
raw = payload['raw']
idx = payload['idx']
universe = payload['universe']

sym0 = list(raw.keys())[0]
dates = sorted(list(set(raw[sym0].index.date)))

print(f"Preloading master data across {len(dates)} dates...")
dt0 = datetime.combine(dates[0], datetime.min.time())
master = Simulator('replay', dt0, starting_capital=500000.0)
master.load(raw, universe, idx)

# Precompute pivots, vwap_std, and tod_mean
pivot_cache = {}
vwap_std_cache = {}
tod_mean_cache = {}

for s, df in raw.items():
    sym = s.replace('.NS', '')
    d_df = df.groupby(df.index.date).agg({'open':'first', 'high':'max', 'low':'min', 'close':'last'})
    pivot_cache[sym] = {}
    d_list = [d for d in dates if d in d_df.index]
    for i in range(1, len(d_list)):
        prev_d = d_list[i-1]
        curr_d = d_list[i]
        p_row = d_df.loc[prev_d]
        pdh, pdl, pdc = float(p_row['high']), float(p_row['low']), float(p_row['close'])
        rng = pdh - pdl
        P = (pdh + pdl + pdc) / 3.0
        R1 = 2.0 * P - pdl; S1 = 2.0 * P - pdh
        R2 = P + rng; S2 = P - rng
        H4 = pdc + 0.55 * rng; H3 = pdc + 0.275 * rng
        L3 = pdc - 0.275 * rng; L4 = pdc - 0.55 * rng
        pivot_cache[sym][curr_d] = {
            'pdh': pdh, 'pdl': pdl, 'pdc': pdc,
            'P': P, 'R1': R1, 'S1': S1, 'R2': R2, 'S2': S2,
            'H4': H4, 'H3': H3, 'L3': L3, 'L4': L4
        }
        
    df_tod = df[['volume']].copy()
    df_tod['tod'] = df.index.time
    df_tod['date_idx'] = df.index.date
    pv_vol = df_tod.pivot(index='date_idx', columns='tod', values='volume')
    pv_mean = pv_vol.shift(1).rolling(20, min_periods=5).mean()
    for d_val, row in pv_mean.iterrows():
        for t_val, m_vol in row.items():
            if not np.isnan(m_vol) and m_vol > 0:
                tod_mean_cache[(sym, d_val, t_val)] = float(m_vol)

    n = len(df)
    dates_arr = df.index.normalize()
    h_arr = df['high'].values.astype(np.float64)
    l_arr = df['low'].values.astype(np.float64)
    c_arr = df['close'].values.astype(np.float64)
    v_arr = df['volume'].values.astype(np.float64)
    tp = (h_arr + l_arr + c_arr) / 3.0
    vwap_std = np.zeros(n, dtype=np.float64)
    j = 0
    while j < n:
        d = dates_arr[j]
        k = j
        while k < n and dates_arr[k] == d:
            k += 1
        sub_tp = tp[j:k]
        sub_v = np.maximum(v_arr[j:k], 1.0)
        cum_v = np.cumsum(sub_v)
        cum_pv = np.cumsum(sub_tp * sub_v)
        cum_pv2 = np.cumsum(sub_tp * sub_tp * sub_v)
        valid = cum_v > 0
        v_vwap = np.zeros(k - j, dtype=np.float64)
        v_vwap[valid] = cum_pv[valid] / cum_v[valid]
        v_var = np.maximum(0.0, (cum_pv2[valid] / cum_v[valid]) - (v_vwap[valid] ** 2))
        vwap_std[j:k] = np.sqrt(v_var)
        j = k
    vwap_std_cache[sym] = vwap_std

print("Pivots, VWAP Std, and TOD Mean precomputed.")

class DualEngineSimulator(Simulator):
    def __init__(self, session_date, starting_capital=50000.0, leverage=1.0, enable_engine2=True):
        super().__init__('replay', session_date, strategy='v6_profit_max', starting_capital=starting_capital)
        self.leverage = leverage
        self.enable_engine2 = enable_engine2

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
                vwap_val = float(sd.vwap[i])
                e30 = float(sd.ema[30][i])
                
                # Validity check
                if o["side"] == "BUY":
                    if cl < vwap_val or cl < op or cl < filled * 0.9985 or cl < e30:
                        del self.pending[sym]; continue
                else:
                    if cl > vwap_val or cl > op or cl > filled * 1.0015 or cl > e30:
                        del self.pending[sym]; continue

                del self.pending[sym]
                curr_d = t.date()
                pivots = pivot_cache.get(sym, {}).get(curr_d)
                sl_pct = 0.0085
                sl_px = round(filled * (1 - sl_pct) if o["side"] == "BUY" else filled * (1 + sl_pct), 2)

                # Target selection: Nearest structural level >= 0.75%
                tgt_name = "R2/H4"
                tgt_px = round(filled * 1.012 if o["side"] == "BUY" else filled * 0.988, 2)
                if pivots is not None:
                    if o["side"] == "BUY":
                        candidates = [("H3", pivots['H3']), ("R1", pivots['R1']), ("PDH", pivots['pdh']), ("H4", pivots['H4']), ("R2", pivots['R2'])]
                        valid = [(name, px) for name, px in sorted(candidates, key=lambda x: x[1]) if (px - filled) / filled >= 0.0075]
                        if valid: tgt_name, tgt_px = valid[0][0], round(valid[0][1], 2)
                        else: tgt_name, tgt_px = "EXT_R2", round(max(pivots['R2'], filled * 1.012), 2)
                    else:
                        candidates = [("L3", pivots['L3']), ("S1", pivots['S1']), ("PDL", pivots['pdl']), ("L4", pivots['L4']), ("S2", pivots['S2'])]
                        valid = [(name, px) for name, px in sorted(candidates, key=lambda x: -x[1]) if (filled - px) / filled >= 0.0075]
                        if valid: tgt_name, tgt_px = valid[0][0], round(valid[0][1], 2)
                        else: tgt_name, tgt_px = "EXT_S2", round(min(pivots['S2'], filled * 0.988), 2)

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

    def _manage_open_profit_max(self, t, decide: dtime):
        n_reg = self.nifty_regime(t)
        is_flat_market = (n_reg == "NEUTRAL")
        mins_left = max(1, ((15 * 60 + 5) - (decide.hour * 60 + decide.minute)))
        bars_left = mins_left // 5

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

            # Late-Session Zero-Slippage Limit Target Harvest
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
            vwap_std_val = float(vwap_std_cache.get(sym, np.zeros(len(sd.df)))[i])
            vwap_top = vwap_val + 2.0 * vwap_std_val
            vwap_bot = vwap_val - 2.0 * vwap_std_val

            if p["bars_held"] >= 2 and p["mfe"] >= 0.70:
                if p["side"] == "BUY" and hi >= vwap_top and cl < (hi + lo) / 2.0:
                    self._close(sym, cl, t, "EXHAUSTION (VWAP +2.0 Sigma Rejection)", market=True); continue
                elif p["side"] == "SELL" and lo <= vwap_bot and cl > (hi + lo) / 2.0:
                    self._close(sym, cl, t, "EXHAUSTION (VWAP -2.0 Sigma Rejection)", market=True); continue

            # 2. Structural Trailing Stop (2-Bar Swing Low/High) + Adaptive Breakeven
            if p["bars_held"] >= 3 and p["mfe"] >= 0.35:
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

            c10, _ = sd.tenmin_state(i)
            e30 = float(sd.ema[30][i])
            if p["bars_held"] >= 3:
                if p["side"] == "BUY" and c10 == "BEARISH" and cl < vwap_val and cl < e30:
                    self._close(sym, cl, t, "STRUCTURAL_REVERSAL (10m+VWAP+EMA30)", market=True); continue
                if p["side"] == "SELL" and c10 == "BULLISH" and cl > vwap_val and cl > e30:
                    self._close(sym, cl, t, "STRUCTURAL_REVERSAL (10m+VWAP+EMA30)", market=True); continue

    def _scan_profit_max(self, t, decide: dtime):
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
                n_prev3 = float(ndf["close"].iloc[ni-3])
                n_15m_ret = (n_cl - n_prev3) / n_prev3 * 100.0
                
            d_norm = t.normalize()
            d_bars = ndf.loc[ndf.index.normalize() == d_norm]
            d_bars_up_to_t = d_bars.loc[d_bars.index <= t]
            if len(d_bars_up_to_t) >= 5:
                n_open = float(d_bars_up_to_t['open'].iloc[0])
                n_high = float(d_bars_up_to_t['high'].max())
                n_low = float(d_bars_up_to_t['low'].min())
                n_ret = (n_cl - n_open) / n_open * 100.0
                n_rng = (n_high - n_low) / n_open * 100.0
                is_nifty_choppy = (abs(n_ret) < 0.35 and n_rng < 0.55) or (abs(n_e10 - n_e30) / n_e30 < 0.0005)

        n_reg = self.nifty_regime(t)
        cands = []

        for sym, sd in self.syms.items():
            if t not in sd.pos:
                continue
            i = sd.pos[t]
            if i < 2:
                continue
            c10, rsi10 = sd.tenmin_state(i)
            ltp = float(sd.c[i])
            adx = float(sd.adx[i])
            prev_adx = float(sd.adx[i - 1])
            vwap_val = float(sd.vwap[i])
            pdi, mdi = float(sd.plus_di[i]), float(sd.minus_di[i])
            e10, e30 = float(sd.ema[10][i]), float(sd.ema[30][i])
            e50 = float(sd.ema[50][i])
            atr_val = float(sd.atr[i]) if not np.isnan(sd.atr[i]) else 0.0
            atr_pct = (atr_val / ltp * 100.0) if ltp > 0 else 0.0
            
            cur_vol = float(sd.v[i])
            rolling_rvol = float(sd.rvol[i]) if not np.isnan(sd.rvol[i]) else 0.0
            tod_mean = tod_mean_cache.get((sym, t.date(), t.time()), 0.0)
            eff_rvol = max(rolling_rvol, cur_vol / tod_mean if tod_mean > 0 else rolling_rvol)
            
            stock_rs_ratio = (ltp / e10) if e10 > 0 else 1.0
            rs_formula_pct = ((stock_rs_ratio / n_rs_ratio) - 1.0) * 100.0

            if sym in self.open or sym in self.pending or sym in self.traded_symbols_today:
                continue
            if np.isnan(rsi10) or np.isnan(adx) or np.isnan(prev_adx) or adx < 20.0 or adx < prev_adx or atr_pct < 0.14:
                continue

            # Check Setup A: Classical 10m Crossover Pullback
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
            # Fired when a stock is in an established trend above VWAP, pulls back to VWAP/EMA10/EMA30,
            # and prints a strong continuation without needing a brand new 10m cross!
            if not setup_type and self.enable_engine2:
                # Strong established trend check
                if ltp > vwap_val and e10 > e30 and pdi > mdi and (48.0 <= rsi10 <= 66.0):
                    # Pullback touch to VWAP or EMA10/30 within last 2 bars, now holding above
                    lo_prev = float(sd.l[i-1])
                    vwap_prev = float(sd.vwap[i-1])
                    if lo_prev <= vwap_prev * 1.002 or lo_prev <= float(sd.ema[30][i-1]):
                        if float(sd.c[i]) > float(sd.o[i]) and float(sd.c[i]) > float(sd.c[i-1]): # green confirmation
                            side = "BUY"
                            limit = float(sd.c[i]) # instant market/limit entry on confirmation
                            period = 10
                            setup_type = "ENGINE2_VWAP_CONT"
                elif ltp < vwap_val and e10 < e30 and mdi > pdi and (34.0 <= rsi10 <= 52.0):
                    hi_prev = float(sd.h[i-1])
                    vwap_prev = float(sd.vwap[i-1])
                    if hi_prev >= vwap_prev * 0.998 or hi_prev >= float(sd.ema[30][i-1]):
                        if float(sd.c[i]) < float(sd.o[i]) and float(sd.c[i]) < float(sd.c[i-1]):
                            side = "SELL"
                            limit = float(sd.c[i])
                            period = 10
                            setup_type = "ENGINE2_VWAP_CONT"

            if not setup_type:
                continue

            # Decoupling filter on flat/contra market
            must_decouple = is_nifty_choppy or (side == "BUY" and n_reg != "BULL") or (side == "SELL" and n_reg != "BEAR")
            if must_decouple:
                if eff_rvol < 1.30: # Calibrated for high-momentum multi-engine
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
                conv_score += 0.5 # boost continuation setups in strong trends
                
            cands.append(dict(symbol=sym, side=side, ltp=ltp, rsi=rsi10, adx=adx, rvol=eff_rvol,
                              atr_pct=atr_pct, conv_score=conv_score, setup=setup_type,
                              limit=limit, ema_period=period, sector=sd.sector))

        cands.sort(key=lambda c: (-c["conv_score"], -c["rvol"], c["symbol"]))
        free = C.MAX_CONCURRENT_POSITIONS - len(self.open) - len(self.pending)
        total_equity = self.starting_capital + self.cash_realized
        committed_cash = sum(p["entry"] * p["qty"] for p in self.open.values()) + \
                         sum(o["limit"] * o["qty"] for o in self.pending.values())
        avail_cash = max(0.0, total_equity * self.leverage - committed_cash)
        n_take = min(len(cands), free)

        min_cash_buffer = max(1000.0, total_equity * 0.05)
        for k, c in enumerate(cands):
            if k >= free or avail_cash < min_cash_buffer:
                continue
            equal_share = avail_cash / max(1, (n_take - k))
            target_alloc = total_equity * self.leverage * 0.48
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

if __name__ == '__main__':
    t0 = time.time()
    cap_levels = [30000.0, 50000.0, 100000.0]
    margin_cases = [
        ("Case A: 1.0x Cash (No Leverage)", 1.0),
        ("Case B: 2.0x MIS Intraday Margin (Standard)", 2.0)
    ]
    
    summary_matrix = []

    for case_label, lev in margin_cases:
        for cap in cap_levels:
            tot_trades, tot_wins = 0, 0
            tot_gross, tot_fees, tot_net = 0.0, 0.0, 0.0
            all_closed = []
            
            for d in dates:
                dt = datetime.combine(d, datetime.min.time())
                sim = DualEngineSimulator(dt, starting_capital=cap, leverage=lev, enable_engine2=True)
                sim.syms = master.syms
                sim.indices = master.indices
                sim.nifty_pos = master.nifty_pos
                sim.nifty_e10 = master.nifty_e10
                sim.nifty_e21 = master.nifty_e21
                sim.nifty_e30 = master.nifty_e30
                for t in sim.session_bar_times():
                    sim.step(t)
                snap = sim.snapshot()
                k = snap['kpi']
                tot_trades += k['trades']
                tot_wins += k['wins']
                tot_gross += k['gross']
                tot_fees += k['fees']
                tot_net += k['realized']
                all_closed.extend(sim.closed)
                
            wr = (tot_wins / tot_trades * 100) if tot_trades else 0
            pct_return = (tot_net / cap) * 100
            trades_per_day = tot_trades / len(dates)
            trades_per_week = trades_per_day * 5.0
            
            summary_matrix.append({
                "case": case_label,
                "capital": cap,
                "trades": tot_trades,
                "trades_per_day": round(trades_per_day, 1),
                "trades_per_week": round(trades_per_week, 1),
                "wins": tot_wins,
                "wr": round(wr, 1),
                "gross": round(tot_gross, 2),
                "fees": round(tot_fees, 2),
                "net": round(tot_net, 2),
                "pct_ret": round(pct_return, 2)
            })

    print("\n" + "="*95)
    print("REVOLUTIONARY DUAL-ENGINE QUANT RESEARCH RESULTS (61 DAYS, 139 STOCKS)")
    print("="*95)
    header = f"{'Case & Margin':<38} | {'Capital':<10} | {'Trades':<7} | {'Trades/Wk':<9} | {'WR (%)':<6} | {'Gross P&L':<12} | {'Fees':<10} | {'Net Realized':<14} | {'% Return':<8}"
    print(header)
    print("-" * len(header))
    for r in summary_matrix:
        print(f"{r['case']:<38} | Rs {r['capital']:>7,.0f} | {r['trades']:>7} | {r['trades_per_week']:>9.1f} | {r['wr']:>5.1f}% | Rs {r['gross']:>9,.2f} | Rs {r['fees']:>8,.2f} | Rs {r['net']:>11,.2f} | {r['pct_ret']:>+7.2f}%")
