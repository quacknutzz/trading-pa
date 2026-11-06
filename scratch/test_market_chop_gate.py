import pickle, pandas as pd
import numpy as np
from datetime import datetime, timedelta, time as dtime
import sys
sys.path.insert(0, '.')
sys.path.insert(0, 'sim')
import config as C
from simulator import Simulator, SymbolData, ema_fast, BAR, breakeven_exit_price
from scratch.fast_decoupling_backtest import DecoupledSimulator, init_worker

class TrueDecoupledLeaderSimulator(DecoupledSimulator):
    """
    True Decoupled Leader Engine:
    Key Principle:
    If Nifty is inside an intraday chop zone (Nifty intraday range < 0.40% or Nifty close within +/-0.25% of open),
    then micro 5m 'BULL' or 'BEAR' regimes are FALSE NOISE.
    In such a market, NO stock gets an 'aligned' pass.
    EVERY stock must strictly satisfy the Decoupled Leader Protocol:
    1. RVOL_TOD >= 1.5 (or rolling RVOL >= 1.5)
    2. Relative Strength RS_EMA >= +0.15% (for BUY) or <= -0.15% (for SELL)
    3. Above VWAP for BUY, Below VWAP for SELL
    4. Adaptive Ratchet: Protect at +0.25% MFE with Tax Shield breakeven.
    """
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
                    self._close(sym, max(p["tgt"], op), t, f"TARGET (Limit 0-Slip +{p.get('tgt_pct', 0.8)}%)", market=False); continue
            else:
                if hi >= p["sl"]:
                    p["mfe"] = max(p["mfe"], bar_mfe); p["mae"] = min(p["mae"], bar_mae)
                    lbl = f"PROFIT_LOCK ({p['sl_state']})" if not p["sl_state"].startswith("INITIAL") else "STOP_LOSS (0.85%)"
                    self._close(sym, max(p["sl"], op), t, lbl, market=True); continue
                if lo <= p["tgt"]:
                    p["mfe"] = max(p["mfe"], bar_mfe); p["mae"] = min(p["mae"], bar_mae)
                    self._close(sym, min(p["tgt"], op), t, f"TARGET (Limit 0-Slip +{p.get('tgt_pct', 0.8)}%)", market=False); continue

            p["mfe"] = max(p["mfe"], bar_mfe)
            p["mae"] = min(p["mae"], bar_mae)

            # Adaptive Tax-Aware Ratchet Curve
            if p["mfe"] >= 0.70:
                lock_pct = (p["mfe"] * 0.68) / 100.0
                new_sl = round(p["entry"] * (1.0 + sgn * lock_pct), 2)
                if (sgn > 0 and new_sl > p["sl"]) or (sgn < 0 and new_sl < p["sl"]):
                    p["sl"] = new_sl
                    p["sl_state"] = f"STAGE3_TRAIL_68% (+{lock_pct*100:.2f}%)"
            elif p["mfe"] >= 0.50:
                lock_pct = (p["mfe"] * 0.58) / 100.0
                new_sl = round(p["entry"] * (1.0 + sgn * lock_pct), 2)
                if (sgn > 0 and new_sl > p["sl"]) or (sgn < 0 and new_sl < p["sl"]):
                    p["sl"] = new_sl
                    p["sl_state"] = f"STAGE2_TRAIL_58% (+{lock_pct*100:.2f}%)"
            elif p["mfe"] >= (0.25 if is_flat_market else 0.34):
                be_px = breakeven_exit_price(p["side"], p["qty"], p["entry"], market_exit=True, min_net_rupees=50.0 if is_flat_market else 160.0)
                if (sgn > 0 and be_px > p["sl"]) or (sgn < 0 and be_px < p["sl"]):
                    p["sl"] = be_px
                    p["sl_state"] = "STAGE1_TAX_SHIELD"
            elif p["mfe"] >= 0.20:
                half_sl = round(p["entry"] * (1.0 - sgn * 0.0038), 2)
                if (sgn > 0 and half_sl > p["sl"]) or (sgn < 0 and half_sl < p["sl"]):
                    p["sl"] = half_sl
                    p["sl_state"] = "STAGE0_RISK_HALVED"

            c10, _ = sd.tenmin_state(i)
            vwap_val = float(sd.vwap[i])
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
            if not c10:
                continue
            side = "BUY" if c10 == "BULLISH" else "SELL"
            ltp = float(sd.c[i])
            adx = float(sd.adx[i])
            prev_adx = float(sd.adx[i - 1])
            vwap_val = float(sd.vwap[i])
            pdi, mdi = float(sd.plus_di[i]), float(sd.minus_di[i])
            e10, e30 = float(sd.ema[10][i]), float(sd.ema[30][i])
            atr_val = float(sd.atr[i]) if not np.isnan(sd.atr[i]) else 0.0
            atr_pct = (atr_val / ltp * 100.0) if ltp > 0 else 0.0
            
            cur_vol = float(sd.v[i])
            rolling_rvol = float(sd.rvol[i]) if not np.isnan(sd.rvol[i]) else 0.0
            tod_mean = self.syms[sym].df.loc[self.syms[sym].df.index.time == t.time(), 'volume'].mean() if len(self.syms[sym].df) else 0.0
            eff_rvol = max(rolling_rvol, cur_vol / tod_mean if tod_mean > 0 else rolling_rvol)
            
            stock_rs_ratio = (ltp / e10) if e10 > 0 else 1.0
            rs_formula_pct = ((stock_rs_ratio / n_rs_ratio) - 1.0) * 100.0
            
            reason = ""
            if sym in self.open or sym in self.pending:
                reason = "already open/pending"
            elif sym in self.traded_symbols_today:
                reason = "Fix #4: per-stock daily cooldown active"
            elif np.isnan(rsi10):
                reason = "10m RSI unavailable"
            elif side == "BUY" and not (50.0 <= rsi10 <= 64.0):
                reason = f"RSI {rsi10:.1f} outside BUY sweet-spot 50-64"
            elif side == "SELL" and not (36.0 <= rsi10 <= 49.0):
                reason = f"RSI {rsi10:.1f} outside SELL sweet-spot 36-49"
            elif np.isnan(adx) or np.isnan(prev_adx) or adx < 20.0:
                reason = f"ADX {adx:.1f} < 20.0"
            elif adx < prev_adx:
                reason = f"Fix #4: ADX falling ({prev_adx:.1f} -> {adx:.1f}, dying trend)"
            elif side == "BUY" and ltp <= vwap_val:
                reason = f"Fix #2: BUY below Intraday VWAP ({ltp:.1f} <= {vwap_val:.1f})"
            elif side == "SELL" and ltp >= vwap_val:
                reason = f"Fix #2: SELL above Intraday VWAP ({ltp:.1f} >= {vwap_val:.1f})"
            elif side == "BUY" and (pdi <= mdi or e10 < e30):
                reason = "5m EMA10/30 or +DI/-DI not aligned with BUY"
            elif side == "SELL" and (mdi <= pdi or e10 > e30):
                reason = "5m EMA10/30 or +DI/-DI not aligned with SELL"
            elif atr_pct < 0.14:
                reason = f"5m ATR ({atr_pct:.2f}%) too low to beat 0.113% tax hurdle"

            # -------------------------------------------------------------
            # INSTITUTIONAL DECOUPLING & CHOP VETO LOGIC
            # -------------------------------------------------------------
            if not reason:
                must_decouple = is_nifty_choppy or (side == "BUY" and n_reg != "BULL") or (side == "SELL" and n_reg != "BEAR")
                
                if must_decouple:
                    if eff_rvol < 1.5:
                        reason = f"Decoupling failed: RVOL {eff_rvol:.2f}x < 1.5x on choppy/contra market"
                    elif side == "BUY" and rs_formula_pct < 0.15:
                        reason = f"Decoupling failed: BUY RS {rs_formula_pct:+.2f}% < +0.15% vs Nifty"
                    elif side == "SELL" and rs_formula_pct > -0.15:
                        reason = f"Decoupling failed: SELL RS {rs_formula_pct:+.2f}% > -0.15% vs Nifty"
                    elif side == "BUY" and n_15m_ret < -0.25 and eff_rvol < 2.0:
                        reason = f"Flash dump protection: Nifty dumping ({n_15m_ret:.2f}%), required RVOL 2.0x > {eff_rvol:.2f}x"
                else:
                    if eff_rvol < 0.8:
                        reason = f"Volume too thin: RVOL {eff_rvol:.2f}x < 0.8x"

            limit, period = sd.pullback_limit(i, side, ltp)
            if not reason and limit <= 0:
                reason = "no 5m EMA on pullback side"
            if reason:
                self._log_signal(str(t + BAR), sym, side, ltp, rsi10, adx, rolling_rvol, "FILTERED", reason, t)
                continue

            conv_score = (eff_rvol * 1.5) + ((adx - 20.0) * 0.1) + (atr_pct * 4.0) + (abs(rs_formula_pct) * 2.0)
            cands.append(dict(symbol=sym, side=side, ltp=ltp, rsi=rsi10, adx=adx, rvol=eff_rvol,
                              atr_pct=atr_pct, conv_score=conv_score,
                              limit=limit, ema_period=period, sector=sd.sector))

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
                self._log_signal(str(t + BAR), c["symbol"], c["side"], c["ltp"], c["rsi"], c["adx"], c["rvol"],
                                 "NO_SLOT", "qualified, but slots/cash currently deployed", t)
                continue
            equal_share = avail_cash / max(1, (n_take - k))
            target_alloc = total_equity * (0.55 if c["conv_score"] >= 3.2 else 0.45)
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
            self._log_signal(str(t + BAR), c["symbol"], c["side"], c["ltp"], c["rsi"], c["adx"], c["rvol"],
                             "ORDER", f"limit {c['side']} {qty} @ {c['limit']} (5m EMA{c['ema_period']} | Conv {c['conv_score']:.1f})", t)

if __name__ == '__main__':
    init_worker()
    with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
        payload = pickle.load(f)
    raw = payload['raw']
    u = payload['universe']
    idx = payload['idx']

    dt = datetime.strptime('2026-09-16', '%Y-%m-%d')
    sim = TrueDecoupledLeaderSimulator('replay', dt, starting_capital=50000.0)
    sim.load(raw, u, idx)
    for t in sim.session_bar_times():
        sim.step(t)

    print("Results for 2026-09-16 (Previously CONCOR lost Rs -267):")
    snap = sim.snapshot()
    k = snap['kpi']
    print(f"Trades: {k['trades']}, Wins: {k['wins']}, Net: Rs {k['realized']:+,.2f}")
    for c in sim.closed:
        print(f"  Trade: {c['symbol']} {c['side']} Net Rs {c['net']:+.2f} Reason: {c['reason']}")
