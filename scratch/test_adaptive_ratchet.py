import pickle, pandas as pd
from datetime import datetime, timedelta, time as dtime
import sys
sys.path.insert(0, '.')
sys.path.insert(0, 'sim')
import config as C
from simulator import Simulator, SymbolData, ema_fast, BAR, breakeven_exit_price
from scratch.fast_decoupling_backtest import DecoupledSimulator, init_worker

init_worker()
with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
    payload = pickle.load(f)
raw = payload['raw']
u = payload['universe']
idx = payload['idx']
ndf = idx['NIFTY 50']

class AdaptiveDecoupledSimulator(DecoupledSimulator):
    """
    DecoupledSimulator with Adaptive Choppy-Day Ratchet:
    On FLAT/NEUTRAL days, the market range is compressed.
    Activate Stage 1 (Tax Shield) at +0.25% MFE instead of waiting for +0.34%,
    moving the stop immediately to breakeven + minimal tax shield (min_net_rupees=50.0).
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
                # Trigger Tax Shield at +0.25% on flat days
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

flat_dates = ['2026-08-10', '2026-09-01', '2026-09-04', '2026-09-10', '2026-09-16', '2026-09-18', '2026-09-29', '2026-10-22', '2026-10-28']
print("\n=== TESTING FLAT DAYS WITH ADAPTIVE RATCHET (+0.25% SHIELD) ===")
tot_net = 0.0
tot_trades = 0
tot_wins = 0
for d_str in flat_dates:
    dt = datetime.strptime(d_str, '%Y-%m-%d')
    sim = AdaptiveDecoupledSimulator('replay', dt, decoupling_mode='decoupled_engine', starting_capital=50000.0)
    sim.load(raw, u, idx)
    for t in sim.session_bar_times():
        sim.step(t)
    snap = sim.snapshot()
    k = snap['kpi']
    tot_trades += k['trades']
    tot_wins += k['wins']
    tot_net += k['realized']
    print(f"{d_str}: Trades {k['trades']}, Wins {k['wins']}, Net Rs {k['realized']:+,.2f}")

print(f"\nTOTAL FLAT DAYS: {tot_trades} Trades, {tot_wins} Wins ({tot_wins/tot_trades*100:.1f}% WR), Net Rs {tot_net:+,.2f}")
