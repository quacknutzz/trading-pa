import sys, pickle, time, os
from datetime import datetime, timedelta, time as dtime
import pandas as pd
import numpy as np
from concurrent.futures import ProcessPoolExecutor

sys.path.insert(0, 'sim')
import datafeed
import config as C
from simulator import Simulator, SymbolData, ema_fast, BAR

# Worker state
G_RAW = None
G_IDX = None
G_UNIVERSE = None
G_TOD_MEAN = None

def init_worker():
    global G_RAW, G_IDX, G_UNIVERSE, G_TOD_MEAN
    with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
        payload = pickle.load(f)
    G_RAW = payload['raw']
    G_IDX = payload['idx']
    G_UNIVERSE = payload['universe']
    
    # Precompute TOD 20-day mean volume table
    G_TOD_MEAN = {}
    for y_sym, df in G_RAW.items():
        sym = y_sym.replace('.NS', '')
        if len(df) < 200:
            continue
        df_copy = df.copy()
        df_copy['tod'] = df_copy.index.time
        df_copy['date'] = df_copy.index.normalize()
        pv = df_copy.pivot(index='date', columns='tod', values='volume')
        pv_mean20 = pv.shift(1).rolling(20, min_periods=5).mean()
        for date_idx, row in pv_mean20.iterrows():
            d_val = date_idx.date()
            for tod_val, mean_vol in row.items():
                if not np.isnan(mean_vol) and mean_vol > 0:
                    G_TOD_MEAN[(sym, d_val, tod_val)] = float(mean_vol)

class DecoupledSimulator(Simulator):
    def __init__(self, mode, session_date, decoupling_mode='decoupled_engine', starting_capital=50000.0):
        super().__init__(mode, session_date, strategy='v6_profit_max', starting_capital=starting_capital)
        self.decoupling_mode = decoupling_mode

    def _scan_profit_max(self, t, decide: dtime):
        n_reg = self.nifty_regime(t)
        ndf = self.indices.get("NIFTY 50")
        
        # Nifty indicators at t
        n_rs_ratio = 1.0
        n_15m_ret = 0.0
        if ndf is not None and t in self.nifty_pos:
            ni = self.nifty_pos[t]
            n_cl = float(ndf["close"].iloc[ni])
            n_e10 = float(self.nifty_e10[ni])
            n_rs_ratio = (n_cl / n_e10) if n_e10 > 0 else 1.0
            if ni >= 3:
                n_prev3 = float(ndf["close"].iloc[ni-3])
                n_15m_ret = (n_cl - n_prev3) / n_prev3 * 100.0

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
            
            # Volume & RVOL
            cur_vol = float(sd.v[i])
            rolling_rvol = float(sd.rvol[i]) if not np.isnan(sd.rvol[i]) else 0.0
            tod_mean = G_TOD_MEAN.get((sym, t.date(), t.time()), 0.0) if G_TOD_MEAN else 0.0
            tod_rvol = (cur_vol / tod_mean) if tod_mean > 0 else rolling_rvol
            
            # Relative Strength (RS) formula: (Stock_Close / Stock_EMA10) / (Nifty_Close / Nifty_EMA10)
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
            # NIFTY FILTER / DECOUPLING ENGINE LOGIC
            # -------------------------------------------------------------
            if not reason:
                if self.decoupling_mode == 'baseline':
                    if rolling_rvol < 1.35 and ((side == "BUY" and n_reg == "BEAR") or (side == "SELL" and n_reg == "BULL")):
                        reason = f"Contra to Nifty 50 ({n_reg}) with low RVOL ({rolling_rvol:.2f}x < 1.35x)"
                
                elif self.decoupling_mode == 'strict_nifty_veto':
                    if side == "BUY" and n_reg != "BULL":
                        reason = f"Nifty {n_reg} != BULL (Strict Veto)"
                    elif side == "SELL" and n_reg != "BEAR":
                        reason = f"Nifty {n_reg} != BEAR (Strict Veto)"
                        
                elif self.decoupling_mode == 'decoupled_engine':
                    is_aligned = (side == "BUY" and n_reg == "BULL") or (side == "SELL" and n_reg == "BEAR")
                    eff_rvol = max(tod_rvol, rolling_rvol)
                    
                    if not is_aligned:
                        # Market is NEUTRAL, FLAT, or CONTRA. Stock MUST prove institutional decoupling!
                        # 1. High Institutional Volume (RVOL >= 1.5)
                        if eff_rvol < 1.5:
                            reason = f"Decoupling failed: RVOL {eff_rvol:.2f}x < 1.5x on non-aligned market ({n_reg})"
                        
                        # 2. Positive Relative Strength outperformance
                        elif side == "BUY" and rs_formula_pct < 0.15:
                            reason = f"Decoupling failed: BUY RS {rs_formula_pct:+.2f}% < +0.15% relative to Nifty"
                        elif side == "SELL" and rs_formula_pct > -0.15:
                            reason = f"Decoupling failed: SELL RS {rs_formula_pct:+.2f}% > -0.15% relative to Nifty"
                            
                        # 3. Anti-Flash Dump Protection:
                        # If Nifty is experiencing a steep intraday drop (15m ret < -0.25%),
                        # BUY is blocked unless stock is exhibiting active heavy absorption (RVOL >= 2.0x)
                        elif side == "BUY" and n_15m_ret < -0.25 and eff_rvol < 2.0:
                            reason = f"Flash dump protection: Nifty dumping ({n_15m_ret:.2f}%), required RVOL 2.0x > {eff_rvol:.2f}x"
                    else:
                        if eff_rvol < 0.8:
                            reason = f"Volume too thin: RVOL {eff_rvol:.2f}x < 0.8x even with aligned market"

            limit, period = sd.pullback_limit(i, side, ltp)
            if not reason and limit <= 0:
                reason = "no 5m EMA on pullback side"
            if reason:
                self._log_signal(str(t + BAR), sym, side, ltp, rsi10, adx, rolling_rvol, "FILTERED", reason, t)
                continue

            eff_rvol = max(tod_rvol, rolling_rvol)
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

def run_single_simulation(args):
    mode, d_str, day_cat, nifty_ret = args
    dt = datetime.strptime(d_str, '%Y-%m-%d')
    sim = DecoupledSimulator('replay', dt, decoupling_mode=mode, starting_capital=50000.0)
    sim.load(G_RAW, G_UNIVERSE, G_IDX)
    for t in sim.session_bar_times():
        sim.step(t)
    snap = sim.snapshot()
    k = snap['kpi']
    return {
        'mode': mode,
        'date': d_str,
        'day_cat': day_cat,
        'nifty_ret': nifty_ret,
        'trades': k['trades'],
        'wins': k['wins'],
        'gross': k['gross'],
        'fees': k['fees'],
        'net': k['realized']
    }

if __name__ == '__main__':
    t0 = time.time()
    # Load once in main to get date list
    with open('sim/cache/nse_5m_60d.pkl', 'rb') as f:
        payload = pickle.load(f)
    ndf = payload['idx']['NIFTY 50']
    days = sorted({ts.normalize() for ts in ndf.index})[3:]
    
    tasks = []
    modes = ['baseline', 'strict_nifty_veto', 'decoupled_engine']
    for d in days:
        day_df = ndf[ndf.index.normalize() == d]
        if len(day_df) < 5:
            continue
        o = day_df['open'].iloc[0]
        c = day_df['close'].iloc[-1]
        ret = (c - o) / o * 100
        day_cat = 'GREEN' if ret > 0.25 else ('RED' if ret < -0.25 else 'FLAT')
        d_str = str(d.date())
        for m in modes:
            tasks.append((m, d_str, day_cat, ret))
            
    print(f"Executing {len(tasks)} simulation runs across {os.cpu_count()} CPU cores...")
    
    with ProcessPoolExecutor(max_workers=8, initializer=init_worker) as executor:
        results = list(executor.map(run_single_simulation, tasks))
        
    df_all = pd.DataFrame(results)
    
    print(f"\nCompleted in {time.time() - t0:.1f} seconds!")
    print("\n" + "="*80)
    print("COMPARATIVE RESEARCH RESULTS ACROSS 58 TRADING DAYS (50K CAPITAL)")
    print("="*80)
    
    for m in modes:
        df_m = df_all[df_all['mode'] == m]
        tot_trades = df_m['trades'].sum()
        tot_wins = df_m['wins'].sum()
        tot_gross = df_m['gross'].sum()
        tot_fees = df_m['fees'].sum()
        tot_net = df_m['net'].sum()
        wr = (tot_wins / tot_trades * 100) if tot_trades else 0
        pf = (tot_gross / tot_fees) if tot_fees else 0 # rough
        
        print(f"\n=================================================================")
        print(f"STRATEGY: {m.upper()}")
        print(f"=================================================================")
        print(f"Total Trades: {tot_trades} | Wins: {tot_wins} ({wr:.1f}% WR)")
        print(f"Gross Profit: Rs {tot_gross:+,.2f}")
        print(f"Taxes / Fees: Rs {tot_fees:,.2f}")
        print(f"Net Realized: Rs {tot_net:+,.2f}  ({tot_net/50000*100:+.2f}% on 50k)")
        
        reg_summary = df_m.groupby('day_cat').agg(
            trades=('trades', 'sum'),
            wins=('wins', 'sum'),
            gross=('gross', 'sum'),
            fees=('fees', 'sum'),
            net=('net', 'sum')
        )
        reg_summary['win_rate'] = (reg_summary['wins'] / reg_summary['trades'] * 100).round(1)
        reg_summary['net_per_trade'] = (reg_summary['net'] / reg_summary['trades']).round(2)
        print("\nPerformance by Market Day Type:")
        print(reg_summary[['trades', 'wins', 'win_rate', 'gross', 'fees', 'net', 'net_per_trade']])
        df_m.to_csv(f'scratch/sim_results_{m}.csv', index=False)

    df_all.to_csv('scratch/comparative_decoupling_results.csv', index=False)
    print("\nSaved all results to scratch/comparative_decoupling_results.csv")
