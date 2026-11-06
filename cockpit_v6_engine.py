"""
cockpit_v6_engine.py -- Automated Signal Discovery, Trade Application, and Management
for the 10m/5m EMA 10/30 Strategy in COCKPIT_V6.

Architecture:
- Runs immediately after AI_BTB_Engine.py completes each 5-minute cycle.
- Zero disruption: standalone file, does not modify any existing scripts.
- Universe: Reads active tickers from AI_BTB_Dashboard.xlsm (Sheet1, Col A).
- Parquet: Reads fresh 5m bars synced by AI_BTB_Engine.
- State: Persists open trades in data/cockpit_v6_state.json across runs.
- Excel: Updates 'COCKPIT_V6' sheet in AI_BTB_Dashboard.xlsm via win32com/openpyxl.
- Broker: Dispatches orders to FYERS API (when --live is passed) or simulates (default paper mode).
"""

import os
import sys
import json
import logging
import argparse
from datetime import datetime, time as dtime
from typing import Dict, Any, List, Optional, Tuple

import numpy as np
import pandas as pd

# ── ROOT DIRECTORY & PATH SETUP ──────────────────────────────────────────────
ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from engine.config import (
    EXCEL_FILENAME,
    EXCEL_SHEET,
    TOKEN_FILENAME,
    get_excel_path,
)
from engine.data_engine import get_root_folder, load_fyers_credentials, master_folder_5m
from engine import numba_kernels as nk

# ── STRATEGY CONFIGURATION ───────────────────────────────────────────────────
COCKPIT_SHEET_NAME = "COCKPIT_V6"
V6_DATA_START_ROW = 5         # Row 4 is headers; Row 5 is first trade row
STATE_FILE = os.path.join(ROOT_DIR, "data", "cockpit_v6_state.json")

EMA_FAST = 10
EMA_SLOW = 30
SL_PCT = 0.0085               # 0.85% Initial Stop-loss
MAX_POSITIONS = 3             # Maximum concurrent active positions (Strict Equal Slot Sizing)
CAPITAL_PER_TRADE = 33333.0   # Equal slot allocation (1/3rd of Rs 1L buying power on Rs 50k @ 2x)
DEFAULT_QTY = 100             # Fallback quantity if capital calculation fails

ENTRY_START_TIME = dtime(9, 25)
ENTRY_CUTOFF_TIME = dtime(13, 15)
FORCE_EXIT_TIME = dtime(15, 5)

# ── LOGGING SETUP ────────────────────────────────────────────────────────────
LOG_FILE = os.path.join(ROOT_DIR, "logs", "cockpit_v6.log")
os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)

logger = logging.getLogger("COCKPIT_V6")
logger.setLevel(logging.INFO)
if not logger.handlers:
    fh = logging.FileHandler(LOG_FILE, encoding="utf-8")
    fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(logging.Formatter("[V6] %(message)s"))
    logger.addHandler(fh)
    logger.addHandler(sh)


# ── INDICATOR MATH ───────────────────────────────────────────────────────────
def ema_series(series: pd.Series, period: int) -> pd.Series:
    """Standard pandas ewm with adjust=False matching TradingView / Zerodha / Fyers."""
    return series.ewm(span=period, adjust=False).mean()


def build_paired_10m(d5: pd.DataFrame) -> pd.DataFrame:
    """Builds sealed 10m bars by pairing 5m bars starting at 09:15."""
    if d5.empty or len(d5) < 2:
        return pd.DataFrame()

    bars = []
    for d, day_df in d5.groupby(d5.index.date):
        n = len(day_df)
        if n < 2:
            continue
        n_pairs = (n // 2) * 2
        day_paired = day_df.iloc[:n_pairs]

        times = day_paired.index[1::2]
        opens = day_paired["open"].values[0::2]
        highs = np.maximum(day_paired["high"].values[0::2], day_paired["high"].values[1::2])
        lows = np.minimum(day_paired["low"].values[0::2], day_paired["low"].values[1::2])
        closes = day_paired["close"].values[1::2]
        volumes = day_paired["volume"].values[0::2] + day_paired["volume"].values[1::2]

        df_day = pd.DataFrame(
            {"open": opens, "high": highs, "low": lows, "close": closes, "volume": volumes},
            index=times,
        )
        bars.append(df_day)

    return pd.concat(bars).sort_index() if bars else pd.DataFrame()


def evaluate_latest_10m_signals(df_5m: pd.DataFrame) -> Dict[str, Any]:
    """
    Evaluates the live forming 10m EMA crossover on the LATEST 5m bar.
    Returns:
      {
        'cross_up': bool,
        'cross_dn': bool,
        'cur_e10': float,
        'cur_e30': float,
        'e10_5m': float,
        'e30_5m': float,
        'cross_up_5m': bool,
        'cross_dn_5m': bool,
        'ltp': float,
        'high': float,
        'low': float,
        'open': float,
        'ts': pd.Timestamp
      }
    """
    if df_5m.empty or len(df_5m) < 40:
        return {}

    # Slice to last 600 bars (~8 trading days) for microsecond evaluation
    if len(df_5m) > 600:
        df_5m = df_5m.tail(600).copy()

    d10 = build_paired_10m(df_5m)
    if d10.empty or len(d10) < 30:
        return {}

    e10_10 = ema_series(d10["close"], EMA_FAST)
    e30_10 = ema_series(d10["close"], EMA_SLOW)

    a10 = 2.0 / (EMA_FAST + 1)
    a30 = 2.0 / (EMA_SLOW + 1)
    times = df_5m.index
    n = len(df_5m)
    closes = df_5m["close"].values

    is_sealed = np.isin(times, d10.index)
    s10_vals = e10_10.reindex(times).values
    s30_vals = e30_10.reindex(times).values

    cross_up = np.zeros(n, dtype=bool)
    cross_dn = np.zeros(n, dtype=bool)
    live_e10 = np.zeros(n, dtype=np.float64)
    live_e30 = np.zeros(n, dtype=np.float64)

    last_s10, last_s30, prev_diff = np.nan, np.nan, np.nan

    for i in range(n):
        c = closes[i]
        if is_sealed[i]:
            s10 = s10_vals[i]
            s30 = s30_vals[i]
            cur10, cur30 = s10, s30
            last_s10, last_s30 = s10, s30
        else:
            if not np.isnan(last_s10):
                cur10 = a10 * c + (1.0 - a10) * last_s10
                cur30 = a30 * c + (1.0 - a30) * last_s30
            else:
                cur10, cur30 = np.nan, np.nan

        live_e10[i] = cur10
        live_e30[i] = cur30
        diff = cur10 - cur30
        if not np.isnan(prev_diff) and not np.isnan(diff):
            cross_up[i] = (prev_diff <= 0) and (diff > 0)
            cross_dn[i] = (prev_diff >= 0) and (diff < 0)
        prev_diff = diff

    # 5m indicators for entry & exit
    e10_5 = ema_series(df_5m["close"], EMA_FAST).values
    e30_5 = ema_series(df_5m["close"], EMA_SLOW).values
    diff5 = e10_5 - e30_5
    cross_up_5m = (np.r_[0, diff5[:-1]] <= 0) & (diff5 > 0)
    cross_dn_5m = (np.r_[0, diff5[:-1]] >= 0) & (diff5 < 0)

    h_arr = df_5m["high"].values.astype(np.float64)
    l_arr = df_5m["low"].values.astype(np.float64)
    c_arr = df_5m["close"].values.astype(np.float64)
    v_arr = df_5m["volume"].values.astype(np.float64)
    pdi, mdi, adx_arr = nk.adx(h_arr, l_arr, c_arr, 14)
    atr_arr = nk.atr(h_arr, l_arr, c_arr, 14)

    # Causal Intraday Session VWAP for today's bars
    last_idx = n - 1
    last_date = times[last_idx].date()
    day_mask = times.date == last_date
    tp_day = (h_arr[day_mask] + l_arr[day_mask] + c_arr[day_mask]) / 3.0
    vol_day = np.maximum(v_arr[day_mask], 1.0)
    vwap_val = float(np.sum(tp_day * vol_day) / np.sum(vol_day)) if len(tp_day) > 0 else float(c_arr[last_idx])

    cur_adx = float(adx_arr[last_idx]) if not np.isnan(adx_arr[last_idx]) else 0.0
    prev_adx = float(adx_arr[last_idx - 1]) if last_idx >= 1 and not np.isnan(adx_arr[last_idx - 1]) else 0.0
    prev2_adx = float(adx_arr[last_idx - 2]) if last_idx >= 2 and not np.isnan(adx_arr[last_idx - 2]) else 0.0
    atr_val = float(atr_arr[last_idx]) if not np.isnan(atr_arr[last_idx]) else 0.0
    ltp_val = float(c_arr[last_idx])
    atr_pct = (atr_val / ltp_val * 100.0) if ltp_val > 0 else 0.0

    sw_low_2 = float(np.min(l_arr[max(0, last_idx - 2):last_idx])) if last_idx >= 1 else float(l_arr[last_idx])
    sw_high_2 = float(np.max(h_arr[max(0, last_idx - 2):last_idx])) if last_idx >= 1 else float(h_arr[last_idx])

    return {
        "cross_up": bool(cross_up[last_idx]),
        "cross_dn": bool(cross_dn[last_idx]),
        "cur_e10": round(float(live_e10[last_idx]), 2),
        "cur_e30": round(float(live_e30[last_idx]), 2),
        "c10_diff": float(live_e10[last_idx] - live_e30[last_idx]),
        "e10_5m": round(float(e10_5[last_idx]), 2),
        "e30_5m": round(float(e30_5[last_idx]), 2),
        "cross_up_5m": bool(cross_up_5m[last_idx]),
        "cross_dn_5m": bool(cross_dn_5m[last_idx]),
        "vwap": round(vwap_val, 2),
        "adx": round(cur_adx, 2),
        "prev_adx": round(prev_adx, 2),
        "prev2_adx": round(prev2_adx, 2),
        "pdi": float(pdi[last_idx]) if not np.isnan(pdi[last_idx]) else 0.0,
        "mdi": float(mdi[last_idx]) if not np.isnan(mdi[last_idx]) else 0.0,
        "atr_pct": round(atr_pct, 3),
        "sw_low_2": round(sw_low_2 * 0.9995, 2),
        "sw_high_2": round(sw_high_2 * 1.0005, 2),
        "ltp": ltp_val,
        "high": float(df_5m["high"].iloc[last_idx]),
        "low": float(df_5m["low"].iloc[last_idx]),
        "open": float(df_5m["open"].iloc[last_idx]),
        "volume": float(df_5m["volume"].iloc[last_idx]),
        "ts": df_5m.index[last_idx],
    }


# ── UNIVERSE & STATE PERSISTENCE ─────────────────────────────────────────────
UNIVERSE_CACHE_FILE = os.path.join(ROOT_DIR, "data", "v6_universe_cache.json")

def load_universe_symbols() -> List[str]:
    """Reads the active universe, cached to avoid re-opening 6MB Excel file every 5 mins."""
    if os.path.exists(UNIVERSE_CACHE_FILE):
        try:
            with open(UNIVERSE_CACHE_FILE, "r") as f:
                return json.load(f)
        except Exception:
            pass

    import openpyxl

    path = get_excel_path()
    symbols = []
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb[EXCEL_SHEET]
        for row in range(4, 250):
            val = ws.cell(row, 1).value
            if val:
                s = str(val).strip().upper()
                if s and s != "SYMBOL" and s != "DATE" and not s.startswith("^"):
                    symbols.append(s)
        wb.close()
    except Exception as e:
        logger.warning(f"Could not read universe from Sheet1 ({e}). Falling back to master parquets.")
        master_dir = os.path.join(ROOT_DIR, "data", "master")
        if os.path.exists(master_dir):
            symbols = [
                f.replace("_NS.parquet", "").replace(".parquet", "")
                for f in os.listdir(master_dir)
                if f.endswith(".parquet")
            ]

    # Filter out market indices like NSEI, NSEBANK
    filtered = sorted(list(set([s for s in symbols if not s.startswith("NSE") and s not in ["NIFTY", "BANKNIFTY"]])))
    try:
        with open(UNIVERSE_CACHE_FILE, "w") as f:
            json.dump(filtered, f)
    except Exception:
        pass
    return filtered



def load_state() -> Dict[str, Any]:
    """Loads active positions and session state from JSON."""
    os.makedirs(os.path.dirname(STATE_FILE), exist_ok=True)
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            logger.error(f"Error reading state file ({e}). Starting fresh.")
    return {
        "positions": {},       # symbol -> position dict
        "closed_today": [],    # list of closed trades today
        "session_date": str(datetime.now().date()),
    }


def save_state(state: Dict[str, Any]) -> None:
    """Saves state atomically to JSON."""
    os.makedirs(os.path.dirname(STATE_FILE), exist_ok=True)
    tmp_path = STATE_FILE + ".tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2, default=str)
    os.replace(tmp_path, STATE_FILE)


from engine.fyers_order_manager import FyersOrderManager


# ── EXCEL COCKPIT_V6 WRITER ──────────────────────────────────────────────────
class CockpitV6Writer:
    """Manages writing and updating rows in the COCKPIT_V6 sheet of AI_BTB_Dashboard.xlsm."""

    def __init__(self):
        self.excel_path = get_excel_path()
        self.sheet_name = COCKPIT_SHEET_NAME

    def sync_to_excel(self, state: Dict[str, Any]) -> bool:
        """
        Synchronizes all open positions and today's closed positions directly
        to the COCKPIT_V6 sheet in AI_BTB_Dashboard.xlsm in a single batch pass.
        Uses openpyxl with keep_vba=True for reliable non-blocking persistence.
        """
        import openpyxl

        try:
            wb = openpyxl.load_workbook(self.excel_path, keep_vba=True)
            ws = wb[self.sheet_name]

            positions = state.get("positions", {})
            closed = state.get("closed_today", [])

            row = V6_DATA_START_ROW
            # 1. Write active positions
            for sym, pos in positions.items():
                pos["excel_row"] = row
                ws.cell(row, 1, pos["symbol"])
                ws.cell(row, 2, pos["side"])
                cur_ltp = pos.get("current_ltp", pos["entry_price"])
                ws.cell(row, 3, cur_ltp)
                ws.cell(row, 17, f"CROSS_{pos['side'][:2]}")
                ws.cell(row, 19, pos["entry_price"])
                ws.cell(row, 20, "-")
                ws.cell(row, 21, pos["sl_price"])
                ws.cell(row, 22, cur_ltp)
                ws.cell(row, 23, str(pos["entry_ts"]))
                ws.cell(row, 25, "OPEN")

                # Bottomline (unrealized PnL)
                pnl_rs = pos.get("unrealized_pnl_rs", 0.0)
                ws.cell(row, 26, round(pnl_rs, 2))
                ws.cell(row, 27, "Active Trade Monitoring")
                ws.cell(row, 28, pos.get("bars_held", 0))
                ws.cell(row, 30, pos.get("peak_price", pos["entry_price"]))
                ws.cell(row, 31, round(pos.get("best_mfe_pct", 0.0), 2))
                ws.cell(row, 35, str(pos.get("last_bar_ts", pos["entry_ts"])))
                row += 1

            # 2. Write today's closed positions underneath
            for pos in closed:
                ws.cell(row, 1, pos["symbol"])
                ws.cell(row, 2, pos["side"])
                exit_price = pos.get("exit_price", pos["entry_price"])
                ws.cell(row, 3, exit_price)
                ws.cell(row, 17, f"CROSS_{pos['side'][:2]}")
                ws.cell(row, 19, pos["entry_price"])
                ws.cell(row, 20, "-")
                ws.cell(row, 21, pos["sl_price"])
                ws.cell(row, 22, exit_price)
                ws.cell(row, 23, str(pos["entry_ts"]))
                ws.cell(row, 24, str(pos.get("exit_ts", "-")))
                outcome = "WIN" if pos.get("net_pnl_rs", 0.0) > 0 else "LOSS"
                ws.cell(row, 25, outcome)
                ws.cell(row, 26, round(pos.get("net_pnl_rs", 0.0), 2))
                ws.cell(row, 27, f"Exit: {pos.get('close_reason', 'CLOSED')}")
                ws.cell(row, 28, pos.get("bars_held", 0))
                ws.cell(row, 30, pos.get("peak_price", pos["entry_price"]))
                ws.cell(row, 31, round(pos.get("best_mfe_pct", 0.0), 2))
                ws.cell(row, 35, str(pos.get("exit_ts", "-")))
                row += 1

            wb.save(self.excel_path)
            wb.close()
            logger.info(f"COCKPIT_V6 synchronized ({len(positions)} open, {len(closed)} closed).")
            return True
        except Exception as e:
            logger.warning(f"Could not sync COCKPIT_V6 to Excel: {e}")
            return False




# ── CORE ORCHESTRATION CYCLE ─────────────────────────────────────────────────
def run_cockpit_cycle(is_live: bool = False, force: bool = False) -> None:
    """
    Executes one 5-minute cycle:
    1. Defense: Evaluate exits on active positions.
    2. Offense: Scan universe for new 10m EMA crosses and enter candidates.
    3. Persist state and update COCKPIT_V6.
    """
    now = datetime.now()
    cur_time = now.time()

    logger.info(f"--- COCKPIT_V6 CYCLE START | {now.strftime('%Y-%m-%d %H:%M:%S')} (Live={is_live}) ---")

    # Time gate (unless forced for testing)
    if not force:
        if cur_time < ENTRY_START_TIME:
            logger.info(f"Market pre-open ({cur_time.strftime('%H:%M')}). Cycle idle.")
            return
        if cur_time > FORCE_EXIT_TIME and cur_time > dtime(15, 30):
            logger.info(f"Market closed ({cur_time.strftime('%H:%M')}). Cycle idle.")
            return

    state = load_state()
    # Reset closed list on new trading day
    today_str = str(now.date())
    if state.get("session_date") != today_str:
        state["session_date"] = today_str
        state["closed_today"] = []

    broker = FyersOrderManager(root=ROOT_DIR, is_live=is_live)
    writer = CockpitV6Writer()
    master_dir = master_folder_5m(ROOT_DIR)

    # Reconcile local state with broker positions
    active_positions = state.setdefault("positions", {})
    if is_live:
        externally_closed = broker.reconcile_broker_positions(active_positions)
        for sym in externally_closed:
            pos = active_positions.pop(sym, None)
            if pos:
                writer.record_exit(pos, exit_price=pos.get("entry_price", 0.0), exit_ts=now, reason="BROKER_RECONCILED")
                state.setdefault("closed_today", []).append(pos)

    # =========================================================================
    # PHASE 1: DEFENSE (Monitor Active Positions)
    # =========================================================================
    symbols_to_close = []

    for sym, pos in list(active_positions.items()):
        pq_path = os.path.join(master_dir, f"{sym}_NS.parquet")
        if not os.path.exists(pq_path):
            pq_path = os.path.join(master_dir, f"{sym}.parquet")
        if not os.path.exists(pq_path):
            continue

        try:
            df_5m = pd.read_parquet(pq_path)
            sig = evaluate_latest_10m_signals(df_5m)
            if not sig:
                continue

            c = sig["ltp"]
            h = sig["high"]
            l = sig["low"]
            op = sig["open"]
            t = sig["ts"]
            entry = pos["entry_price"]
            side = pos["side"]
            sl = pos["sl_price"]

            # Update telemetry
            pos["bars_held"] = pos.get("bars_held", 0) + 1
            if side == "BUY":
                pos["peak_price"] = max(pos.get("peak_price", entry), h)
                mfe = (pos["peak_price"] - entry) / entry * 100.0
            else:
                pos["peak_price"] = min(pos.get("peak_price", entry), l)
                mfe = (entry - pos["peak_price"]) / entry * 100.0
            pos["best_mfe_pct"] = max(pos.get("best_mfe_pct", 0.0), mfe)

            # Fix #3: 2-Stage Post-Tax Breakeven Ratchet + 2-Bar Swing Trail (after 3 bars & MFE >= 0.35%)
            if pos["bars_held"] >= 3 and pos["best_mfe_pct"] >= 0.35:
                sgn = 1.0 if side == "BUY" else -1.0
                # Statutory fee (~0.16%) + 0.05% proportional net profit buffer (0.21% total)
                be_px = round(entry * (1.0 + sgn * 0.0021), 2)
                if (side == "BUY" and sl < be_px < c) or (side == "SELL" and sl > be_px > c):
                    sl = be_px
                    pos["sl_price"] = sl
                if side == "BUY" and sl < sig["sw_low_2"] < c:
                    sl = sig["sw_low_2"]
                    pos["sl_price"] = sl
                elif side == "SELL" and sl > sig["sw_high_2"] > c:
                    sl = sig["sw_high_2"]
                    pos["sl_price"] = sl

            # Check Exit Conditions
            exit_price, reason = None, None

            # 1. Stop-Loss / Trailing Profit-Lock Hit
            if side == "BUY" and l <= sl:
                exit_price = min(op, sl)
                reason = "STOP_LOSS"
            elif side == "SELL" and h >= sl:
                exit_price = max(op, sl)
                reason = "STOP_LOSS"

            # 2. Fix #6: Hybrid Multi-Factor Structural Reversal Exit (replaces ADX < 18 & naked 5m cross)
            elif pos["bars_held"] >= 3 and side == "BUY" and (
                (sig["cross_dn"] or sig["c10_diff"] < 0 or sig["cross_dn_5m"])
                and c < sig["vwap"]
                and c < sig["e30_5m"]
            ):
                exit_price = c
                reason = "STRUCTURAL_REVERSAL (Contra+VWAP+EMA30)"
            elif pos["bars_held"] >= 3 and side == "SELL" and (
                (sig["cross_up"] or sig["c10_diff"] > 0 or sig["cross_up_5m"])
                and c > sig["vwap"]
                and c > sig["e30_5m"]
            ):
                exit_price = c
                reason = "STRUCTURAL_REVERSAL (Contra+VWAP+EMA30)"

            # 3. 3:05 PM Force Exit
            elif cur_time >= FORCE_EXIT_TIME:
                exit_price = c
                reason = "3_PM_EXIT"

            if exit_price is not None:
                # Trigger Square Off via Order Manager
                exit_res = broker.place_exit_order(sym, side, pos["qty"], estimated_price=exit_price, reason=reason)
                actual_exit = exit_res.get("fill_price") or exit_price
                pnl_pct = (actual_exit - entry) / entry * 100.0 if side == "BUY" else (entry - actual_exit) / entry * 100.0
                pnl_rs = (pnl_pct / 100.0) * entry * pos["qty"]

                symbols_to_close.append({
                    "symbol": sym,
                    "side": side,
                    "entry_price": entry,
                    "sl_price": sl,
                    "exit_price": round(actual_exit, 2),
                    "entry_ts": pos["entry_ts"],
                    "exit_ts": str(t),
                    "qty": pos["qty"],
                    "bars_held": pos.get("bars_held", 0),
                    "peak_price": pos.get("peak_price", entry),
                    "best_mfe_pct": pos.get("best_mfe_pct", 0.0),
                    "close_reason": reason,
                    "net_pnl_rs": round(pnl_rs, 2),
                })
            else:
                pnl_pct = (c - entry) / entry * 100.0 if side == "BUY" else (entry - c) / entry * 100.0
                pnl_rs = (pnl_pct / 100.0) * entry * pos["qty"]
                pos["current_ltp"] = c
                pos["unrealized_pnl_rs"] = round(pnl_rs, 2)
                pos["last_bar_ts"] = str(t)

        except Exception as e:
            logger.error(f"Error checking position for {sym}: {e}")

    # Remove closed positions from active dict
    for closed_item in symbols_to_close:
        sym = closed_item["symbol"]
        active_positions.pop(sym, None)
        state.setdefault("closed_today", []).append(closed_item)

    # =========================================================================
    # PHASE 2: OFFENSE (Scan Universe for New Entries)
    # =========================================================================
    available_slots = MAX_POSITIONS - len(active_positions)
    can_enter = (cur_time < ENTRY_CUTOFF_TIME) and (available_slots > 0)
    traded_today = {item["symbol"] for item in state.get("closed_today", [])} | set(active_positions.keys())

    if can_enter:
        universe = load_universe_symbols()
        candidates = []

        for sym in universe:
            # Fix #4b: Per-Stock Cooldown (skip if already open or traded earlier today)
            if sym in traded_today:
                continue

            pq_path = os.path.join(master_dir, f"{sym}_NS.parquet")
            if not os.path.exists(pq_path):
                pq_path = os.path.join(master_dir, f"{sym}.parquet")
            if not os.path.exists(pq_path):
                continue

            try:
                df_5m = pd.read_parquet(pq_path)
                sig = evaluate_latest_10m_signals(df_5m)
                if not sig:
                    continue

                # Fix #5a (Min ATR >= 0.14%) & Fix #4a (ADX >= 20 and 2-bar smoothed rising ADX)
                if sig["adx"] < 20.0 or sig["atr_pct"] < 0.14:
                    continue
                if sig["adx"] < sig["prev_adx"] and sig["adx"] < sig["prev2_adx"]:
                    continue

                # Fix #2: Causal VWAP Side & Directional Movement confirmation
                if sig["cross_up"] and sig["ltp"] > sig["vwap"] and sig["pdi"] > sig["mdi"] and sig["e10_5m"] >= sig["e30_5m"]:
                    candidates.append({
                        "symbol": sym,
                        "side": "BUY",
                        "price": sig["ltp"],
                        "ts": sig["ts"],
                        "vol": sig["volume"],
                    })
                elif sig["cross_dn"] and sig["ltp"] < sig["vwap"] and sig["mdi"] > sig["pdi"] and sig["e10_5m"] <= sig["e30_5m"]:
                    candidates.append({
                        "symbol": sym,
                        "side": "SELL",
                        "price": sig["ltp"],
                        "ts": sig["ts"],
                        "vol": sig["volume"],
                    })
            except Exception:
                continue

        # Sort candidates by volume (liquidity priority)
        candidates.sort(key=lambda x: x["vol"], reverse=True)
        admitted = candidates[:available_slots]

        for cand in admitted:
            sym = cand["symbol"]
            side = cand["side"]
            price = cand["price"]
            ts = cand["ts"]

            qty = int(CAPITAL_PER_TRADE / price) if price > 0 else DEFAULT_QTY
            qty = max(qty, 1)

            sl_price = price * (1.0 - SL_PCT) if side == "BUY" else price * (1.0 + SL_PCT)

            pos_data = {
                "symbol": sym,
                "side": side,
                "entry_price": price,
                "sl_price": round(sl_price, 2),
                "entry_ts": str(ts),
                "qty": qty,
                "bars_held": 0,
                "peak_price": price,
                "best_mfe_pct": 0.0,
                "excel_row": 0,
            }

            # Place broker entry order
            order_res = broker.place_entry_order(sym, side, qty, estimated_price=price)

            if order_res.get("status") in ["FILLED", "SUBMITTED"]:
                actual_entry = order_res.get("fill_price") or price
                pos_data["entry_price"] = round(actual_entry, 2)
                pos_data["sl_price"] = round(
                    actual_entry * (1.0 - SL_PCT) if side == "BUY" else actual_entry * (1.0 + SL_PCT), 2
                )
                active_positions[sym] = pos_data
                logger.info(f"[NEW ENTRY ADMITTED] {side} {sym} @ {actual_entry:.2f} (Qty: {qty}, SL: {pos_data['sl_price']:.2f})")
            else:
                logger.warning(f"[ENTRY REJECTED] {sym}: {order_res.get('message')}")

    else:
        if cur_time >= ENTRY_CUTOFF_TIME:
            logger.info(f"Entry cutoff reached ({cur_time.strftime('%H:%M')}). No new entries permitted.")
        elif available_slots <= 0:
            logger.info(f"Max positions capacity reached ({MAX_POSITIONS}/{MAX_POSITIONS}).")

    # =========================================================================
    # PHASE 3: PERSISTENCE & EXCEL SYNCHRONIZATION
    # =========================================================================
    writer.sync_to_excel(state)
    save_state(state)
    logger.info(f"--- COCKPIT_V6 CYCLE COMPLETE | Active Positions: {len(active_positions)} ---")



# ── CLI ENTRY POINT ──────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="COCKPIT_V6 Engine Orchestrator")
    parser.add_argument("--live", action="store_true", help="Enable live order placement via FYERS API (default is paper mode)")
    parser.add_argument("--force", action="store_true", help="Bypass market hours check (for testing)")
    parser.add_argument("--status", action="store_true", help="Print current active positions and exit")
    args = parser.parse_args()

    if args.status:
        st = load_state()
        pos = st.get("positions", {})
        print(f"\nCOCKPIT_V6 ACTIVE POSITIONS ({len(pos)}):")
        print("=" * 70)
        for sym, p in pos.items():
            print(f"  {sym:<12} | {p['side']:<4} | Entry: {p['entry_price']:7.2f} | SL: {p['sl_price']:7.2f} | Qty: {p['qty']}")
        print("=" * 70)
        return

    run_cockpit_cycle(is_live=args.live, force=args.force)


if __name__ == "__main__":
    main()
