"""
sim/server.py -- Runs the paper-trading simulator & Net-Profit Maximizer web terminal.

  python sim/server.py --mode live                 # today, as the market moves (09:15-15:30 IST)
  python sim/server.py --mode replay               # replay the last session
  python sim/server.py --mode replay --date 2026-10-30 --strategy v6_profit_max

Then open http://localhost:8765
No broker keys. No money. Orders are never sent anywhere.
"""
import argparse
import csv
import json
import os
import pickle
import sys
import threading
import time
import webbrowser
from datetime import datetime, timedelta, time as dtime
from typing import Optional

import pandas as pd

SIM_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SIM_DIR)

import config as C
import datafeed
from simulator import Simulator, BAR

from fastapi import FastAPI, Request
from fastapi.responses import FileResponse, JSONResponse
import uvicorn

OUT_DIR = os.path.join(SIM_DIR, "journal")
CACHE_FILE = os.path.join(SIM_DIR, "cache", "nse_5m_60d.pkl")
os.makedirs(OUT_DIR, exist_ok=True)

app = FastAPI()
SIM: Simulator = None
UNIVERSE = None
CACHED_DATA = None
CACHED_IDX = None
LIVE_THREAD_ACTIVE = False


@app.get("/")
def index():
    return FileResponse(os.path.join(SIM_DIR, "static", "index.html"))


@app.get("/api/state")
def state():
    return JSONResponse(SIM.snapshot() if SIM else {"status": "STARTING"})


@app.get("/api/chart/{symbol}")
def chart(symbol: str):
    return JSONResponse(SIM.chart(symbol) if SIM else {})


@app.get("/api/backtest_report")
def backtest_report():
    rpt_path = os.path.join(OUT_DIR, "backtest_12week_report.json")
    if os.path.exists(rpt_path):
        with open(rpt_path, "r", encoding="utf-8") as f:
            return JSONResponse(json.load(f))
    return JSONResponse({"error": "Report not generated yet"})


@app.post("/api/control")
async def control(req: Request):
    """Switch strategy preset, replay date, capital, or live/replay mode on the fly from the browser UI."""
    global SIM, LIVE_THREAD_ACTIVE
    body = await req.json()
    new_strategy = body.get("strategy", SIM.strategy if SIM else C.DEFAULT_STRATEGY)
    new_mode = body.get("mode", SIM.mode if SIM else "replay")
    new_date_str = body.get("date")
    new_capital = float(body.get("capital", SIM.starting_capital if SIM else C.STARTING_CAPITAL))
    speed = float(body.get("speed", 0.0))

    if new_mode == "live":
        LIVE_THREAD_ACTIVE = True
        threading.Thread(target=lambda: run_live(new_strategy, new_capital), daemon=True).start()
        return JSONResponse({"ok": True, "mode": "live", "strategy": new_strategy, "capital": new_capital})
    else:
        LIVE_THREAD_ACTIVE = False
        dt = datetime.strptime(new_date_str, "%Y-%m-%d") if new_date_str else (SIM.session_date.to_pydatetime() if SIM else None)
        threading.Thread(target=lambda: run_replay(dt, speed, new_strategy, new_capital), daemon=True).start()
        return JSONResponse({"ok": True, "mode": "replay", "date": new_date_str, "strategy": new_strategy, "capital": new_capital})


def write_journal(sim: Simulator):
    day = str(sim.session_date.date())
    snap = sim.snapshot()
    prefix = f"{sim.mode}_{day}_{sim.strategy}"
    with open(os.path.join(OUT_DIR, f"{prefix}_summary.json"), "w", encoding="utf-8") as f:
        json.dump({"kpi": snap["kpi"], "strategy": sim.strategy, "config": snap["config"], "closed": snap["closed"]}, f, indent=1)
    # Also write the standard filename for backward compatibility
    with open(os.path.join(OUT_DIR, f"{sim.mode}_{day}_summary.json"), "w", encoding="utf-8") as f:
        json.dump({"kpi": snap["kpi"], "strategy": sim.strategy, "config": snap["config"], "closed": snap["closed"]}, f, indent=1)
    if sim.closed:
        keys = [k for k in sim.closed[0] if k != "cost_detail"]
        with open(os.path.join(OUT_DIR, f"{prefix}_trades.csv"), "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=keys, extrasaction="ignore")
            w.writeheader(); w.writerows(sim.closed)
    with open(os.path.join(OUT_DIR, f"{prefix}_signals.csv"), "w", newline="", encoding="utf-8") as f:
        if sim.signals:
            w = csv.DictWriter(f, fieldnames=list(sim.signals[0]))
            w.writeheader(); w.writerows(sim.signals)


def fetch_all(universe, period="5d", start=None, end=None, use_cache=True):
    global CACHED_DATA, CACHED_IDX
    if use_cache and CACHED_DATA is not None and CACHED_IDX is not None:
        return dict(CACHED_DATA), dict(CACHED_IDX)
    if use_cache and os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "rb") as f:
                payload = pickle.load(f)
            CACHED_DATA = payload["raw"]
            CACHED_IDX = payload["idx"]
            return dict(CACHED_DATA), dict(CACHED_IDX)
        except Exception:
            pass
    ys = [u["yahoo"] for u in universe] + list(C.INDEX_TICKERS.values())
    raw = datafeed.fetch(ys, period=period, start=start, end=end)
    idx = {name: raw.pop(t) for name, t in C.INDEX_TICKERS.items() if t in raw}
    CACHED_DATA, CACHED_IDX = dict(raw), dict(idx)
    return raw, idx


def ensure_sim_loaded(mode: str, date: Optional[datetime], strategy: str, capital: Optional[float] = None):
    global SIM, UNIVERSE
    if UNIVERSE is None:
        UNIVERSE = datafeed.load_universe()
    data, idx = fetch_all(UNIVERSE, period="60d", use_cache=True)
    if date is None and data:
        last_dt = max(df.index[-1] for df in data.values()).normalize()
        date = last_dt.to_pydatetime()
    cap = capital or (SIM.starting_capital if SIM else C.STARTING_CAPITAL)
    if SIM is None:
        SIM = Simulator(mode, date or datetime.now(), strategy=strategy, starting_capital=cap)
        SIM.status = "LOADING 60-DAY NSE DATASET"
        SIM.load(data, UNIVERSE, idx)
    else:
        SIM.mode = mode
        SIM.reset_state(session_date=date or datetime.now(), strategy=strategy, capital=cap)
        if not SIM.syms:
            SIM.load(data, UNIVERSE, idx)
    return data, idx, date


def run_replay(date: Optional[datetime], speed: float, strategy: str = C.DEFAULT_STRATEGY, capital: Optional[float] = None):
    global SIM
    data, idx, date = ensure_sim_loaded("replay", date, strategy, capital)
    SIM.data_note = f"Yahoo Finance 5m bars (60-day history) | {len(SIM.syms)}/{len(UNIVERSE)} symbols ready"
    times = SIM.session_bar_times()
    if not times:
        SIM.status = f"NO DATA FOR {SIM.session_date.date()}"
        return
    SIM.status = f"REPLAYING ({C.STRATEGY_LABELS.get(strategy, strategy)})"
    for t in times:
        SIM.step(t)
        if speed > 0:
            time.sleep(speed)
    SIM.compute_session_comparison()
    SIM.status = "REPLAY COMPLETE"
    write_journal(SIM)
    k = SIM.snapshot()["kpi"]
    print(f"\nREPLAY {SIM.session_date.date()} [{strategy}] done | cap Rs {SIM.starting_capital:,.0f} | trades {k['trades']} | wins {k['wins']} | "
          f"gross Rs {k['gross']:,.2f} | fees Rs {k['fees']:,.2f} | net P&L Rs {k['realized']:,.2f}")


def run_live(strategy: str = C.DEFAULT_STRATEGY, capital: Optional[float] = None):
    global SIM, LIVE_THREAD_ACTIVE, UNIVERSE
    LIVE_THREAD_ACTIVE = True
    today = datetime.now()
    data, idx, _ = ensure_sim_loaded("live", today, strategy, capital)
    while LIVE_THREAD_ACTIVE and SIM.mode == "live":
        now = datetime.now()
        if now.weekday() >= 5:
            SIM.status = "MARKET CLOSED (Weekend) - Switch to Replay Mode above to inspect any past day"
            time.sleep(60); continue
        if now.time() < dtime(9, 14):
            SIM.status = f"WAITING FOR NSE MARKET OPEN (09:15 IST) - Current time {now:%H:%M:%S}"
            time.sleep(15); continue
        try:
            ys = [u["yahoo"] for u in UNIVERSE] + list(C.INDEX_TICKERS.values())
            fresh = datafeed.fetch(ys, period="1d")
            fidx = {name: fresh.pop(t) for name, t in C.INDEX_TICKERS.items() if t in fresh}
            for y_sym, df in fresh.items():
                old = data.get(y_sym)
                data[y_sym] = df if old is None else pd.concat([old, df]).pipe(lambda x: x[~x.index.duplicated(keep="last")]).sort_index()
            for k, v in fidx.items():
                old = idx.get(k)
                idx[k] = v if old is None else pd.concat([old, v]).pipe(lambda x: x[~x.index.duplicated(keep="last")]).sort_index()
            SIM.load(data, UNIVERSE, idx)
            SIM.data_note = f"LIVE NSE 5m Feed | {len(fresh)}/{len(UNIVERSE)} stocks refreshed at {now:%H:%M:%S}"
        except Exception as e:
            SIM.data_note = f"Live refresh warning: {e}"
        cutoff = pd.Timestamp(datetime.now()) - BAR - timedelta(seconds=C.BAR_GRACE_SECONDS)
        for t in SIM.session_bar_times():
            if t <= cutoff:
                SIM.step(t)
        SIM.refresh_ltps()
        if datetime.now().time() > dtime(15, 31):
            SIM.compute_session_comparison()
            SIM.status = "SESSION CLOSED (15:30 IST) - Final Journal Saved"
            write_journal(SIM)
            time.sleep(120); continue
        SIM.status = f"LIVE PAPER TRADING ({C.STRATEGY_LABELS.get(SIM.strategy, SIM.strategy)})"
        write_journal(SIM)
        time.sleep(C.LIVE_POLL_SECONDS)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--mode", choices=["live", "replay"], default="live")
    ap.add_argument("--date", help="replay date YYYY-MM-DD (default: last session)")
    ap.add_argument("--strategy", choices=list(C.STRATEGY_LABELS.keys()), default=C.DEFAULT_STRATEGY)
    ap.add_argument("--speed", type=float, default=0.0, help="replay seconds per 5m bar (0 = instant)")
    ap.add_argument("--no-browser", action="store_true")
    ap.add_argument("--no-server", action="store_true", help="replay only, print summary, exit")
    a = ap.parse_args()

    if a.mode == "replay":
        date = datetime.strptime(a.date, "%Y-%m-%d") if a.date else None
        target = lambda: run_replay(date, 0 if a.no_server else a.speed, a.strategy)
    else:
        target = lambda: run_live(a.strategy)

    if a.no_server:
        target(); return
    threading.Thread(target=target, daemon=True).start()
    if not a.no_browser:
        threading.Timer(2.0, lambda: webbrowser.open(f"http://localhost:{C.SERVER_PORT}")).start()
    print(f"Paper-trading terminal: http://localhost:{C.SERVER_PORT}  (Ctrl+C to stop)")
    uvicorn.run(app, host="127.0.0.1", port=C.SERVER_PORT, log_level="warning")


if __name__ == "__main__":
    main()
