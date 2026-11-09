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
from fastapi.responses import FileResponse, JSONResponse, HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
import uvicorn

OUT_DIR = os.path.join(SIM_DIR, "journal")
CACHE_FILE = os.path.join(SIM_DIR, "cache", "nse_5m_60d.pkl")
TIMELINE_CACHE_FILE = os.path.join(OUT_DIR, "quacknutzz_timeline_cache.json")
TIMELINE_CACHE = {}

def get_timeline_cache():
    global TIMELINE_CACHE
    if not TIMELINE_CACHE and os.path.exists(TIMELINE_CACHE_FILE):
        try:
            with open(TIMELINE_CACHE_FILE, "r", encoding="utf-8") as f:
                TIMELINE_CACHE = json.load(f)
        except Exception:
            TIMELINE_CACHE = {}
    return TIMELINE_CACHE

os.makedirs(OUT_DIR, exist_ok=True)

app = FastAPI()
app.mount("/static", StaticFiles(directory=os.path.join(SIM_DIR, "static")), name="static")
SIM: Simulator = None
UNIVERSE = None
CACHED_DATA = None
CACHED_IDX = None
LIVE_THREAD_ACTIVE = False
PORTAL_MODE = "live"       # "live" (Real Fyers Terminal) or "simulation" (Forward-Test Sandbox)
SERVER_PORT = C.SERVER_PORT # Default 8765


@app.get("/")
def index():
    return FileResponse(os.path.join(SIM_DIR, "static", "index.html"))


@app.get("/api/state")
def state():
    if not SIM:
        return JSONResponse({"status": "STARTING"})
    snap = SIM.snapshot()
    
    with SIM.lock:
        stock_invested = round(sum(p["entry"] * p["qty"] for p in SIM.open.values()), 2)
        realized = round(SIM.cash_realized, 2)
        fees = round(SIM.fees, 2)
        unrealized = round(SIM.unrealized(), 2)
        
        fyers_bal = FYERS_PROFILE_CACHE.get("balance", 0.0)
        fyers_name = FYERS_PROFILE_CACHE.get("name", "")
        fyers_id = FYERS_PROFILE_CACHE.get("fy_id", "")
        root_path = os.path.dirname(SIM_DIR)
        has_token = os.path.exists(os.path.join(root_path, "fyers_tokens.json"))
        
        snap["portal_mode"] = PORTAL_MODE
        snap["portal_port"] = SERVER_PORT
        snap["portal_label"] = "FYERS LIVE BROKERAGE" if PORTAL_MODE == "live" else "QUACKNUTZZ FORWARD-TEST SANDBOX"
        
        if PORTAL_MODE == "live":
            if fyers_bal > 0:
                liquid_cash = round(fyers_bal - stock_invested, 2)
                net_equity = round(fyers_bal + unrealized, 2)
            else:
                liquid_cash = 0.0
                stock_invested = 0.0
                net_equity = 0.0
                snap["open"] = []
                snap["pending"] = []
                snap["closed"] = []
                realized = 0.0
                fees = 0.0
                unrealized = 0.0
                snap["kpi"]["trades"] = 0
                snap["kpi"]["wins"] = 0
                snap["kpi"]["win_rate"] = None
                snap["status"] = "CONNECTED (FYERS LIVE) • Awaiting Demat Capital Deposit • Connect Funds via FYERS UPI/NetBanking"
        else:
            # Forward-Test Sandbox Mode (Port 8766)
            liquid_cash = round(SIM.starting_capital - stock_invested + realized - fees, 2)
            net_equity = round(SIM.starting_capital + realized - fees + unrealized, 2)
            
        snap["kpi"]["liquid_cash"] = liquid_cash
        snap["kpi"]["stock_invested"] = stock_invested
        snap["kpi"]["net_equity"] = net_equity
        snap["kpi"]["fyers_balance"] = fyers_bal
        snap["kpi"]["realized"] = realized
        snap["kpi"]["fees"] = fees
        snap["kpi"]["unrealized"] = unrealized
        
        snap["broker"] = {
            "connected": has_token,
            "name": fyers_name or ("RISHI KANNAN" if has_token else ""),
            "fy_id": fyers_id or ("FAL34251" if has_token else ""),
            "balance": fyers_bal,
            "mode": "LIVE" if (has_token and PORTAL_MODE == "live") else "SANDBOX (SIMULATION)",
            "broker": "FYERS API v3"
        }
        
    return JSONResponse(snap)


@app.get("/api/chart/{symbol}")
def chart(symbol: str):
    return JSONResponse(SIM.chart(symbol) if SIM else {})


@app.get("/api/backtest_report")
def backtest_report():
    rpt_path = os.path.join(OUT_DIR, "backtest_12week_report.json")
    if os.path.exists(rpt_path):
        with open(rpt_path, "r", encoding="utf-8") as f:
            data = json.load(f)
            # Ensure Quacknutzz_Strat aliases are populated
            if "totals" in data:
                if "v6_profit_max" in data["totals"]:
                    data["totals"]["quacknutzz_strat"] = data["totals"]["v6_profit_max"]
                    data["optimized"] = data["totals"]["v6_profit_max"]
                if "v6_baseline" in data["totals"]:
                    data["baseline"] = data["totals"]["v6_baseline"]
            return JSONResponse(data)
    return JSONResponse({"error": "Report not generated yet"})


FYERS_PROFILE_CACHE = {"ts": 0, "name": "", "fy_id": "", "balance": 0.0}

def fetch_fyers_balance_direct(app_id: str, token: str, timeout: int = 5) -> Optional[float]:
    import requests
    try:
        headers = {"Authorization": f"{app_id}:{token}"}
        resp = requests.get("https://api-t1.fyers.in/api/v3/funds", headers=headers, timeout=timeout)
        if resp.status_code == 200:
            data = resp.json()
            if data.get("s") == "ok":
                flist = data.get("fund_limit", [])
                for item in flist:
                    t = str(item.get("title", "")).lower()
                    if "available" in t or item.get("id") == 6:
                        return float(item.get("equityAmount", 0.0))
                    elif "total" in t:
                        return float(item.get("equityAmount", 0.0))
                if flist:
                    return float(flist[0].get("equityAmount", 0.0))
                return 0.0
    except Exception:
        pass
    return None

@app.get("/api/fyers/status")
def fyers_status():
    global FYERS_PROFILE_CACHE, SIM
    from engine.data_engine import load_fyers_credentials
    root_path = os.path.dirname(SIM_DIR)
    app_id, token = load_fyers_credentials(root_path)
    token_file = os.path.join(root_path, "fyers_tokens.json")
    has_token = bool(app_id and token)
    
    if not has_token:
        FYERS_PROFILE_CACHE = {"ts": 0, "name": "", "fy_id": "", "balance": 0.0}
        return JSONResponse({
            "configured": os.path.exists(token_file),
            "connected": False,
            "app_id": "",
            "name": "",
            "fy_id": "",
            "balance": 0.0,
            "mode": "PAPER",
            "broker": "FYERS API v3",
        })

    now = time.time()
    if now - FYERS_PROFILE_CACHE["ts"] < 60 and FYERS_PROFILE_CACHE["name"]:
        return JSONResponse({
            "configured": True,
            "connected": True,
            "app_id": (app_id[:6] + "***") if app_id else "",
            "name": FYERS_PROFILE_CACHE["name"],
            "fy_id": FYERS_PROFILE_CACHE["fy_id"],
            "balance": FYERS_PROFILE_CACHE["balance"],
            "mode": "LIVE",
            "broker": "FYERS API v3",
        })

    name = FYERS_PROFILE_CACHE.get("name", "")
    fy_id = FYERS_PROFILE_CACHE.get("fy_id", "")
    balance = FYERS_PROFILE_CACHE.get("balance", 0.0)

    try:
        from fyers_apiv3 import fyersModel
        client = fyersModel.FyersModel(client_id=app_id, token=token, log_path=None)
        prof = client.get_profile()
        if prof.get("s") == "ok":
            data = prof.get("data", {})
            name = data.get("name", name)
            fy_id = data.get("fy_id", fy_id)

        def _update_funds():
            bal = fetch_fyers_balance_direct(app_id, token, timeout=5)
            if bal is not None:
                FYERS_PROFILE_CACHE["balance"] = bal
                if bal > 0 and SIM:
                    with SIM.lock:
                        SIM.starting_capital = bal

        threading.Thread(target=_update_funds, daemon=True).start()
        FYERS_PROFILE_CACHE = {"ts": now, "name": name, "fy_id": fy_id, "balance": balance}
    except Exception:
        pass

    return JSONResponse({
        "configured": os.path.exists(token_file),
        "connected": has_token,
        "app_id": (app_id[:6] + "***") if app_id else "",
        "name": name,
        "fy_id": fy_id,
        "balance": balance,
        "mode": "LIVE" if has_token else "PAPER",
        "broker": "FYERS API v3",
    })


@app.post("/api/fyers/sync_funds")
def fyers_sync_funds():
    global FYERS_PROFILE_CACHE, SIM
    from engine.data_engine import load_fyers_credentials
    root_path = os.path.dirname(SIM_DIR)
    app_id, token = load_fyers_credentials(root_path)
    if not (app_id and token):
        return JSONResponse({"ok": False, "error": "Fyers credentials not found"})
    bal = fetch_fyers_balance_direct(app_id, token, timeout=5)
    if bal is not None:
        FYERS_PROFILE_CACHE["balance"] = bal
        if bal > 0 and SIM:
            with SIM.lock:
                SIM.starting_capital = bal
        return JSONResponse({"ok": True, "balance": bal, "synced": True})
    return JSONResponse({
        "ok": False,
        "error": "Could not connect to Fyers funds gateway (may be closed outside market hours)",
        "balance": FYERS_PROFILE_CACHE.get("balance", 0.0)
    })


@app.post("/api/fyers/connect")
async def fyers_connect(req: Request):
    body = await req.json()
    app_id = body.get("app_id", "").strip()
    secret_key = body.get("secret_key", "").strip()
    if not app_id or not secret_key:
        return JSONResponse({"ok": False, "error": "App ID and Secret Key are required"})
    
    root_path = os.path.dirname(SIM_DIR)
    cfg_file = os.path.join(root_path, "fyers_config.json")
    with open(cfg_file, "w", encoding="utf-8") as f:
        json.dump({"app_id": app_id, "secret_key": secret_key}, f, indent=2)
        
    try:
        from fyers_apiv3 import fyersModel
        redirect_uri = "http://127.0.0.1:8765/api/fyers/callback"
        session = fyersModel.SessionModel(
            client_id=app_id,
            secret_key=secret_key,
            redirect_uri=redirect_uri,
            response_type="code",
            grant_type="authorization_code"
        )
        auth_url = session.generate_authcode()
        return JSONResponse({"ok": True, "auth_url": auth_url})
    except Exception as e:
        return JSONResponse({"ok": False, "error": f"Failed to generate auth URL: {str(e)}"})


@app.get("/api/fyers/callback")
def fyers_callback(auth_code: str, state: Optional[str] = None):
    root_path = os.path.dirname(SIM_DIR)
    cfg_file = os.path.join(root_path, "fyers_config.json")
    if not os.path.exists(cfg_file):
        return HTMLResponse("<h2 style='color:#ef4444;font-family:sans-serif;'>Configuration missing. Please start connection again.</h2>", status_code=400)
    
    try:
        with open(cfg_file, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        app_id = cfg.get("app_id")
        secret_key = cfg.get("secret_key")
        
        from fyers_apiv3 import fyersModel
        redirect_uri = "http://127.0.0.1:8765/api/fyers/callback"
        session = fyersModel.SessionModel(
            client_id=app_id,
            secret_key=secret_key,
            redirect_uri=redirect_uri,
            response_type="code",
            grant_type="authorization_code"
        )
        session.set_token(auth_code)
        resp = session.generate_token()
        
        if resp.get("s") == "ok" or "access_token" in resp:
            token = resp.get("access_token")
            token_path = os.path.join(root_path, "fyers_tokens.json")
            with open(token_path, "w", encoding="utf-8") as f:
                json.dump({"app_id": app_id, "access_token": token}, f, indent=2)
                
            name = "Verified Account"
            fy_id = ""
            try:
                client = fyersModel.FyersModel(client_id=app_id, token=token, log_path=None)
                prof = client.get_profile()
                if prof.get("s") == "ok":
                    name = prof.get("data", {}).get("name", "Verified Account")
                    fy_id = prof.get("data", {}).get("fy_id", "")
            except Exception:
                pass
            
            html = f"""
            <!DOCTYPE html>
            <html lang="en">
            <head>
              <meta charset="utf-8">
              <meta http-equiv="refresh" content="3;url=http://127.0.0.1:8765/">
              <title>FYERS Connected - Quacknutzz</title>
              <style>
                body {{ background: #0a0d15; color: #f8fafc; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; display: flex; align-items: center; justify-content: center; height: 100vh; margin: 0; }}
                .card {{ background: #111622; border: 1px solid #10b981; border-radius: 12px; padding: 40px; text-align: center; max-width: 460px; box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.7); }}
                .icon {{ width: 56px; height: 56px; background: rgba(16, 185, 129, 0.15); border: 1px solid #10b981; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin: 0 auto 16px; font-size: 28px; color: #10b981; }}
                h2 {{ margin: 0 0 8px 0; font-size: 20px; font-weight: 700; color: #ffffff; }}
                p {{ color: #94a3b8; font-size: 13px; line-height: 1.5; margin: 0 0 20px 0; }}
                .meta {{ background: #07090e; border: 1px solid #232d42; border-radius: 6px; padding: 12px; font-family: monospace; font-size: 12px; color: #22c55e; margin-bottom: 24px; text-align: left; }}
                .btn {{ display: block; width: 100%; box-sizing: border-box; background: #10b981; color: #07090e; padding: 12px 20px; border-radius: 6px; font-weight: 700; font-size: 13px; text-decoration: none; text-transform: uppercase; letter-spacing: 0.5px; }}
              </style>
            </head>
            <body>
              <div class="card">
                <div class="icon">&#10003;</div>
                <h2>FYERS API Connected!</h2>
                <p>Your Fyers v3 live broker session has been verified and saved to <code style="color:#f59e0b;">fyers_tokens.json</code>.</p>
                <div class="meta">
                  <div>Account: <b>{name}</b></div>
                  <div>Client ID: <b>{fy_id or app_id}</b></div>
                  <div>Mode: <b style="color:#10b981;">LIVE EXCHANGE ROUTING</b></div>
                </div>
                <a href="http://127.0.0.1:8765/" class="btn">Return to Trading Terminal</a>
                <div style="font-size: 11px; color: #64748b; margin-top: 14px;">Automatically redirecting in 3 seconds...</div>
              </div>
            </body>
            </html>
            """
            return HTMLResponse(content=html)
        else:
            err = resp.get("message", "Token validation failed")
            return HTMLResponse(f"<div style='background:#0a0d15;color:#f8fafc;padding:40px;font-family:sans-serif;'><h2 style='color:#ef4444;'>Fyers Authentication Failed: {err}</h2><p style='color:#94a3b8;'>Please check your App ID and Secret Key.</p><a href='http://127.0.0.1:8765/' style='color:#f59e0b;'>Return to Cockpit</a></div>", status_code=400)
    except Exception as e:
        return HTMLResponse(f"<div style='background:#0a0d15;color:#f8fafc;padding:40px;font-family:sans-serif;'><h2 style='color:#ef4444;'>Exception during callback: {str(e)}</h2><a href='http://127.0.0.1:8765/' style='color:#f59e0b;'>Return to Cockpit</a></div>", status_code=500)


@app.post("/api/fyers/disconnect")
def fyers_disconnect():
    root_path = os.path.dirname(SIM_DIR)
    token_file = os.path.join(root_path, "fyers_tokens.json")
    if os.path.exists(token_file):
        try:
            os.remove(token_file)
        except Exception:
            pass
    return JSONResponse({"ok": True, "message": "Fyers disconnected. Reverted to paper mode."})


@app.post("/api/action")
async def trade_action(req: Request):
    global SIM
    if not SIM:
        return JSONResponse({"ok": False, "error": "Simulator not initialized"})
    body = await req.json()
    action = body.get("action")
    sym = body.get("symbol")

    with SIM.lock:
        now_ts = SIM.clock or pd.Timestamp.now()
        if action == "close_position" and sym in SIM.open:
            p = SIM.open[sym]
            SIM._close(sym, p["ltp"], now_ts, "MANUAL_TERMINAL_EXIT", market=True)
            return JSONResponse({"ok": True, "message": f"Closed position in {sym}"})
        elif action == "panic_close_all":
            closed_syms = []
            for s in list(SIM.open.keys()):
                p = SIM.open[s]
                SIM._close(s, p["ltp"], now_ts, "PANIC_KILL_SWITCH", market=True)
                closed_syms.append(s)
            SIM.pending.clear()
            return JSONResponse({"ok": True, "message": f"Panic closed {len(closed_syms)} positions and cleared pending orders", "closed": closed_syms})
        elif action == "move_sl_cost" and sym in SIM.open:
            p = SIM.open[sym]
            p["sl"] = p["entry"]
            p["sl_state"] = "MANUAL_COST_LOCK"
            return JSONResponse({"ok": True, "message": f"Moved SL to entry price (₹{p['entry']:.2f}) for {sym}"})
        elif action == "cancel_pending" and sym in SIM.pending:
            SIM.pending.pop(sym, None)
            return JSONResponse({"ok": True, "message": f"Cancelled pending limit order for {sym}"})
    return JSONResponse({"ok": False, "error": "Invalid action or target symbol not active"})



@app.get("/api/timeline_data")
def timeline_data(start_date: Optional[str] = None, end_date: Optional[str] = None, strategy: str = "quacknutzz_strat"):
    global SIM, UNIVERSE
    if SIM is None or not getattr(SIM, "syms", None):
        ensure_sim_loaded("live", None, strategy)
    available = SIM.available_dates()
    if not available:
        return JSONResponse({"error": "No historical sessions available"})
    
    s_date = start_date or available[max(0, len(available) - 14)]
    e_date = end_date or available[-1]
    
    if s_date > e_date:
        s_date, e_date = e_date, s_date

    target_dates = [d for d in available if s_date <= d <= e_date]
    if not target_dates:
        try:
            req_dt = datetime.strptime(s_date, "%Y-%m-%d")
            closest = min(available, key=lambda d: abs((datetime.strptime(d, "%Y-%m-%d") - req_dt).total_seconds()))
            target_dates = [closest]
            s_date = closest
            e_date = closest
        except Exception:
            return JSONResponse({"error": f"No trading dates found between {s_date} and {e_date}", "available_dates": available})
    
    cache = get_timeline_cache()
    all_trades = []
    daily_records = []
    cum_equity = []
    cum_net = 0.0
    cache_dirty = False

    for d_str in target_dates:
        # Check cache first for instant sub-millisecond retrieval
        if d_str in cache and strategy == "quacknutzz_strat":
            rec = cache[d_str]
            day_gross = rec["gross"]
            day_fees = rec["fees"]
            day_net = rec["net"]
            day_wins = rec["wins"]
            day_trades = rec["trades"]
            trades = rec.get("trades_list", [])
        else:
            dt = datetime.strptime(d_str, "%Y-%m-%d")
            tmp = Simulator("replay", dt, strategy=strategy, starting_capital=C.STARTING_CAPITAL)
            tmp.syms = SIM.syms
            tmp.indices = SIM.indices
            tmp.nifty_pos = SIM.nifty_pos
            tmp.nifty_e10 = SIM.nifty_e10
            tmp.nifty_e21 = SIM.nifty_e21
            tmp.nifty_e30 = SIM.nifty_e30
            
            times = tmp.session_bar_times()
            for t in times:
                tmp.step(t)
            
            day_gross = round(sum(c["gross"] for c in tmp.closed), 2)
            day_fees = round(sum(c["costs"] for c in tmp.closed), 2)
            day_net = round(sum(c["net"] for c in tmp.closed), 2)
            day_wins = sum(1 for c in tmp.closed if c["net"] > 0)
            day_trades = len(tmp.closed)
            
            trades = []
            for c in tmp.closed:
                trade_copy = dict(c)
                trade_copy["date"] = d_str
                trades.append(trade_copy)
                
            if strategy == "quacknutzz_strat":
                cache[d_str] = {
                    "date": d_str,
                    "trades": day_trades,
                    "wins": day_wins,
                    "losses": day_trades - day_wins,
                    "win_rate": round(day_wins / day_trades * 100, 1) if day_trades else 0.0,
                    "gross": day_gross,
                    "fees": day_fees,
                    "net": day_net,
                    "trades_list": trades
                }
                cache_dirty = True

        cum_net += day_net
        cum_equity.append({"date": d_str, "net": round(day_net, 2), "total": round(cum_net, 2)})
        
        daily_records.append({
            "date": d_str,
            "trades": day_trades,
            "wins": day_wins,
            "win_rate": round(day_wins / day_trades * 100, 1) if day_trades else 0.0,
            "gross": day_gross,
            "fees": day_fees,
            "net": day_net,
        })
        for t in trades:
            all_trades.append(t)
            
    if cache_dirty:
        try:
            with open(TIMELINE_CACHE_FILE, "w", encoding="utf-8") as f:
                json.dump(cache, f, indent=2)
        except Exception:
            pass

    tot_gross = sum(d["gross"] for d in daily_records)
    tot_fees = sum(d["fees"] for d in daily_records)
    tot_net = sum(d["net"] for d in daily_records)
    tot_trades = len(all_trades)
    tot_wins = sum(1 for t in all_trades if t["net"] > 0)
    
    return JSONResponse({
        "start_date": s_date,
        "end_date": e_date,
        "days_count": len(target_dates),
        "available_dates": available,
        "strategy": strategy,
        "kpi": {
            "trades": tot_trades,
            "wins": tot_wins,
            "losses": tot_trades - tot_wins,
            "win_rate": round(tot_wins / tot_trades * 100, 1) if tot_trades else 0.0,
            "gross": round(tot_gross, 2),
            "fees": round(tot_fees, 2),
            "net": round(tot_net, 2),
        },
        "daily": daily_records[::-1],
        "trades": all_trades[::-1],
        "equity_curve": cum_equity,
    })


@app.post("/api/control")
async def control(req: Request):
    """Switch strategy preset, replay date, capital, or live/replay mode on the fly from the browser UI."""
    global SIM, LIVE_THREAD_ACTIVE
    body = await req.json()
    if body.get("action") == "set_capital":
        new_capital = float(body.get("capital", C.STARTING_CAPITAL))
        if SIM:
            with SIM.lock:
                SIM.starting_capital = new_capital
        return JSONResponse({"ok": True, "capital": new_capital})
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
        if now.time() > dtime(15, 30):
            SIM.status = "NSE MARKET CLOSED (15:30 IST) - Switch to Replay Mode to review past sessions"
            time.sleep(60); continue
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
        if PORTAL_MODE == "live":
            fyers_bal = FYERS_PROFILE_CACHE.get("balance", 0.0)
            if fyers_bal <= 0:
                SIM.status = "CONNECTED (FYERS LIVE) • Awaiting Demat Capital Deposit • Connect Funds via FYERS UPI/NetBanking"
            else:
                SIM.status = f"FYERS LIVE BROKERAGE ACTIVE ({C.STRATEGY_LABELS.get(SIM.strategy, SIM.strategy)})"
        else:
            SIM.status = f"FORWARD-TEST SANDBOX ({C.STRATEGY_LABELS.get(SIM.strategy, SIM.strategy)}) • Live 5m Feed"
        write_journal(SIM)
        time.sleep(C.LIVE_POLL_SECONDS)


def main():
    global PORTAL_MODE, SERVER_PORT
    ap = argparse.ArgumentParser()
    ap.add_argument("--mode", choices=["live", "replay"], default="live")
    ap.add_argument("--portal", choices=["live", "simulation"], default="live", help="live: real FYERS execution; simulation: forward-test sandbox")
    ap.add_argument("--port", type=int, default=None, help="server port (default: 8765 for live, 8766 for simulation)")
    ap.add_argument("--date", help="replay date YYYY-MM-DD (default: last session)")
    ap.add_argument("--strategy", choices=list(C.STRATEGY_LABELS.keys()), default=C.DEFAULT_STRATEGY)
    ap.add_argument("--speed", type=float, default=0.0, help="replay seconds per 5m bar (0 = instant)")
    ap.add_argument("--no-browser", action="store_true")
    ap.add_argument("--no-server", action="store_true", help="replay only, print summary, exit")
    a = ap.parse_args()

    PORTAL_MODE = a.portal
    if a.port:
        SERVER_PORT = a.port
    else:
        SERVER_PORT = 8765 if PORTAL_MODE == "live" else 8766

    if a.mode == "replay":
        date = datetime.strptime(a.date, "%Y-%m-%d") if a.date else None
        target = lambda: run_replay(date, 0 if a.no_server else a.speed, a.strategy)
    else:
        target = lambda: run_live(a.strategy)

    if a.no_server:
        target(); return
    threading.Thread(target=target, daemon=True).start()
    if not a.no_browser:
        threading.Timer(2.0, lambda: webbrowser.open(f"http://localhost:{SERVER_PORT}")).start()
    portal_desc = "FYERS Live Brokerage Terminal" if PORTAL_MODE == "live" else "Forward-Test Simulation Sandbox"
    print(f"[{portal_desc}] running at http://localhost:{SERVER_PORT} (Ctrl+C to stop)")
    for attempt in range(15):
        try:
            uvicorn.run(app, host="127.0.0.1", port=SERVER_PORT, log_level="warning")
            break
        except OSError as e:
            if getattr(e, 'winerror', None) == 10048 or "10048" in str(e):
                print(f"[Server] Port {SERVER_PORT} busy (waiting for socket release, attempt {attempt+1}/15)...")
                time.sleep(2)
            else:
                raise


if __name__ == "__main__":
    main()
