"""
fyers_order_manager.py -- Live Order Application & Position Management for FYERS API v3.

Handles:
1. Authentication leveraging existing fyers_tokens.json.
2. Market entry order execution with fill verification & traded price capture.
3. Square-off order execution (Stop-Loss, 5m reverse cross, 3 PM cutoff).
4. Broker position reconciliation (syncs local state with fyers.positions()).
5. Real-time available margin checks before entry.
6. Clean simulation fallback for Paper Trading.
"""

import os
import sys
import time
import json
import logging
import socket
import re
import requests
from typing import Dict, Any, List, Optional, Tuple

# Force IPv4 socket resolution to ensure all Fyers API traffic routes via the whitelisted public IPv4
try:
    _orig_getaddrinfo = socket.getaddrinfo
    def _getaddrinfo_ipv4(host, port, family=0, type=0, proto=0, flags=0):
        return _orig_getaddrinfo(host, port, socket.AF_INET, type, proto, flags)
    socket.getaddrinfo = _getaddrinfo_ipv4
except Exception:
    pass

logger = logging.getLogger("FYERS_ORDER_MGR")
if not logger.handlers:
    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(logging.Formatter("[FYERS_OM] %(asctime)s %(message)s"))
    logger.addHandler(sh)
    
    log_file_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "logs", "fyers_orders.log")
    try:
        fh = logging.FileHandler(log_file_path, encoding="utf-8")
        fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
        logger.addHandler(fh)
    except Exception:
        pass
    logger.setLevel(logging.INFO)

# Path setup
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from engine.data_engine import load_fyers_credentials

FYERS_API_BASE = "https://api-t1.fyers.in/api/v3"


class FyersOrderManager:
    """
    Production-grade order application and position management bridge for FYERS API v3.
    """

    def __init__(self, root: Optional[str] = None, is_live: bool = False):
        self.root = root or ROOT_DIR
        self.is_live = is_live
        self.client = None
        self.profile = None

        if self.is_live:
            self._connect()

    def _connect(self) -> bool:
        """Initializes connection to Fyers API using existing credentials."""
        try:
            from fyers_apiv3 import fyersModel
            app_id, access_token = load_fyers_credentials(self.root)
            if not app_id or not access_token:
                logger.error("Credentials missing in fyers_tokens.json! Operating in PAPER mode.")
                self.is_live = False
                return False

            self.client = fyersModel.FyersModel(
                client_id=app_id, token=access_token, log_path=None
            )
            prof_resp = self.client.get_profile()
            if prof_resp.get("s") == "ok":
                self.profile = prof_resp.get("data", {})
                logger.info(f"Connected to FYERS live account: {self.profile.get('name')} ({self.profile.get('fy_id')})")
                return True
            else:
                logger.error(f"Fyers profile error: {prof_resp}. Falling back to PAPER mode.")
                self.is_live = False
                return False
        except Exception as e:
            logger.error(f"Failed to connect to Fyers API ({e}). Falling back to PAPER mode.")
            self.is_live = False
            return False

    def get_available_margin(self) -> float:
        """Fetches available intraday equity margin from Fyers."""
        if not self.is_live or not self.client:
            return 999999.0  # Simulated unlimited paper margin

        try:
            resp = self.client.funds()
            if resp.get("s") == "ok":
                fund_limits = resp.get("fund_limit", [])
                for item in fund_limits:
                    # Item 10 is 'Available Balance' or check title
                    if "Available" in item.get("title", "") or item.get("id") == 10:
                        return float(item.get("equityAmount", 0.0))
                    if "Total Balance" in item.get("title", "") or item.get("id") == 1:
                        total = float(item.get("equityAmount", 0.0))
                return total
        except Exception as e:
            logger.warning(f"Could not fetch available margin: {e}")
        return 0.0

    def _auth_header(self) -> str:
        """Builds the Fyers Authorization header ('app_id:access_token')."""
        app_id, access_token = load_fyers_credentials(self.root)
        if not app_id or not access_token:
            return ""
        return f"{app_id}:{access_token}"

    def get_margin_per_share(self, symbol: str, side: str) -> float:
        """
        Returns the intraday margin (in Rs.) required for a SINGLE share of `symbol`
        via the Fyers multiorder margin calculator. This inherently encodes the
        per-stock leverage. Returns 0.0 on failure (caller must then SKIP the trade).
        """
        if not self.is_live or not self.client:
            return 0.0
        auth = self._auth_header()
        if not auth:
            return 0.0

        clean_sym = symbol.upper().replace("NSE:", "").replace("-EQ", "")
        fy_symbol = f"NSE:{clean_sym}-EQ"
        fy_side = 1 if str(side).upper() == "BUY" else -1
        payload = {
            "data": [{
                "symbol": fy_symbol,
                "qty": 1,
                "side": fy_side,
                "type": 2,
                "productType": "INTRADAY",
                "limitPrice": 0,
                "stopLoss": 0,
            }]
        }
        try:
            resp = requests.post(
                f"{FYERS_API_BASE}/multiorder/margin",
                headers={
                    "Authorization": auth,
                    "Content-Type": "application/json",
                    "version": "3",
                },
                json=payload,
                timeout=8,
            )
            body = resp.json()
            if body.get("s") != "ok":
                logger.warning(f"Margin calc failed for {fy_symbol}: {body.get('message')}")
                return 0.0
            data = body.get("data", {})
            margin = 0.0
            # `margin_total` = margin required for THIS order block (qty=1 -> per share).
            # `margin_new_order` is the account-wide total (prevMargin + this order), NOT per-order.
            if isinstance(data, dict):
                margin = float(data.get("margin_total") or 0.0)
            elif isinstance(data, list) and data and isinstance(data[0], dict):
                margin = float(data[0].get("margin_total") or 0.0)
            return margin
        except Exception as e:
            logger.warning(f"Margin calc exception for {fy_symbol}: {e}")
            return 0.0

    def compute_qty(self, symbol: str, side: str, cash_budget: float) -> Tuple[int, float]:
        """
        Sizes a trade from a cash budget: qty = floor(budget / margin_per_share).
        The margin already applies per-stock leverage. Returns (qty, margin_per_share).
        qty is 0 when sizing is impossible (no margin info) -> caller skips the trade.
        """
        if cash_budget <= 0:
            return 0, 0.0
        mps = self.get_margin_per_share(symbol, side)
        if mps <= 0:
            return 0, 0.0
        qty = int(cash_budget // mps)
        return (qty if qty >= 1 else 0), mps

    def get_order_fill_details(self, order_id: str, max_wait_sec: float = 2.0) -> Dict[str, Any]:
        """
        Polls the orderbook to confirm if an order was FILLED, REJECTED, or TRANSIT,
        and extracts the exact traded price.
        """
        if not self.is_live or not self.client:
            return {"status": "FILLED", "traded_price": None, "message": "Simulated fill"}

        start_t = time.time()
        while time.time() - start_t <= max_wait_sec:
            try:
                ob_resp = self.client.orderbook()
                if ob_resp.get("s") == "ok":
                    order_list = ob_resp.get("orderBook", [])
                    for order in order_list:
                        if str(order.get("id")) == str(order_id):
                            order_status = order.get("status")
                            # Fyers Status codes: 2 = Filled, 5 = Rejected, 6 = Cancelled, 1 = Transit/Open
                            if order_status == 2:
                                traded_price = float(order.get("tradedPrice", 0.0))
                                traded_qty = int(order.get("filledQty", 0))
                                return {
                                    "status": "FILLED",
                                    "traded_price": traded_price,
                                    "traded_qty": traded_qty,
                                    "message": "Order filled successfully",
                                }
                            elif order_status in [5, 6]:
                                return {
                                    "status": "REJECTED",
                                    "traded_price": 0.0,
                                    "traded_qty": 0,
                                    "message": order.get("message", "Order rejected/cancelled"),
                                }
            except Exception as e:
                logger.warning(f"Error querying orderbook for {order_id}: {e}")
            time.sleep(0.4)

        return {"status": "SUBMITTED", "traded_price": None, "message": "Order submitted, pending fill confirmation"}

    def place_entry_order(
        self, symbol: str, side: str, qty: int, estimated_price: float, tag: str = "V6_ENTRY",
        sl_pct: float = 0.0, tp_pct: float = 0.0
    ) -> Dict[str, Any]:
        """
        Applies a live market entry order (BUY or SELL).
        Returns a dict containing: status, order_id, fill_price, qty, message.
        """
        fy_symbol = f"NSE:{symbol}-EQ"
        fy_side = 1 if side == "BUY" else -1

        logger.info(f"Applying Entry Order: {side} {qty} {fy_symbol} (Est. LTP: {estimated_price:.2f})...")

        if not self.is_live or not self.client:
            return {
                "status": "FILLED",
                "order_id": f"PAPER_{int(time.time()*1000)}",
                "fill_price": estimated_price,
                "qty": qty,
                "side": side,
                "simulated": True,
                "message": "Paper order simulated at LTP",
            }

        # Check margin using the REAL per-stock intraday margin (not a fixed 5x assumption).
        avail_margin = self.get_available_margin()
        mps = self.get_margin_per_share(symbol, side)
        required_est = mps * int(qty) if mps > 0 else (estimated_price * int(qty))
        if avail_margin > 0 and required_est > 0 and avail_margin < required_est:
            logger.error(f"Insufficient margin for {fy_symbol}! Required: ~Rs. {required_est:.2f}, Avail: Rs. {avail_margin:.2f}")
            return {
                "status": "REJECTED",
                "order_id": None,
                "fill_price": None,
                "qty": qty,
                "message": f"Insufficient margin (Avail: Rs. {avail_margin:.2f}, Req: Rs. {required_est:.2f})",
            }

        order_data = {
            "symbol": fy_symbol,
            "qty": int(qty),
            "type": 2,              # 2 = Market order
            "side": fy_side,        # 1 = Buy, -1 = Sell
            "productType": "INTRADAY",
            "limitPrice": 0,
            "stopPrice": 0,
            "validity": "DAY",
            "disclosedQty": 0,
            "offlineOrder": False,
            "orderTag": re.sub(r'[^a-zA-Z0-9]', '', str(tag))[:20] or "V6CLI",
        }

        # Smart Stop-loss / Target overlay (replaces deprecated BO/CO).
        # Values are OFFSETS from entry fill; legType 2 = percent of entry.
        if sl_pct and sl_pct > 0:
            order_data["stopLoss"] = round(float(sl_pct), 2)
        if tp_pct and tp_pct > 0:
            order_data["takeProfit"] = round(float(tp_pct), 2)
        if (sl_pct and sl_pct > 0) or (tp_pct and tp_pct > 0):
            order_data["legType"] = 2

        try:
            resp = self.client.place_order(data=order_data)
            if resp.get("s") == "ok":
                order_id = resp.get("id")
                logger.info(f"Entry order submitted to FYERS. Order ID: {order_id}. Confirming fill...")
                fill_details = self.get_order_fill_details(order_id)
                fill_price = fill_details.get("traded_price") or estimated_price

                # Safety: confirm the SL/Target overlay is attached to the position.
                # Fyers stores the OCO on the POSITION (not as separate leg orders),
                # so we read it back from positions().
                if (sl_pct and sl_pct > 0) or (tp_pct and tp_pct > 0):
                    try:
                        o_sym = symbol.upper().replace("NSE:", "").replace("-EQ", "")
                        pos_resp = self.client.positions()
                        attached = False
                        for p in pos_resp.get("netPositions", []):
                            p_sym = str(p.get("symbol", "")).upper().replace("NSE:", "").replace("-EQ", "")
                            if p_sym == o_sym and (float(p.get("stopLoss") or 0) > 0 or float(p.get("takeProfit") or 0) > 0):
                                attached = True
                                logger.info(f"[LEG CHECK] {symbol}: position carries SL={p.get('stopLoss')} TP={p.get('takeProfit')} (legType={p.get('legType')}).")
                                break
                        if not attached:
                            logger.error(f"[LEG CHECK] {symbol}: position has NO SL/TP attached! Position may be unprotected.")
                    except Exception as le:
                        logger.warning(f"[LEG CHECK] {symbol}: could not verify overlay ({le})")

                return {
                    "status": fill_details.get("status"),
                    "order_id": order_id,
                    "fill_price": fill_price,
                    "qty": qty,
                    "side": side,
                    "simulated": False,
                    "message": fill_details.get("message"),
                }
            else:
                logger.error(f"FYERS Order Placement Error: {resp.get('message')}")
                return {
                    "status": "REJECTED",
                    "order_id": None,
                    "fill_price": None,
                    "qty": qty,
                    "message": resp.get("message", "Order placement failed"),
                }
        except Exception as e:
            logger.error(f"Exception placing entry order for {fy_symbol}: {e}")
            return {
                "status": "ERROR",
                "order_id": None,
                "fill_price": None,
                "qty": qty,
                "message": str(e),
            }

    def place_exit_order(
        self, symbol: str, open_side: str, qty: int, estimated_price: float, reason: str = "EXIT"
    ) -> Dict[str, Any]:
        """
        Squares off an existing open position by submitting an opposite market order.
        """
        exit_side = "SELL" if open_side == "BUY" else "BUY"
        fy_symbol = f"NSE:{symbol}-EQ"
        fy_side = -1 if open_side == "BUY" else 1

        logger.info(f"Applying Square-Off Order: {exit_side} {qty} {fy_symbol} (Reason: {reason}, Est. LTP: {estimated_price:.2f})...")

        if not self.is_live or not self.client:
            return {
                "status": "FILLED",
                "order_id": f"PAPER_EXIT_{int(time.time()*1000)}",
                "fill_price": estimated_price,
                "qty": qty,
                "simulated": True,
                "message": f"Paper exit simulated at LTP ({reason})",
            }

        order_data = {
            "symbol": fy_symbol,
            "qty": int(qty),
            "type": 2,              # 2 = Market order
            "side": fy_side,
            "productType": "INTRADAY",
            "limitPrice": 0,
            "stopPrice": 0,
            "validity": "DAY",
            "disclosedQty": 0,
            "offlineOrder": False,
            "orderTag": re.sub(r'[^a-zA-Z0-9]', '', f"V6{reason}")[:20] or "V6EXIT",
        }

        try:
            resp = self.client.place_order(data=order_data)
            if resp.get("s") == "ok":
                order_id = resp.get("id")
                logger.info(f"Square-off order submitted to FYERS. Order ID: {order_id}. Confirming fill...")
                fill_details = self.get_order_fill_details(order_id)
                fill_price = fill_details.get("traded_price") or estimated_price

                return {
                    "status": fill_details.get("status"),
                    "order_id": order_id,
                    "fill_price": fill_price,
                    "qty": qty,
                    "simulated": False,
                    "message": fill_details.get("message"),
                }
            else:
                logger.error(f"FYERS Square-Off Error: {resp.get('message')}")
                return {
                    "status": "ERROR",
                    "order_id": None,
                    "fill_price": estimated_price,
                    "qty": qty,
                    "message": resp.get("message"),
                }
        except Exception as e:
            logger.error(f"Exception during square-off for {fy_symbol}: {e}")
            return {
                "status": "ERROR",
                "order_id": None,
                "fill_price": estimated_price,
                "qty": qty,
                "message": str(e),
            }

    def _cancel_pending_legs(self, symbol: str) -> int:
        """
        Cancels any pending (open) SL/Target legs for a symbol so no trace of the
        order block is left behind before/after a square-off.
        Returns the number of orders cancelled.
        """
        if not self.is_live or not self.client:
            return 0
        clean_sym = symbol.upper().replace("NSE:", "").replace("-EQ", "")
        cancelled = 0
        try:
            ob_resp = self.client.orderbook()
            if ob_resp.get("s") != "ok":
                return 0
            for order in ob_resp.get("orderBook", []):
                o_sym = str(order.get("symbol", "")).upper().replace("NSE:", "").replace("-EQ", "")
                o_status = order.get("status")
                # 6 = Pending, 1 = Cancelled, 2 = Filled, 5 = Rejected, 4 = Transit
                if o_sym == clean_sym and o_status in [4, 6]:
                    oid = order.get("id")
                    try:
                        c_resp = self.client.cancel_order({"id": oid})
                        if c_resp.get("s") == "ok":
                            cancelled += 1
                            logger.info(f"Cancelled pending leg {oid} for {clean_sym}.")
                        else:
                            logger.warning(f"Could not cancel leg {oid} for {clean_sym}: {c_resp.get('message')}")
                    except Exception as ce:
                        logger.warning(f"Exception cancelling leg {oid}: {ce}")
        except Exception as e:
            logger.warning(f"Error sweeping pending legs for {clean_sym}: {e}")
        return cancelled

    def reconcile_broker_positions(self, local_positions: Dict[str, Any]) -> List[str]:
        """
        Reconciles active local state positions against actual FYERS broker positions.
        Returns a list of symbols that were closed externally (e.g. broker RMS or manual app exit).
        """
        if not self.is_live or not self.client:
            return []

        externally_closed = []
        try:
            pos_resp = self.client.positions()
            if pos_resp.get("s") != "ok":
                return []

            net_positions = pos_resp.get("netPositions", [])
            # Map symbol -> netQty
            broker_map = {}
            for p in net_positions:
                clean_sym = p.get("symbol", "").replace("NSE:", "").replace("-EQ", "")
                broker_map[clean_sym] = int(p.get("netQty", 0))

            for sym, pos in list(local_positions.items()):
                broker_qty = broker_map.get(sym, 0)
                # If local state says OPEN, but broker net quantity is 0, it was squared off externally
                if broker_qty == 0:
                    logger.warning(f"[RECONCILIATION] {sym} is 0 at broker! Position was closed externally.")
                    externally_closed.append(sym)
        except Exception as e:
            logger.warning(f"Error during broker position reconciliation: {e}")

        return externally_closed


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="FYERS Order Dispatcher CLI")
    parser.add_argument("--symbol", type=str, default="", help="Stock ticker symbol (e.g. UPL)")
    parser.add_argument("--action", type=str, default="", choices=["BUY", "SELL", ""], help="Order action (BUY or SELL)")
    parser.add_argument("--qty", type=int, default=1, help="Quantity to trade (0 = auto-detect open broker position for square-off)")
    parser.add_argument("--tag", type=str, default="V6_CLI", help="Order tag / Reason")
    parser.add_argument("--sl-pct", type=float, default=0.0, help="Stop-loss offset in percent of entry (e.g. 1.0)")
    parser.add_argument("--tp-pct", type=float, default=0.0, help="Take-profit offset in percent of entry (e.g. 1.5)")
    parser.add_argument("--calc-qty", action="store_true", help="Sizing mode: emit SIZE|TICKER=qty lines from live capital / per-stock margin")
    parser.add_argument("--slots", type=int, default=0, help="Free slots available (sizing mode)")
    parser.add_argument("--symbols", type=str, default="", help="Comma list TICK:ACTION,... (sizing mode)")
    args = parser.parse_args()

    # ---- SIZING MODE: dynamic qty from live available capital + per-stock margin ----
    # budget per admitted trade = (available - committed) / remaining_free_slots
    # qty = floor(budget / margin_per_share)   (margin already applies per-stock leverage)
    if args.calc_qty:
        fom = FyersOrderManager(is_live=True)
        if not fom.is_live or fom.client is None:
            logger.error("Sizing requested but broker not live. Emitting no sizes.")
            sys.exit(1)
        available = fom.get_available_margin()
        if available <= 0 or args.slots < 1 or not args.symbols:
            logger.error(f"Sizing impossible (avail={available}, slots={args.slots}). Emitting no sizes.")
            sys.exit(1)
        committed = 0.0
        pairs = [p.strip() for p in args.symbols.split(",") if p.strip()]
        for i, pair in enumerate(pairs):
            remaining_slots = args.slots - i
            if remaining_slots < 1:
                break
            if ":" not in pair:
                continue
            tick, act = pair.split(":", 1)
            tick, act = tick.strip().upper(), act.strip().upper()
            budget = (available - committed) / remaining_slots
            qty, marg_per_share = fom.compute_qty(tick, act, budget)
            logger.info(f"[SIZE] {tick} {act}: budget={budget:.2f}, margin/share={marg_per_share:.2f} -> qty={qty}")
            print(f"SIZE|{tick}={qty}")
            committed += qty * marg_per_share
        sys.exit(0)

    if not args.symbol or not args.action:
        logger.error("--symbol and --action are required (unless --calc-qty is used).")
        sys.exit(2)

    logger.info(f"CLI Dispatcher called: symbol={args.symbol}, action={args.action}, qty={args.qty}, tag={args.tag}, sl%={args.sl_pct}, tp%={args.tp_pct}")

    is_exit = args.qty <= 0
    clean_target = args.symbol.upper().replace("NSE:", "").replace("-EQ", "")
    fom = FyersOrderManager(is_live=True)

    trade_qty = args.qty
    # If qty <= 0, auto-detect open position from Fyers to square off
    if is_exit and fom.client:
        try:
            pos_resp = fom.client.positions()
            for p in pos_resp.get("netPositions", []):
                sym_name = p.get("symbol", "").replace("NSE:", "").replace("-EQ", "")
                if sym_name == clean_target:
                    trade_qty = abs(int(p.get("netQty", 0)))
                    logger.info(f"Auto-detected open broker position for {clean_target}: {trade_qty} shares.")
                    break
        except Exception as e:
            logger.warning(f"Error querying broker positions for {clean_target}: {e}")

    if trade_qty <= 0:
        logger.warning(f"No open broker position found for {args.symbol} to square off (or qty was 0). Order not placed.")
        # Nothing open locally; sweep any orphan legs just in case.
        fom._cancel_pending_legs(args.symbol)
        sys.exit(0)

    if is_exit:
        # Clear SL/Target legs first so they cannot fire while we square off.
        fom._cancel_pending_legs(args.symbol)
        res = fom.place_entry_order(args.symbol, args.action, trade_qty, estimated_price=0.0, tag=args.tag)
        print("Order Dispatch Result:", res)
        # Verify the whole block is gone: position flat + no leftover pending legs.
        try:
            fom._cancel_pending_legs(args.symbol)
            remaining = 0
            if fom.client:
                pos_resp = fom.client.positions()
                for p in pos_resp.get("netPositions", []):
                    sym_name = p.get("symbol", "").replace("NSE:", "").replace("-EQ", "")
                    if sym_name == clean_target:
                        remaining = int(p.get("netQty", 0))
            logger.info(f"[EXIT VERIFY] {clean_target} netQty after square-off = {remaining} (0 = flat, no traces).")
        except Exception as e:
            logger.warning(f"[EXIT VERIFY] could not confirm flat for {clean_target}: {e}")
    else:
        res = fom.place_entry_order(args.symbol, args.action, trade_qty, estimated_price=0.0, tag=args.tag,
                                    sl_pct=args.sl_pct, tp_pct=args.tp_pct)
        print("Order Dispatch Result:", res)
