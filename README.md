# Cockpit V6 — Algorithmic Intraday Trading Engine & Live Paper Simulator

[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![Platform](https://img.shields.io/badge/platform-NSE%20India-orange.svg)]()
[![Broker](https://img.shields.io/badge/broker-FYERS%20API%20v3-green.svg)](https://fyers.in/)
[![Engine](https://img.shields.io/badge/engine-Numba%20JIT%20Accelerated-yellow.svg)]()

A high-performance algorithmic trading platform and zero-risk simulation engine designed for National Stock Exchange of India (NSE) intraday equities. It combines a robust 10m/5m EMA 10/30 strategy with institutional order flow filters, multi-factor structural exits, dynamic slot sizing, and real-time telemetry.

---

## 🚀 Key Strategy Features

1. **Dual-Timeframe Trend Crossover**:
   - **Entry**: 10-minute EMA 10 vs 30 crossover on forming bars between 09:25 and 13:15 IST.
   - **Entry Order**: Resting limit order placed at the nearest 5m EMA pullback level (10/30/50/100/200) with a 3-bar cancellation policy.
2. **True Momentum Gate (09:15 – 09:45 AM IST)**:
   - Eliminates early-morning retail noise without sacrificing explosive morning trends.
   - Dynamic validation: Relative Volume ($\text{RVOL} \ge 1.25$), Adverse Wick Filtering ($\le 35\%$ wick rejection), and Macro NIFTY 50 trend alignment.
3. **Multi-Factor Structural Exits**:
   - Dynamic target scaling: $\max(\text{Brokerage Breakeven Target}, 1.5 \times \text{ATR}_{5m})$.
   - 2-Bar Swing Pivot Trailing Stop: Protects accumulated paper gains while allowing runners to breathe.
   - Fast momentum exhaustion & contra-cross protection.
4. **Capital & Fee Efficiency**:
   - Strict equal slot sizing (default 3 concurrent slots at 2.0x MIS margin).
   - Optimized for Stop-Loss Limit (`SL-L`) execution to prevent aggressive market-order slippage drag.
   - Comprehensive Indian statutory cost model (STT, NSE turnover, SEBI, GST, Stamp Duty, FYERS flat brokerage).

---

## 📂 Repository Architecture

```text
Trading Pa/
├── AI_BTB_Dashboard.xlsm      # Master Excel workbook with live Sheet1 feeds and trade logging
├── AI_BTB_Engine.py           # 5-minute cycle data synchronization & baseline runner
├── cockpit_v6_engine.py       # Live automated signal discovery, sizing & FYERS execution
├── run_live_simulation.bat    # 1-click launcher for live paper-trading web terminal
├── run_replay_last_session.bat# 1-click launcher for session replay terminal
├── engine/                    # Production quantitative execution package
│   ├── config.py              # Broker and engine environment configuration
│   ├── data_engine.py         # Historical & live market data pipelines
│   ├── fyers_order_manager.py # FYERS API v3 order execution & orderbook state manager
│   ├── indicators.py          # Vectorized technical indicator computations
│   ├── logger.py              # Centralized structured audit logger
│   ├── numba_kernels.py       # C-speed JIT Numba compiled indicator kernels
│   └── ta_engine.py           # Full-universe multi-timeframe analysis engine
├── sim/                       # In-memory paper-trading simulator & web dashboard
│   ├── config.py              # Strategy parameters, risk, slot sizes & costs
│   ├── datafeed.py            # Real-time multi-stock NSE data polling
│   ├── server.py              # Asynchronous HTTP/WebSocket telemetry daemon
│   ├── simulator.py           # Dual-Engine backtester & real-time simulation engine
│   ├── universe.json          # Curated NSE universe of 139 liquid symbols
│   ├── static/                # Frontend dashboard assets (HTML5, Canvas, WebSockets)
│   └── journal/               # Session trade logs, signal histories, and summary JSONs
├── docs/                      # Comprehensive technical architecture & strategy guides
│   ├── AI_BTB_Cockpit_V6_Complete_System_Master_Guide.html / .pdf
│   ├── AI_BTB_Cockpit_V6_Modernization_Report.html / .pdf
│   └── Trading_Terminology_and_System_Guide.html
├── vba/                       # Original Excel VBA macro modules & sheet definitions
│   ├── Module_Cockpit_V6.bas
│   └── ... (All workbook VBA modules and Sheet classes)
└── scratch/                   # Quantitative research, parameter optimization & audit scripts
```

---

## 🛠️ Quick Start

### 1. Prerequisites
- Python 3.10 or higher
- Windows OS (for Excel COM automation, or any OS for pure Python simulation mode)

Install dependencies:
```bash
pip install pandas numpy numba yfinance openpyxl fyers-apiv3
```

### 2. Launch the Web Simulation Cockpit (Zero Money Risk)
Double-click `run_live_simulation.bat` or run:
```bash
python sim/server.py --mode live
```
Open your browser at: **[http://127.0.0.1:8765](http://127.0.0.1:8765)**

### 3. Replay a Historical Trading Session
To review bar-by-bar execution from the previous trading day:
Double-click `run_replay_last_session.bat` or run:
```bash
python sim/server.py --mode replay --speed 0.25
```

### 4. Run Quantitative Backtests
To run the full 62-day (~12-week) backtest across all 139 stocks:
```bash
python scratch/test_dynamic_target_formula.py
```

---

## 🛡️ License
Private & Proprietary. Developed for algorithmic intraday trading research on NSE India.
