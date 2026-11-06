@echo off
title COCKPIT V6 - Live Paper Trading Terminal (0 Real Money)
cd /d "%~dp0"
echo ======================================================================
echo   COCKPIT V6 // LIVE PAPER-TRADING TERMINAL (SIMULATION MODE)
echo   - Zero real money, zero broker API keys required.
echo   - Polls live NSE 5-minute bars and runs V6 strategy in memory.
echo   - Opening web dashboard at http://localhost:8765 ...
echo ======================================================================
python sim\server.py --mode live
pause
