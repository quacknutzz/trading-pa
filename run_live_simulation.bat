@echo off
title QUACKNUTZZ - Live Intraday Trading Terminal
cd /d "%~dp0"
echo ======================================================================
echo   QUACKNUTZZ TERMINAL // LIVE INTRADAY TRADING ENGINE
echo   - Strategy: Quacknutzz_Strat (Production Ready)
echo   - 10m/5m EMA 10/30 Trend + 2-Tier Tax-Aware Ratchet
echo   - Opening terminal dashboard at http://localhost:8765 ...
echo ======================================================================
python sim\server.py --mode live --strategy quacknutzz_strat
pause
