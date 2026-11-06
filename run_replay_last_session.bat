@echo off
title QUACKNUTZZ - Session Replay Terminal
cd /d "%~dp0"
echo ======================================================================
echo   QUACKNUTZZ TERMINAL // REPLAY TRADING ENGINE
echo   - Strategy: Quacknutzz_Strat (Production Ready)
echo   - Replaying NSE trading session bar-by-bar with tape audit
echo   - Opening terminal dashboard at http://localhost:8765 ...
echo ======================================================================
python sim\server.py --mode replay --speed 0.25 --strategy quacknutzz_strat
pause
