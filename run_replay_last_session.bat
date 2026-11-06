@echo off
title COCKPIT V6 - Replay Last Trading Session (0 Real Money)
cd /d "%~dp0"
echo ======================================================================
echo   COCKPIT V6 // SESSION REPLAY TERMINAL (SIMULATION MODE)
echo   - Replaying the most recent NSE trading session bar-by-bar.
echo   - Opening web dashboard at http://localhost:8765 ...
echo ======================================================================
python sim\server.py --mode replay --speed 0.25
pause
