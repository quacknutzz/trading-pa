@echo off
title QUACKNUTZZ // Quantitative Trading Terminal
cd /d "%~dp0"

echo ======================================================================
echo    QUACKNUTZZ TERMINAL // QUANTITATIVE COMMAND STATION
echo    Strategy: Quacknutzz_Strat (Production Ready)
echo    Booting backend engine & opening browser dashboard...
echo ======================================================================

:: Start the Python server. server.py will automatically open http://localhost:8765 in your default browser!
echo Starting trading engine on http://localhost:8765 ...
echo (Keep this console window open or minimized while trading)
echo.

python sim\server.py --mode live --strategy quacknutzz_strat
pause
