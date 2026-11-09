@echo off
setlocal
cd /d "%~dp0"

echo =====================================================================
echo  QUACKNUTZZ AUTONOMOUS TRADING ENGINE - TASK SCHEDULER SETUP
echo =====================================================================
echo.
echo Setting up automated schedule:
echo   - Trigger: Every Monday, Tuesday, Wednesday, Thursday, Friday
echo   - Time:    09:00 AM IST (15 minutes before 09:15 AM market open)
echo   - Mode:    100%% Silent / Headless (No command window popups)
echo   - Strategy: Quacknutzz_Strat (Production Engine)
echo   - Terminal: http://localhost:8765
echo.

schtasks /Create /TN "QuacknutzzTradingBot" /TR "wscript.exe \"%~dp0scripts\start_headless_live.vbs\"" /SC WEEKLY /D MON,TUE,WED,THU,FRI /ST 09:00 /F /RL HIGHEST

if %ERRORLEVEL% EQU 0 (
    echo.
    echo [SUCCESS] Windows Task 'QuacknutzzTradingBot' successfully created!
    echo The engine will wake up and start running automatically at 9:00 AM on weekdays.
    echo You can open http://localhost:8765 anytime to check live trading.
) else (
    echo.
    echo [NOTE] If permission error occurred, right-click this file and choose 'Run as administrator'.
)

echo.
pause
