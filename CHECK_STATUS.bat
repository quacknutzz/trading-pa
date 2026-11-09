@echo off
setlocal
echo =====================================================================
echo  QUACKNUTZZ TRADING SYSTEM - DUAL-PORTAL HEALTH CHECK
echo =====================================================================
echo.

:: Check Port 8765 (FYERS Live Terminal)
netstat -ano | findstr ":8765" | findstr "LISTENING" >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [PORT 8765 - LIVE TERMINAL]       ONLINE [Listening]
) else (
    echo [PORT 8765 - LIVE TERMINAL]       OFFLINE [Not running]
)

:: Check Port 8766 (Forward-Test Sandbox)
netstat -ano | findstr ":8766" | findstr "LISTENING" >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [PORT 8766 - SIMULATION SANDBOX]  ONLINE [Listening]
) else (
    echo [PORT 8766 - SIMULATION SANDBOX]  OFFLINE [Not running]
)

echo.
echo ---------------------------------------------------------------------
echo  BROKER AND SCHEDULER STATUS
echo ---------------------------------------------------------------------
if exist "fyers_tokens.json" (
    echo [FYERS BROKER] Credentials installed [fyers_tokens.json found]
) else (
    echo [FYERS BROKER] Missing fyers_tokens.json
)

:: Check Task Scheduler
schtasks /query /tn "QuacknutzzTradingBot" >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [AUTO-SCHEDULER] Windows Task Scheduler 'QuacknutzzTradingBot' is READY!
    echo                 Scheduled Mon-Fri at 09:00 AM IST.
) else (
    echo [AUTO-SCHEDULER] Task not found. Run SETUP_AUTOSTART.bat to register.
)

echo.
echo =====================================================================
echo  QUICK COMMANDS:
echo   - START_BOTH_PORTALS.bat           : Launch both Live and Sandbox
echo   - LAUNCH_SIMULATION_SANDBOX.bat    : Launch Sandbox only [port 8766]
echo   - START_ENGINE.bat                 : Launch Live Terminal [port 8765]
echo   - STOP_ENGINE.bat                  : Stop all servers
echo =====================================================================
echo.
pause
