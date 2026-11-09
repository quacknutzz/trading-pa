@echo off
setlocal
cd /d "%~dp0"

:: Detect Python executable
where pythonw >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set "PY=pythonw"
) else (
    where python >nul 2>&1
    if %ERRORLEVEL% EQU 0 (
        set "PY=python"
    ) else (
        if exist "C:\Users\rishi\AppData\Local\Programs\Python\Python313\pythonw.exe" (
            set "PY=C:\Users\rishi\AppData\Local\Programs\Python\Python313\pythonw.exe"
        ) else (
            echo [ERROR] Python not found on PATH! Please install Python.
            pause
            exit /b 1
        )
    )
)

echo =====================================================================
echo  STARTING BOTH QUACKNUTZZ PORTALS
echo   1. FYERS Live Brokerage Terminal [Port 8765]
echo   2. Forward-Test Simulation Sandbox [Port 8766]
echo =====================================================================
echo.

:: 1. Start Live Brokerage Terminal on Port 8765
netstat -ano | findstr ":8765" | findstr "LISTENING" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [1/2] Starting FYERS Live Terminal on port 8765...
    start "" %PY% sim\server.py --mode live --portal live --port 8765 --strategy quacknutzz_strat --no-browser
) else (
    echo [1/2] Port 8765 already running.
)

:: 2. Start Forward-Test Simulation Sandbox on Port 8766
netstat -ano | findstr ":8766" | findstr "LISTENING" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [2/2] Starting Simulation Sandbox on port 8766...
    start "" %PY% sim\server.py --mode live --portal simulation --port 8766 --strategy quacknutzz_strat --no-browser
) else (
    echo [2/2] Port 8766 already running.
)

ping 127.0.0.1 -n 4 >nul
echo.
echo Opening portals in browser...
start http://localhost:8765
start http://localhost:8766
echo.
echo [DONE] Both portals active!
