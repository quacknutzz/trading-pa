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
echo  LAUNCHING QUACKNUTZZ FORWARD-TEST SIMULATION SANDBOX [PORT 8766]
echo  Runs live strategy simulation on 5m NSE market feed without real risk
echo =====================================================================
echo.

:: Check if already running on port 8766
netstat -ano | findstr ":8766" | findstr "LISTENING" >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [OK] Simulation sandbox is already running on port 8766!
    start http://localhost:8766
    exit /b 0
)

echo Starting Forward-Test Sandbox on port 8766...
start "" %PY% sim\server.py --mode live --portal simulation --port 8766 --strategy quacknutzz_strat --no-browser

ping 127.0.0.1 -n 3 >nul
start http://localhost:8766
echo Forward-Test Sandbox running at http://localhost:8766
