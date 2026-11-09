@echo off
setlocal
cd /d "%~dp0\.."

if not exist "logs" mkdir "logs"

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
            echo [%DATE% %TIME%] [ERROR] Python not found on PATH! >> logs\engine_auto.log
            exit /b 1
        )
    )
)

:: 1. Check & start Live Brokerage Terminal on Port 8765
netstat -ano | findstr ":8765" | findstr "LISTENING" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [%DATE% %TIME%] Auto-starting FYERS Live Terminal (Port 8765)... >> logs\engine_auto.log
    start "" %PY% sim\server.py --mode live --portal live --port 8765 --strategy quacknutzz_strat --no-browser
) else (
    echo [%DATE% %TIME%] Port 8765 already listening. >> logs\engine_auto.log
)

:: 2. Check & start Forward-Test Simulation Sandbox on Port 8766
netstat -ano | findstr ":8766" | findstr "LISTENING" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [%DATE% %TIME%] Auto-starting Simulation Sandbox (Port 8766)... >> logs\engine_auto.log
    start "" %PY% sim\server.py --mode live --portal simulation --port 8766 --strategy quacknutzz_strat --no-browser
) else (
    echo [%DATE% %TIME%] Port 8766 already listening. >> logs\engine_auto.log
)

exit /b 0
