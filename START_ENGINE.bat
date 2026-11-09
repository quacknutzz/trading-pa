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

echo Starting Quacknutzz Live Engine in background...
start "" %PY% sim\server.py --mode live --portal live --port 8765 --strategy quacknutzz_strat --no-browser

ping 127.0.0.1 -n 3 >nul
start http://localhost:8765
echo Engine running at http://localhost:8765
