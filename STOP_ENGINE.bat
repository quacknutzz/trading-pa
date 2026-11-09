@echo off
setlocal
echo =====================================================================
echo  STOPPING ALL QUACKNUTZZ SERVERS - PORTS 8765 AND 8766
echo =====================================================================

:: Terminate Port 8765
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":8765" ^| findstr "LISTENING"') do (
    taskkill /F /PID %%a >nul 2>&1
)

:: Terminate Port 8766
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":8766" ^| findstr "LISTENING"') do (
    taskkill /F /PID %%a >nul 2>&1
)

:: Also terminate any stray python/pythonw processes running sim\server.py
wmic process where "commandline like '%%sim\\server.py%%'" call terminate >nul 2>&1

echo [OK] All Quacknutzz trading servers stopped.
