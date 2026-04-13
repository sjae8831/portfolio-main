@echo off
title Certificate Auto v2.0

echo.
echo  ==========================================
echo   MES Certificate Automation v2.0
echo  ==========================================
echo.

python --version >/dev/null 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed.
    echo Please install Python from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during install.
    echo.
    pause
    exit /b 1
)

echo [OK] Python found:
python --version
echo.

echo [INFO] Installing required packages...
pip install requests openpyxl urllib3 -q
echo [OK] All packages ready.
echo.

echo [START] Launching program...
echo.

cd /d "%~dp0"
python mes_auto_v2.py

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Program exited with error.
    echo Please check the message above.
    echo.
    pause
)
