@echo off
rem MektepApp AI - one-click launcher for Windows
cd /d "%~dp0"
title MektepApp AI

if not exist ".venv\Scripts\python.exe" (
    echo [1/3] Creating virtual environment .venv ...
    python -m venv .venv 2>nul || py -3 -m venv .venv
)
if not exist ".venv\Scripts\python.exe" (
    echo.
    echo Python 3.10+ was not found.
    echo Install it from https://www.python.org/downloads/
    echo and tick "Add python.exe to PATH" during installation.
    pause
    exit /b 1
)

echo [2/3] Installing dependencies - the first run takes a few minutes ...
".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -q -r requirements.txt
if errorlevel 1 (
    echo.
    echo Could not install dependencies. Check the internet connection and that Python is 3.10 or newer.
    pause
    exit /b 1
)

echo [3/3] Starting MektepApp AI - the browser will open automatically.
echo Press Ctrl+C in this window to stop the app.
".venv\Scripts\python.exe" -m streamlit run app.py
pause
