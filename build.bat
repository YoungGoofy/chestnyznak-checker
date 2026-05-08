@echo off
title CISChecker - Build

echo ================================================
echo  CISChecker - Build .exe
echo ================================================
echo.

if not exist ".venv\Scripts\python.exe" (
    echo [1/4] Creating virtual environment...
    python -m venv .venv
    if errorlevel 1 (
        echo ERROR: Python not found. Install Python 3.8+ first.
        pause
        exit /b 1
    )
) else (
    echo [1/4] Virtual environment already exists.
)

echo [2/4] Installing dependencies...
.venv\Scripts\pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo ERROR: Failed to install dependencies.
    pause
    exit /b 1
)

echo [3/4] Installing pyinstaller...
.venv\Scripts\pip install pyinstaller --quiet

echo [4/4] Building .exe...
.venv\Scripts\pyinstaller --onefile --windowed --hidden-import openpyxl --name CISChecker gui_app.py
if errorlevel 1 (
    echo ERROR: Build failed.
    pause
    exit /b 1
)

echo.
echo ================================================
echo  DONE!
echo.
echo    dist\CISChecker.exe
echo.
echo  Put .env next to .exe (or set token via GUI).
echo ================================================
pause
