@echo off
title ChestnyZnakChecker - Setup

echo ================================================
echo  ChestnyZnakChecker - Dev Setup
echo ================================================
echo.

if not exist ".venv\Scripts\python.exe" (
    echo Creating virtual environment...
    python -m venv .venv
    if errorlevel 1 (
        echo ERROR: Python not found.
        pause
        exit /b 1
    )
)

echo Installing dependencies...
.venv\Scripts\pip install -r requirements.txt --quiet

echo.
echo ================================================
echo  Ready!
echo.
echo  Launch GUI:
echo    run.bat
echo    (or: .venv\Scripts\python gui_app.py)
echo.
echo  Launch CLI:
echo    .venv\Scripts\python check_codes.py --true -f codes.txt -o result.xlsx
echo ================================================

if "%1"=="gui" goto run_gui
if "%1"=="run" goto run_gui
pause
goto :eof

:run_gui
echo.
echo Launching GUI...
.venv\Scripts\python gui_app.py
goto :eof
