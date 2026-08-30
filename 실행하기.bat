@echo off
cd /d "%~dp0"

REM pythonw = console-less Python. Without it a black cmd window
REM stays open behind the app.

where pythonw >nul 2>&1
if %errorlevel%==0 (
    start "" pythonw run.py
    exit /b 0
)

where pyw >nul 2>&1
if %errorlevel%==0 (
    start "" pyw run.py
    exit /b 0
)

REM Fallback: console Python. The window will stay open.
python run.py
if errorlevel 1 pause
