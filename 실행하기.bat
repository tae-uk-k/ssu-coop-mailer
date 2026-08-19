@echo off
cd /d "%~dp0"

where python >nul 2>&1
if %errorlevel%==0 (
    python run.py
) else (
    py run.py
)

if errorlevel 1 (
    echo.
    pause
)
