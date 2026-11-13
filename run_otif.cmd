@echo off
REM Task Scheduler entry point for the daily OTIF build.
REM Any arguments passed here are forwarded to main.py.
setlocal
cd /d "%~dp0"

if not exist "logs" mkdir "logs"

if not exist ".venv\Scripts\python.exe" (
    echo [%DATE% %TIME%] .venv is missing - run: python -m venv .venv ^&^& .venv\Scripts\pip install -r requirements.txt >> "logs\failures.log"
    exit /b 1
)

".venv\Scripts\python.exe" main.py %*
set RC=%ERRORLEVEL%

if %RC% NEQ 0 (
    echo [%DATE% %TIME%] OTIF run failed with exit code %RC% >> "logs\failures.log"
)

exit /b %RC%
