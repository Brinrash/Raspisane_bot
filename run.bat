@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

echo =======================================
echo        MAX Schedule Bot
echo =======================================

where py >nul 2>nul
if not errorlevel 1 (
    set "PYTHON_CMD=py -3"
) else (
    set "PYTHON_CMD=python"
)

if not exist ".venv\Scripts\python.exe" (
    echo Creating virtual environment...
    %PYTHON_CMD% -m venv .venv
    if errorlevel 1 goto :error
)

set "PYTHON_EXE=%CD%\.venv\Scripts\python.exe"

"%PYTHON_EXE%" -m pip install --upgrade pip
if errorlevel 1 goto :error
"%PYTHON_EXE%" -m pip install -r requirements.txt
if errorlevel 1 goto :error

if not exist ".env" (
    copy ".env.example" ".env" >nul
    echo.
    echo Created .env. Open it and insert MAX_BOT_TOKEN.
    echo Then run run.bat again.
    pause
    exit /b 0
)

"%PYTHON_EXE%" max_bot.py
echo.
echo Bot stopped.
pause
exit /b 0

:error
echo.
echo Setup failed. Check Python installation and network connection.
pause
exit /b 1
