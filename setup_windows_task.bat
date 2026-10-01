@echo off
setlocal

echo =========================================================
echo Setting up Daily Restaurant Ratings Task in Windows...
echo =========================================================

set SCRIPT_DIR=%~dp0
set PYTHON_EXE=python.exe
set TASK_NAME=RestaurantRatingsDailyMailer

:: Create Windows Scheduled Task to run daily at 09:00 AM
schtasks /create /tn "%TASK_NAME%" /tr "\"%PYTHON_EXE%\" \"%SCRIPT_DIR%main.py\" --run-now" /sc daily /st 09:00 /f

if %ERRORLEVEL% equ 0 (
    echo.
    echo [SUCCESS] Windows Scheduled Task '%TASK_NAME%' created successfully!
    echo It will trigger automatically every day at 09:00 AM.
) else (
    echo.
    echo [ERROR] Failed to create scheduled task. Try running this batch file as Administrator.
)

echo.
pause
