@echo off
:: Windows: double-click to refresh the dashboard (same pipeline as Update-Dashboard.command on macOS).
:: All logic lives in build\update.js
title MSP-CCB Dashboard Update
cd /d "%~dp0"
where node >nul 2>&1
if errorlevel 1 (
    echo ERROR: Node.js is not installed or not in PATH. Install it from https://nodejs.org
    pause
    exit /b 1
)
node build\update.js
if errorlevel 1 (
    echo.
    echo Update failed. See the error above.
    pause
    exit /b 1
)
exit /b 0
