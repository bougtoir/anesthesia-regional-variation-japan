@echo off
REM ProjectHub - Double-click this file to start the server
REM Requires Node.js v18+ installed and available in PATH

where node >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Node.js is not installed.
    echo Please install from https://nodejs.org/
    pause
    exit /b 1
)

cd /d "%~dp0"
node start.mjs %*
if %errorlevel% neq 0 pause
