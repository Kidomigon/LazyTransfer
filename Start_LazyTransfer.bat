@echo off
:: LazyTransfer - Start GUI
:: This batch file launches the PowerShell GUI with admin privileges

:: Check for admin rights
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Requesting administrator privileges...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

:: Run the PowerShell GUI
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0LazyTransfer-GUI.ps1"
