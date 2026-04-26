@echo off
title Industry Order Tool - Upgrade Helper
echo.
echo ============================================================
echo   Industry Order Tool - Upgrade Helper
echo ============================================================
echo.
echo Starting PowerShell upgrade script...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0upgrade.ps1" %*

echo.
echo ============================================================
pause
