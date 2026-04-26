@echo off
chcp 65001 >nul
title 工厂订单转换工具 - 升级助手
echo.
echo ============================================================
echo   工厂订单转换工具 - 升级助手
echo ============================================================
echo.
echo 启动 PowerShell 升级脚本...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0升级.ps1"

echo.
echo ============================================================
pause
