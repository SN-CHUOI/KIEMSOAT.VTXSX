@echo off
chcp 65001 >nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Quan_ly_nguon.ps1"
echo.
pause
