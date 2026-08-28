@echo off
chcp 65001 >nul
echo ============================================
echo   DANG CAP NHAT DU LIEU DASHBOARD TU EXCEL
echo ============================================
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0update_data.ps1"
echo.
echo ============================================
echo   XONG! Hay mo lai (hoac F5) file dashboard.html
echo ============================================
pause
