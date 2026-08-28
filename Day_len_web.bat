@echo off
chcp 65001 >nul
echo ============================================
echo   DOC EXCEL VA DAY DU LIEU LEN WEB
echo ============================================
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0auto_publish.ps1"
echo.
echo ============================================
echo   XONG! Web se tu cap nhat sau ~30-60 giay.
echo   Link: https://sn-chuoi.github.io/KIEMSOATVATU.XSX/
echo ============================================
pause
