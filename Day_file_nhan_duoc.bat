@echo off
chcp 65001 >nul
echo ============================================
echo   DAY FILE data.js NHAN DUOC LEN WEB
echo ============================================
echo.
echo Script nay se day dung file data.js dang co
echo trong folder nay len web - KHONG doc lai Excel.
echo.
echo Truoc khi tiep tuc, hay chac chan ban da:
echo   1. Nhan file data.js tu dong nghiep (qua Zalo/email)
echo   2. Copy DE LEN file data.js cu trong folder nay
echo.
pause
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Day_file_nhan_duoc.ps1"
echo.
echo ============================================
echo   XONG! Web se tu cap nhat sau ~30-60 giay.
echo   Link: https://sn-chuoi.github.io/KIEMSOAT.VTXSX/
echo ============================================
pause
