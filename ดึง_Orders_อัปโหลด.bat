@echo off
chcp 65001 >nul
title ดึง Orders + อัปโหลด Cloud
echo.
echo ========================================
echo   ดึง Orders จาก Facebook
echo   + อัปโหลดขึ้น Cloud อัตโนมัติ
echo ========================================
echo.
cd /d "%~dp0"
python fetch_and_upload.py --month 3
echo.
echo ========================================
echo   เสร็จแล้ว! เปิด Dashboard ได้ที่:
echo   https://fb-orders-bot.onrender.com
echo ========================================
echo.
pause
