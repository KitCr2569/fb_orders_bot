@echo off
chcp 65001 >nul
title อัปโหลด JSON ขึ้น Cloud
echo.
echo ========================================
echo   อัปโหลด JSON ล่าสุดขึ้น Cloud
echo   (ไม่ดึงใหม่จาก Facebook)
echo ========================================
echo.
cd /d "%~dp0"
python fetch_and_upload.py --upload-only
echo.
echo ========================================
echo   เสร็จแล้ว! เปิด Dashboard ได้ที่:
echo   https://fb-orders-bot.onrender.com
echo ========================================
echo.
pause
