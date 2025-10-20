@echo off
chcp 65001 >nul
title نظام البحث في قواعد بيانات الناخبين العراقيين

echo ============================================
echo   نظام البحث في قواعد بيانات الناخبين
echo   Iraqi Voter Database Search System
echo ============================================
echo.

echo 🔍 Checking Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python is not installed!
    echo Please download and install Python from: https://www.python.org/downloads/
    pause
    exit /b 1
)
echo ✅ Python found!
echo.

echo 📦 Installing required packages...
pip install Flask pyodbc --quiet
if errorlevel 1 (
    echo ⚠️  Warning: Package installation had issues, but continuing...
)
echo ✅ Packages installed!
echo.

echo ============================================
echo   🚀 STARTING APPLICATION...
echo ============================================
echo.
echo 📍 Access at: http://127.0.0.1:5000
echo 📍 Press Ctrl+C to stop the server
echo.

python app.py

pause
