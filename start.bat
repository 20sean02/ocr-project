@echo off
chcp 65001 >nul 2>nul
title OCR System
cd /d "%~dp0"

echo ==============================
echo   OCR System - Starting...
echo ==============================
echo.

:: Check Python
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Python not found!
    echo Please install Python 3.10+:
    echo   https://www.python.org/downloads/
    echo.
    echo Check "Add Python to PATH" during install
    echo.
    pause
    exit /b 1
)

for /f "tokens=*" %%i in ('python --version 2^>^&1') do echo Python: %%i

:: Check Tesseract
where tesseract >nul 2>nul
if %errorlevel% neq 0 (
    :: Check common install paths
    if exist "C:\Program Files\Tesseract-OCR\tesseract.exe" (
        set "PATH=%PATH%;C:\Program Files\Tesseract-OCR"
    ) else if exist "C:\Program Files (x86)\Tesseract-OCR\tesseract.exe" (
        set "PATH=%PATH%;C:\Program Files (x86)\Tesseract-OCR"
    ) else (
        echo.
        echo [ERROR] Tesseract OCR not found!
        echo Please install:
        echo   https://github.com/UB-Mannheim/tesseract/wiki
        echo.
        echo During install:
        echo   1. Select "Chinese Traditional" language pack
        echo   2. Select "Add to PATH"
        echo.
        pause
        exit /b 1
    )
)

:: Create venv if needed
if not exist ".venv" (
    echo.
    echo First run - creating virtual environment...
    python -m venv .venv
)

:: Activate venv
call .venv\Scripts\activate.bat

:: Install dependencies
echo Installing packages...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Failed to install packages
    pause
    exit /b 1
)

echo.
echo Starting server...
echo Browser will open automatically. If not, go to:
echo   http://127.0.0.1:5050
echo.
echo Close this window to stop the server.
echo ==============================
echo.

:: Open browser after short delay
start "" http://127.0.0.1:5050

:: Start the app
python app.py
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Application crashed
)
pause
