@echo off
chcp 65001 >nul
title 公文 OCR 辨識系統
cd /d "%~dp0"

echo ==============================
echo   公文 OCR 辨識系統
echo ==============================
echo.

:: Check Python
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [錯誤] 找不到 Python！
    echo 請先安裝 Python 3.10+：
    echo   https://www.python.org/downloads/
    echo.
    echo 安裝時請勾選 "Add Python to PATH"
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
        echo [錯誤] 找不到 Tesseract OCR！
        echo 請先安裝：
        echo   https://github.com/UB-Mannheim/tesseract/wiki
        echo.
        echo 安裝時請：
        echo   1. 勾選 "Chinese Traditional" 語言包
        echo   2. 勾選 "Add to PATH"
        echo.
        pause
        exit /b 1
    )
)

:: Create venv if needed
if not exist ".venv" (
    echo.
    echo 首次啟動，正在建立虛擬環境...
    python -m venv .venv
)

:: Activate venv
call .venv\Scripts\activate.bat

:: Install dependencies
echo 正在檢查套件...
pip install -q -r requirements.txt 2>nul

echo.
echo 啟動伺服器中...
echo 瀏覽器將自動開啟，如未開啟請手動前往：
echo   http://127.0.0.1:5050
echo.
echo 關閉此視窗可停止伺服器
echo ==============================
echo.

:: Open browser after short delay
start "" http://127.0.0.1:5050

:: Start the app
python app.py
if %errorlevel% neq 0 (
    echo.
    echo [錯誤] 程式異常結束
)
pause
