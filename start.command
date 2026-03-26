#!/bin/bash
# 公文 OCR 辨識系統 — Mac 啟動腳本
# 雙擊此檔案即可啟動

cd "$(dirname "$0")"

echo "=============================="
echo "  公文 OCR 辨識系統"
echo "=============================="
echo ""

# Check Python
if command -v python3 &>/dev/null; then
    PY=python3
elif command -v python &>/dev/null; then
    PY=python
else
    echo "[錯誤] 找不到 Python！"
    echo "請先安裝 Python 3.10+："
    echo "  https://www.python.org/downloads/"
    echo ""
    echo "按 Enter 關閉..."
    read
    exit 1
fi

echo "Python: $($PY --version)"

# Check Tesseract
if ! command -v tesseract &>/dev/null; then
    echo ""
    echo "[錯誤] 找不到 Tesseract OCR！"
    echo "請先安裝："
    echo "  brew install tesseract tesseract-lang"
    echo ""
    echo "如果沒有 Homebrew，先安裝 Homebrew："
    echo "  /bin/bash -c \"\$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)\""
    echo ""
    echo "按 Enter 關閉..."
    read
    exit 1
fi

echo "Tesseract: $(tesseract --version 2>&1 | head -1)"

# Create venv if needed
if [ ! -d ".venv" ]; then
    echo ""
    echo "首次啟動，正在建立虛擬環境..."
    $PY -m venv .venv
fi

# Activate venv
source .venv/bin/activate

# Install dependencies
echo "正在檢查套件..."
pip install -q -r requirements.txt 2>/dev/null

echo ""
echo "啟動伺服器中..."
echo "瀏覽器將自動開啟，如未開啟請手動前往："
echo "  http://127.0.0.1:5050"
echo ""
echo "按 Ctrl+C 可停止伺服器"
echo "=============================="
echo ""

# Open browser after short delay
(sleep 2 && open "http://127.0.0.1:5050") &

# Start the app
DEBUG=1 $PY app.py
