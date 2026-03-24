# 公文 OCR 辨識系統

上傳公文圖片，自動辨識欄位（來文機關、字號、收文文號、日期、事由），可編輯後匯出 CSV / Excel。

## 安裝步驟

### 1. 安裝 Python（3.10 以上）

- **Mac**: 前往 https://www.python.org/downloads/ 下載安裝
- **Windows**: 前往 https://www.python.org/downloads/ 下載安裝，**安裝時務必勾選「Add Python to PATH」**

### 2. 安裝 Tesseract OCR

**Mac:**
```bash
# 先安裝 Homebrew（如果還沒有）
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

# 安裝 Tesseract 及中文語言包
brew install tesseract tesseract-lang
```

**Windows:**
1. 前往 https://github.com/UB-Mannheim/tesseract/wiki
2. 下載最新版安裝程式
3. 安裝時勾選 **Chinese Traditional** 語言包
4. 安裝時勾選 **Add to PATH**（或手動將安裝路徑加入系統 PATH）

### 3. 下載本程式

```bash
git clone https://github.com/20sean02/ocr-project.git
cd ocr-project
```

或直接在 GitHub 頁面點擊 **Code → Download ZIP**，解壓縮後進入資料夾。

## 啟動方式

### Mac
雙擊 `start.command` 檔案即可。

> 首次執行可能需要右鍵 → 打開，並允許執行。

### Windows
雙擊 `start.bat` 檔案即可。

---

啟動後瀏覽器會自動開啟 http://127.0.0.1:5050 ，即可開始使用。

## 使用方式

1. 拖曳或選擇公文圖片上傳
2. 系統自動 OCR 辨識各欄位
3. 點擊欄位可直接編輯修正
4. 匯出 CSV 或 Excel

### 匯入之前的資料

如果你之前已經匯出過 CSV，可以用「匯入已有 CSV」功能載入，再繼續上傳新圖片。

## 關閉

- **Mac**: 在終端機按 `Ctrl+C`
- **Windows**: 直接關閉命令列視窗
