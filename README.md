# Data Processing Tools — Excel/CSV 清理與比對工具

![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)
![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)
![Status](https://img.shields.io/badge/Release-v2.0.0-orange.svg)

**中文說明**  
這是一個針對 Excel/CSV 報表的一站式處理工具，支援 **清理、比對、模糊比對**，可用 GUI 或命令列執行，幫助自動化處理公司內部資料。  

**English Summary**  
One-stop Excel/CSV Cleaning & Matching toolkit (GUI/CLI). Supports auto header/encoding detection, field normalization, multi-key exact & fuzzy matching.

---

## ✨ 功能 Features
- **自動偵測**：表頭、分隔符、編碼（含 Big5/GBK）
- **清理與正規化**：去空白、全形轉半形、電話/Email/日期格式化
- **比對工具**：多鍵值精確比對＋姓名模糊比對（RapidFuzz）
- **雙模式**：GUI（圖形介面）＋ CLI（命令列工具）

---

## 📦 安裝 Install
```bash
git clone https://github.com/Avon46/data-processing-tools.git
cd data-processing-tools
pip install -r requirements.txt
```

## 🚀 快速開始 Quick Start

### GUI
```bash
python main.py
```

### CLI
```bash
excel-tools --left samples/master.xlsx --right samples/new.xlsx \
  --left-key "身分證字號,電話" --right-key "身分證字號,電話" \
  --output out --fuzzy-col 姓名 --fuzzy-threshold 85
```

## 📂 輸出 Output
- `left_clean.xlsx` / `right_clean.xlsx` → 清理後的資料
- `match_inner.xlsx` → 兩邊完全匹配
- `left_only.xlsx` / `right_only.xlsx` → 僅左／僅右出現的資料
- `fuzzy_matches.xlsx` → 模糊比對結果（若啟用）
- `summary.html` → 總結報告

## ⚙️ 設定檔 Config
```json
{
  "預設設定": {
    "標準報表": {
      "skiprows": 1,
      "columns": "all",
      "encoding": "auto",
      "description": "一般報表的通用模板"
    }
  },
  "自定義設定": {
    "公司內部報表": {
      "skiprows": 3,
      "columns": "1,2,4,5",
      "encoding": "UTF-8",
      "description": "公司內部格式",
      "注意事項": "跳過前三行標題"
    }
  }
}
```

## 📄 授權 License
MIT License © 2025 鄭桓
