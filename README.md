# Data Processing Tools v2.0.0

[![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Version](https://img.shields.io/badge/Version-2.0.0-orange.svg)](https://github.com/Avon46/data-processing-tools/releases)

## 📋 專案簡介

Data Processing Tools 是一個強大的 Excel/CSV 資料清理與比對工具包，提供圖形化介面(GUI)和命令列介面(CLI)兩種使用方式。

## ✨ 主要功能

### 🔧 資料清理
- 自動檢測和處理編碼問題
- 移除重複資料和空行
- 標準化資料格式
- 支援多種檔案格式 (Excel, CSV)

### 🔍 資料比對
- 精確比對：完全匹配的資料識別
- 模糊比對：使用 RapidFuzz 演算法進行相似度比對
- 可調整的相似度閾值
- 支援多欄位比對

### ⚙️ 設定檔管理
- JSON 格式的設定檔
- 可自定義比對規則和欄位映射
- 靈活的輸出格式設定

## 🚀 安裝方式

### 方法一：從 GitHub 安裝
```bash
git clone https://github.com/Avon46/data-processing-tools.git
cd data-processing-tools
pip install -r requirements.txt
```

### 方法二：使用 pip 安裝
```bash
pip install data-processing-tools
```

## 📖 使用方式

### 🖥️ GUI 模式
```bash
# 啟動圖形化介面
python main.py
```

### ⌨️ CLI 模式
```bash
# 使用命令列工具
python dptools/cli.py

# 直接執行比對
python excel_matcher.py --config config.json
```

### 📝 CLI 參數說明
```bash
python excel_matcher.py [選項]

選項:
  --config FILE    設定檔路徑 (預設: config.json)
  --master FILE    主檔案路徑
  --new FILE       新檔案路徑
  --output FILE    輸出檔案路徑
  --threshold N    相似度閾值 (0-100, 預設: 80)
  --help           顯示說明
```

## ⚙️ 設定檔範例

### config.json
```json
{
  "master_file": "master.xlsx",
  "new_file": "new.xlsx",
  "output_file": "output.xlsx",
  "similarity_threshold": 80,
  "columns": {
    "name": "姓名",
    "id": "身分證字號",
    "phone": "電話",
    "email": "email",
    "birthday": "生日",
    "created_date": "建立日期"
  },
  "matching_rules": [
    {
      "type": "exact",
      "columns": ["身分證字號"]
    },
    {
      "type": "fuzzy",
      "columns": ["姓名", "電話"],
      "threshold": 85
    }
  ]
}
```

## 📊 輸出檔案說明

### 比對結果包含：
- **完全匹配**：100% 相同的記錄
- **模糊匹配**：相似度達到閾值的記錄
- **新增記錄**：在新檔案中但主檔案中沒有的記錄
- **刪除記錄**：在主檔案中但新檔案中沒有的記錄

### 輸出格式：
- Excel 格式 (.xlsx)
- 多個工作表分別顯示不同類型的比對結果
- 包含相似度分數和匹配類型標記

## 📁 專案結構

```
data-processing-tools/
├── main.py                 # GUI 主程式
├── excel_matcher.py        # 核心比對邏輯
├── excel_processor.py      # Excel 處理模組
├── universal_processor.py  # 通用處理模組
├── dptools/
│   └── cli.py             # CLI 入口點
├── samples/                # 範例檔案
│   ├── master.xlsx
│   ├── new.xlsx
│   └── readme.txt
├── config_template.json    # 設定檔範本
├── requirements.txt        # 依賴套件
├── pyproject.toml         # 專案配置
└── README.md              # 專案說明
```

## 🧪 測試範例

專案包含 `samples/` 資料夾，內有測試用的 Excel 檔案：

1. 將 `samples/master.xlsx` 和 `samples/new.xlsx` 複製到專案根目錄
2. 使用預設設定檔執行比對
3. 查看輸出結果

詳細測試說明請參考 `samples/readme.txt`。

## 🤝 貢獻指南

歡迎提交 Issue 和 Pull Request！

1. Fork 本專案
2. 創建功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 開啟 Pull Request

## 📄 授權條款

本專案採用 MIT 授權條款 - 詳見 [LICENSE](LICENSE) 檔案。

## 👨‍💻 作者

**鄭桓** - [GitHub](https://github.com/Avon46)

---

⭐ 如果這個專案對你有幫助，請給個 Star！
