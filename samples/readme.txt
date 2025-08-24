# 測試範例說明

這個資料夾包含測試用的 Excel 檔案，可以用來測試 excel_matcher.py 的功能。

## 檔案說明

### master.xlsx
- 主檔案，包含 5 筆測試資料
- 欄位：姓名、身分證字號、電話、email、生日、建立日期
- 用於比對的基準資料

### new.xlsx  
- 新檔案，包含 5 筆測試資料
- 與 master.xlsx 有部分重複，部分新增
- 用於測試比對功能

## 測試步驟

1. 將 samples/ 資料夾中的 master.xlsx 和 new.xlsx 複製到專案根目錄
2. 確保 config_template.json 中的檔案路徑設定正確
3. 執行比對：
   ```bash
   python excel_matcher.py --config config_template.json
   ```
4. 查看輸出的 output.xlsx 檔案，檢查比對結果

## 預期結果

- 完全匹配：相同的記錄（身分證字號相同）
- 模糊匹配：相似的記錄（姓名或電話相似）
- 新增記錄：只在 new.xlsx 中存在的記錄
- 刪除記錄：只在 master.xlsx 中存在的記錄

## 注意事項

- 這些是測試用的假資料，僅供功能測試
- 實際使用時請替換為真實的資料檔案
- 建議先用小檔案測試，確認設定正確後再處理大檔案
