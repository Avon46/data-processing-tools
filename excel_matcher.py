# excel_matcher.py - Excel 資料比對模組
# 功能：提供兩個 Excel 檔案之間的資料比對功能，類似 VLOOKUP 但更強大
# 支援精確比對和模糊比對，可處理文字清理和相似度計算

import pandas as pd  # 用於資料處理和分析
import re  # 用於正則表達式處理
import os  # 用於檔案路徑操作
from datetime import datetime  # 用於時間處理
# 導入 PyQt5 的 GUI 元件
from PyQt5.QtWidgets import (
    QWidget, QLabel, QPushButton, QVBoxLayout, QHBoxLayout, QTableWidget,
    QTableWidgetItem, QFileDialog, QComboBox, QMessageBox,
    QCheckBox, QSpinBox, QLineEdit
)
from PyQt5.QtGui import QColor  # 用於設定顏色
# 導入模糊比對相關功能
from rapidfuzz import process, fuzz
import numpy as np  # 用於數值運算

class ExcelMatcherApp(QWidget):
    """
    Excel 資料比對應用程式類別
    提供圖形化介面來比對兩個 Excel 檔案的資料
    支援精確比對和模糊比對兩種模式
    """
    def __init__(self):
        super().__init__()
        # === 初始化成員變數 ===
        self.file_main = ""        # 主檔案路徑（通常是商品排行榜）
        self.file_clear = ""       # 清理檔案路徑（已經清理過的資料）
        self.df_main = None        # 主檔案的資料框架
        self.df_clear = None       # 清理檔案的資料框架
        self.initUI()  # 初始化使用者介面

    def initUI(self):
        """
        初始化使用者介面方法
        建立所有的 GUI 元件和佈局
        """
        # 設定視窗標題和基本屬性
        self.setWindowTitle("Excel 資料比對")
        self.setGeometry(100, 100, 600, 500)

        # 建立主要的垂直佈局容器
        layout = QVBoxLayout()

        # === 主檔案選擇區域 ===
        main_file_layout = QHBoxLayout()
        # 主檔案選擇按鈕
        self.btn_select_main = QPushButton("選擇 商品排行榜檔案")
        self.btn_select_main.setToolTip("選擇要作為主檔的 Excel 檔案，通常是排行榜或原始資料")
        self.btn_select_main.clicked.connect(self.select_main_file)
        main_file_layout.addWidget(self.btn_select_main)
        
        # 主檔案路徑顯示框
        self.main_file_path = QLineEdit()
        self.main_file_path.setReadOnly(True)  # 設為唯讀，防止手動編輯
        self.main_file_path.setPlaceholderText("未選擇 商品排行榜檔案")
        main_file_layout.addWidget(self.main_file_path)
        layout.addLayout(main_file_layout)

        # 主檔案比對欄位選擇
        layout.addWidget(QLabel("選擇主檔比對欄位 (將用於與清理檔比對):"))
        self.combo_main = QComboBox()
        self.combo_main.setToolTip("選擇主檔中要用來比對的欄位，通常是商品名稱或編號")
        layout.addWidget(self.combo_main)

        # === 清理檔案選擇區域 ===
        clear_file_layout = QHBoxLayout()
        # 清理檔案選擇按鈕
        self.btn_select_clear = QPushButton("選擇 清理後檔案")
        self.btn_select_clear.setToolTip("選擇已經清理過的 Excel 檔案，將與主檔進行比對")
        self.btn_select_clear.clicked.connect(self.select_clear_file)
        clear_file_layout.addWidget(self.btn_select_clear)
        
        # 清理檔案路徑顯示框
        self.clear_file_path = QLineEdit()
        self.clear_file_path.setReadOnly(True)  # 設為唯讀，防止手動編輯
        self.clear_file_path.setPlaceholderText("未選擇 清理後檔案")
        clear_file_layout.addWidget(self.clear_file_path)
        layout.addLayout(clear_file_layout)

        # 清理檔案比對欄位選擇
        layout.addWidget(QLabel("選擇清理檔比對欄位 (將與主檔欄位進行比對):"))
        self.combo_clear = QComboBox()
        self.combo_clear.setToolTip("選擇清理檔中要用來比對的欄位，通常是商品名稱或編號")
        layout.addWidget(self.combo_clear)

        # === 比對設定區域 ===
        # 模糊比對開關
        self.fuzzy_checkbox = QCheckBox("啟用模糊比對")
        self.fuzzy_checkbox.setChecked(False)  # 預設關閉

        # 相似度門檻設定
        self.threshold_spinner = QSpinBox()
        self.threshold_spinner.setRange(0, 100)  # 設定範圍 0-100
        self.threshold_spinner.setValue(85)  # 預設值 85%

        # === 執行按鈕 ===
        self.btn_match = QPushButton("執行比對")
        self.btn_match.setToolTip("開始進行主檔與清理檔的欄位比對，並產生結果預覽與檔案")
        self.btn_match.clicked.connect(self.run_matching)

        # === 結果預覽表格 ===
        self.table_preview = QTableWidget()

        # 將所有元件加入到佈局中
        layout.addWidget(self.fuzzy_checkbox)
        layout.addWidget(QLabel("相似度門檻:"))
        layout.addWidget(self.threshold_spinner)
        layout.addWidget(self.btn_match)
        layout.addWidget(QLabel("比對結果預覽:"))
        layout.addWidget(self.table_preview)

        # 設定佈局
        self.setLayout(layout)

    def select_main_file(self):
        """
        選擇主檔案方法
        開啟檔案對話框讓使用者選擇主檔案（通常是商品排行榜）
        """
        file_name, _ = QFileDialog.getOpenFileName(self, "選擇 商品排行榜檔案", "", "Excel Files (*.xls *.xlsx)")
        if file_name:
            try:
                # 儲存檔案路徑
                self.file_main = file_name
                self.main_file_path.setText(file_name)
                
                # 讀取 Excel 檔案
                self.df_main = pd.read_excel(file_name)
                
                # 清空並重新填充欄位選擇下拉選單
                self.combo_main.clear()
                self.combo_main.addItems([str(col) for col in self.df_main.columns])
                
            except Exception as e:
                # 錯誤處理：顯示讀取失敗的詳細訊息
                QMessageBox.critical(self, "錯誤", f"讀取主檔案失敗:\n{str(e)}")
                # 重置相關變數
                self.df_main = None
                self.combo_main.clear()
                self.main_file_path.setText("")

    def select_clear_file(self):
        """
        選擇清理檔案方法
        開啟檔案對話框讓使用者選擇清理檔案（已經處理過的資料）
        """
        file_name, _ = QFileDialog.getOpenFileName(self, "選擇 清理後檔案", "", "Excel Files (*.xls *.xlsx)")
        if file_name:
            try:
                # 儲存檔案路徑
                self.file_clear = file_name
                self.clear_file_path.setText(file_name)
                
                # 讀取 Excel 檔案
                self.df_clear = pd.read_excel(file_name)
                
                # 清空並重新填充欄位選擇下拉選單
                self.combo_clear.clear()
                self.combo_clear.addItems([str(col) for col in self.df_clear.columns])
                
            except Exception as e:
                # 錯誤處理：顯示讀取失敗的詳細訊息
                QMessageBox.critical(self, "錯誤", f"讀取清理檔案失敗:\n{str(e)}")
                # 重置相關變數
                self.df_clear = None
                self.combo_clear.clear()
                self.clear_file_path.setText("")

    def clean_text(self, text):
        """
        文字清理方法
        將文字進行標準化處理，移除空白字元、全形空格等，並轉為小寫
        這樣可以提高比對的準確性
        
        參數：
            text: 要清理的文字
        
        返回：
            清理後的文字字串
        """
        if pd.isna(text):  # 檢查是否為空值
            return ""
        # 移除空白字元、全形空格，並轉為小寫
        return re.sub(r'\s+|\u3000', '', str(text).strip().lower())

    def run_matching(self):
        """
        執行比對方法
        主要的比對邏輯，包含精確比對和模糊比對兩種模式
        """
        # 檢查是否已選擇兩個檔案
        if not self.file_main or not self.file_clear or self.df_main is None or self.df_clear is None:
            QMessageBox.warning(self, "錯誤", "請先選擇兩個 Excel 檔案！")
            return

        try:
            # 獲取選擇的比對欄位
            main_key = self.combo_main.currentText()
            clear_key = self.combo_clear.currentText()

            # 驗證欄位是否存在
            if main_key not in self.df_main.columns or clear_key not in self.df_clear.columns:
                QMessageBox.warning(self, "錯誤", "請選擇正確的比對欄位！")
                return

            # 複製資料框架以避免修改原始資料
            df_main = self.df_main.copy()
            df_clear = self.df_clear.copy()

            # 為兩個資料框架添加清理後的比對鍵值欄位
            df_main['__clean_key__'] = df_main[main_key].apply(self.clean_text)
            df_clear['__clean_key__'] = df_clear[clear_key].apply(self.clean_text)

            # 獲取比對設定
            use_fuzzy = self.fuzzy_checkbox.isChecked()  # 是否啟用模糊比對
            threshold = self.threshold_spinner.value()    # 相似度門檻

            # 初始化結果儲存變數
            matched_rows = []  # 儲存比對成功的行
            fuzzy_flags = []   # 儲存是否為模糊比對的標記

            # 獲取清理檔的比對鍵值列表和集合（用於加速查詢）
            clear_keys = df_clear['__clean_key__'].tolist()
            clear_key_set = set(clear_keys)

            # === 主要的比對迴圈 ===
            for idx, row in df_main.iterrows():
                val = row['__clean_key__']
                
                # 處理複雜的資料類型（如 Series、list、tuple）
                if isinstance(val, (pd.Series, list, tuple)):
                    v = val[0] if len(val) > 0 else ""
                    is_na = pd.isna(v)
                    if not isinstance(is_na, (bool, np.bool_)):
                        is_na = np.all(is_na)
                    key_main = str(v) if not is_na else ""
                else:
                    is_na = pd.isna(val)
                    if not isinstance(is_na, (bool, np.bool_)):
                        is_na = np.all(is_na)
                    key_main = str(val) if not is_na else ""

                # === 比對邏輯 ===
                # 1. 精確比對：使用 pandas 的 mask 進行快速比對
                mask = df_clear['__clean_key__'] == key_main
                is_match = mask.sum() > 0
                
                if is_match:
                    # 找到精確匹配，取第一筆結果
                    exact_match = df_clear[mask]
                    matched_rows.append(exact_match.iloc[0].to_dict())
                    fuzzy_flags.append(False)  # 標記為精確比對
                    
                elif key_main in clear_key_set:
                    # 在集合中找到匹配（備用檢查）
                    matched_rows.append({col: None for col in df_clear.columns})
                    fuzzy_flags.append(False)
                    
                elif use_fuzzy and len(clear_keys) > 0:
                    # 2. 模糊比對：使用 rapidfuzz 進行相似度計算
                    # process.extractOne 需要傳入 list[str] 格式
                    match_result = process.extractOne(key_main, list(map(str, clear_keys)), scorer=fuzz.ratio)
                    
                    if match_result is not None:
                        match, score, match_idx = match_result
                        
                        if score >= threshold:
                            # 相似度達到門檻，記錄匹配結果
                            matched_rows.append(df_clear.iloc[match_idx].to_dict())
                            fuzzy_flags.append(True)  # 標記為模糊比對
                        else:
                            # 相似度未達門檻，記錄為無匹配
                            matched_rows.append({col: None for col in df_clear.columns})
                            fuzzy_flags.append(False)
                    else:
                        # 模糊比對失敗，記錄為無匹配
                        matched_rows.append({col: None for col in df_clear.columns})
                        fuzzy_flags.append(False)
                else:
                    # 無匹配且未啟用模糊比對，記錄為無匹配
                    matched_rows.append({col: None for col in df_clear.columns})
                    fuzzy_flags.append(False)

            # === 建立結果資料框架 ===
            # 確保 main_key 是單一欄位名稱，並重新命名為 "KEY"
            main_col = df_main[[main_key]].copy()
            main_col.columns = ["KEY"]
            
            # 將主檔的比對欄位和清理檔的匹配結果合併
            df_result = pd.concat([
                main_col.reset_index(drop=True),
                pd.DataFrame(matched_rows).reset_index(drop=True)
            ], axis=1)
            
            # 添加模糊比對標記欄位
            df_result["是否模糊比對"] = fuzzy_flags

            # === 建立結果儲存資料夾 ===
            output_dir = os.path.join(os.path.dirname(self.file_clear), "比對結果")
            os.makedirs(output_dir, exist_ok=True)

            # === 設定檔案名稱 ===
            base_name = os.path.splitext(os.path.basename(self.file_main))[0]
            result_file = os.path.join(output_dir, f"{base_name}_比對結果.xlsx")
            unmatched_file = os.path.join(output_dir, f"{base_name}_未匹配.xlsx")

            # === 儲存結果檔案 ===
            # 儲存完整的比對結果
            df_result.to_excel(result_file, index=False)

            # 找出未匹配的記錄（在清理檔中但不在主檔中的記錄）
            unmatched_mask = ~df_clear['__clean_key__'].isin(df_main['__clean_key__'])
            df_unmatched = df_clear[unmatched_mask].drop(columns=['__clean_key__'])
            df_unmatched.to_excel(unmatched_file, index=False)

            # === 顯示預覽和結果 ===
            # 在表格中預覽前 10 筆結果
            self.preview_result(df_result)

            # 顯示成功訊息和統計資訊
            QMessageBox.information(
                self,
                "成功",
                f"比對完成！\n結果已儲存於：\n{output_dir}\n\n未匹配筆數: {df_unmatched.shape[0]}"
            )

        except Exception as e:
            # 錯誤處理：顯示詳細的錯誤訊息和堆疊追蹤
            import traceback
            QMessageBox.critical(self, "錯誤", f"執行過程發生錯誤:\n{str(e)}\n{traceback.format_exc()}")

    def preview_result(self, df):
        """
        結果預覽方法
        在表格中顯示比對結果的前 10 筆，並用顏色標示模糊比對的結果
        
        參數：
            df: 要預覽的結果資料框架
        """
        # 取前 10 筆資料進行預覽
        preview_df = df.head(10)
        
        # 設定表格的行數和列數
        self.table_preview.setRowCount(len(preview_df))
        self.table_preview.setColumnCount(len(preview_df.columns))
        # 設定表格標題
        self.table_preview.setHorizontalHeaderLabels([str(col) for col in preview_df.columns])

        # 填充表格資料
        for row in range(len(preview_df)):
            # 檢查該行是否為模糊比對
            is_fuzzy = preview_df.at[row, "是否模糊比對"]
            
            for col in range(len(preview_df.columns)):
                # 建立表格項目
                item = QTableWidgetItem(str(preview_df.iat[row, col]))
                
                # 如果是模糊比對且不是標記欄位，用黃色背景標示
                if is_fuzzy and preview_df.columns[col] != "是否模糊比對":
                    item.setBackground(QColor("yellow"))  # 用底色標示模糊比對
                
                # 將項目加入到表格中
                self.table_preview.setItem(row, col, item)
        
        # 自動調整欄寬以適應內容
        self.table_preview.resizeColumnsToContents()
