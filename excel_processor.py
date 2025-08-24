# excel_processor.py - Excel 資料處理模組
# 功能：提供 Excel 和 CSV 檔案的讀取、清理和轉換功能
# 支援多種檔案格式和編碼，可自定義欄位選擇和跳過行數

import pandas as pd  # 用於資料處理和分析
import chardet  # 用於自動偵測檔案編碼
# 導入 PyQt5 的 GUI 元件
from PyQt5.QtWidgets import QTableWidget, QTableWidgetItem, QMessageBox
from PyQt5.QtWidgets import QVBoxLayout, QHBoxLayout, QComboBox, QSpinBox, QLineEdit, QPushButton, QFileDialog, QLabel, QWidget

class ExcelProcessor(QWidget):
    """
    Excel 資料處理器類別
    提供圖形化介面來處理 Excel 和 CSV 檔案
    支援多種預設設定和自定義參數
    """
    def __init__(self):
        super().__init__()
        # 設定視窗標題和基本屬性
        self.setWindowTitle('Excel 清理資料')
        self.setGeometry(300, 300, 600, 500)

        # 建立主要的垂直佈局容器
        layout = QVBoxLayout()

        # === 檔案選擇區域 ===
        file_select_layout = QHBoxLayout()
        # 檔案選擇按鈕
        self.select_button = QPushButton('選擇檔案', self)
        self.select_button.setToolTip('選擇要處理的 Excel 或 CSV 檔案')
        self.select_button.clicked.connect(self.select_file)
        file_select_layout.addWidget(self.select_button)
        
        # 檔案路徑顯示框
        self.file_path_edit = QLineEdit(self)
        self.file_path_edit.setReadOnly(True)  # 設為唯讀，防止手動編輯
        self.file_path_edit.setPlaceholderText("請選擇要讀取的檔案")
        file_select_layout.addWidget(self.file_path_edit)
        layout.addLayout(file_select_layout)

        # === 檔案類型選擇區域 ===
        self.file_type_label = QLabel("選擇檔案類型:", self)
        layout.addWidget(self.file_type_label)

        # 檔案類型下拉選單，包含預設的處理設定
        self.file_type_combo = QComboBox(self)
        self.file_type_combo.addItems(["自訂", '香連', "甘妹", "周照子", "炎弟", "威宇", "麻煮", "小旺號", "扶旺號"])
        self.file_type_combo.currentIndexChanged.connect(self.update_settings)  # 當選擇改變時自動更新設定
        layout.addWidget(self.file_type_combo)

        # === 跳過行數設定區域 ===
        self.skiprows_label = QLabel("跳過的列數 (skiprows):", self)
        layout.addWidget(self.skiprows_label)

        # 跳過行數輸入框，設定為數字輸入
        self.skiprows_input = QSpinBox(self)
        self.skiprows_input.setMinimum(0)  # 最小值為 0
        self.skiprows_input.setMaximum(100)  # 最大值為 100
        self.skiprows_input.setValue(7)  # 預設值為 7
        layout.addWidget(self.skiprows_input)

        # === 欄位選擇區域 ===
        self.columns_label = QLabel("要保留的欄位 (用逗號分隔):", self)
        layout.addWidget(self.columns_label)

        # 欄位選擇輸入框，用逗號分隔多個欄位編號
        self.columns_input = QLineEdit(self)
        self.columns_input.setText("2,3")  # 預設選擇第 2 和第 3 欄
        self.columns_input.setPlaceholderText("請輸入欄位數字，用逗號分隔")
        layout.addWidget(self.columns_input)

        # === 編碼選擇區域 ===
        self.encoding_label = QLabel("選擇檔案編碼:", self)
        layout.addWidget(self.encoding_label)

        # 編碼選擇下拉選單
        self.encoding_combo = QComboBox(self)
        self.encoding_combo.addItems(["UTF-8", "UTF-16", "Big5", "GBK", "自動偵測"])
        layout.addWidget(self.encoding_combo)

        # === 處理按鈕 ===
        self.process_button = QPushButton('開始處理', self)
        self.process_button.setToolTip('根據設定的欄位與跳過列數，讀取並預覽檔案內容')
        self.process_button.clicked.connect(self.process_data)
        self.process_button.setEnabled(False)  # 初始狀態為禁用，需要先選擇檔案
        layout.addWidget(self.process_button)

        # === 預覽表格 ===
        self.preview_table = QTableWidget()
        layout.addWidget(self.preview_table)

        # === 儲存按鈕 ===
        self.save_button = QPushButton('確認並儲存', self)
        self.save_button.setToolTip('將處理後的資料儲存為新的 Excel 檔案')
        self.save_button.clicked.connect(self.save_file)
        self.save_button.setEnabled(False)  # 初始狀態為禁用，需要先處理資料
        layout.addWidget(self.save_button)

        # 設定佈局
        self.setLayout(layout)

        # === 初始化成員變數 ===
        self.input_file = None      # 輸入檔案路徑
        self.output_file = None     # 輸出檔案路徑
        self.df_processed = None    # 處理後的資料框架

        # === 預設檔案處理設定 ===
        # 針對不同類型的檔案，預先設定好跳過行數和要保留的欄位
        self.file_settings = {
            "香連": {"skiprows": 17, "columns": [0, 1, 3]},      # 香連檔案：跳過 17 行，保留第 0、1、3 欄
            "甘妹": {"skiprows": 7, "columns": [2, 3]},          # 甘妹檔案：跳過 7 行，保留第 2、3 欄
            "周照子": {"skiprows": 7, "columns": [2, 3]},        # 周照子檔案：跳過 7 行，保留第 2、3 欄
            "炎弟": {"skiprows": 7, "columns": [2, 3]},          # 炎弟檔案：跳過 7 行，保留第 2、3 欄
            "威宇": {"skiprows": 7, "columns": [2, 3]},          # 威宇檔案：跳過 7 行，保留第 2、3 欄
            "麻煮": {"skiprows": 7, "columns": [2, 3]},          # 麻煮檔案：跳過 7 行，保留第 2、3 欄
            "小旺號": {"skiprows": 8, "columns": [2, 3, 6]},     # 小旺號檔案：跳過 8 行，保留第 2、3、6 欄
            "扶旺號": {"skiprows": 10, "columns": [1, 5]}        # 扶旺號檔案：跳過 10 行，保留第 1、5 欄
        }

    def select_file(self):
        """
        檔案選擇方法
        開啟檔案對話框讓使用者選擇要處理的檔案
        """
        options = QFileDialog.Options()
        # 支援 Excel 和 CSV 檔案格式
        self.input_file, _ = QFileDialog.getOpenFileName(
            self, 
            "選擇要讀取的檔案", 
            "", 
            "Excel Files (*.xls *.xlsx);;CSV Files (*.csv)", 
            options=options
        )

        if self.input_file:
            # 如果成功選擇檔案，更新介面狀態
            self.file_path_edit.setText(self.input_file)
            self.process_button.setEnabled(True)  # 啟用處理按鈕
        else:
            # 如果沒有選擇檔案，清空路徑並顯示警告
            self.file_path_edit.setText("")
            QMessageBox.warning(self, "錯誤", "未選擇檔案！")

    def update_settings(self):
        """
        更新設定方法
        當使用者選擇不同的檔案類型時，自動更新相關的處理參數
        """
        selected_type = self.file_type_combo.currentText()
        if selected_type in self.file_settings:
            # 獲取預設設定
            settings = self.file_settings[selected_type]
            # 自動設定跳過行數
            self.skiprows_input.setValue(settings["skiprows"])
            # 自動設定要保留的欄位
            self.columns_input.setText(",".join(map(str, settings["columns"])))

    def process_data(self):
        """
        資料處理方法
        根據使用者設定的參數讀取和處理檔案
        """
        if not self.input_file:
            QMessageBox.warning(self, "錯誤", "請先選擇檔案！")
            return

        # 獲取跳過行數設定
        skiprows = self.skiprows_input.value()
        
        # 解析欄位選擇設定
        try:
            columns = list(map(int, self.columns_input.text().split(',')))
        except ValueError:
            QMessageBox.warning(self, "錯誤", "欄位輸入錯誤，請輸入數字並用逗號分隔！")
            return

        # 獲取編碼設定
        selected_encoding = self.encoding_combo.currentText()
        encoding = None if selected_encoding == "自動偵測" else selected_encoding

        try:
            # 根據檔案類型選擇不同的讀取方法
            if self.input_file.endswith(".csv"):
                # CSV 檔案處理
                if encoding is None:
                    # 如果選擇自動偵測編碼，使用 chardet 來偵測
                    with open(self.input_file, 'rb') as f:
                        encoding = chardet.detect(f.read())['encoding']
                # 讀取 CSV 檔案
                df = pd.read_csv(self.input_file, encoding=encoding, skiprows=skiprows)
            else:
                # Excel 檔案處理
                df = pd.read_excel(self.input_file, skiprows=skiprows)
            
            # 根據使用者選擇的欄位來篩選資料
            self.df_processed = df.iloc[:, columns]
            
            # 顯示處理結果預覽
            self.preview_result()
            # 啟用儲存按鈕
            self.save_button.setEnabled(True)
            
        except Exception as e:
            # 錯誤處理：顯示詳細的錯誤訊息
            QMessageBox.critical(self, "錯誤", f"處理檔案時發生錯誤：{e}")

    def preview_result(self):
        """
        結果預覽方法
        在表格中顯示處理後的資料，讓使用者確認結果
        """
        df = self.df_processed
        if df is None:
            QMessageBox.warning(self, "錯誤", "尚未有可預覽的資料！")
            return
        
        # 設定表格的行數和列數
        self.preview_table.setRowCount(len(df))
        self.preview_table.setColumnCount(len(df.columns))
        # 設定表格標題
        self.preview_table.setHorizontalHeaderLabels([str(col) for col in df.columns])

        # 填充表格資料
        for i in range(len(df)):
            for j in range(len(df.columns)):
                self.preview_table.setItem(i, j, QTableWidgetItem(str(df.iat[i, j])))
        
        # 自動調整欄寬以適應內容
        self.preview_table.resizeColumnsToContents()

    def save_file(self):
        """
        檔案儲存方法
        將處理後的資料儲存為新的 Excel 檔案
        """
        options = QFileDialog.Options()
        # 開啟儲存檔案對話框
        self.output_file, _ = QFileDialog.getSaveFileName(
            self, 
            "儲存檔案", 
            "processed.xlsx", 
            "Excel Files (*.xlsx)", 
            options=options
        )

        if self.output_file:
            # 檢查是否有可儲存的資料
            if self.df_processed is None:
                QMessageBox.warning(self, "錯誤", "沒有可儲存的資料！")
                return
            
            # 確保檔案副檔名為 .xlsx
            if not self.output_file.endswith(".xlsx"):
                self.output_file += ".xlsx"
            
            try:
                # 將資料儲存為 Excel 檔案
                self.df_processed.to_excel(self.output_file, index=False)
                QMessageBox.information(self, "成功", f"資料已儲存到: {self.output_file}")
            except Exception as e:
                # 錯誤處理：顯示儲存失敗的詳細訊息
                QMessageBox.critical(self, "錯誤", f"儲存檔案時發生錯誤：{e}")
