# universal_processor.py - 通用報表處理器
# 功能：提供智慧化的報表格式偵測和處理功能
# 支援多種報表格式，可自動偵測表頭位置和欄位類型

import pandas as pd
import chardet
import json
import os
import re
from typing import Dict, List, Tuple, Optional
from PyQt5.QtWidgets import (
    QWidget, QLabel, QPushButton, QVBoxLayout, QHBoxLayout, QComboBox, 
    QSpinBox, QLineEdit, QPushButton, QFileDialog, QTextEdit, QMessageBox,
    QTableWidget, QTableWidgetItem, QCheckBox, QGroupBox, QGridLayout
)
from PyQt5.QtGui import QFont, QColor
from PyQt5.QtCore import Qt

class UniversalProcessor(QWidget):
    """
    通用報表處理器類別
    提供智慧化的報表格式偵測和處理功能
    支援多種報表格式，可自動偵測表頭位置和欄位類型
    """
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle('通用報表處理器')
        self.setGeometry(100, 100, 1000, 700)
        
        # 初始化變數
        self.input_file = None
        self.df_raw = None
        self.df_processed = None
        self.config = self.load_config()
        
        # 初始化介面
        self.initUI()
        
    def load_config(self) -> Dict:
        """
        載入設定檔案
        如果設定檔案不存在，則使用預設設定
        """
        config_file = "config_template.json"
        if os.path.exists(config_file):
            try:
                with open(config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print(f"載入設定檔案失敗: {e}")
        
        # 返回預設設定
        return {
            "預設設定": {
                "通用報表": {"skiprows": 0, "columns": "all", "encoding": "auto"},
                "有表頭的報表": {"skiprows": 1, "columns": "all", "encoding": "auto"},
                "複雜報表": {"skiprows": 5, "columns": "2,3,4", "encoding": "auto"}
            }
        }
    
    def initUI(self):
        """初始化使用者介面"""
        layout = QVBoxLayout()
        
        # === 檔案選擇區域 ===
        file_group = QGroupBox("檔案選擇")
        file_layout = QGridLayout()
        
        # 檔案選擇按鈕
        self.select_button = QPushButton('選擇報表檔案')
        self.select_button.clicked.connect(self.select_file)
        file_layout.addWidget(self.select_button, 0, 0)
        
        # 檔案路徑顯示
        self.file_path_edit = QLineEdit()
        self.file_path_edit.setReadOnly(True)
        self.file_path_edit.setPlaceholderText("請選擇要處理的報表檔案")
        file_layout.addWidget(self.file_path_edit, 0, 1)
        
        # 檔案資訊顯示
        self.file_info_label = QLabel("未選擇檔案")
        self.file_info_label.setStyleSheet("color: gray;")
        file_layout.addWidget(self.file_info_label, 1, 0, 1, 2)
        
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        # === 智慧偵測區域 ===
        detect_group = QGroupBox("智慧偵測結果")
        detect_layout = QVBoxLayout()
        
        # 偵測結果顯示
        self.detect_text = QTextEdit()
        self.detect_text.setMaximumHeight(150)
        self.detect_text.setReadOnly(True)
        detect_layout.addWidget(self.detect_text)
        
        # 自動偵測按鈕
        self.detect_button = QPushButton('開始智慧偵測')
        self.detect_button.clicked.connect(self.smart_detect)
        self.detect_button.setEnabled(False)
        detect_layout.addWidget(self.detect_button)
        
        detect_group.setLayout(detect_layout)
        layout.addWidget(detect_group)
        
        # === 處理設定區域 ===
        settings_group = QGroupBox("處理設定")
        settings_layout = QGridLayout()
        
        # 跳過行數設定
        settings_layout.addWidget(QLabel("跳過行數:"), 0, 0)
        self.skiprows_input = QSpinBox()
        self.skiprows_input.setRange(0, 100)
        self.skiprows_input.setValue(0)
        settings_layout.addWidget(self.skiprows_input, 0, 1)
        
        # 欄位選擇設定
        settings_layout.addWidget(QLabel("選擇欄位:"), 1, 0)
        self.columns_input = QLineEdit()
        self.columns_input.setText("all")
        self.columns_input.setPlaceholderText("輸入欄位編號或用 'all' 選擇全部")
        settings_layout.addWidget(self.columns_input, 1, 1)
        
        # 編碼設定
        settings_layout.addWidget(QLabel("檔案編碼:"), 2, 0)
        self.encoding_combo = QComboBox()
        self.encoding_combo.addItems(["自動偵測", "UTF-8", "UTF-16", "Big5", "GBK"])
        settings_layout.addWidget(self.encoding_combo, 2, 1)
        
        # 預設設定選擇
        settings_layout.addWidget(QLabel("預設設定:"), 3, 0)
        self.preset_combo = QComboBox()
        self.preset_combo.addItems(list(self.config["預設設定"].keys()))
        self.preset_combo.currentTextChanged.connect(self.apply_preset)
        settings_layout.addWidget(self.preset_combo, 3, 1)
        
        settings_group.setLayout(settings_layout)
        layout.addWidget(settings_group)
        
        # === 處理按鈕 ===
        button_layout = QHBoxLayout()
        
        self.process_button = QPushButton('開始處理')
        self.process_button.clicked.connect(self.process_data)
        self.process_button.setEnabled(False)
        button_layout.addWidget(self.process_button)
        
        self.save_button = QPushButton('儲存結果')
        self.save_button.clicked.connect(self.save_file)
        self.save_button.setEnabled(False)
        button_layout.addWidget(self.save_button)
        
        layout.addLayout(button_layout)
        
        # === 預覽區域 ===
        preview_group = QGroupBox("資料預覽")
        preview_layout = QVBoxLayout()
        
        self.preview_table = QTableWidget()
        preview_layout.addWidget(self.preview_table)
        
        preview_group.setLayout(preview_layout)
        layout.addWidget(preview_group)
        
        self.setLayout(layout)
    
    def select_file(self):
        """選擇檔案"""
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(
            self, 
            "選擇報表檔案", 
            "", 
            "所有支援的檔案 (*.xlsx *.xls *.csv *.txt);;Excel 檔案 (*.xlsx *.xls);;CSV 檔案 (*.csv);;文字檔案 (*.txt)", 
            options=options
        )
        
        if file_name:
            self.input_file = file_name
            self.file_path_edit.setText(file_name)
            self.file_info_label.setText(f"已選擇: {os.path.basename(file_name)}")
            self.detect_button.setEnabled(True)
            self.process_button.setEnabled(True)
        else:
            self.file_info_label.setText("未選擇檔案")
    
    def smart_detect(self):
        """智慧偵測報表格式"""
        if not self.input_file:
            return
            
        try:
            # 讀取檔案進行偵測
            self.df_raw = self.read_file_smart()
            
            if self.df_raw is None:
                return
            
            # 分析檔案結構
            analysis = self.analyze_file_structure()
            
            # 顯示偵測結果
            self.display_detection_results(analysis)
            
            # 自動套用最佳設定
            self.apply_best_settings(analysis)
            
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"智慧偵測失敗: {str(e)}")
    
    def read_file_smart(self) -> Optional[pd.DataFrame]:
        """智慧讀取檔案"""
        try:
            # 自動偵測編碼
            if self.encoding_combo.currentText() == "自動偵測":
                with open(self.input_file, 'rb') as f:
                    result = chardet.detect(f.read())
                    encoding = result['encoding']
            else:
                encoding = self.encoding_combo.currentText()
            
            # 根據檔案類型讀取
            if self.input_file.endswith('.csv'):
                # 嘗試不同的分隔符號
                for sep in [',', '\t', ';', '|']:
                    try:
                        df = pd.read_csv(self.input_file, encoding=encoding, sep=sep, nrows=20)
                        if len(df.columns) > 1:  # 成功分割
                            break
                    except:
                        continue
                else:
                    # 如果都失敗，使用預設逗號
                    df = pd.read_csv(self.input_file, encoding=encoding, nrows=20)
            else:
                df = pd.read_excel(self.input_file, nrows=20)
            
            return df
            
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"讀取檔案失敗: {str(e)}")
            return None
    
    def analyze_file_structure(self) -> Dict:
        """分析檔案結構"""
        analysis = {
            "總行數": len(self.df_raw),
            "總欄數": len(self.df_raw.columns),
            "建議跳過行數": 0,
            "建議欄位": [],
            "欄位類型": {},
            "資料品質": {}
        }
        
        # 分析每一行，找出可能的表頭位置
        header_scores = []
        for i in range(min(10, len(self.df_raw))):
            score = self.calculate_header_score(self.df_raw.iloc[i])
            header_scores.append((i, score))
        
        # 找出最佳表頭位置
        best_header_row = max(header_scores, key=lambda x: x[1])[0]
        analysis["建議跳過行數"] = best_header_row
        
        # 分析欄位類型
        for col in self.df_raw.columns:
            col_data = self.df_raw.iloc[best_header_row:][col]
            col_type = self.infer_column_type(col_data)
            analysis["欄位類型"][col] = col_type
        
        # 建議要保留的欄位
        analysis["建議欄位"] = list(range(len(self.df_raw.columns)))
        
        return analysis
    
    def calculate_header_score(self, row) -> float:
        """計算某一行作為表頭的可能性分數"""
        score = 0
        
        for val in row:
            val_str = str(val).lower()
            
            # 檢查是否包含常見的表頭關鍵字
            header_keywords = ['名稱', '編號', '數量', '價格', '日期', '時間', 'name', 'id', 'qty', 'price', 'date']
            if any(keyword in val_str for keyword in header_keywords):
                score += 2
            
            # 檢查是否為文字（表頭通常是文字）
            if isinstance(val, str) and not val.isdigit():
                score += 1
            
            # 檢查是否為空值（表頭通常不為空）
            if pd.notna(val):
                score += 1
        
        return score
    
    def infer_column_type(self, data) -> str:
        """推斷欄位類型"""
        # 移除空值
        clean_data = data.dropna()
        
        if len(clean_data) == 0:
            return "未知"
        
        # 檢查是否為數字
        numeric_count = sum(pd.to_numeric(clean_data, errors='coerce').notna())
        if numeric_count / len(clean_data) > 0.8:
            return "數值"
        
        # 檢查是否為日期
        date_count = sum(pd.to_datetime(clean_data, errors='coerce').notna())
        if date_count / len(clean_data) > 0.8:
            return "日期"
        
        # 檢查是否為布林值
        bool_count = sum(clean_data.astype(str).str.lower().isin(['true', 'false', '是', '否', '1', '0']))
        if bool_count / len(clean_data) > 0.8:
            return "布林值"
        
        return "文字"
    
    def display_detection_results(self, analysis: Dict):
        """顯示偵測結果"""
        result_text = f"""
=== 智慧偵測結果 ===

📊 檔案基本資訊:
   - 總行數: {analysis['總行數']}
   - 總欄數: {analysis['總欄數']}

🎯 建議設定:
   - 建議跳過行數: {analysis['建議跳過行數']}
   - 建議保留欄位: 全部欄位

📋 欄位類型分析:
"""
        
        for col, col_type in analysis['欄位類型'].items():
            result_text += f"   - {col}: {col_type}\n"
        
        result_text += "\n💡 建議: 系統已自動套用最佳設定，您也可以手動調整。"
        
        self.detect_text.setText(result_text)
    
    def apply_best_settings(self, analysis: Dict):
        """套用最佳設定"""
        self.skiprows_input.setValue(analysis['建議跳過行數'])
        self.columns_input.setText("all")
    
    def apply_preset(self):
        """套用預設設定"""
        preset_name = self.preset_combo.currentText()
        if preset_name in self.config["預設設定"]:
            preset = self.config["預設設定"][preset_name]
            self.skiprows_input.setValue(preset.get("skiprows", 0))
            self.columns_input.setText(str(preset.get("columns", "all")))
            encoding = preset.get("encoding", "auto")
            if encoding == "auto":
                self.encoding_combo.setCurrentText("自動偵測")
            else:
                self.encoding_combo.setCurrentText(encoding)
    
    def process_data(self):
        """處理資料"""
        if not self.input_file or self.df_raw is None:
            QMessageBox.warning(self, "錯誤", "請先選擇檔案並進行智慧偵測！")
            return
        
        try:
            # 獲取設定
            skiprows = self.skiprows_input.value()
            columns_input = self.columns_input.text()
            encoding = self.encoding_combo.currentText()
            
            # 讀取完整檔案
            if self.input_file.endswith('.csv'):
                if encoding == "自動偵測":
                    with open(self.input_file, 'rb') as f:
                        result = chardet.detect(f.read())
                        encoding = result['encoding']
                df = pd.read_csv(self.input_file, encoding=encoding, skiprows=skiprows)
            else:
                df = pd.read_excel(self.input_file, skiprows=skiprows)
            
            # 處理欄位選擇
            if columns_input.lower() == "all":
                self.df_processed = df
            else:
                try:
                    columns = [int(x.strip()) for x in columns_input.split(',')]
                    self.df_processed = df.iloc[:, columns]
                except ValueError:
                    QMessageBox.warning(self, "錯誤", "欄位格式錯誤，請使用數字並用逗號分隔！")
                    return
            
            # 顯示預覽
            self.preview_result()
            self.save_button.setEnabled(True)
            
            QMessageBox.information(self, "成功", "資料處理完成！")
            
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"處理資料時發生錯誤: {str(e)}")
    
    def preview_result(self):
        """預覽結果"""
        if self.df_processed is None:
            return
        
        # 設定表格
        preview_df = self.df_processed.head(20)  # 顯示前20行
        self.preview_table.setRowCount(len(preview_df))
        self.preview_table.setColumnCount(len(preview_df.columns))
        self.preview_table.setHorizontalHeaderLabels([str(col) for col in preview_df.columns])
        
        # 填充資料
        for i in range(len(preview_df)):
            for j in range(len(preview_df.columns)):
                item = QTableWidgetItem(str(preview_df.iat[i, j]))
                self.preview_table.setItem(i, j, item)
        
        self.preview_table.resizeColumnsToContents()
    
    def save_file(self):
        """儲存檔案"""
        if self.df_processed is None:
            QMessageBox.warning(self, "錯誤", "沒有可儲存的資料！")
            return
        
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(
            self, 
            "儲存處理結果", 
            "processed_report.xlsx", 
            "Excel 檔案 (*.xlsx);;CSV 檔案 (*.csv)", 
            options=options
        )
        
        if file_name:
            try:
                if file_name.endswith('.csv'):
                    self.df_processed.to_csv(file_name, index=False, encoding='utf-8-sig')
                else:
                    if not file_name.endswith('.xlsx'):
                        file_name += '.xlsx'
                    self.df_processed.to_excel(file_name, index=False)
                
                QMessageBox.information(self, "成功", f"檔案已儲存至: {file_name}")
                
            except Exception as e:
                QMessageBox.critical(self, "錯誤", f"儲存檔案時發生錯誤: {str(e)}")
