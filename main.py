# main.py - Excel 工具集主程式
# 功能：建立一個包含三個主要功能的桌面應用程式
# 1. Excel 資料處理 - 清理和轉換 Excel/CSV 檔案（針對特定公司）
# 2. Excel VLOOKUP 比對 - 比對兩個 Excel 檔案的資料
# 3. 通用報表處理器 - 智慧化處理各種格式的報表檔案

import sys
# 導入 PyQt5 的 GUI 元件
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QTabWidget, QWidget
# 導入自定義的 Excel 處理和比對模組
from excel_processor import ExcelProcessor
from excel_matcher import ExcelMatcherApp
# 導入新的通用報表處理器
from universal_processor import UniversalProcessor

class MainWindow(QMainWindow):
    """
    主視窗類別
    繼承自 QMainWindow，提供應用程式的主要介面
    """
    def __init__(self):
        super().__init__()

        # 設定視窗標題和基本屬性
        self.setWindowTitle('Excel 工具集 - 通用版')
        self.setGeometry(100, 100, 1200, 800)  # 增加視窗大小以容納新功能

        # 建立主要的垂直佈局容器
        layout = QVBoxLayout()

        # 建立分頁元件，讓使用者可以在不同功能間切換
        self.tabs = QTabWidget()
        
        # 添加第一個分頁：通用報表處理器（推薦新使用者使用）
        self.tabs.addTab(UniversalProcessor(), "🚀 通用報表處理器")
        
        # 添加第二個分頁：Excel 資料處理功能（針對特定公司）
        self.tabs.addTab(ExcelProcessor(), "🏢 特定公司 Excel 處理")
        
        # 添加第三個分頁：Excel VLOOKUP 比對功能
        self.tabs.addTab(ExcelMatcherApp(), "🔍 Excel VLOOKUP 比對")

        # 將分頁元件加入到主佈局中
        layout.addWidget(self.tabs)

        # 建立容器元件來承載佈局
        container = QWidget()
        container.setLayout(layout)
        # 設定為中央元件
        self.setCentralWidget(container)

# 程式進入點
if __name__ == '__main__':
    # 建立 PyQt5 應用程式實例
    app = QApplication(sys.argv)
    # 建立主視窗實例
    window = MainWindow()
    # 顯示主視窗
    window.show()
    # 進入應用程式的事件迴圈，直到使用者關閉視窗
    sys.exit(app.exec_())
