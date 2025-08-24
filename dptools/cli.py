#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Data Processing Tools CLI Entry Point
"""

import sys
import subprocess
import os

def main():
    """主函數：根據參數決定啟動 GUI 或 CLI 模式"""
    
    # 獲取專案根目錄
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    
    if len(sys.argv) == 1:
        # 無參數：啟動 GUI 模式
        print("🚀 啟動 Data Processing Tools GUI 模式...")
        gui_script = os.path.join(project_root, "main.py")
        subprocess.run([sys.executable, gui_script], check=False)
    else:
        # 有參數：啟動 CLI 模式
        print("⌨️  啟動 Data Processing Tools CLI 模式...")
        cli_script = os.path.join(project_root, "excel_matcher.py")
        subprocess.run([sys.executable, cli_script] + sys.argv[1:], check=False)

if __name__ == "__main__":
    main()
