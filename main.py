import tkinter as tk
import ctypes
import sys
import os

from src.ui.main_window import ExcelComparatorApp

def set_app_user_model_id():
    """设置 Windows 任务栏图标 AppUserModelID"""
    try:
        # 格式：公司名.产品名.子模块.版本号
        myappid = 'ExcelCompareTool.Main.V1_6_2' 
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except Exception:
        pass

def get_version():
    """从 updates_notes.txt 中读取最新版本号"""
    version = "V1.6.2"
    try:
        if os.path.exists("updates_notes.txt"):
            with open("updates_notes.txt", "r", encoding="utf-8") as f:
                first_line = f.readline().strip()
                if first_line.startswith('V'):
                    # 格式：V1.2.0 2025-12-26 -> 提取 V1.2.0
                    version = first_line.split(' ')[0]
    except Exception:
        pass
    return version

if __name__ == "__main__":
    set_app_user_model_id()
    
    root = tk.Tk()
    version = get_version()
    root.title(f"Excel 差异比对工具 - {version}")
    
    app = ExcelComparatorApp(root)
    root.mainloop()
