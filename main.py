import tkinter as tk
import ctypes
import sys
import os

from src.ui.main_window import ExcelComparatorApp

def set_app_user_model_id():
    """设置 Windows 任务栏图标 AppUserModelID"""
    try:
        # 格式：公司名.产品名.子模块.版本号
        myappid = 'ExcelCompareTool.Main.V1_0_2' 
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except Exception:
        pass

def get_version():
    """从 updates_notes.txt 中读取最新版本号"""
    version = "V1.0.2"
    try:
        if os.path.exists("updates_notes.txt"):
            with open("updates_notes.txt", "r", encoding="utf-8") as f:
                lines = f.readlines()
                for line in lines:
                    if line.startswith("版本:"):
                        version = line.split(":")[1].strip()
                        break
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
