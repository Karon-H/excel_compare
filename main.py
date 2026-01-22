import ctypes
import sys
import os

# 1. 强制设置日志输出，确保能看到报错
print("--- 程序启动调试 ---")

def setup_qt_environment():
    """手动设置 Qt 插件路径，防止 'windows' 平台插件找不到的错误"""
    try:
        import PyQt5
        base_path = os.path.dirname(PyQt5.__file__)
        # 常见路径 1: PyQt5/Qt5/plugins
        path1 = os.path.join(base_path, "Qt5", "plugins")
        # 常见路径 2: PyQt5/Qt/plugins
        path2 = os.path.join(base_path, "Qt", "plugins")
        
        plugin_path = None
        if os.path.exists(path1):
            plugin_path = path1
        elif os.path.exists(path2):
            plugin_path = path2
        else:
            try:
                import PyQt5_Qt5
                path3 = os.path.join(os.path.dirname(PyQt5_Qt5.__file__), "Qt", "plugins")
                if os.path.exists(path3):
                    plugin_path = path3
            except ImportError:
                pass
            
        if plugin_path:
            os.environ["QT_QPA_PLATFORM_PLUGIN_PATH"] = plugin_path
            print(f"DEBUG: 成功设置 Qt 插件路径: {plugin_path}")
        else:
            print("DEBUG: 警告：未找到 Qt 插件路径")
            
    except Exception as e:
        print(f"DEBUG: 设置环境时出错: {e}")

# 执行环境设置
setup_qt_environment()

try:
    print("DEBUG: 正在导入 PyQt5 模块...")
    from PyQt5 import QtWidgets, QtCore, QtGui
    print("DEBUG: 导入 PyQt5 成功")
    
    print("DEBUG: 正在导入 MainWindow...")
    from src.ui.main_window import MainWindow
    print("DEBUG: 导入 MainWindow 成功")
except Exception as e:
    print(f"DEBUG: 导入失败: {e}")
    sys.exit(1)

def set_app_user_model_id():
    try:
        myappid = 'ExcelCompareTool.Main.V3_0_0'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except Exception:
        pass

def get_version():
    version = "V1.9.0"
    try:
        if os.path.exists("updates_notes.txt"):
            with open("updates_notes.txt", "r", encoding="utf-8") as f:
                first_line = f.readline().strip()
                if first_line.startswith('V'):
                    version = first_line.split(' ')[0]
    except Exception:
        pass
    return version

if __name__ == "__main__":
    print("DEBUG: 进入入口函数")
    set_app_user_model_id()
    version = get_version()
    print(f"DEBUG: 当前版本: {version}")
    
    try:
        # 设置高 DPI 支持
        QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
        QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps)
        
        app = QtWidgets.QApplication(sys.argv)
        print("DEBUG: QApplication 实例创建成功")
        
        window = MainWindow(version)
        print("DEBUG: MainWindow 实例创建成功")
        
        window.resize(1100, 850)
        window.show()
        print("DEBUG: Window 已执行 show()")
        
        print("DEBUG: 进入事件循环...")
        sys.exit(app.exec_())
    except Exception as e:
        print(f"DEBUG: 运行时发生严重错误: {e}")
        import traceback
        traceback.print_exc()
        input("按回车键退出...") # 保持窗口以便查看错误
