为了方便其他项目参考，我将本项目中这套**“全平台兼容的 Python GUI 图标设置方案”**总结为以下三个关键步骤。这套方案特别解决了 Python 程序打包后图标丢失、以及 Windows 任务栏图标显示为 Python 默认“大路灯”的问题。

### **第一步：设置 AppUserModelID (解决任务栏分组问题)**
在程序入口文件（如 `main.py`）的最开始处调用。这能告诉 Windows 该进程是一个独立的应用程序，不要把它和 Python 解释器合并。

```python
import ctypes

def set_app_user_model_id():
    try:
        # 格式：公司名.产品名.子模块.版本号
        myappid = 'mycompany.myproduct.v1_0' 
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except Exception:
        pass

if __name__ == "__main__":
    set_app_user_model_id()
    # 启动 GUI 逻辑...
```
参考代码：[main.py:L34-50](file:///e:/code/需求文档出厂参数项提取脚本/main.py#L34-50)

---

### **第二步：编写资源路径兼容函数 (解决路径找不到问题)**
为了让代码在“直接运行”和“PyInstaller 打包后运行”都能找到图标，需要一个自适应函数：

```python
import os
import sys

def get_resource_path(relative_path):
    """获取资源文件的绝对路径"""
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller 打包后的临时目录
        return os.path.join(sys._MEIPASS, relative_path)
    
    # 开发环境：假设图标在项目根目录或 assets 目录下
    base_path = os.path.dirname(os.path.abspath(__file__))
    # 根据实际目录结构调整层级
    project_root = os.path.dirname(os.path.dirname(base_path)) 
    return os.path.join(project_root, "assets", relative_path)
```
参考代码：[main_window.py:L174-201](file:///e:/code/需求文档出厂参数项提取脚本/src/ui/main_window.py#L174-201)

---

### **第三步：三重图标设置逻辑 (核心兼容代码)**
在 GUI 初始化时，按顺序执行以下三种设置方式。建议封装成一个方法：

```python
from PIL import Image, ImageTk

def set_window_icon(root, icon_path):
    if not os.path.exists(icon_path):
        return

    # 1. 基础方式：设置标题栏图标
    try:
        root.iconbitmap(icon_path)
    except Exception: pass

    # 2. 现代方式：设置任务栏和多分辨率图标 (需安装 Pillow)
    try:
        img = Image.open(icon_path)
        photo = ImageTk.PhotoImage(img)
        root.iconphoto(True, photo)
        # 注意：必须保留 photo 的引用，否则会被垃圾回收导致图标消失
        root._icon_photo = photo 
    except Exception: pass

    # 3. Windows 底层方式：强行刷新任务栏图标
    if os.name == 'nt':
        try:
            import ctypes
            hwnd = root.winfo_id()
            # 加载图标资源
            hicon = ctypes.windll.user32.LoadImageW(
                None, icon_path, 1, 0, 0, 0x0010 | 0x0040
            )
            if hicon:
                # 发送消息设置大图标(1)和小图标(0)
                ctypes.windll.user32.SendMessageW(hwnd, 0x0080, 0, hicon)
                ctypes.windll.user32.SendMessageW(hwnd, 0x0080, 1, hicon)
        except Exception: pass
```
参考代码：[main_window.py:L203-251](file:///e:/code/需求文档出厂参数项提取脚本/src/ui/main_window.py#L203-251)

---

### **经验总结**
- **文件格式**：务必使用标准 `.ico` 格式（包含多尺寸分辨率的包），不要简单地把 `.png` 改名。
- **依赖库**：推荐使用 `Pillow` 库来处理图标，它对透明度和多尺寸支持最好。
- **打包注意**：如果使用 PyInstaller，记得在 `.spec` 文件或命令行中添加 `--add-data "assets/favicon.ico;assets"`。