import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import sys
import ctypes
from PIL import Image, ImageTk
from src.logic.excel_processor import ExcelDiffer

class ExcelComparatorApp:
    """Excel 比对工具主窗口类"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 差异比对工具")
        self.root.geometry("1000x800")

        # 设置窗口图标
        self.set_window_icon()

        # 变量存储
        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.handle_merged = tk.BooleanVar(value=True) # 默认开启合并单元格处理
        self.sync_h_scroll = tk.BooleanVar(value=False) # 默认不开启左右同步滚动
        self.row_height = tk.IntVar(value=30)         # 默认行高
        self.sheet_list1 = []
        self.sheet_list2 = []
        self.df1 = None
        self.df2 = None

        # 样式设置
        self.style = ttk.Style()
        self.style.configure("Treeview", font=('Arial', 10), rowheight=self.row_height.get())
        self.style.configure("Treeview.Heading", font=('Arial', 10, 'bold'))

        self._scrolling = False  # 用于防止同步滚动时的递归循环
        self.create_widgets()

    def create_widgets(self):
        # 顶部控制面板
        control_frame = ttk.LabelFrame(self.root, text="文件选择与设置", padding=10)
        control_frame.pack(fill=tk.X, padx=10, pady=5)

        # 文件 1 选择区域 (使用蓝色系背景)
        f1_frame = tk.Frame(control_frame, bg='#E6F3FF', padx=5, pady=5)
        f1_frame.grid(row=0, column=0, columnspan=5, sticky='ew', pady=(0, 5))
        
        tk.Label(f1_frame, text="Excel 文件 1 (旧):", bg='#E6F3FF', fg='#0056b3', font=('Arial', 9, 'bold')).grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(f1_frame, textvariable=self.file1_path, width=60).grid(row=0, column=1, padx=5)
        ttk.Button(f1_frame, text="浏览...", command=lambda: self.browse_file(1)).grid(row=0, column=2)
        
        tk.Label(f1_frame, text="选择 Sheet:", bg='#E6F3FF').grid(row=0, column=3, padx=10)
        self.sheet_combo1 = ttk.Combobox(f1_frame, state="readonly", width=20)
        self.sheet_combo1.grid(row=0, column=4)

        # 文件 2 选择区域 (使用绿色系背景)
        f2_frame = tk.Frame(control_frame, bg='#E6FFFA', padx=5, pady=5)
        f2_frame.grid(row=1, column=0, columnspan=5, sticky='ew')
        
        tk.Label(f2_frame, text="Excel 文件 2 (新):", bg='#E6FFFA', fg='#28a745', font=('Arial', 9, 'bold')).grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(f2_frame, textvariable=self.file2_path, width=60).grid(row=0, column=1, padx=5)
        ttk.Button(f2_frame, text="浏览...", command=lambda: self.browse_file(2)).grid(row=0, column=2)

        tk.Label(f2_frame, text="选择 Sheet:", bg='#E6FFFA').grid(row=0, column=3, padx=10)
        self.sheet_combo2 = ttk.Combobox(f2_frame, state="readonly", width=20)
        self.sheet_combo2.grid(row=0, column=4)

        # 比对按钮控制区
        action_frame = ttk.Frame(self.root, padding=5)
        action_frame.pack(fill=tk.X, padx=10)
        
        ttk.Button(action_frame, text="开始比对", command=self.compare_files).pack(side=tk.LEFT)
        
        # 合并单元格选项
        ttk.Checkbutton(action_frame, text="自动填充合并单元格", variable=self.handle_merged).pack(side=tk.LEFT, padx=10)
        
        # 左右滚动同步选项
        ttk.Checkbutton(action_frame, text="左右滚动同步", variable=self.sync_h_scroll).pack(side=tk.LEFT, padx=10)
        
        # 行高设置
        ttk.Label(action_frame, text="行高:").pack(side=tk.LEFT, padx=(10, 2))
        row_height_spin = ttk.Spinbox(action_frame, from_=20, to=100, width=5, textvariable=self.row_height, command=self.update_row_height)
        row_height_spin.pack(side=tk.LEFT)
        row_height_spin.bind("<Return>", lambda e: self.update_row_height())
        
        # 添加图例
        legend_frame = ttk.Frame(action_frame)
        legend_frame.pack(side=tk.LEFT, padx=20)
        
        ttk.Label(legend_frame, text="图例: ").pack(side=tk.LEFT)
        ttk.Label(legend_frame, text=" 修改 (左右均高亮) ", background='#FFCCCC').pack(side=tk.LEFT, padx=2)
        ttk.Label(legend_frame, text=" 新增 (仅右侧) ", background='#CCFFCC').pack(side=tk.LEFT, padx=2)
        ttk.Label(legend_frame, text=" 删除 (仅左侧) ", background='#FFFFCC').pack(side=tk.LEFT, padx=2)

        # 结果展示区域 (双栏显示)
        result_main_frame = ttk.Frame(self.root, padding=10)
        result_main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # 使用 PanedWindow 允许调整左右占比
        self.paned = ttk.PanedWindow(result_main_frame, orient=tk.HORIZONTAL)
        self.paned.pack(fill=tk.BOTH, expand=True)

        # 左侧表格 (文件 1)
        self.left_frame = ttk.LabelFrame(self.paned, text="文件 1 (旧)")
        self.paned.add(self.left_frame, weight=1)
        
        self.tree_left = ttk.Treeview(self.left_frame)
        self.hsb_l = ttk.Scrollbar(self.left_frame, orient="horizontal", command=lambda *args: self.sync_scroll_x(1, *args))
        self.hsb_l.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree_left.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.tree_left.configure(xscrollcommand=lambda *args: self._on_tree_x_scroll(1, *args))
        
        # 右侧表格 (文件 2)
        self.right_frame = ttk.LabelFrame(self.paned, text="文件 2 (新)")
        self.paned.add(self.right_frame, weight=1)
        
        self.tree_right = ttk.Treeview(self.right_frame)
        self.hsb_r = ttk.Scrollbar(self.right_frame, orient="horizontal", command=lambda *args: self.sync_scroll_x(2, *args))
        self.hsb_r.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree_right.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.tree_right.configure(xscrollcommand=lambda *args: self._on_tree_x_scroll(2, *args))

        # 共享垂直滚动条
        self.vsb = ttk.Scrollbar(result_main_frame, orient="vertical", command=self.sync_scroll_y)
        self.vsb.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.tree_left.configure(yscrollcommand=self._on_tree_scroll)
        self.tree_right.configure(yscrollcommand=self._on_tree_scroll)

        # 定义差异 tag 颜色
        for tree in [self.tree_left, self.tree_right]:
            tree.tag_configure('modified', background='#FFCCCC')
            tree.tag_configure('added', background='#CCFFCC')
            tree.tag_configure('deleted', background='#FFFFCC')
            tree.tag_configure('equal', background='#FFFFFF')

    def sync_scroll_y(self, *args):
        """当拉动滚动条时，同步两个表格的滚动位置"""
        if self._scrolling:
            return
        self._scrolling = True
        try:
            self.tree_left.yview(*args)
            self.tree_right.yview(*args)
        finally:
            self._scrolling = False

    def _on_tree_scroll(self, *args):
        """当通过鼠标滚轮或键盘滚动某个表格时，同步另一个表格和滚动条"""
        if self._scrolling:
            return
        self._scrolling = True
        try:
            self.vsb.set(*args)
            self.tree_left.yview_moveto(args[0])
            self.tree_right.yview_moveto(args[0])
        finally:
            self._scrolling = False

    def sync_scroll_x(self, source_id, *args):
        """当拉动水平滚动条时，根据配置同步两个表格的水平位置"""
        if self._scrolling:
            return
        self._scrolling = True
        try:
            if source_id == 1:
                self.tree_left.xview(*args)
                if self.sync_h_scroll.get():
                    self.tree_right.xview(*args)
            else:
                self.tree_right.xview(*args)
                if self.sync_h_scroll.get():
                    self.tree_left.xview(*args)
        finally:
            self._scrolling = False

    def _on_tree_x_scroll(self, source_id, *args):
        """当通过鼠标滚轮或键盘水平滚动某个表格时，同步另一个表格和对应的滚动条"""
        if self._scrolling:
            return
        self._scrolling = True
        try:
            if source_id == 1:
                self.hsb_l.set(*args)
                self.tree_left.xview_moveto(args[0])
                if self.sync_h_scroll.get():
                    self.hsb_r.set(*args)
                    self.tree_right.xview_moveto(args[0])
            else:
                self.hsb_r.set(*args)
                self.tree_right.xview_moveto(args[0])
                if self.sync_h_scroll.get():
                    self.hsb_l.set(*args)
                    self.tree_left.xview_moveto(args[0])
        finally:
            self._scrolling = False

    def update_row_height(self):
        """动态更新表格行高"""
        try:
            new_height = self.row_height.get()
            self.style.configure("Treeview", rowheight=new_height)
        except Exception:
            pass

    def browse_file(self, file_num):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            if file_num == 1:
                self.file1_path.set(filename)
                self.load_sheets(filename, self.sheet_combo1)
            else:
                self.file2_path.set(filename)
                self.load_sheets(filename, self.sheet_combo2)

    def load_sheets(self, filepath, combo_box):
        try:
            sheets = ExcelDiffer.load_sheets(filepath)
            combo_box['values'] = sheets
            if sheets:
                combo_box.current(0)
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def compare_files(self):
        file1 = self.file1_path.get()
        file2 = self.file2_path.get()
        sheet1 = self.sheet_combo1.get()
        sheet2 = self.sheet_combo2.get()

        if not file1 or not file2:
            messagebox.showwarning("提示", "请先选择两个 Excel 文件")
            return
        
        if not sheet1 or not sheet2:
            messagebox.showwarning("提示", "请选择要比对的 Sheet")
            return

        try:
            # 使用逻辑层提供的原始读取方法，以保持合并单元格等格式的准确性
            df1 = ExcelDiffer.read_excel_raw(file1, sheet1, handle_merged=self.handle_merged.get())
            df2 = ExcelDiffer.read_excel_raw(file2, sheet2, handle_merged=self.handle_merged.get())

            # 调用逻辑层进行比对
            columns, results = ExcelDiffer.compare_dataframes(df1, df2)
            
            # 展示结果
            self.show_diff(columns, results)

        except Exception as e:
            messagebox.showerror("错误", f"比对过程中发生错误: {e}")

    def show_diff(self, columns, results):
        """在两个 Treeview 中展示比对结果"""
        # 清空现有内容
        for tree in [self.tree_left, self.tree_right]:
            for item in tree.get_children():
                tree.delete(item)
            
            # 设置列头
            tree["columns"] = columns
            tree["show"] = "headings"
            
            for col in columns:
                tree.heading(col, text=col)
                # 初始列宽，稍后根据内容调整
                tree.column(col, width=100, anchor=tk.W)

        # 填充数据
        for left_vals, right_vals, tags in results:
            self.tree_left.insert("", tk.END, values=left_vals, tags=tags)
            self.tree_right.insert("", tk.END, values=right_vals, tags=tags)

        # 自动调整列宽
        self.auto_adjust_columns(columns, results)

    def auto_adjust_columns(self, columns, results):
        """根据内容自动调整两个表格的列宽"""
        # 限制最大列宽，防止某些单元格内容过长导致表格过宽
        MAX_COL_WIDTH = 400
        
        for idx, col in enumerate(columns):
            # 计算该列所有值的最大宽度
            # 考虑左右两个表格中该列的内容
            all_vals_in_col = []
            for left_vals, right_vals, _ in results:
                all_vals_in_col.append(str(left_vals[idx]))
                all_vals_in_col.append(str(right_vals[idx]))
            
            if not all_vals_in_col:
                max_w = 100
            else:
                max_w = max([len(val) for val in all_vals_in_col]) * 10
            
            # 加上列头宽度
            header_w = len(col) * 12
            final_w = min(max(max_w, header_w, 80), MAX_COL_WIDTH)
            
            for tree in [self.tree_left, self.tree_right]:
                tree.column(col, width=final_w, stretch=False)

    def get_resource_path(self, relative_path):
        """获取资源文件的绝对路径，兼容开发环境和 PyInstaller 打包环境"""
        if hasattr(sys, '_MEIPASS'):
            # PyInstaller 打包后的临时目录
            return os.path.join(sys._MEIPASS, relative_path)
        
        # 开发环境：图标在项目根目录的 assets 目录下
        # 当前文件在 src/ui/，向上两级到项目根目录
        base_path = os.path.dirname(os.path.abspath(__file__))
        project_root = os.path.dirname(os.path.dirname(base_path)) 
        return os.path.join(project_root, "assets", relative_path)

    def set_window_icon(self):
        """根据全平台兼容方案设置窗口图标"""
        icon_path = self.get_resource_path("excel.ico")
        
        if not os.path.exists(icon_path):
            return

        # 1. 基础方式：设置标题栏图标
        try:
            self.root.iconbitmap(icon_path)
        except Exception:
            pass

        # 2. 现代方式：设置任务栏和多分辨率图标 (需安装 Pillow)
        try:
            img = Image.open(icon_path)
            photo = ImageTk.PhotoImage(img)
            self.root.iconphoto(True, photo)
            # 注意：必须保留 photo 的引用，否则会被垃圾回收导致图标消失
            self.root._icon_photo = photo 
        except Exception:
            pass

        # 3. Windows 底层方式：强行刷新任务栏图标
        if os.name == 'nt':
            try:
                hwnd = self.root.winfo_id()
                # 加载图标资源
                hicon = ctypes.windll.user32.LoadImageW(
                    None, icon_path, 1, 0, 0, 0x0010 | 0x0040
                )
                if hicon:
                    # 发送消息设置大图标(1)和小图标(0)
                    ctypes.windll.user32.SendMessageW(hwnd, 0x0080, 0, hicon)
                    ctypes.windll.user32.SendMessageW(hwnd, 0x0080, 1, hicon)
            except Exception:
                pass
