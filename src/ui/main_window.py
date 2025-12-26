import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import sys
import ctypes
import threading
from PIL import Image, ImageTk
from src.logic.excel_processor import ExcelDiffer

class ScrollableTable(tk.Frame):
    """自定义的可滚动表格组件，支持单元格级高亮和虚拟化渲染（优化性能）"""
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        self.canvas = tk.Canvas(self, bg='white', highlightthickness=0)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.hsb = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        
        # 内部容器：实际上不放所有数据，只用于撑开滚动条
        self.scrollable_frame = tk.Frame(self.canvas, bg='white')
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        
        self.canvas.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.vsb.grid(row=0, column=1, sticky="ns")
        self.hsb.grid(row=1, column=0, sticky="ew")
        
        # 数据存储
        self.columns = []
        self.data = []
        self.tags_list = []
        self.row_height = 30
        self.col_widths = []
        
        # 缓存的控件
        self.header_widgets = []
        self.cell_widgets = {} # (row_idx, col_idx) -> label
        self.visible_rows = 0
        
        # 绑定事件
        self.canvas.bind("<Configure>", self._on_configure)
        self._bind_mouse_wheel(self.canvas)

    def _on_configure(self, event):
        """窗口大小改变时触发"""
        self._update_view()

    def _bind_mouse_wheel(self, widget):
        widget.bind("<MouseWheel>", self._on_mouse_wheel)
        widget.bind("<Button-4>", self._on_mouse_wheel)
        widget.bind("<Button-5>", self._on_mouse_wheel)

    def _on_mouse_wheel(self, event):
        if event.num == 4 or event.delta > 0:
            delta = -1
        elif event.num == 5 or event.delta < 0:
            delta = 1
        else:
            delta = 0

        if event.state & 0x0001: # Shift key
            self.canvas.xview_scroll(delta, "units")
        else:
            self.canvas.yview_scroll(delta, "units")
        
        self._update_view() # 滚动时即时更新视图
        
        if hasattr(self, 'on_scroll_callback'):
            self.on_scroll_callback()
        return "break"

    def clear(self):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.header_widgets = []
        self.cell_widgets = {}
        self.data = []
        self.tags_list = []

    def set_data(self, columns, data, tags_list, row_height=30):
        self.clear()
        self.columns = columns
        self.data = data
        self.tags_list = tags_list
        self.row_height = row_height
        
        if not data:
            return

        # 预计算每列的大致宽度
        def get_display_width(s):
            """计算字符串在显示时的近似宽度（中文字符占 1.8 倍宽度）"""
            width = 0
            for char in str(s):
                if ord(char) > 127:
                    width += 1.8
                else:
                    width += 1
            return width

        self.col_widths = [80] * len(columns)
        for i, col in enumerate(columns):
            max_w = get_display_width(col)
            # 采样前 300 行计算宽度，平衡性能与准确性
            sample_size = min(len(data), 300)
            for row in data[:sample_size]:
                val = str(row[i])
                if val.startswith(">>> "): val = val[4:-4]
                max_w = max(max_w, get_display_width(val))
            # 增加 padding，并提高最大宽度至 600
            self.col_widths[i] = min(max_w * 8 + 30, 600)

        # 设置内部框架大小以撑开滚动条
        total_width = sum(self.col_widths)
        total_height = (len(data) + 1) * self.row_height
        self.scrollable_frame.config(width=total_width, height=total_height)
        self.canvas.config(scrollregion=(0, 0, total_width, total_height))
        
        self._update_view()

    def _update_view(self):
        """核心：虚拟化渲染，只创建/显示可见区域的控件"""
        if not self.data:
            return

        # 获取当前视图范围
        v_start = self.canvas.yview()[0]
        v_end = self.canvas.yview()[1]
        h_start = self.canvas.xview()[0]
        h_end = self.canvas.xview()[1]

        total_rows = len(self.data)
        total_height = (total_rows + 1) * self.row_height
        
        # 计算可见行范围
        start_row = int(v_start * total_height / self.row_height) - 1
        end_row = int(v_end * total_height / self.row_height) + 1
        
        start_row = max(0, start_row)
        end_row = min(total_rows, end_row)

        # 计算可见列范围 (水平虚拟化)
        total_width = sum(self.col_widths)
        start_col = 0
        current_w = 0
        for i, w in enumerate(self.col_widths):
            if (current_w + w) / total_width < h_start:
                start_col = i
            if current_w / total_width > h_end:
                end_col = i
                break
            current_w += w
        else:
            end_col = len(self.columns)
        
        start_col = max(0, start_col - 1)
        end_col = min(len(self.columns), end_col + 1)

        # 1. 处理表头 (始终可见，但水平位置随 hsb 移动)
        # 清理不可见的表头
        current_visible_headers = set(range(start_col, end_col))
        for j in list(range(len(self.header_widgets))):
            if j not in current_visible_headers and self.header_widgets[j]:
                self.header_widgets[j].destroy()
                self.header_widgets[j] = None
        
        if not hasattr(self, 'header_widgets') or not self.header_widgets:
            self.header_widgets = [None] * len(self.columns)

        for j in current_visible_headers:
            if self.header_widgets[j] is None:
                lbl = tk.Label(self.scrollable_frame, text=self.columns[j], font=('Arial', 10, 'bold'),
                             relief="raised", bg='#F0F0F0', padx=10, borderwidth=1)
                lbl.place(x=sum(self.col_widths[:j]), y=0, width=self.col_widths[j], height=self.row_height)
                self.header_widgets[j] = lbl

        # 2. 处理数据行
        current_visible_keys = set()
        for i in range(start_row, end_row):
            row_vals = self.data[i]
            tags = self.tags_list[i]
            status = row_vals[0] if row_vals else ""
            
            for j in range(start_col, end_col):
                val = row_vals[j]
                key = (i, j)
                current_visible_keys.add(key)
                
                if key not in self.cell_widgets:
                    val_str = str(val)
                    cell_bg = "white"
                    fg_color = "black"
                    
                    # 状态颜色填充
                    if j == 0: # 状态列
                        if status == "新增": cell_bg = '#CCFFCC' # 浅绿
                        elif status == "删除": cell_bg = '#FFFFCC' # 浅黄
                        elif status == "修改": cell_bg = '#E6F3FF' # 浅蓝
                        elif status == "一致": cell_bg = '#F8F8F8'
                    else:
                        # 差异行背景
                        if 'added' in tags: cell_bg = '#F0FFF0' # 极浅绿
                        elif 'deleted' in tags: cell_bg = '#FFFFF0' # 极浅黄
                    
                    # 差异单元格高亮 (>>> <<<)
                    if val_str.startswith(">>> ") and val_str.endswith(" <<<"):
                        val_str = val_str[4:-4]
                        cell_bg = '#FFCCCC' # 浅红色填充
                        fg_color = '#CC0000' # 深红色文字
                    
                    lbl = tk.Label(self.scrollable_frame, text=val_str, font=('Arial', 10),
                                 bg=cell_bg, fg=fg_color, padx=10, 
                                 borderwidth=1, relief="groove", anchor="w")
                    lbl.place(x=sum(self.col_widths[:j]), y=(i+1)*self.row_height, 
                             width=self.col_widths[j], height=self.row_height)
                    self.cell_widgets[key] = lbl
                    self._bind_mouse_wheel(lbl)

        # 3. 清理不可见区域的控件以节省内存和提升性能
        to_remove = []
        for key in self.cell_widgets:
            if key not in current_visible_keys:
                # 范围判定：行列都要考虑
                r_idx, c_idx = key
                if (r_idx < start_row - 5 or r_idx > end_row + 5 or 
                    c_idx < start_col - 2 or c_idx > end_col + 2):
                    to_remove.append(key)
        
        for key in to_remove:
            self.cell_widgets[key].destroy()
            del self.cell_widgets[key]

class LoadingDialog:
    """比对过程中的加载弹窗"""
    def __init__(self, parent, title="请稍候", message="正在处理中..."):
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.geometry("350x120")
        self.top.resizable(False, False)
        self.top.transient(parent)  # 设置为父窗口的临时窗口
        self.top.grab_set()         # 模态弹窗
        
        # 居中显示
        parent_x = parent.winfo_x()
        parent_y = parent.winfo_y()
        parent_w = parent.winfo_width()
        parent_h = parent.winfo_height()
        x = parent_x + (parent_w - 350) // 2
        y = parent_y + (parent_h - 120) // 2
        self.top.geometry(f"+{x}+{y}")

        # 界面元素
        frame = ttk.Frame(self.top, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)

        self.label = ttk.Label(frame, text=message, font=('Arial', 10))
        self.label.pack(pady=(0, 10))

        self.progress = ttk.Progressbar(frame, mode='indeterminate', length=280)
        self.progress.pack(pady=5)
        self.progress.start(10)

        # 禁用关闭按钮
        self.top.protocol("WM_DELETE_WINDOW", lambda: None)

    def close(self):
        self.top.grab_release()
        self.top.destroy()

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
        self.sync_v_scroll = tk.BooleanVar(value=True) # 默认开启垂直滚动同步
        self.sync_h_scroll = tk.BooleanVar(value=True) # 默认开启左右同步滚动
        self.sheet_list1 = []
        self.sheet_list2 = []
        self.df1 = None
        self.df2 = None

        # 样式设置
        self.style = ttk.Style()
        self.style.configure("Treeview.Heading", font=('Arial', 10, 'bold'))

        self._scrolling = False  # 用于防止同步滚动时的递归循环
        self.create_widgets()
        self.create_menu()

    def create_menu(self):
        """创建菜单栏"""
        menubar = tk.Menu(self.root)
        
        # 帮助菜单
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="使用说明", command=self.show_usage)
        help_menu.add_command(label="版本说明", command=self.show_version)
        
        menubar.add_cascade(label="帮助", menu=help_menu)
        
        self.root.config(menu=menubar)

    def show_usage(self):
        """显示使用说明"""
        usage_text = """
Excel 差异比对工具 使用说明：

1. 选择文件：点击“浏览...”按钮选择需要比对的两个 Excel 文件（旧版本和新版本）。
2. 选择 Sheet：在下拉框中选择要比对的工作表。
3. 设置选项：
   - 垂直/左右滚动同步：开启后两侧表格将同步滚动。
4. 开始比对：点击“开始比对”按钮，程序将分析差异并在下方展示。
5. 结果说明：
   - 单元格浅红填充：表示内容发生了变化。
   - 行浅绿填充：表示该行为新增行。
   - 行浅黄填充（仅左侧）：表示该行为删除行。
6. 导出结果：比对完成后，点击“导出比对结果”可将差异保存为 Excel 文件。

注意：程序已默认开启合并单元格自动填充功能，确保比对结果的准确性。

快捷键：
- 鼠标滚轮：垂直滚动。
- Shift + 鼠标滚轮：水平滚动。
"""
        messagebox.showinfo("使用说明", usage_text)

    def show_version(self):
        """显示版本说明"""
        version = "V1.1.1"
        try:
            if os.path.exists("updates_notes.txt"):
                with open("updates_notes.txt", "r", encoding="utf-8") as f:
                    content = f.read()
                    first_line = content.split('\n')[0]
                    if "版本:" in first_line:
                        version = first_line.split(":")[1].strip()
        except Exception:
            pass
            
        version_text = f"""
Excel 差异比对工具

当前版本：{version}
更新日期：2025-12-26

主要功能：
- 支持多 Sheet 选择比对
- 单元格级差异精准识别与高亮
- 支持合并单元格自动填充处理
- 左右同步滚动与虚拟化渲染
- 支持导出高亮标记的 Excel 结果
"""
        messagebox.showinfo("版本说明", version_text)

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
        
        self.btn_compare_ref = ttk.Button(action_frame, text="开始比对", command=self.compare_files)
        self.btn_compare_ref.pack(side=tk.LEFT)
        
        # 导出按钮 (默认禁用)
        self.btn_export = ttk.Button(action_frame, text="导出比对结果", command=self.export_results, state=tk.DISABLED)
        self.btn_export.pack(side=tk.LEFT, padx=10)
        
        # 垂直滚动同步选项
        ttk.Checkbutton(action_frame, text="垂直滚动同步", variable=self.sync_v_scroll).pack(side=tk.LEFT, padx=10)
        
        # 左右滚动同步选项
        ttk.Checkbutton(action_frame, text="左右滚动同步", variable=self.sync_h_scroll).pack(side=tk.LEFT, padx=10)
        
        # 添加图例
        legend_frame = ttk.Frame(action_frame)
        legend_frame.pack(side=tk.LEFT, padx=20)
        
        ttk.Label(legend_frame, text="状态图例: ").pack(side=tk.LEFT)
        ttk.Label(legend_frame, text=" 修改 ", background='#E6F3FF').pack(side=tk.LEFT, padx=2)
        ttk.Label(legend_frame, text=" 新增 ", background='#CCFFCC').pack(side=tk.LEFT, padx=2)
        ttk.Label(legend_frame, text=" 删除 ", background='#FFFFCC').pack(side=tk.LEFT, padx=2)
        ttk.Label(legend_frame, text=" | 单元格差异: >>> <<< ", background='#FFCCCC', foreground='#CC0000').pack(side=tk.LEFT, padx=(10, 2))

        # 结果展示区域 (双栏显示)
        result_main_frame = ttk.Frame(self.root, padding=10)
        result_main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # 使用 PanedWindow 允许调整左右占比
        self.paned = ttk.PanedWindow(result_main_frame, orient=tk.HORIZONTAL)
        self.paned.pack(fill=tk.BOTH, expand=True)

        # 左侧表格 (文件 1)
        self.left_frame = ttk.LabelFrame(self.paned, text="文件 1 (旧)")
        self.paned.add(self.left_frame, weight=1)
        
        self.table_left = ScrollableTable(self.left_frame)
        self.table_left.pack(fill=tk.BOTH, expand=True)
        
        # 右侧表格 (文件 2)
        self.right_frame = ttk.LabelFrame(self.paned, text="文件 2 (新)")
        self.paned.add(self.right_frame, weight=1)
        
        self.table_right = ScrollableTable(self.right_frame)
        self.table_right.pack(fill=tk.BOTH, expand=True)

        # 绑定同步滚动逻辑
        self.table_left.canvas.bind("<Configure>", lambda e: self.sync_scroll_setup())
        self.table_right.canvas.bind("<Configure>", lambda e: self.sync_scroll_setup())
        
        # 缓存比对结果用于导出
        self.last_columns = None
        self.last_results = None

    def sync_scroll_setup(self):
        """设置同步滚动绑定"""
        self.table_left.canvas.configure(yscrollcommand=lambda *args: self.sync_scroll_y(self.table_left, *args))
        self.table_right.canvas.configure(yscrollcommand=lambda *args: self.sync_scroll_y(self.table_right, *args))
        self.table_left.canvas.configure(xscrollcommand=lambda *args: self.sync_scroll_x(self.table_left, *args))
        self.table_right.canvas.configure(xscrollcommand=lambda *args: self.sync_scroll_x(self.table_right, *args))

    def sync_scroll_y(self, source_table, *args):
        """同步垂直滚动"""
        if self._scrolling:
            return
        
        self._scrolling = True
        try:
            # 更新源表格的滚动条
            source_table.vsb.set(*args)
            source_table._update_view()
            
            # 如果开启了同步，更新另一个表格
            if self.sync_v_scroll.get():
                target_table = self.table_right if source_table == self.table_left else self.table_left
                target_table.vsb.set(*args)
                target_table.canvas.yview_moveto(args[0])
                target_table._update_view()
        finally:
            self._scrolling = False

    def sync_scroll_x(self, source_table, *args):
        """同步水平滚动（如果开启）"""
        if self._scrolling:
            return
            
        self._scrolling = True
        try:
            # 更新源表格的滚动条
            source_table.hsb.set(*args)
            source_table._update_view()
            
            # 如果开启了同步，更新另一个表格
            if self.sync_h_scroll.get():
                target_table = self.table_right if source_table == self.table_left else self.table_left
                target_table.hsb.set(*args)
                target_table.canvas.xview_moveto(args[0])
                target_table._update_view()
        finally:
            self._scrolling = False

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

        # 显示加载弹窗
        self.loading = LoadingDialog(self.root, message="正在读取文件并比对差异，请稍候...")
        
        # 禁用按钮
        self.btn_compare_ref.config(state=tk.DISABLED)
        self.btn_export.config(state=tk.DISABLED)

        # 开启线程执行比对
        thread = threading.Thread(target=self._compare_thread_task, args=(file1, file2, sheet1, sheet2))
        thread.daemon = True
        thread.start()

    def _compare_thread_task(self, file1, file2, sheet1, sheet2):
        """后台线程执行比对任务"""
        try:
            # 1. 读取文件（默认开启合并单元格处理）
            df1 = ExcelDiffer.read_excel_raw(file1, sheet1, handle_merged=True)
            df2 = ExcelDiffer.read_excel_raw(file2, sheet2, handle_merged=True)

            # 2. 调用逻辑层进行比对
            columns, results = ExcelDiffer.compare_dataframes(df1, df2)
            
            # 3. 回到主线程更新 UI
            self.root.after(0, lambda: self._update_ui_after_compare(columns, results))

        except Exception as e:
            self.root.after(0, lambda: self._handle_compare_error(str(e)))

    def _update_ui_after_compare(self, columns, results):
        """比对完成后更新 UI（主线程）"""
        self.last_columns = columns
        self.last_results = results
        
        # 关闭加载弹窗
        if hasattr(self, 'loading'):
            self.loading.close()
        
        # 恢复按钮状态
        self.btn_compare_ref.config(state=tk.NORMAL)
        self.btn_export.config(state=tk.NORMAL)
        
        # 展示结果
        self.show_diff(columns, results)

    def _handle_compare_error(self, error_msg):
        """处理比对过程中的错误（主线程）"""
        # 关闭加载弹窗
        if hasattr(self, 'loading'):
            self.loading.close()
            
        self.btn_compare_ref.config(state=tk.NORMAL)
        messagebox.showerror("错误", f"比对过程中发生错误: {error_msg}")

    def show_diff(self, columns, results):
        """在两个自定义表格中展示比对结果"""
        left_data = [r[0] for r in results]
        right_data = [r[1] for r in results]
        tags_list = [r[2] for r in results]
        
        self.table_left.set_data(columns, left_data, tags_list)
        self.table_right.set_data(columns, right_data, tags_list)

    def export_results(self):
        """将缓存的比对结果导出到 Excel"""
        if self.last_columns is None or self.last_results is None:
            messagebox.showwarning("提示", "没有可导出的比对结果")
            return
            
        output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="比对结果.xlsx"
        )
        
        if not output_path:
            return
            
        try:
            ExcelDiffer.export_diff(output_path, self.last_columns, self.last_results)
            messagebox.showinfo("成功", f"结果已成功导出至:\n{output_path}")
        except Exception as e:
            messagebox.showerror("错误", f"导出失败: {e}")

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
