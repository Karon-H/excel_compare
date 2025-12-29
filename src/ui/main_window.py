import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import MessageDialog
import pandas as pd
import os
import sys
import ctypes
import threading
import windnd
from PIL import Image, ImageTk
from src.logic.excel_processor import ExcelDiffer

class ScrollableTable(tk.Frame):
    """自定义的可滚动表格组件，支持单元格级高亮和虚拟化渲染（优化性能）"""
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        self.canvas = tk.Canvas(self, bg='white', highlightthickness=0)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self._on_vscroll)
        self.hsb = ttk.Scrollbar(self, orient="horizontal", command=self._on_hscroll)
        
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
        self.base_row_height = 30
        self.row_height = 30
        self.row_heights = [] # 新增：记录每行的高度
        self.base_col_widths = []
        self.col_widths = []
        self.freeze_panes = tk.BooleanVar(value=True) # 默认开启冻结
        self.freeze_rows = tk.IntVar(value=1)  # 冻结前 N 行 (含表头)
        self.freeze_cols = tk.IntVar(value=1)  # 冻结前 M 列
        self.selected_cell = None # 新增：选中的单元格 (row_idx, col_idx)
        
        # 缓存的控件
        self.header_widgets = []
        self.cell_widgets = {} # (row_idx, col_idx) -> label
        self.visible_rows = 0
        
        # 性能优化：对象池
        self._widget_pool = []
        
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
        
        self._update_view() 
        
        if hasattr(self, 'on_scroll_callback'):
            self.on_scroll_callback()
        return "break"

    def _on_vscroll(self, *args):
        """处理垂直滚动条拖动"""
        self.canvas.yview(*args)
        self._update_view()
        if hasattr(self, 'on_scroll_callback'):
            self.on_scroll_callback()

    def _on_hscroll(self, *args):
        """处理水平滚动条拖动"""
        self.canvas.xview(*args)
        self._update_view()
        if hasattr(self, 'on_scroll_callback'):
            self.on_scroll_callback()

    def clear(self):
        # 将所有控件归还对象池
        for widget in self.cell_widgets.values():
            widget.place_forget()
            self._widget_pool.append(widget)
        self.cell_widgets = {}
        
        for widget in self.header_widgets:
            if widget:
                widget.place_forget()
                self._widget_pool.append(widget)
        self.header_widgets = []
        
        # 清理剩余可能的子控件
        for widget in self.scrollable_frame.winfo_children():
            if widget not in self._widget_pool:
                widget.destroy()

        self.data = []
        self.tags_list = []
        self.row_heights = []
        self.col_widths = []

    def set_data(self, columns, data, tags_list, row_height=30):
        self.clear()
        self.columns = columns
        self.data = data
        self.tags_list = tags_list
        self.base_row_height = row_height
        self.row_height = row_height
        self.selected_cell = None
        
        if not data:
            return
        
        # 1. 初始化列表长度
        self.header_widgets = [None] * len(columns)
        
        # 2. 初始化每行高度 (表头 + 数据行)
        self.row_heights = [row_height] * (len(data) + 1)

        # 3. 预计算每列的大致基础宽度
        def get_display_width(s):
            """计算字符串在显示时的近似宽度（中文字符占 1.8 倍宽度）"""
            width = 0
            for char in str(s):
                if ord(char) > 127:
                    width += 1.8
                else:
                    width += 1
            return width

        self.base_col_widths = [80] * len(columns)
        for i, col in enumerate(columns):
            max_w = get_display_width(col)
            sample_size = min(len(data), 300)
            for row in data[:sample_size]:
                val = str(row[i])
                if val.startswith(">>> "): val = val[4:-4]
                max_w = max(max_w, get_display_width(val))
            self.base_col_widths[i] = min(max_w * 8 + 30, 600)
        
        # 初始列宽 = 基础宽度
        self.col_widths = [int(w) for w in self.base_col_widths]

        self._update_scroll_region()
        self._update_view()

    def _get_row_y(self, row_idx):
        """计算第 row_idx 行的起始 y 坐标"""
        return sum(self.row_heights[:row_idx])

    def _get_row_at_y(self, y_coord):
        """根据 y 坐标查找所在的行索引"""
        current_y = 0
        for i, h in enumerate(self.row_heights):
            if current_y <= y_coord < current_y + h:
                return i
            current_y += h
        return len(self.row_heights) - 1

    def _get_col_x(self, col_idx):
        """计算第 col_idx 列的起始 x 坐标"""
        return sum(self.col_widths[:col_idx])

    def _get_col_at_x(self, x_coord):
        """根据 x 坐标查找所在的列索引"""
        current_x = 0
        for i, w in enumerate(self.col_widths):
            if current_x <= x_coord < current_x + w:
                return i
            current_x += w
        return len(self.col_widths) - 1

    def _update_scroll_region(self):
        """更新滚动区域大小"""
        if not self.data:
            return
        
        total_width = sum(self.col_widths)
        total_height = sum(self.row_heights)
        self.scrollable_frame.config(width=total_width, height=total_height)
        self.canvas.config(scrollregion=(0, 0, total_width, total_height))

    def _on_cell_click(self, event, row_idx, col_idx):
        """单元格点击事件"""
        self.selected_cell = (row_idx, col_idx)
        self._update_view()
        if hasattr(self, 'on_select_callback'):
            self.on_select_callback(row_idx, col_idx)

    def _update_view(self):
        """核心：虚拟化渲染 + 对象池优化"""
        if not self.data:
            return

        # 1. 计算布局参数
        total_width = sum(self.col_widths)
        total_height = sum(self.row_heights)
        
        h_start, h_end = self.canvas.xview()
        v_start, v_end = self.canvas.yview()
        
        x_offset = h_start * total_width
        y_offset = v_start * total_height
        
        view_w = self.canvas.winfo_width()
        view_h = self.canvas.winfo_height()

        # 计算可见行/列范围
        start_row = self._get_row_at_y(y_offset)
        end_row = self._get_row_at_y(y_offset + view_h) + 1
        start_col = self._get_col_at_x(x_offset)
        end_col = self._get_col_at_x(x_offset + view_w) + 1
        
        start_row = max(0, start_row)
        end_row = min(len(self.data) + 1, end_row)
        start_col = max(0, start_col)
        end_col = min(len(self.columns), end_col)

        f_rows = self.freeze_rows.get()
        f_cols = self.freeze_cols.get()

        current_visible_keys = set() # (row, col)

        # 2. 渲染表头 (row 0)
        header_cols = set(range(start_col, end_col))
        for c in range(f_cols): header_cols.add(c)
        
        for j in header_cols:
            key = (0, j)
            current_visible_keys.add(key)
            
            if self.header_widgets[j] is None:
                if self._widget_pool:
                    lbl = self._widget_pool.pop()
                else:
                    lbl = tk.Label(self.scrollable_frame, padx=5, highlightthickness=0)
                    self._bind_mouse_wheel(lbl)
                
                self.header_widgets[j] = lbl

            # 统一配置属性（防止回收时样式残留）
            lbl = self.header_widgets[j]
            bg = '#CCE8FF' if self.selected_cell == (0, j) else '#F0F0F0'
            lbl.config(
                text=self.columns[j], 
                font=('Microsoft YaHei', 10, 'bold'), 
                relief="raised",
                borderwidth=1,
                bg=bg
            )
            lbl.bind("<Button-1>", lambda e, c=j: self._on_cell_click(e, 0, c))

            # 布局
            tx = self._get_col_x(j)
            ty = 0
            if f_rows > 0: ty = y_offset
            if j < f_cols: tx = x_offset + self._get_col_x(j)
            
            self.header_widgets[j].place(x=tx, y=ty, width=self.col_widths[j], height=self.row_heights[0])
            if f_rows > 0 or j < f_cols: self.header_widgets[j].lift()

        # 3. 渲染数据行
        visible_data_rows = set(range(max(1, start_row), end_row))
        for r in range(1, f_rows): 
            if r < len(self.row_heights): visible_data_rows.add(r)
            
        for r_idx in visible_data_rows:
            i = r_idx - 1
            row_vals = self.data[i]
            tags = self.tags_list[i]
            status = row_vals[0] if row_vals else ""
            
            data_cols = set(range(start_col, end_col))
            for c in range(f_cols): data_cols.add(c)
            
            for j in data_cols:
                key = (r_idx, j)
                current_visible_keys.add(key)
                
                if key not in self.cell_widgets:
                    if self._widget_pool:
                        lbl = self._widget_pool.pop()
                    else:
                        lbl = tk.Label(self.scrollable_frame, padx=5, highlightthickness=0, anchor="w")
                        self._bind_mouse_wheel(lbl)
                    self.cell_widgets[key] = lbl
                
                lbl = self.cell_widgets[key]
                val = str(row_vals[j])
                cell_bg, fg = "white", "black"
                
                if j == 0:
                    if status == "新增": cell_bg = '#CCFFCC'
                    elif status == "删除": cell_bg = '#FFFFCC'
                    elif status == "修改": cell_bg = '#E6F3FF'
                    else: cell_bg = '#F8F8F8'
                else:
                    if 'added' in tags: cell_bg = '#F0FFF0'
                    elif 'deleted' in tags: cell_bg = '#FFFFF0'
                
                if val.startswith(">>> ") and val.endswith(" <<<"):
                    val = val[4:-4]; cell_bg = '#FFCCCC'; fg = '#CC0000'
                
                # 统一配置属性
                final_bg = '#CCE8FF' if self.selected_cell == (r_idx, j) else cell_bg
                lbl.config(
                    text=val, 
                    font=('Microsoft YaHei', 10), 
                    bg=final_bg, 
                    fg=fg,
                    relief="groove",
                    borderwidth=1
                )
                lbl.bind("<Button-1>", lambda e, r=r_idx, c=j: self._on_cell_click(e, r, c))
                lbl.original_bg = cell_bg

                # 布局
                tx = self._get_col_x(j)
                ty = self._get_row_y(r_idx)
                if r_idx < f_rows: ty = y_offset + self._get_row_y(r_idx)
                if j < f_cols: tx = x_offset + self._get_col_x(j)
                
                self.cell_widgets[key].place(x=tx, y=ty, width=self.col_widths[j], height=self.row_heights[r_idx])
                if r_idx < f_rows or j < f_cols: self.cell_widgets[key].lift()

        # 4. 回收不再可见的控件
        # 处理表头回收
        for j in range(len(self.columns)):
            if (0, j) not in current_visible_keys and self.header_widgets[j] is not None:
                self.header_widgets[j].place_forget()
                self._widget_pool.append(self.header_widgets[j])
                self.header_widgets[j] = None
        
        # 处理数据行回收
        to_remove = []
        for key, widget in self.cell_widgets.items():
            if key not in current_visible_keys:
                widget.place_forget()
                self._widget_pool.append(widget)
                to_remove.append(key)
        for key in to_remove: del self.cell_widgets[key]

class LoadingDialog:
    """比对过程中的加载弹窗 - 升级带百分比和详细进度"""
    def __init__(self, parent, title="请稍候", message="正在处理中..."):
        self.top = ttk.Toplevel(parent)
        self.top.title(title)
        self.top.geometry("400x180")
        self.top.resizable(False, False)
        self.top.transient(parent)
        self.top.grab_set()
        
        # 居中显示
        parent_x = parent.winfo_x()
        parent_y = parent.winfo_y()
        parent_w = parent.winfo_width()
        parent_h = parent.winfo_height()
        x = parent_x + (parent_w - 400) // 2
        y = parent_y + (parent_h - 180) // 2
        self.top.geometry(f"+{x}+{y}")

        frame = ttk.Frame(self.top, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)

        self.label = ttk.Label(frame, text=message, font=('Microsoft YaHei', 10))
        self.label.pack(pady=(0, 5), anchor=W)

        self.progress_var = tk.DoubleVar(value=0)
        self.progress = ttk.Progressbar(frame, variable=self.progress_var, maximum=100, bootstyle=INFO)
        self.progress.pack(fill=X, pady=10)

        self.detail_label = ttk.Label(frame, text="准备开始...", font=('Microsoft YaHei', 9), bootstyle=SECONDARY)
        self.detail_label.pack(anchor=W)

        # 禁用关闭按钮
        self.top.protocol("WM_DELETE_WINDOW", lambda: None)

    def update_progress(self, percent, detail_text):
        """更新进度条和详情文本"""
        self.progress_var.set(percent)
        self.detail_label.config(text=detail_text)
        self.top.update_idletasks()

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

        # 拖拽支持
        windnd.hook_dropfiles(self.root, func=self.on_file_drop)

        # 变量存储
        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.sync_v_scroll = tk.BooleanVar(value=True) # 默认开启垂直滚动同步
        self.sync_h_scroll = tk.BooleanVar(value=True) # 默认开启左右同步滚动
        self.sheet_list1 = []
        self.sheet_list2 = []
        self.df1 = None
        self.df2 = None

        # 界面控制变量
        self.freeze_rows = tk.IntVar(value=1)
        self.freeze_cols = tk.IntVar(value=1)
        self.row_h_var = tk.DoubleVar(value=30)
        self.col_w_var = tk.DoubleVar(value=1.0)
        self.sel_row_h = tk.IntVar(value=30)
        self.sel_col_w = tk.IntVar(value=100)
        self.key_columns = [] # 选中的关键列

        # 样式设置
        self.setup_styles()

        self._scrolling = False  # 用于防止同步滚动时的递归循环
        self.setup_ui()
        self.create_menu()

    def on_file_drop(self, files):
        """处理文件拖拽"""
        for file_path in files:
            path_str = file_path.decode('gbk') if isinstance(file_path, bytes) else file_path
            if path_str.lower().endswith(('.xlsx', '.xls')):
                # 根据拖拽位置或当前状态决定填入哪个框
                # 这里简单策略：如果第一个为空填第一个，否则填第二个
                if not self.file1_path.get():
                    self.file1_path.set(path_str)
                    self.load_sheets(path_str, self.sheet_combo1)
                else:
                    self.file2_path.set(path_str)
                    self.load_sheets(path_str, self.sheet_combo2)

    def setup_styles(self):
        """配置自定义样式"""
        self.style = ttk.Style()
        self.style.configure("Treeview.Heading", font=('Microsoft YaHei', 10, 'bold'))
        # 自定义一些特定样式
        self.style.configure('Custom.TFrame', background='#f8f9fa')
        self.style.configure('Header.TLabel', font=('Microsoft YaHei', 11, 'bold'))
        self.style.configure('Action.TButton', font=('Microsoft YaHei', 10, 'bold'))

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
        version = "V1.9.2"
        try:
            if os.path.exists("updates_notes.txt"):
                with open("updates_notes.txt", "r", encoding="utf-8") as f:
                    first_line = f.readline().strip()
                    if first_line.startswith('V'):
                        # 格式：V1.2.0 2025-12-26 -> 提取 V1.2.0
                        version = first_line.split(' ')[0]
        except Exception:
            pass
            
        version_text = f"""
Excel 差异比对工具

当前版本：{version}
更新日期：2025-12-29

更新亮点：
- 升级：全面现代化 UI 改造，支持主题切换与深色模式。
- 优化：操作 UI 深度整合，将所有设置集中于顶部，提升操作效率。
- 新增：恢复并优化“选中精调”功能，支持对选中行列尺寸进行微调。
- 新增：支持文件拖拽 (Drag & Drop) 快速录入数据源。
- 核心：支持关键列（主键）比对，精准定位乱序数据差异。
"""
        messagebox.showinfo("版本说明", version_text)

    def setup_ui(self):
        # 使用 ttkbootstrap 的布局容器
        main_container = ttk.Frame(self.root, padding=10)
        main_container.pack(fill=tk.BOTH, expand=True)

        # 1. 顶部操作面板：整合所有配置与设置
        top_panel = ttk.Frame(main_container)
        top_panel.pack(fill=tk.X, pady=(0, 10))

        # --- 第一行：数据源配置 ---
        ds_frame = ttk.Labelframe(top_panel, text="数据源配置", padding=10, bootstyle=PRIMARY)
        ds_frame.pack(fill=tk.X, pady=(0, 5))
        
        # 配置列权重
        for i in range(5): ds_frame.columnconfigure(i, weight=1)

        # 文件 1
        ttk.Label(ds_frame, text="文件 1 (旧):", font=('Microsoft YaHei', 9, 'bold')).grid(row=0, column=0, sticky=tk.W, pady=2)
        ttk.Entry(ds_frame, textvariable=self.file1_path, width=40).grid(row=0, column=1, columnspan=2, sticky=tk.EW, padx=5)
        ttk.Button(ds_frame, text="浏览...", command=lambda: self.browse_file(1), bootstyle=SECONDARY).grid(row=0, column=3, padx=5)
        self.sheet_combo1 = ttk.Combobox(ds_frame, state="readonly", width=15)
        self.sheet_combo1.grid(row=0, column=4, sticky=tk.EW, padx=5)

        # 文件 2
        ttk.Label(ds_frame, text="文件 2 (新):", font=('Microsoft YaHei', 9, 'bold')).grid(row=1, column=0, sticky=tk.W, pady=2)
        ttk.Entry(ds_frame, textvariable=self.file2_path, width=40).grid(row=1, column=1, columnspan=2, sticky=tk.EW, padx=5)
        ttk.Button(ds_frame, text="浏览...", command=lambda: self.browse_file(2), bootstyle=SECONDARY).grid(row=1, column=3, padx=5)
        self.sheet_combo2 = ttk.Combobox(ds_frame, state="readonly", width=15)
        self.sheet_combo2.grid(row=1, column=4, sticky=tk.EW, padx=5)

        # --- 第二行：功能设置与比对选项 (并排) ---
        middle_settings = ttk.Frame(top_panel)
        middle_settings.pack(fill=tk.X, pady=5)

        # A. 比对模式与关键列
        opt_group = ttk.Labelframe(middle_settings, text="比对模式", padding=10, bootstyle=INFO)
        opt_group.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        self.key_mode_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt_group, text="启用主键模式", variable=self.key_mode_var, command=self.toggle_key_mode, bootstyle="round-toggle").pack(side=tk.LEFT, padx=5)
        self.btn_select_keys = ttk.Button(opt_group, text="选择关键列...", command=self.select_key_columns, state=tk.DISABLED, bootstyle="outline-info")
        self.btn_select_keys.pack(side=tk.LEFT, padx=5)
        self.key_cols_label = ttk.Label(opt_group, text="（未选择）", font=('Microsoft YaHei', 8), bootstyle=SECONDARY)
        self.key_cols_label.pack(side=tk.LEFT, padx=5)

        # B. 滚动同步
        sync_group = ttk.Labelframe(middle_settings, text="视图同步", padding=10, bootstyle=INFO)
        sync_group.pack(side=tk.LEFT, fill=tk.Y, padx=5)
        ttk.Checkbutton(sync_group, text="纵滚", variable=self.sync_v_scroll, bootstyle="round-toggle").pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(sync_group, text="横滚", variable=self.sync_h_scroll, bootstyle="round-toggle").pack(side=tk.LEFT, padx=5)

        # C. 视图精调 (行高列宽冻结)
        view_group = ttk.Labelframe(middle_settings, text="视图缩放/冻结", padding=10, bootstyle=INFO)
        view_group.pack(side=tk.LEFT, fill=tk.Y, padx=5)
        
        # 冻结
        f_box = ttk.Frame(view_group)
        f_box.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(f_box, text="冻结R/C:").pack(side=tk.LEFT)
        tk.Spinbox(f_box, from_=0, to=50, textvariable=self.freeze_rows, width=2).pack(side=tk.LEFT, padx=2)
        tk.Spinbox(f_box, from_=0, to=20, textvariable=self.freeze_cols, width=2).pack(side=tk.LEFT, padx=2)
        self.freeze_rows.trace_add("write", self.sync_freeze_rows)
        self.freeze_cols.trace_add("write", self.sync_freeze_cols)

        # 缩放滑块
        ttk.Label(view_group, text="行高:").pack(side=tk.LEFT)
        ttk.Scale(view_group, from_=20, to=100, variable=self.row_h_var, orient=tk.HORIZONTAL, length=60, command=self.on_row_height_change).pack(side=tk.LEFT, padx=2)
        self.row_height_label = ttk.Label(view_group, text="30px", font=('Microsoft YaHei', 8))
        self.row_height_label.pack(side=tk.LEFT, padx=(0, 5))

        ttk.Label(view_group, text="列宽:").pack(side=tk.LEFT)
        ttk.Scale(view_group, from_=0.5, to=3.0, variable=self.col_w_var, orient=tk.HORIZONTAL, length=60, command=self.on_col_width_change).pack(side=tk.LEFT, padx=2)
        self.col_width_label = ttk.Label(view_group, text="100%", font=('Microsoft YaHei', 8))
        self.col_width_label.pack(side=tk.LEFT)

        # D. 选中项精调
        sel_group = ttk.Labelframe(middle_settings, text="选中精调", padding=10, bootstyle=INFO)
        sel_group.pack(side=tk.LEFT, fill=tk.Y, padx=5)
        
        ttk.Label(sel_group, text="行高:").pack(side=tk.LEFT)
        tk.Spinbox(sel_group, from_=10, to=500, textvariable=self.sel_row_h, width=3, command=self.adjust_selected_row_height).pack(side=tk.LEFT, padx=2)
        self.sel_row_h.trace_add("write", lambda *args: self.adjust_selected_row_height())

        ttk.Label(sel_group, text="列宽:").pack(side=tk.LEFT, padx=(5, 0))
        tk.Spinbox(sel_group, from_=10, to=1000, textvariable=self.sel_col_w, width=4, command=self.adjust_selected_col_width).pack(side=tk.LEFT, padx=2)
        self.sel_col_w.trace_add("write", lambda *args: self.adjust_selected_col_width())

        # --- 第三行：操作按钮、主题切换与图例 (整合) ---
        action_bar = ttk.Frame(top_panel)
        action_bar.pack(fill=tk.X, pady=(5, 0))

        # 核心按钮
        self.btn_compare = ttk.Button(action_bar, text="🚀 开始比对差异", command=self.compare_files, bootstyle=SUCCESS, width=18)
        self.btn_compare.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(action_bar, text="📊 导出结果", command=self.export_results, bootstyle=INFO, width=15).pack(side=tk.LEFT, padx=5)

        # 主题
        ttk.Label(action_bar, text="🎨 主题:", font=('Microsoft YaHei', 9)).pack(side=tk.LEFT, padx=(15, 5))
        self.theme_combo = ttk.Combobox(action_bar, values=['cosmo', 'flatly', 'litera', 'minty', 'lumen', 'sandstone', 'yeti', 'pulse', 'united', 'morph', 'journal', 'darkly', 'superhero', 'solar', 'cyborg', 'vapor'], state="readonly", width=10)
        self.theme_combo.set('cosmo')
        self.theme_combo.pack(side=tk.LEFT)
        self.theme_combo.bind('<<ComboboxSelected>>', self.change_theme)

        # 图例说明 (放在最右侧)
        legend_box = ttk.Frame(action_bar)
        legend_box.pack(side=tk.RIGHT)
        ttk.Label(legend_box, text="修改", background='#E6F3FF', foreground='black', width=4, anchor=tk.CENTER).pack(side=tk.LEFT, padx=2)
        ttk.Label(legend_box, text="新增", background='#CCFFCC', foreground='black', width=4, anchor=tk.CENTER).pack(side=tk.LEFT, padx=2)
        ttk.Label(legend_box, text="删除", background='#FFFFCC', foreground='black', width=4, anchor=tk.CENTER).pack(side=tk.LEFT, padx=2)
        ttk.Label(legend_box, text="单元格差异", background='#FFCCCC', foreground='#CC0000', padding=(5, 0)).pack(side=tk.LEFT, padx=5)

        # 2. 结果展示区域 (中间填满)
        result_container = ttk.Frame(main_container)
        result_container.pack(fill=tk.BOTH, expand=True)
        
        self.paned = ttk.Panedwindow(result_container, orient=tk.HORIZONTAL)
        self.paned.pack(fill=tk.BOTH, expand=True)

        self.left_frame = ttk.Labelframe(self.paned, text="文件 1 (旧)", bootstyle=SECONDARY)
        self.paned.add(self.left_frame, weight=1)
        self.table_left = ScrollableTable(self.left_frame)
        self.table_left.pack(fill=tk.BOTH, expand=True)
        self.table_left.on_select_callback = lambda r, c: self.on_cell_selected(r, c, source="left")
        
        self.right_frame = ttk.Labelframe(self.paned, text="文件 2 (新)", bootstyle=SECONDARY)
        self.paned.add(self.right_frame, weight=1)
        self.table_right = ScrollableTable(self.right_frame)
        self.table_right.pack(fill=tk.BOTH, expand=True)
        self.table_right.on_select_callback = lambda r, c: self.on_cell_selected(r, c, source="right")

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

    def on_cell_selected(self, row_idx, col_idx, source="left"):
        """处理单元格选中事件"""
        target_table = self.table_right if source == "left" else self.table_left
        source_table = self.table_left if source == "left" else self.table_right
        
        # 同步选中状态
        target_table.selected_cell = (row_idx, col_idx)
        target_table._update_view()
        
        # 更新调整控件的值
        self.sel_row_h.set(source_table.row_heights[row_idx])
        self.sel_col_w.set(source_table.col_widths[col_idx])

    def adjust_selected_row_height(self):
        """调整选中行的高度"""
        if not self.table_left.data: return
        
        row_idx = self.table_left.selected_cell[0] if self.table_left.selected_cell else None
        if row_idx is None:
            row_idx = self.table_right.selected_cell[0] if self.table_right.selected_cell else None
        
        if row_idx is not None:
            try:
                new_h = int(self.sel_row_h.get())
                if 10 <= new_h <= 500:
                    self.table_left.row_heights[row_idx] = new_h
                    self.table_right.row_heights[row_idx] = new_h
                    self.update_tables_view()
            except (ValueError, tk.TclError):
                pass

    def adjust_selected_col_width(self):
        """调整选中列的宽度"""
        if not self.table_left.data: return
        
        col_idx = self.table_left.selected_cell[1] if self.table_left.selected_cell else None
        if col_idx is None:
            col_idx = self.table_right.selected_cell[1] if self.table_right.selected_cell else None
            
        if col_idx is not None:
            try:
                new_w = int(self.sel_col_w.get())
                if 10 <= new_w <= 1000:
                    self.table_left.col_widths[col_idx] = new_w
                    self.table_right.col_widths[col_idx] = new_w
                    self.update_tables_view()
            except (ValueError, tk.TclError):
                pass

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

    def sync_freeze_rows(self, *args):
        """同步左右表格的冻结行设置"""
        val = self.freeze_rows.get()
        self.table_left.freeze_rows.set(val)
        self.table_right.freeze_rows.set(val)
        self.update_tables_view()

    def sync_freeze_cols(self, *args):
        """同步左右表格的冻结列设置"""
        val = self.freeze_cols.get()
        self.table_left.freeze_cols.set(val)
        self.table_right.freeze_cols.set(val)
        self.update_tables_view()

    def on_row_height_change(self, value):
        """行高改变"""
        h = int(float(value))
        self.row_height_label.config(text=f"{h}px")
        
        # 更新所有行的基础行高
        if hasattr(self.table_left, 'row_heights') and self.table_left.row_heights:
            self.table_left.row_heights = [h] * len(self.table_left.row_heights)
            self.table_right.row_heights = [h] * len(self.table_right.row_heights)
            
        self.table_left.row_height = h
        self.table_right.row_height = h
        self.update_tables_view()

    def on_col_width_change(self, value):
        """全局列宽缩放改变"""
        factor = float(value)
        self.col_width_label.config(text=f"{int(factor * 100)}%")
        
        # 基于原始基础宽度进行缩放
        if hasattr(self.table_left, 'base_col_widths') and self.table_left.base_col_widths:
            self.table_left.col_widths = [int(w * factor) for w in self.table_left.base_col_widths]
            self.table_right.col_widths = [int(w * factor) for w in self.table_right.base_col_widths]
            self.update_tables_view()

    def update_tables_view(self):
        """更新左右表格的视图"""
        # 更新滚动区域
        self.table_left._update_scroll_region()
        self.table_right._update_scroll_region()
        # 强制重新渲染可见区域
        self.table_left._update_view()
        self.table_right._update_view()

    def toggle_key_mode(self):
        """切换关键列比对模式"""
        if self.key_mode_var.get():
            self.btn_select_keys.config(state=tk.NORMAL)
        else:
            self.btn_select_keys.config(state=tk.DISABLED)
            self.key_columns = []
            self.key_cols_label.config(text="未选择关键列")

    def select_key_columns(self):
        """打开对话框选择关键列"""
        file1 = self.file1_path.get()
        sheet1 = self.sheet_combo1.get()
        
        if not file1 or not sheet1:
            messagebox.showwarning("提示", "请先选择文件和 Sheet 以获取列信息")
            return
            
        try:
            # 临时读取表头
            df = ExcelDiffer.read_excel_raw(file1, sheet1, handle_merged=False)
            columns = list(df.columns)
            
            # 创建选择窗口
            top = tk.Toplevel(self.root)
            top.title("选择关键列 (可多选)")
            top.geometry("600x500")
            top.transient(self.root)
            top.grab_set()
            
            # 居中
            x = self.root.winfo_x() + (self.root.winfo_width() - 600) // 2
            y = self.root.winfo_y() + (self.root.winfo_height() - 500) // 2
            top.geometry(f"+{x}+{y}")
            
            main_v_frame = ttk.Frame(top, padding=10)
            main_v_frame.pack(fill=tk.BOTH, expand=True)
            
            ttk.Label(main_v_frame, text="请勾选作为主键的列：", font=('Microsoft YaHei', 10, 'bold')).pack(pady=(0, 10))
            
            # 使用带滚动条的列表显示复选框
            list_frame = ttk.Frame(main_v_frame)
            list_frame.pack(fill=tk.BOTH, expand=True)

            canvas = tk.Canvas(list_frame, highlightthickness=0)
            scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
            scrollable_frame = ttk.Frame(canvas)

            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )

            # 让 canvas 宽度自适应
            canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            def on_canvas_configure(event):
                canvas.itemconfig(canvas_window, width=event.width)
            canvas.bind('<Configure>', on_canvas_configure)

            canvas.configure(yscrollcommand=scrollbar.set)

            check_vars = {}
            # 每行显示 3 列
            num_cols = 3
            for i, col in enumerate(columns):
                var = tk.BooleanVar(value=(col in self.key_columns))
                check_vars[col] = var
                cb = ttk.Checkbutton(scrollable_frame, text=col, variable=var)
                cb.grid(row=i // num_cols, column=i % num_cols, sticky=tk.W, padx=10, pady=5)
            
            # 设置列权重，使列等宽
            for c in range(num_cols):
                scrollable_frame.grid_columnconfigure(c, weight=1)

            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")

            def on_confirm():
                selected = [col for col, var in check_vars.items() if var.get()]
                if not selected:
                    if not messagebox.askyesno("提示", "未选择任何关键列，是否关闭主键模式？"):
                        return
                    self.key_mode_var.set(False)
                    self.toggle_key_mode()
                else:
                    self.key_columns = selected
                    self.key_cols_label.config(text=f"已选: {', '.join(selected[:2])}{'...' if len(selected) > 2 else ''}")
                top.destroy()

            ttk.Button(main_v_frame, text="确定", command=on_confirm).pack(pady=10)
            
        except Exception as e:
            messagebox.showerror("错误", f"获取列信息失败: {e}")

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

        # 获取关键列配置
        key_cols = self.key_columns if self.key_mode_var.get() else None
        if self.key_mode_var.get() and not key_cols:
            messagebox.showwarning("提示", "您开启了主键比对模式，但未选择任何关键列。")
            return

        # 显示加载弹窗
        self.loading = LoadingDialog(self.root, message="正在读取文件并比对差异，请稍候...")
        
        # 禁用按钮
        self.btn_compare.config(state=tk.DISABLED)

        # 开启线程执行比对
        thread = threading.Thread(target=self._compare_thread_task, args=(file1, file2, sheet1, sheet2, key_cols))
        thread.daemon = True
        thread.start()

    def change_theme(self, event=None):
        """切换界面主题"""
        theme = self.theme_combo.get()
        style = ttk.Style()
        style.theme_use(theme)
        # 切换主题后重新设置一些自定义样式
        self.setup_styles()

    def _compare_thread_task(self, file1, file2, sheet1, sheet2, key_cols=None):
        """后台线程执行比对任务 - 增加进度反馈"""
        try:
            # 1. 读取数据
            self.loading.update_progress(10, "正在读取文件 1...")
            df1 = ExcelDiffer.read_excel_raw(file1, sheet1, handle_merged=True)
            
            self.loading.update_progress(30, "正在读取文件 2...")
            df2 = ExcelDiffer.read_excel_raw(file2, sheet2, handle_merged=True)
            
            # 2. 执行比对
            self.loading.update_progress(50, f"正在进行{'主键' if key_cols else '序列'}比对...")
            columns, results = ExcelDiffer.compare_dataframes(df1, df2, key_columns=key_cols)
            
            # 3. 准备渲染
            total_rows = len(results)
            self.loading.update_progress(80, f"比对完成，正在准备渲染 {total_rows} 行数据...")
            
            # 回到主线程更新 UI
            self.root.after(0, lambda: self._update_ui_after_compare(columns, results))

        except Exception as e:
            self.root.after(0, lambda: self._handle_compare_error(str(e)))

    def _update_ui_after_compare(self, columns, results):
        """比对完成后在主线程更新 UI"""
        self.last_columns = columns
        self.last_results = results
        
        # 关闭加载弹窗
        if hasattr(self, 'loading'):
            self.loading.close()
        
        # 恢复按钮状态
        self.btn_compare.config(state=tk.NORMAL)
        
        # 展示结果
        self.show_diff(columns, results)
        
        messagebox.showinfo("完成", f"比对完成！共发现 {len(results)} 行差异。")

    def _handle_compare_error(self, error_msg):
        """处理比对过程中的错误（主线程）"""
        # 关闭加载弹窗
        if hasattr(self, 'loading'):
            self.loading.close()
            
        self.btn_compare.config(state=tk.NORMAL)
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
