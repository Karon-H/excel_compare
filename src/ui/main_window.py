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
        if hasattr(self, 'header_widgets'):
            for widget in self.header_widgets:
                if widget: widget.destroy()
        self.header_widgets = []
        self.cell_widgets = {}
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
        """核心：虚拟化渲染，支持自由冻结和个别行列调整"""
        if not self.data:
            return

        # 获取滚动偏移（像素）
        total_width = sum(self.col_widths)
        total_height = sum(self.row_heights)
        
        h_start_ratio = self.canvas.xview()[0]
        v_start_ratio = self.canvas.yview()[0]
        h_end_ratio = self.canvas.xview()[1]
        v_end_ratio = self.canvas.yview()[1]
        
        x_offset = h_start_ratio * total_width
        y_offset = v_start_ratio * total_height

        # 计算可见行范围 (包含表头，表头是第 0 行)
        start_row = self._get_row_at_y(y_offset)
        end_row = self._get_row_at_y(v_end_ratio * total_height) + 1
        
        start_row = max(0, start_row)
        end_row = min(len(self.data) + 1, end_row)

        # 计算可见列范围
        start_col = self._get_col_at_x(x_offset)
        end_col = self._get_col_at_x(h_end_ratio * total_width) + 1
        
        start_col = max(0, start_col)
        end_col = min(len(self.columns), end_col)

        f_rows = self.freeze_rows.get()
        f_cols = self.freeze_cols.get()

        # 1. 处理表头 (行索引为 0)
        cols_to_render_header = set(range(start_col, end_col))
        for c in range(f_cols):
            cols_to_render_header.add(c)
            
        for j in range(len(self.columns)):
            if j in cols_to_render_header:
                if self.header_widgets[j] is None:
                    lbl = tk.Label(self.scrollable_frame, text=self.columns[j], 
                                 font=('Arial', 10, 'bold'), bg='#F0F0F0',
                                 borderwidth=1, relief="raised", padx=5)
                    self._bind_mouse_wheel(lbl)
                    lbl.bind("<Button-1>", lambda e, c=j: self._on_cell_click(e, 0, c))
                    self.header_widgets[j] = lbl
                
                # 冻结位置计算
                target_x = self._get_col_x(j)
                target_y = 0
                
                if f_rows > 0:
                    target_y = y_offset
                if j < f_cols:
                    target_x = x_offset + self._get_col_x(j)
                
                # 选中状态背景
                bg_color = '#CCE8FF' if self.selected_cell == (0, j) else '#F0F0F0'
                self.header_widgets[j].config(bg=bg_color)
                
                self.header_widgets[j].place(x=target_x, y=target_y, 
                                           width=self.col_widths[j], height=self.row_heights[0])
                if f_rows > 0 or j < f_cols:
                    self.header_widgets[j].lift()
            elif self.header_widgets[j] is not None:
                self.header_widgets[j].destroy()
                self.header_widgets[j] = None

        # 2. 处理数据行 (数据行索引 i 从 0 开始，在 row_heights 中对应 i+1)
        current_visible_keys = set()
        
        rows_to_render = set(range(max(1, start_row), end_row))
        for r in range(1, f_rows): # 冻结数据行
            if r < len(self.row_heights):
                rows_to_render.add(r)
        
        for r_idx in rows_to_render:
            i = r_idx - 1 # 数据索引
            row_vals = self.data[i]
            tags = self.tags_list[i]
            status = row_vals[0] if row_vals else ""
            
            cols_to_render = set(range(start_col, end_col))
            for c in range(f_cols):
                cols_to_render.add(c)

            for j in cols_to_render:
                val = row_vals[j]
                key = (i, j)
                current_visible_keys.add(key)
                
                if key not in self.cell_widgets:
                    val_str = str(val)
                    cell_bg = "white"
                    fg_color = "black"
                    
                    if j == 0:
                        if status == "新增": cell_bg = '#CCFFCC'
                        elif status == "删除": cell_bg = '#FFFFCC'
                        elif status == "修改": cell_bg = '#E6F3FF'
                        elif status == "一致": cell_bg = '#F8F8F8'
                    else:
                        if 'added' in tags: cell_bg = '#F0FFF0'
                        elif 'deleted' in tags: cell_bg = '#FFFFF0'
                    
                    if val_str.startswith(">>> ") and val_str.endswith(" <<<"):
                        val_str = val_str[4:-4]
                        cell_bg = '#FFCCCC'
                        fg_color = '#CC0000'
                    
                    lbl = tk.Label(self.scrollable_frame, text=val_str, font=('Arial', 10),
                                 bg=cell_bg, fg=fg_color, padx=5, 
                                 borderwidth=1, relief="groove", anchor="w")
                    self.cell_widgets[key] = lbl
                    self._bind_mouse_wheel(lbl)
                    lbl.bind("<Button-1>", lambda e, r=r_idx, c=j: self._on_cell_click(e, r, c))
                    lbl.original_bg = cell_bg

                # 冻结位置计算
                target_x = self._get_col_x(j)
                target_y = self._get_row_y(r_idx)
                
                if r_idx < f_rows:
                    target_y = y_offset + self._get_row_y(r_idx)
                if j < f_cols:
                    target_x = x_offset + self._get_col_x(j)
                
                # 选中状态高亮
                if self.selected_cell == (r_idx, j):
                    self.cell_widgets[key].config(bg='#CCE8FF')
                else:
                    self.cell_widgets[key].config(bg=self.cell_widgets[key].original_bg)
                
                self.cell_widgets[key].place(x=target_x, y=target_y, 
                                           width=self.col_widths[j], height=self.row_heights[r_idx])
                if r_idx < f_rows or j < f_cols:
                    self.cell_widgets[key].lift()

        # 3. 清理
        to_remove = []
        for key in self.cell_widgets:
            if key not in current_visible_keys:
                i_idx, c_idx = key
                r_idx = i_idx + 1
                if (r_idx < start_row - 10 or r_idx > end_row + 10 or 
                    c_idx < start_col - 5 or c_idx > end_col + 5):
                    if not (r_idx < f_rows or c_idx < f_cols):
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

        # 界面控制变量
        self.freeze_rows = tk.IntVar(value=1)
        self.freeze_cols = tk.IntVar(value=1)
        self.row_h_var = tk.DoubleVar(value=30)
        self.col_w_var = tk.DoubleVar(value=1.0)
        self.sel_row_h = tk.IntVar(value=30)
        self.sel_col_w = tk.IntVar(value=100)

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
        version = "V1.3.0"
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
更新日期：2025-12-26

更新亮点：
- 智能行对齐算法：支持识别中间插入行和删除行，比对更精准
- 表格冻结功能：支持首行首列冻结，方便大数据量查看
- 单元格大小调整：支持 50% - 200% 的比例缩放
- 性能优化：支持大数据量虚拟化滚动
"""
        messagebox.showinfo("版本说明", version_text)

    def create_widgets(self):
        # 1. 结果展示区域 (双栏显示) - 先创建表格
        result_main_frame = ttk.Frame(self.root, padding=10)
        result_main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.paned = ttk.PanedWindow(result_main_frame, orient=tk.HORIZONTAL)
        self.paned.pack(fill=tk.BOTH, expand=True)

        self.left_frame = ttk.LabelFrame(self.paned, text="文件 1 (旧)")
        self.paned.add(self.left_frame, weight=1)
        self.table_left = ScrollableTable(self.left_frame)
        self.table_left.pack(fill=tk.BOTH, expand=True)
        self.table_left.on_select_callback = lambda r, c: self.on_cell_selected(r, c, source="left")
        
        self.right_frame = ttk.LabelFrame(self.paned, text="文件 2 (新)")
        self.paned.add(self.right_frame, weight=1)
        self.table_right = ScrollableTable(self.right_frame)
        self.table_right.pack(fill=tk.BOTH, expand=True)
        self.table_right.on_select_callback = lambda r, c: self.on_cell_selected(r, c, source="right")

        # 2. 顶部控制面板
        control_frame = ttk.LabelFrame(self.root, text="文件选择与设置", padding=10)
        control_frame.pack(fill=tk.X, padx=10, pady=5, before=result_main_frame) # 放在结果区域上方

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

        # 底部控制区容器
        bottom_frame = ttk.Frame(self.root, padding=5)
        bottom_frame.pack(fill=tk.X, padx=10, pady=5)

        # 1. 核心操作区 (左侧)
        action_group = ttk.LabelFrame(bottom_frame, text="操作控制", padding=5)
        action_group.pack(side=tk.LEFT, fill=tk.Y, padx=5)
        
        self.btn_compare_ref = ttk.Button(action_group, text="开始比对", command=self.compare_files, width=12)
        self.btn_compare_ref.pack(side=tk.LEFT, padx=5)
        
        self.btn_export = ttk.Button(action_group, text="导出结果", command=self.export_results, state=tk.DISABLED, width=12)
        self.btn_export.pack(side=tk.LEFT, padx=5)
        
        sync_frame = ttk.Frame(action_group)
        sync_frame.pack(side=tk.LEFT, padx=10)
        ttk.Checkbutton(sync_frame, text="同步纵滚", variable=self.sync_v_scroll).pack(side=tk.TOP, anchor=tk.W)
        ttk.Checkbutton(sync_frame, text="同步横滚", variable=self.sync_h_scroll).pack(side=tk.TOP, anchor=tk.W)

        # 2. 显示设置区 (中间)
        display_group = ttk.LabelFrame(bottom_frame, text="显示设置", padding=5)
        display_group.pack(side=tk.LEFT, fill=tk.Y, padx=5)
        
        # 冻结设置
        freeze_f = ttk.Frame(display_group)
        freeze_f.pack(side=tk.LEFT, padx=5)
        
        f_row_f = ttk.Frame(freeze_f)
        f_row_f.pack(fill=tk.X)
        ttk.Label(f_row_f, text="冻结行:").pack(side=tk.LEFT)
        tk.Spinbox(f_row_f, from_=0, to=50, textvariable=self.freeze_rows, width=3).pack(side=tk.LEFT, padx=2)
        self.freeze_rows.trace_add("write", self.sync_freeze_rows)

        f_col_f = ttk.Frame(freeze_f)
        f_col_f.pack(fill=tk.X, pady=(2, 0))
        ttk.Label(f_col_f, text="冻结列:").pack(side=tk.LEFT)
        tk.Spinbox(f_col_f, from_=0, to=20, textvariable=self.freeze_cols, width=3).pack(side=tk.LEFT, padx=2)
        self.freeze_cols.trace_add("write", self.sync_freeze_cols)

        # 尺寸滑块
        slider_f = ttk.Frame(display_group)
        slider_f.pack(side=tk.LEFT, padx=10)
        
        h_f = ttk.Frame(slider_f)
        h_f.pack(fill=tk.X)
        ttk.Label(h_f, text="行高:").pack(side=tk.LEFT)
        ttk.Scale(h_f, from_=20, to=100, variable=self.row_h_var, orient=tk.HORIZONTAL, length=80, command=self.on_row_height_change).pack(side=tk.LEFT, padx=5)
        self.row_height_label = ttk.Label(h_f, text="30px", width=5)
        self.row_height_label.pack(side=tk.LEFT)

        w_f = ttk.Frame(slider_f)
        w_f.pack(fill=tk.X, pady=(2, 0))
        ttk.Label(w_f, text="列宽:").pack(side=tk.LEFT)
        ttk.Scale(w_f, from_=0.5, to=3.0, variable=self.col_w_var, orient=tk.HORIZONTAL, length=80, command=self.on_col_width_change).pack(side=tk.LEFT, padx=5)
        self.col_width_label = ttk.Label(w_f, text="100%", width=5)
        self.col_width_label.pack(side=tk.LEFT)

        # 3. 选中调整区 (右侧)
        selection_group = ttk.LabelFrame(bottom_frame, text="选中项精调", padding=5)
        selection_group.pack(side=tk.LEFT, fill=tk.Y, padx=5)
        
        sel_h_f = ttk.Frame(selection_group)
        sel_h_f.pack(fill=tk.X)
        ttk.Label(sel_h_f, text="选中行高:").pack(side=tk.LEFT)
        self.sel_row_spin = tk.Spinbox(sel_h_f, from_=10, to=500, textvariable=self.sel_row_h, width=5)
        self.sel_row_spin.pack(side=tk.LEFT, padx=5)
        self.sel_row_h.trace_add("write", lambda *args: self.adjust_selected_row_height())
        
        sel_w_f = ttk.Frame(selection_group)
        sel_w_f.pack(fill=tk.X, pady=(2, 0))
        ttk.Label(sel_w_f, text="选中列宽:").pack(side=tk.LEFT)
        self.sel_col_spin = tk.Spinbox(sel_w_f, from_=10, to=1000, textvariable=self.sel_col_w, width=5)
        self.sel_col_spin.pack(side=tk.LEFT, padx=5)
        self.sel_col_w.trace_add("write", lambda *args: self.adjust_selected_col_width())

        # 4. 图例区 (最右侧)
        legend_group = ttk.LabelFrame(bottom_frame, text="图例说明", padding=5)
        legend_group.pack(side=tk.LEFT, fill=tk.Y, padx=5, expand=True)
        
        l_row1 = ttk.Frame(legend_group)
        l_row1.pack(fill=tk.X)
        ttk.Label(l_row1, text="修改", background='#E6F3FF', width=6, anchor=tk.CENTER).pack(side=tk.LEFT, padx=2)
        ttk.Label(l_row1, text="新增", background='#CCFFCC', width=6, anchor=tk.CENTER).pack(side=tk.LEFT, padx=2)
        ttk.Label(l_row1, text="删除", background='#FFFFCC', width=6, anchor=tk.CENTER).pack(side=tk.LEFT, padx=2)
        
        l_row2 = ttk.Frame(legend_group)
        l_row2.pack(fill=tk.X, pady=(2, 0))
        ttk.Label(l_row2, text="单元格差异: >>>内容<<<", background='#FFCCCC', foreground='#CC0000', font=('Arial', 8)).pack(side=tk.LEFT, padx=2)

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
