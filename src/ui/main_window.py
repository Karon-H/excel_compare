import os
from PyQt5 import QtWidgets, QtCore, QtGui
from src.logic.excel_processor import ExcelDiffer
from src.logic.csv_processor import CSVDiffer
from src.ui.statistics_dialog import StatisticsDialog
from src.ui.full_compare_dialog import FullCompareSummaryDialog
from src.ui.frozen_table_view import FrozenTableView
from src.ui.diff_table_model import DiffTableModel


class DifferFactory:
    """比对引擎工厂类"""
    @staticmethod
    def get_differ(filepath):
        ext = os.path.splitext(filepath)[1].lower()
        if ext in ['.csv']:
            return CSVDiffer()
        else:
            # 默认为 Excel 比对引擎
            return ExcelDiffer()

class CompareWorker(QtCore.QObject):
    progress = QtCore.pyqtSignal(int, str)
    finished = QtCore.pyqtSignal(list, list)
    error = QtCore.pyqtSignal(str)

    def __init__(self, file1, file2, sheet1, sheet2, key_cols, header_row=None, has_header=True):
        super().__init__()
        self.file1 = file1
        self.file2 = file2
        self.sheet1 = sheet1
        self.sheet2 = sheet2
        self.key_cols = key_cols
        self.header_row = header_row
        self.has_header = has_header

    @QtCore.pyqtSlot()
    def run(self):
        try:
            differ = DifferFactory.get_differ(self.file1)
            self.progress.emit(10, "正在读取文件 1...")
            df1 = differ.read_data(self.file1, self.sheet1, handle_merged=True, header_row=self.header_row, has_header=self.has_header)
            self.progress.emit(30, "正在读取文件 2...")
            df2 = differ.read_data(self.file2, self.sheet2, handle_merged=True, header_row=self.header_row, has_header=self.has_header)
            mode_text = "主键" if self.key_cols else "序列"
            self.progress.emit(50, f"正在进行{mode_text}比对...")
            columns, results = differ.compare(df1, df2, key_columns=self.key_cols)
            total_rows = len(results)
            self.progress.emit(80, f"比对完成，正在准备渲染 {total_rows} 行数据...")
            self.finished.emit(columns, results)
        except Exception as e:
            self.error.emit(str(e))


class FullCompareWorker(QtCore.QObject):
    progress = QtCore.pyqtSignal(int, str)
    finished = QtCore.pyqtSignal(list)
    error = QtCore.pyqtSignal(str)

    def __init__(self, file1, file2, key_cols, header_row=None, has_header=True):
        super().__init__()
        self.file1 = file1
        self.file2 = file2
        self.key_cols = key_cols
        self.header_row = header_row
        self.has_header = has_header

    @QtCore.pyqtSlot()
    def run(self):
        try:
            differ = DifferFactory.get_differ(self.file1)
            summary = differ.compare_all(
                self.file1, self.file2, 
                key_columns=self.key_cols, 
                header_row=self.header_row, 
                has_header=self.has_header,
                progress_callback=self.progress.emit
            )
            self.finished.emit(summary)
        except Exception as e:
            self.error.emit(str(e))


class KeyColumnsDialog(QtWidgets.QDialog):
    def __init__(self, columns, selected, parent=None):
        super().__init__(parent)
        self.setWindowTitle("选择关键列")
        self.resize(520, 420)
        layout = QtWidgets.QVBoxLayout(self)
        label = QtWidgets.QLabel("请勾选作为主键的列：")
        layout.addWidget(label)
        self.list_widget = QtWidgets.QListWidget()
        for col in columns:
            item = QtWidgets.QListWidgetItem(col)
            item.setFlags(item.flags() | QtCore.Qt.ItemIsUserCheckable)
            item.setCheckState(QtCore.Qt.Checked if col in selected else QtCore.Qt.Unchecked)
            self.list_widget.addItem(item)
        layout.addWidget(self.list_widget)
        buttons = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def selected_columns(self):
        selected = []
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if item.checkState() == QtCore.Qt.Checked:
                selected.append(item.text())
        return selected


from src.logic.excel_processor import ExcelDiffer

class HTMLDelegate(QtWidgets.QStyledItemDelegate):
    """
    用于在单元格内渲染 HTML 的委派类
    """
    def paint(self, painter, option, index):
        options = QtWidgets.QStyleOptionViewItem(option)
        self.initStyleOption(options, index)

        painter.save()
        
        doc = QtGui.QTextDocument()
        # 处理可能存在的 None
        text = options.text if options.text else ""
        # 如果包含 HTML 标签，则按 HTML 渲染，否则按普通文本
        if "<span" in text or "<b" in text or "<br" in text:
            doc.setHtml(text)
        else:
            doc.setPlainText(text)

        # 绘制背景
        options.text = ""
        style = options.widget.style()
        style.drawControl(QtWidgets.QStyle.CE_ItemViewItem, options, painter)

        # 计算文本区域
        textRect = style.subElementRect(QtWidgets.QStyle.SE_ItemViewItemText, options)
        
        # 居中对齐处理 (可选)
        painter.translate(textRect.topLeft())
        
        # 设置绘制上下文
        ctx = QtGui.QAbstractTextDocumentLayout.PaintContext()
        # 处理选中时的文字颜色
        if option.state & QtWidgets.QStyle.State_Selected:
            ctx.palette.setColor(QtGui.QPalette.Text, option.palette.color(QtGui.QPalette.HighlightedText))
        
        # 绘制
        doc.documentLayout().draw(painter, ctx)

        painter.restore()

    def sizeHint(self, option, index):
        options = QtWidgets.QStyleOptionViewItem(option)
        self.initStyleOption(options, index)
        doc = QtGui.QTextDocument()
        if "<span" in options.text:
            doc.setHtml(options.text)
        else:
            doc.setPlainText(options.text)
        return QtCore.QSize(int(doc.idealWidth()) + 10, int(doc.size().height()) + 5)

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self, version):
        super().__init__()
        self.version = version
        self.setWindowTitle(f"Excel 差异比对工具 - {version}")
        self.setAcceptDrops(True)
        self._syncing_scroll = False
        self._syncing_selection = False
        self.last_columns = None
        self.all_results = None
        self.display_results = None
        self.key_columns = []
        self.base_col_widths = []
        self.left_model = None
        self.right_model = None
        self.progress_dialog = None
        self.setup_ui()
        self.setup_menu()
        self.set_window_icon()
        self.load_settings()

    def load_settings(self):
        """从 QSettings 加载用户偏好设置"""
        settings = QtCore.QSettings("MyCompany", "ExcelDiffer")
        
        # 窗口几何状态
        geom = settings.value("geometry")
        if geom:
            self.restoreGeometry(geom)
        state = settings.value("windowState")
        if state:
            self.restoreState(state)

        # 文件路径 (只恢复目录，防止文件不存在报错)
        f1 = settings.value("file1", "")
        f2 = settings.value("file2", "")
        if f1 and os.path.exists(f1):
            self.file1_edit.setText(f1)
            self.load_sheets(f1, self.sheet_combo1)
            s1 = settings.value("sheet1", "")
            if s1:
                idx = self.sheet_combo1.findText(s1)
                if idx >= 0:
                    self.sheet_combo1.setCurrentIndex(idx)

        if f2 and os.path.exists(f2):
            self.file2_edit.setText(f2)
            self.load_sheets(f2, self.sheet_combo2)
            s2 = settings.value("sheet2", "")
            if s2:
                idx = self.sheet_combo2.findText(s2)
                if idx >= 0:
                    self.sheet_combo2.setCurrentIndex(idx)

        # 比对配置
        self.key_mode_cb.setChecked(settings.value("key_mode", False, type=bool))
        self.key_columns = settings.value("key_columns", [])
        if self.key_columns:
            self.key_label.setText("已选: " + ", ".join(self.key_columns[:2]))
            self.key_label.setStyleSheet("color: black; font-weight: bold;")
        
        self.no_header_cb.setChecked(settings.value("no_header", False, type=bool))
        self.header_row_spin.setValue(settings.value("header_row", 1, type=int))

        # 视图配置
        self.sync_v_cb.setChecked(settings.value("sync_v", True, type=bool))
        self.sync_h_cb.setChecked(settings.value("sync_h", True, type=bool))
        self.freeze_row_spin.setValue(settings.value("freeze_row", 0, type=int))
        self.freeze_col_spin.setValue(settings.value("freeze_col", 0, type=int))
        self.row_height_spin.setValue(settings.value("row_height", 30, type=int))
        self.col_width_spin.setValue(settings.value("col_width", 1.0, type=float))

    def save_settings(self):
        """保存用户偏好设置到 QSettings"""
        settings = QtCore.QSettings("MyCompany", "ExcelDiffer")
        
        # 窗口几何状态
        settings.setValue("geometry", self.saveGeometry())
        settings.setValue("windowState", self.saveState())

        # 文件路径
        settings.setValue("file1", self.file1_edit.text())
        settings.setValue("file2", self.file2_edit.text())
        settings.setValue("sheet1", self.sheet_combo1.currentText())
        settings.setValue("sheet2", self.sheet_combo2.currentText())

        # 比对配置
        settings.setValue("key_mode", self.key_mode_cb.isChecked())
        settings.setValue("key_columns", self.key_columns)
        settings.setValue("no_header", self.no_header_cb.isChecked())
        settings.setValue("header_row", self.header_row_spin.value())

        # 视图配置
        settings.setValue("sync_v", self.sync_v_cb.isChecked())
        settings.setValue("sync_h", self.sync_h_cb.isChecked())
        settings.setValue("freeze_row", self.freeze_row_spin.value())
        settings.setValue("freeze_col", self.freeze_col_spin.value())
        settings.setValue("row_height", self.row_height_spin.value())
        settings.setValue("col_width", self.col_width_spin.value())

    def closeEvent(self, event):
        """窗口关闭前保存设置"""
        self.save_settings()
        super().closeEvent(event)

    def setup_menu(self):
        menubar = self.menuBar()
        help_menu = menubar.addMenu("帮助")
        usage_action = QtWidgets.QAction("使用说明", self)
        version_action = QtWidgets.QAction("版本说明", self)
        usage_action.triggered.connect(self.show_usage)
        version_action.triggered.connect(self.show_version)
        help_menu.addAction(usage_action)
        help_menu.addAction(version_action)

    def setup_ui(self):
        central = QtWidgets.QWidget()
        self.setCentralWidget(central)
        main_layout = QtWidgets.QVBoxLayout(central)
        main_layout.setContentsMargins(5, 5, 5, 5)
        main_layout.setSpacing(5)

        # 1. 顶部配置区域 (精简布局)
        config_panel = QtWidgets.QWidget()
        config_layout = QtWidgets.QVBoxLayout(config_panel)
        config_layout.setContentsMargins(0, 0, 0, 0)
        config_layout.setSpacing(2)

        # 数据源行
        data_group = QtWidgets.QGroupBox("数据源与比对配置")
        data_layout = QtWidgets.QGridLayout(data_group)
        data_layout.setContentsMargins(10, 5, 10, 5)
        data_layout.setSpacing(8)
        
        self.file1_edit = QtWidgets.QLineEdit()
        self.file2_edit = QtWidgets.QLineEdit()
        self.sheet_combo1 = QtWidgets.QComboBox()
        self.sheet_combo2 = QtWidgets.QComboBox()
        self.file1_edit.setPlaceholderText("选择或拖入旧 Excel 文件...")
        self.file2_edit.setPlaceholderText("选择或拖入新 Excel 文件...")
        
        browse_btn1 = QtWidgets.QPushButton("浏览...")
        browse_btn2 = QtWidgets.QPushButton("浏览...")
        browse_btn1.setFixedWidth(60)
        browse_btn2.setFixedWidth(60)
        browse_btn1.clicked.connect(lambda: self.browse_file(1))
        browse_btn2.clicked.connect(lambda: self.browse_file(2))

        data_layout.addWidget(QtWidgets.QLabel("旧文件:"), 0, 0)
        data_layout.addWidget(self.file1_edit, 0, 1)
        data_layout.addWidget(browse_btn1, 0, 2)
        data_layout.addWidget(self.sheet_combo1, 0, 3)
        data_layout.addWidget(QtWidgets.QLabel("新文件:"), 0, 4)
        data_layout.addWidget(self.file2_edit, 0, 5)
        data_layout.addWidget(browse_btn2, 0, 6)
        data_layout.addWidget(self.sheet_combo2, 0, 7)
        data_layout.setColumnStretch(1, 2)
        data_layout.setColumnStretch(5, 2)
        data_layout.setColumnStretch(3, 1)
        data_layout.setColumnStretch(7, 1)
        
        config_layout.addWidget(data_group)
        
        # 选项行
        options_layout = QtWidgets.QHBoxLayout()
        options_layout.setSpacing(15)

        # 比对模式与主键
        key_layout = QtWidgets.QHBoxLayout()
        self.key_mode_cb = QtWidgets.QCheckBox("主键模式")
        self.key_btn = QtWidgets.QPushButton("选择关键列...")
        self.key_label = QtWidgets.QLabel("未选")
        self.key_label.setStyleSheet("color: gray;")
        self.key_btn.setEnabled(False)
        self.key_mode_cb.toggled.connect(self.toggle_key_mode)
        self.key_btn.clicked.connect(self.select_key_columns)
        
        # 表头行设置
        header_row_layout = QtWidgets.QHBoxLayout()
        self.no_header_cb = QtWidgets.QCheckBox("无表头")
        self.header_row_spin = QtWidgets.QSpinBox()
        self.header_row_spin.setRange(1, 100)
        self.header_row_spin.setValue(1)
        self.header_row_spin.setFixedWidth(50)
        self.no_header_cb.toggled.connect(lambda checked: self.header_row_spin.setEnabled(not checked))
        header_row_layout.addWidget(self.no_header_cb)
        header_row_layout.addWidget(QtWidgets.QLabel("表头行:"))
        header_row_layout.addWidget(self.header_row_spin)
        
        key_layout.addWidget(self.key_mode_cb)
        key_layout.addWidget(self.key_btn)
        key_layout.addWidget(self.key_label)
        key_layout.addStretch()
        key_layout.addLayout(header_row_layout)
        options_layout.addLayout(key_layout)

        # 搜索过滤
        filter_layout = QtWidgets.QHBoxLayout()
        self.search_edit = QtWidgets.QLineEdit()
        self.search_edit.setPlaceholderText("搜索内容...")
        self.search_edit.setFixedWidth(150)
        self.filter_combo = QtWidgets.QComboBox()
        self.filter_combo.addItems(["全部状态", "只看修改", "只看新增", "只看删除", "一致项"])
        self.search_edit.textChanged.connect(self.apply_filter)
        self.filter_combo.currentIndexChanged.connect(self.apply_filter)
        filter_layout.addWidget(QtWidgets.QLabel("搜索:"))
        filter_layout.addWidget(self.search_edit)
        filter_layout.addWidget(self.filter_combo)
        options_layout.addLayout(filter_layout)

        # 视图控制
        view_layout = QtWidgets.QHBoxLayout()
        self.sync_v_cb = QtWidgets.QCheckBox("纵滚同步")
        self.sync_h_cb = QtWidgets.QCheckBox("横滚同步")
        
        # 冻结设置
        self.freeze_row_spin = QtWidgets.QSpinBox()
        self.freeze_col_spin = QtWidgets.QSpinBox()
        self.freeze_row_spin.setRange(0, 20)
        self.freeze_col_spin.setRange(0, 20)
        self.freeze_row_spin.setFixedWidth(45)
        self.freeze_col_spin.setFixedWidth(45)
        self.freeze_row_spin.valueChanged.connect(self.update_freeze)
        self.freeze_col_spin.valueChanged.connect(self.update_freeze)
        
        self.sync_v_cb.setChecked(True)
        self.sync_h_cb.setChecked(True)
        self.detail_toggle = QtWidgets.QCheckBox("差异明细")
        
        self.detail_toggle.toggled.connect(self.toggle_detail_panel)
        
        view_layout.addWidget(self.sync_v_cb)
        view_layout.addWidget(self.sync_h_cb)
        view_layout.addWidget(QtWidgets.QLabel("冻结行:"))
        view_layout.addWidget(self.freeze_row_spin)
        view_layout.addWidget(QtWidgets.QLabel("冻结列:"))
        view_layout.addWidget(self.freeze_col_spin)
        view_layout.addWidget(self.detail_toggle)
        options_layout.addLayout(view_layout)
        
        options_layout.addStretch()
        
        # 操作按钮
        self.compare_btn = QtWidgets.QPushButton("开始比对")
        self.full_compare_btn = QtWidgets.QPushButton("全表比对")
        self.export_btn = QtWidgets.QPushButton("导出结果")
        self.compare_btn.setMinimumWidth(100)
        self.full_compare_btn.setMinimumWidth(100)
        self.compare_btn.setStyleSheet("background-color: #0078d7; color: white; font-weight: bold; padding: 5px;")
        self.full_compare_btn.setStyleSheet("background-color: #28a745; color: white; font-weight: bold; padding: 5px;")
        self.compare_btn.clicked.connect(self.compare_files)
        self.full_compare_btn.clicked.connect(self.compare_all_sheets)
        self.export_btn.clicked.connect(self.export_results)
        options_layout.addWidget(self.compare_btn)
        options_layout.addWidget(self.full_compare_btn)
        options_layout.addWidget(self.export_btn)
        
        config_layout.addLayout(options_layout)
        main_layout.addWidget(config_panel)

        # 2. 中间主体区域 (Tab 切换)
        self.tab_widget = QtWidgets.QTabWidget()
        
        # Tab 1: 双窗同步比对
        self.compare_tab = QtWidgets.QWidget()
        compare_tab_layout = QtWidgets.QVBoxLayout(self.compare_tab)
        compare_tab_layout.setContentsMargins(0, 0, 0, 0)
        
        self.splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        self.left_table = FrozenTableView()
        self.right_table = FrozenTableView()
        self.diff_table = FrozenTableView()
        
        for t in [self.left_table, self.right_table, self.diff_table]:
            t.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectItems)
            t.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
            t.setAlternatingRowColors(True)
            t.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Interactive)
            t.verticalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Interactive)
            t.horizontalHeader().setStretchLastSection(True)
            t.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
            t.customContextMenuRequested.connect(lambda pos, table=t: self.show_table_context_menu(pos, table))
            
        self.splitter.addWidget(self.left_table)
        self.splitter.addWidget(self.right_table)
        
        # 差异明细面板
        self.detail_container = QtWidgets.QGroupBox("当前行差异明细")
        detail_layout = QtWidgets.QVBoxLayout(self.detail_container)
        self.detail_text = QtWidgets.QTextEdit()
        self.detail_text.setReadOnly(True)
        detail_layout.addWidget(self.detail_text)
        self.splitter.addWidget(self.detail_container)
        self.detail_container.hide()
        
        compare_tab_layout.addWidget(self.splitter)
        self.tab_widget.addTab(self.compare_tab, "双窗同步比对")
        
        # Tab 2: 差异项列表 (只显示有差异的行)
        self.diff_tab = QtWidgets.QWidget()
        diff_tab_layout = QtWidgets.QVBoxLayout(self.diff_tab)
        diff_tab_layout.addWidget(self.diff_table)
        self.tab_widget.addTab(self.diff_tab, "差异项总览")
        
        main_layout.addWidget(self.tab_widget)

        # 3. 底部状态与图例
        bottom_layout = QtWidgets.QHBoxLayout()
        legend_layout = QtWidgets.QHBoxLayout()
        legend_layout.addWidget(self.create_legend_label("新增", "#CCFFCC"))
        legend_layout.addWidget(self.create_legend_label("删除", "#FFFFCC"))
        legend_layout.addWidget(self.create_legend_label("修改", "#E6F3FF"))
        legend_layout.addWidget(self.create_legend_label("单元格差异", "#FFCCCC", "#CC0000"))
        bottom_layout.addLayout(legend_layout)
        
        bottom_layout.addStretch()
        
        # 缩放控制
        scale_layout = QtWidgets.QHBoxLayout()
        self.row_height_spin = QtWidgets.QSpinBox()
        self.row_height_spin.setRange(20, 150)
        self.row_height_spin.setValue(30)
        self.col_width_spin = QtWidgets.QDoubleSpinBox()
        self.col_width_spin.setRange(0.5, 3.0)
        self.col_width_spin.setSingleStep(0.1)
        self.col_width_spin.setValue(1.0)
        self.row_height_spin.valueChanged.connect(self.apply_row_height)
        self.col_width_spin.valueChanged.connect(self.apply_column_widths)
        scale_layout.addWidget(QtWidgets.QLabel("行高:"))
        scale_layout.addWidget(self.row_height_spin)
        scale_layout.addWidget(QtWidgets.QLabel("列宽系数:"))
        scale_layout.addWidget(self.col_width_spin)
        bottom_layout.addLayout(scale_layout)
        
        main_layout.addLayout(bottom_layout)

        # 滚动同步绑定
        self.left_table.verticalScrollBar().valueChanged.connect(
            lambda v: self.sync_scroll(self.left_table, self.right_table, v, "v")
        )
        self.right_table.verticalScrollBar().valueChanged.connect(
            lambda v: self.sync_scroll(self.right_table, self.left_table, v, "v")
        )
        self.left_table.horizontalScrollBar().valueChanged.connect(
            lambda v: self.sync_scroll(self.left_table, self.right_table, v, "h")
        )
        self.right_table.horizontalScrollBar().valueChanged.connect(
            lambda v: self.sync_scroll(self.right_table, self.left_table, v, "h")
        )

    def set_window_icon(self):
        icon_path = self.get_resource_path("excel.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QtGui.QIcon(icon_path))

    def create_legend_label(self, text, bg_color, fg_color="#000000"):
        label = QtWidgets.QLabel(text)
        label.setAlignment(QtCore.Qt.AlignCenter)
        label.setStyleSheet(f"background:{bg_color}; color:{fg_color}; padding:2px 6px; border:1px solid #cccccc;")
        return label

    def get_resource_path(self, relative_path):
        base_path = os.path.dirname(os.path.abspath(__file__))
        project_root = os.path.dirname(os.path.dirname(base_path))
        return os.path.join(project_root, "assets", relative_path)

    def toggle_key_mode(self, checked):
        self.key_btn.setEnabled(checked)
        if not checked:
            self.key_columns = []
            self.key_label.setText("未选择关键列")

    def select_key_columns(self):
        file1 = self.file1_edit.text().strip()
        sheet1 = self.sheet_combo1.currentText()
        if not file1 or not sheet1:
            QtWidgets.QMessageBox.warning(self, "提示", "请先选择文件和 Sheet 以获取列信息")
            return
        try:
            header_row = self.header_row_spin.value()
            has_header = not self.no_header_cb.isChecked()
            differ = DifferFactory.get_differ(file1)
            df = differ.read_data(file1, sheet1, handle_merged=False, header_row=header_row, has_header=has_header)
            columns = df.columns.tolist()
            dialog = KeyColumnsDialog(columns, self.key_columns, self)
            if dialog.exec_() == QtWidgets.QDialog.Accepted:
                selected = dialog.selected_columns()
                if not selected:
                    self.key_mode_cb.setChecked(False)
                    self.toggle_key_mode(False)
                else:
                    self.key_columns = selected
                    label_text = "已选: " + ", ".join(selected[:2])
                    if len(selected) > 2:
                        label_text += "..."
                    self.key_label.setText(label_text)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "错误", f"获取列信息失败: {e}")

    def browse_file(self, index):
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "选择文件", "", "Data files (*.xlsx *.xls *.csv)")
        if not file_path:
            return
        if index == 1:
            self.file1_edit.setText(file_path)
            self.load_sheets(file_path, self.sheet_combo1)
        else:
            self.file2_edit.setText(file_path)
            self.load_sheets(file_path, self.sheet_combo2)

    def load_sheets(self, filepath, combo):
        try:
            differ = DifferFactory.get_differ(filepath)
            sheets = differ.get_containers(filepath)
            combo.clear()
            combo.addItems(sheets)
            if sheets:
                combo.setCurrentIndex(0)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "错误", str(e))

    def compare_files(self):
        file1 = self.file1_edit.text().strip()
        file2 = self.file2_edit.text().strip()
        sheet1 = self.sheet_combo1.currentText()
        sheet2 = self.sheet_combo2.currentText()
        if not file1 or not file2:
            QtWidgets.QMessageBox.warning(self, "提示", "请先选择两个 Excel 文件")
            return
        if not sheet1 or not sheet2:
            QtWidgets.QMessageBox.warning(self, "提示", "请选择要比对的 Sheet")
            return
        key_cols = self.key_columns if self.key_mode_cb.isChecked() else None
        if self.key_mode_cb.isChecked() and not key_cols:
            QtWidgets.QMessageBox.warning(self, "提示", "已开启主键比对模式，但未选择关键列")
            return

        self.compare_btn.setEnabled(False)
        self.full_compare_btn.setEnabled(False)
        self.progress_dialog = QtWidgets.QProgressDialog("正在处理中...", "", 0, 100, self)
        self.progress_dialog.setWindowTitle("请稍候")
        self.progress_dialog.setWindowModality(QtCore.Qt.ApplicationModal)
        self.progress_dialog.setAutoClose(False)
        self.progress_dialog.setAutoReset(False)
        self.progress_dialog.show()

        header_row = self.header_row_spin.value()
        has_header = not self.no_header_cb.isChecked()
        self.worker_thread = QtCore.QThread()
        self.worker = CompareWorker(file1, file2, sheet1, sheet2, key_cols, header_row=header_row, has_header=has_header)
        self.worker.moveToThread(self.worker_thread)
        self.worker_thread.started.connect(self.worker.run)
        self.worker.progress.connect(self.update_progress)
        self.worker.finished.connect(self.on_compare_finished)
        self.worker.error.connect(self.on_compare_error)
        self.worker.finished.connect(self.worker_thread.quit)
        self.worker.error.connect(self.worker_thread.quit)
        self.worker_thread.finished.connect(self.worker.deleteLater)
        self.worker_thread.finished.connect(self.worker_thread.deleteLater)
        self.worker_thread.start()

    def compare_all_sheets(self):
        file1 = self.file1_edit.text().strip()
        file2 = self.file2_edit.text().strip()
        if not file1 or not file2:
            QtWidgets.QMessageBox.warning(self, "提示", "请先选择两个 Excel 文件")
            return
        
        key_cols = self.key_columns if self.key_mode_cb.isChecked() else None
        # 全表比对时，如果开启了主键模式但没有选择列，提示用户
        if self.key_mode_cb.isChecked() and not key_cols:
            QtWidgets.QMessageBox.warning(self, "提示", "已开启主键比对模式，但未选择关键列。\n全表比对将尝试对所有 Sheet 应用相同的关键列，如果某些 Sheet 不包含这些列，可能会报错。")
            # 允许继续，因为不同 Sheet 列名可能不同

        self.compare_btn.setEnabled(False)
        self.full_compare_btn.setEnabled(False)
        self.progress_dialog = QtWidgets.QProgressDialog("正在开始全表比对...", "", 0, 100, self)
        self.progress_dialog.setWindowTitle("全表比对中")
        self.progress_dialog.setWindowModality(QtCore.Qt.ApplicationModal)
        self.progress_dialog.setAutoClose(False)
        self.progress_dialog.setAutoReset(False)
        self.progress_dialog.show()

        header_row = self.header_row_spin.value()
        has_header = not self.no_header_cb.isChecked()
        self.worker_thread = QtCore.QThread()
        self.worker = FullCompareWorker(file1, file2, key_cols, header_row=header_row, has_header=has_header)
        self.worker.moveToThread(self.worker_thread)
        self.worker_thread.started.connect(self.worker.run)
        self.worker.progress.connect(self.update_progress)
        self.worker.finished.connect(self.on_full_compare_finished)
        self.worker.error.connect(self.on_full_compare_error)
        self.worker.finished.connect(self.worker_thread.quit)
        self.worker.error.connect(self.worker_thread.quit)
        self.worker_thread.finished.connect(self.worker.deleteLater)
        self.worker_thread.finished.connect(self.worker_thread.deleteLater)
        self.worker_thread.start()

    def on_full_compare_finished(self, summary):
        if self.progress_dialog:
            self.progress_dialog.close()
        self.compare_btn.setEnabled(True)
        self.full_compare_btn.setEnabled(True)
        
        dialog = FullCompareSummaryDialog(summary, self)
        dialog.exec_()

    def on_full_compare_error(self, message):
        if self.progress_dialog:
            self.progress_dialog.close()
        self.compare_btn.setEnabled(True)
        self.full_compare_btn.setEnabled(True)
        QtWidgets.QMessageBox.critical(self, "全表比对错误", f"比对过程中发生错误: {message}")

    def update_progress(self, value, text):
        if self.progress_dialog:
            self.progress_dialog.setValue(value)
            self.progress_dialog.setLabelText(text)

    def on_compare_finished(self, columns, results):
        if self.progress_dialog:
            self.progress_dialog.close()
        self.compare_btn.setEnabled(True)
        self.search_edit.setText("")
        self.filter_combo.setCurrentIndex(0)
        self.update_tables_data(columns, results, update_cache=True)
        
        # 统计比对结果
        stats = {
            'added': 0,
            'deleted': 0,
            'modified_rows': 0,
            'modified_cells': 0,
            'equal': 0
        }
        
        for r in results:
            info = r[2]
            status = info.get('status', 'equal')
            if status == 'added':
                stats['added'] += 1
            elif status == 'deleted':
                stats['deleted'] += 1
            elif status == 'modified':
                stats['modified_rows'] += 1
                stats['modified_cells'] += len(info.get('diff_cols', []))
            elif status == 'equal':
                stats['equal'] += 1
        
        # 弹出统计仪表盘
        dialog = StatisticsDialog(stats, self)
        dialog.exec_()

    def on_compare_error(self, message):
        if self.progress_dialog:
            self.progress_dialog.close()
        self.compare_btn.setEnabled(True)
        QtWidgets.QMessageBox.critical(self, "错误", f"比对过程中发生错误: {message}")

    def update_tables_data(self, columns, results, update_cache):
        if update_cache:
            self.last_columns = columns
            self.all_results = results
        self.display_results = results

        row_count = len(results)
        col_count = len(columns)
        
        # 1. 提取数据和差异信息
        left_data_list = [r[0] for r in results]
        right_data_list = [r[1] for r in results]
        diff_infos = [r[2] for r in results]

        # 2. 创建并设置自定义模型 (Tab 1)
        left_model = DiffTableModel(left_data_list, columns, diff_infos, role_type="left", parent=self)
        right_model = DiffTableModel(right_data_list, columns, diff_infos, role_type="right", parent=self)
        
        self.left_table.setModel(left_model)
        self.right_table.setModel(right_model)
        self.left_model = left_model
        self.right_model = right_model

        # 3. 准备差异汇总表 (Tab 2) - 换成更直观的“变更列表”模式
        diff_only_results = [r for r in results if r[2]['status'] != "equal"]
        
        # 差异汇总表列：[主键/位置, 状态, 修改项, 旧值, 新值]
        diff_summary_columns = ["主键/位置", "状态", "修改项", "旧值 (旧文件)", "新值 (新文件)"]
        
        # 获取主键列的索引
        key_indices = []
        if self.key_columns:
            for kc in self.key_columns:
                if kc in columns:
                    key_indices.append(columns.index(kc))
        
        diff_summary_data = []
        diff_summary_infos = []
        
        for left_vals, right_vals, info in diff_only_results:
            status = info['status']
            status_text = "修改" if status == 'modified' else ("新增" if status == 'added' else "删除")
            
            # 确定行标识信息 (Key Info)
            if key_indices:
                # 优先从新值中获取主键（处理新增），新值没有则从旧值获取（处理删除）
                # 注意：excel_processor 保证了主键一致时 left_vals/right_vals 的主键位相同
                base_vals = right_vals if status != 'deleted' else left_vals
                key_vals = [str(base_vals[i]) for i in key_indices if i < len(base_vals)]
                key_info = " | ".join(key_vals)
            else:
                # 无主键模式，使用数据在结果集中的序号（大致对应行号）
                try:
                    orig_idx = results.index((left_vals, right_vals, info))
                    key_info = f"第 {orig_idx + 1} 行"
                except:
                    key_info = "未知行"

            if status == 'modified':
                diff_cols = info.get('diff_cols', [])
                for c_idx in diff_cols:
                    col_name = columns[c_idx]
                    l_val = left_vals[c_idx]
                    r_val = right_vals[c_idx]
                    
                    # 生成微观差异 HTML
                    diff_html = ExcelDiffer.get_text_diff_html(l_val, r_val)
                    
                    # 差异汇总表列：[主键/位置, 状态, 修改项, 旧值, 新值]
                    # 我们在“新值”列显示微观差异，或者保持原样但在“新值”里高亮？
                    # 用户的意思是高亮。我们直接把 diff_html 放在新值列。
                    diff_summary_data.append([key_info, status_text, col_name, str(l_val), diff_html])
                    
                    # 高亮“旧值”和“新值”列 (索引 3 和 4)
                    diff_summary_infos.append({
                        'status': 'modified', 
                        'diff_cols': [3, 4]
                    })
            elif status == 'added':
                # 新增行：展示一个概览，提示去主视图查看完整行
                diff_summary_data.append([key_info, status_text, "整行", "", "(查看主视图详情)"])
                diff_summary_infos.append({'status': 'added', 'diff_cols': []})
            elif status == 'deleted':
                # 删除行
                diff_summary_data.append([key_info, status_text, "整行", "(查看主视图详情)", ""])
                diff_summary_infos.append({'status': 'deleted', 'diff_cols': []})

        diff_model = DiffTableModel(diff_summary_data, diff_summary_columns, diff_summary_infos, role_type="diff", parent=self)
        self.diff_table.setModel(diff_model)
        
        # 为“新值”列设置 HTML 委派
        self.diff_table.setItemDelegateForColumn(4, HTMLDelegate(self.diff_table))
        
        # 4. 界面调整
        header = self.diff_table.horizontalHeader()
        header.setSectionResizeMode(QtWidgets.QHeaderView.Interactive)  # 允许手动调整
        self.diff_table.resizeColumnsToContents()
        
        # 限制某些列的最大宽度，防止被撑开得太大导致“显示不全”
        max_col_width = 350
        for i in range(header.count()):
            if self.diff_table.columnWidth(i) > max_col_width:
                self.diff_table.setColumnWidth(i, max_col_width)
        
        # 针对“旧值”和“新值”列给予足够的最小宽度，并让最后一列自动填充
        if header.count() >= 5:
            # 索引 3 和 4 分别是旧值和新值
            self.diff_table.setColumnWidth(3, max(self.diff_table.columnWidth(3), 150))
            self.diff_table.setColumnWidth(4, max(self.diff_table.columnWidth(4), 150))
        
        header.setStretchLastSection(True)
        
        self.apply_row_height(self.row_height_spin.value())
        self.compute_base_widths(columns, results)
        self.apply_column_widths(self.col_width_spin.value())
        self.connect_selection()
        self.clear_detail_panel()

    def compute_base_widths(self, columns, results):
        if not columns:
            self.base_col_widths = []
            return
        metrics = QtGui.QFontMetrics(self.left_table.font())
        max_widths = [metrics.horizontalAdvance(str(col)) for col in columns]
        sample_size = min(len(results), 300)
        for i in range(sample_size):
            left_vals, right_vals, _ = results[i]
            for col_idx in range(len(columns)):
                l_text = str(left_vals[col_idx] if col_idx < len(left_vals) else "")
                r_text = str(right_vals[col_idx] if col_idx < len(right_vals) else "")
                max_widths[col_idx] = max(max_widths[col_idx], metrics.horizontalAdvance(l_text))
                max_widths[col_idx] = max(max_widths[col_idx], metrics.horizontalAdvance(r_text))
        self.base_col_widths = [min(w + 30, 600) for w in max_widths]

    def apply_row_height(self, value):
        row_height = int(value)
        for t in [self.left_table, self.right_table, self.diff_table]:
            t.verticalHeader().setDefaultSectionSize(row_height)
            # 强制刷新所有行高（处理已经手动调整过的行）
            for i in range(t.model().rowCount() if t.model() else 0):
                t.setRowHeight(i, row_height)
            if hasattr(t, 'update_frozen_geometry'):
                t.update_frozen_geometry()

    def apply_column_widths(self, value):
        if not self.base_col_widths:
            return
        factor = float(value)
        for i, base_w in enumerate(self.base_col_widths):
            width = int(base_w * factor)
            self.left_table.setColumnWidth(i, width)
            self.right_table.setColumnWidth(i, width)
            self.diff_table.setColumnWidth(i, width)
        
        for t in [self.left_table, self.right_table, self.diff_table]:
            if hasattr(t, 'update_frozen_geometry'):
                t.update_frozen_geometry()

    def show_table_context_menu(self, pos, table):
        menu = QtWidgets.QMenu()
        auto_fit_col_action = menu.addAction("自动调整此列宽度")
        auto_fit_all_cols_action = menu.addAction("自动调整所有列宽")
        menu.addSeparator()
        reset_row_height_action = menu.addAction("重置行高")
        
        action = menu.exec_(table.viewport().mapToGlobal(pos))
        if action == auto_fit_col_action:
            index = table.indexAt(pos)
            if index.isValid():
                table.resizeColumnToContents(index.column())
        elif action == auto_fit_all_cols_action:
            table.resizeColumnsToContents()
        elif action == reset_row_height_action:
            self.apply_row_height(self.row_height_spin.value())

    def connect_selection(self):
        if self.left_table.selectionModel():
            self.left_table.selectionModel().selectionChanged.connect(self.on_left_selection)
        if self.right_table.selectionModel():
            self.right_table.selectionModel().selectionChanged.connect(self.on_right_selection)

    def on_left_selection(self, selected, deselected):
        if self._syncing_selection:
            return
        indexes = self.left_table.selectionModel().selectedIndexes()
        if not indexes:
            return
        index = indexes[0]
        self._syncing_selection = True
        if self.right_model:
            self.right_table.setCurrentIndex(self.right_model.index(index.row(), index.column()))
        self._syncing_selection = False
        self.update_detail_panel(index.row())

    def on_right_selection(self, selected, deselected):
        if self._syncing_selection:
            return
        indexes = self.right_table.selectionModel().selectedIndexes()
        if not indexes:
            return
        index = indexes[0]
        self._syncing_selection = True
        if self.left_model:
            self.left_table.setCurrentIndex(self.left_model.index(index.row(), index.column()))
        self._syncing_selection = False
        self.update_detail_panel(index.row())

    def sync_scroll(self, source, target, value, axis):
        if self._syncing_scroll:
            return
        if axis == "v" and not self.sync_v_cb.isChecked():
            return
        if axis == "h" and not self.sync_h_cb.isChecked():
            return
        self._syncing_scroll = True
        if axis == "v":
            target.verticalScrollBar().setValue(value)
        else:
            target.horizontalScrollBar().setValue(value)
        self._syncing_scroll = False

    def apply_filter(self):
        if not self.all_results or not self.last_columns:
            return
        search_kw = self.search_edit.text().strip().lower()
        filter_status = self.filter_combo.currentText()
        filtered = []
        for left_vals, right_vals, info in self.all_results:
            status = left_vals[0] if left_vals else ""
            if filter_status != "全部状态" and status != filter_status.replace("只看", "").replace("项", ""):
                # 处理 "全部状态" 和 "只看修改" 等文本匹配
                current_filter = filter_status.replace("只看", "").replace("项", "")
                if current_filter != "全部状态" and status != current_filter:
                    continue
            
            if search_kw:
                found = False
                for val in left_vals:
                    if search_kw in str(val).lower():
                        found = True
                        break
                if not found:
                    for val in right_vals:
                        if search_kw in str(val).lower():
                            found = True
                            break
                if not found:
                    continue
            filtered.append((left_vals, right_vals, info))
        self.update_tables_data(self.last_columns, filtered, update_cache=False)

    def update_freeze(self):
        if not hasattr(self, 'left_table') or not hasattr(self, 'right_table') or not hasattr(self, 'diff_table'):
            return
        rows = self.freeze_row_spin.value()
        cols = self.freeze_col_spin.value()
        for t in [self.left_table, self.right_table, self.diff_table]:
            t.set_frozen(rows, cols)

    def toggle_detail_panel(self, checked):
        if checked:
            self.detail_container.show()
        else:
            self.detail_container.hide()
            self.clear_detail_panel()

    def clear_detail_panel(self):
        self.detail_text.clear()

    def update_detail_panel(self, row_idx):
        if not self.display_results or row_idx >= len(self.display_results):
            return
        left_data, right_data, info = self.display_results[row_idx]
        columns = self.last_columns or []
        status = left_data[0] if left_data else ""
        diff_cols = info.get('diff_cols', [])
        
        html_parts = []
        html_parts.append(f"<b>行索引:</b> {row_idx + 1} | <b>状态:</b> {status}<br>")
        html_parts.append("<hr>")
        
        has_diff = False
        for i in range(1, len(columns)):
            col_name = columns[i]
            val1 = left_data[i] if i < len(left_data) else ""
            val2 = right_data[i] if i < len(right_data) else ""
            
            is_cell_diff = i in diff_cols
            is_row_diff = status in ["新增", "删除"]
            
            if not (is_cell_diff or is_row_diff):
                continue
                
            has_diff = True
            html_parts.append(f"<b>列:</b> {col_name}<br>")
            if status == "删除":
                html_parts.append(f"<span style='color:#CC0000'>旧值:</span> <span style='background:#FFFFCC'>{val1}</span><br><br>")
            elif status == "新增":
                html_parts.append(f"<span style='color:#CC0000'>新值:</span> <span style='background:#CCFFCC'>{val2}</span><br><br>")
            else:
                # 单元格微观差异高亮
                diff_html = ExcelDiffer.get_text_diff_html(val1, val2)
                html_parts.append(f"<span style='color:#CC0000'>旧值:</span> <span style='background:#FFCCCC'>{val1}</span><br>")
                html_parts.append(f"<span style='color:#CC0000'>对比:</span> <span style='background:#FFCCCC'>{diff_html}</span><br><br>")
        
        if not has_diff:
            html_parts.append("<i>(该行所有列内容一致)</i>")
        
        self.detail_text.setHtml("".join(html_parts))

    def export_results(self):
        if not self.last_columns or self.all_results is None:
            QtWidgets.QMessageBox.warning(self, "提示", "没有可导出的比对结果")
            return
        output_path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "导出比对结果", "比对结果.xlsx", "Excel files (*.xlsx);;CSV files (*.csv)")
        if not output_path:
            return
        try:
            key_cols = self.key_columns if self.key_mode_cb.isChecked() else None
            differ = DifferFactory.get_differ(output_path if output_path.endswith('.csv') else self.file1_edit.text())
            differ.export(output_path, self.last_columns, self.all_results, key_columns=key_cols)
            QtWidgets.QMessageBox.information(self, "成功", f"结果已成功导出至:\n{output_path}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "错误", f"导出失败: {e}")

    def show_usage(self):
        usage_text = (
            "Excel 差异比对工具 使用说明：\n\n"
            "1. 选择文件：点击“浏览...”按钮选择需要比对的两个 Excel 文件（旧版本和新版本）。\n"
            "2. 选择 Sheet：在下拉框中选择要比对的工作表。\n"
            "3. 设置选项：\n"
            "   - 纵滚/横滚同步：开启后两侧表格将同步滚动。\n"
            "4. 开始比对：点击“开始比对差异”按钮，程序将分析差异并在下方展示。\n"
            "5. 结果说明：\n"
            "   - 单元格浅红填充：表示内容发生了变化。\n"
            "   - 行浅绿填充：表示该行为新增行。\n"
            "   - 行浅黄填充：表示该行为删除行。\n"
            "6. 导出结果：比对完成后，点击“导出结果”可将差异保存为 Excel 文件。\n\n"
            "注意：程序默认开启合并单元格自动填充功能，确保比对结果的准确性。"
        )
        QtWidgets.QMessageBox.information(self, "使用说明", usage_text)

    def show_version(self):
        version = self.version
        details = ""
        if os.path.exists("updates_notes.txt"):
            with open("updates_notes.txt", "r", encoding="utf-8") as f:
                lines = f.read().splitlines()
            details = "\n".join(lines[:8]).strip()
        if not details:
            details = "未找到更新记录。"
        text = f"Excel 差异比对工具\n\n当前版本：{version}\n\n更新说明：\n{details}"
        QtWidgets.QMessageBox.information(self, "版本说明", text)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        file_paths = [u.toLocalFile() for u in urls]
        for path in file_paths:
            if path.lower().endswith((".xlsx", ".xls")):
                if not self.file1_edit.text().strip():
                    self.file1_edit.setText(path)
                    self.load_sheets(path, self.sheet_combo1)
                else:
                    self.file2_edit.setText(path)
                    self.load_sheets(path, self.sheet_combo2)
