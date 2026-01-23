from PyQt5 import QtWidgets, QtCore, QtGui

class FrozenTableView(QtWidgets.QTableView):
    """
    支持冻结指定行数和列数的 TableView。
    通过在主表格上方叠加三个从属表格来实现：
    1. left_view: 冻结列区域
    2. top_view: 冻结行区域
    3. corner_view: 冻结行与列的交集区域
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        
        self._frozen_rows = 0
        self._frozen_cols = 0
        
        # 创建辅助视图
        self.left_view = QtWidgets.QTableView(self)
        self.top_view = QtWidgets.QTableView(self)
        self.corner_view = QtWidgets.QTableView(self)
        
        self.init_aux_views()
        
        # 连接信号：同步垂直滚动
        self.verticalScrollBar().valueChanged.connect(self._sync_v_scroll)
        # 连接信号：同步水平滚动
        self.horizontalScrollBar().valueChanged.connect(self._sync_h_scroll)
        
        # 连接信号：同步表头缩放
        self.verticalHeader().sectionResized.connect(self._on_v_section_resized)
        self.horizontalHeader().sectionResized.connect(self._on_h_section_resized)
        
    def init_aux_views(self):
        for view in [self.left_view, self.top_view, self.corner_view]:
            view.setFocusPolicy(QtCore.Qt.NoFocus)
            view.verticalHeader().hide()
            view.horizontalHeader().hide()
            view.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
            view.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
            view.setStyleSheet("QTableView { border: none; background-color: white; selection-background-color: #0078d7; }")
            # view.setSelectionModel(self.selectionModel()) # 在 setModel 中同步
            view.setAlternatingRowColors(True)
            view.hide()
            self.viewport().stackUnder(view)
            # 安装事件过滤器，以便统一处理鼠标滚轮等事件
            view.installEventFilter(self)

    def eventFilter(self, obj, event):
        if event.type() == QtCore.QEvent.Wheel:
            if obj in [self.left_view, self.top_view, self.corner_view]:
                # 将子视图的滚轮事件转发给主视图处理
                self.wheelEvent(event)
                return True
        return super().eventFilter(obj, event)

    def wheelEvent(self, event):
        if event.modifiers() & QtCore.Qt.ShiftModifier:
            # Shift + 滚轮 -> 横向滚动
            delta = event.angleDelta().y()
            h_bar = self.horizontalScrollBar()
            # 滚动步长，可以根据需要调整，这里直接使用 delta
            h_bar.setValue(h_bar.value() - delta)
            event.accept()
        else:
            super().wheelEvent(event)

    def set_frozen(self, rows: int, cols: int):
        """设置冻结的行数和列数"""
        self._frozen_rows = max(0, rows)
        self._frozen_cols = max(0, cols)
        
        if self._frozen_rows > 0 or self._frozen_cols > 0:
            self._update_all_frozen_views()
        else:
            self.left_view.hide()
            self.top_view.hide()
            self.corner_view.hide()

    def setModel(self, model):
        super().setModel(model)
        self.left_view.setModel(model)
        self.top_view.setModel(model)
        self.corner_view.setModel(model)
        
        # 同步选择模型
        if self.selectionModel():
            for view in [self.left_view, self.top_view, self.corner_view]:
                view.setSelectionModel(self.selectionModel())
                
        if self._frozen_rows > 0 or self._frozen_cols > 0:
            self._update_all_frozen_views()

    def _sync_v_scroll(self, value):
        self.left_view.verticalScrollBar().setValue(value)

    def _sync_h_scroll(self, value):
        self.top_view.horizontalScrollBar().setValue(value)

    def _on_v_section_resized(self, logicalIndex, oldSize, newSize):
        # 同步行高
        self.left_view.setRowHeight(logicalIndex, newSize)
        self.corner_view.setRowHeight(logicalIndex, newSize)
        if logicalIndex < self._frozen_rows:
            self.top_view.setRowHeight(logicalIndex, newSize)
        self.update_frozen_geometry()

    def _on_h_section_resized(self, logicalIndex, oldSize, newSize):
        # 同步列宽
        self.top_view.setColumnWidth(logicalIndex, newSize)
        self.corner_view.setColumnWidth(logicalIndex, newSize)
        if logicalIndex < self._frozen_cols:
            self.left_view.setColumnWidth(logicalIndex, newSize)
        self.update_frozen_geometry()

    def _update_all_frozen_views(self):
        if not self.model():
            return
            
        # 同步基础配置
        for view in [self.left_view, self.top_view, self.corner_view]:
            view.verticalHeader().setDefaultSectionSize(self.verticalHeader().defaultSectionSize())
            view.horizontalHeader().setDefaultSectionSize(self.horizontalHeader().defaultSectionSize())

        # 1. 配置 left_view (冻结列)
        if self._frozen_cols > 0:
            for col in range(self.model().columnCount()):
                self.left_view.setColumnHidden(col, col >= self._frozen_cols)
                if col < self._frozen_cols:
                    self.left_view.setColumnWidth(col, self.columnWidth(col))
            self.left_view.show()
        else:
            self.left_view.hide()

        # 2. 配置 top_view (冻结行)
        if self._frozen_rows > 0:
            for row in range(self.model().rowCount()):
                self.top_view.setRowHidden(row, row >= self._frozen_rows)
                if row < self._frozen_rows:
                    self.top_view.setRowHeight(row, self.rowHeight(row))
            self.top_view.show()
        else:
            self.top_view.hide()

        # 3. 配置 corner_view (交集)
        if self._frozen_rows > 0 and self._frozen_cols > 0:
            for col in range(self.model().columnCount()):
                self.corner_view.setColumnHidden(col, col >= self._frozen_cols)
            for row in range(self.model().rowCount()):
                self.corner_view.setRowHidden(row, row >= self._frozen_rows)
            self.corner_view.show()
        else:
            self.corner_view.hide()

        self.update_frozen_geometry()

    def update_frozen_geometry(self):
        if self._frozen_rows == 0 and self._frozen_cols == 0:
            return

        v_header_width = self.verticalHeader().width() if self.verticalHeader().isVisible() else 0
        h_header_height = self.horizontalHeader().height() if self.horizontalHeader().isVisible() else 0
        frame_width = self.frameWidth()
        
        # 计算冻结宽度和高度
        frozen_width = 0
        for col in range(self._frozen_cols):
            frozen_width += self.columnWidth(col)
            
        frozen_height = 0
        for row in range(self._frozen_rows):
            frozen_height += self.rowHeight(row)

        # left_view (冻结列)
        if self._frozen_cols > 0:
            self.left_view.setGeometry(
                v_header_width + frame_width,
                h_header_height + frame_width,
                frozen_width,
                self.viewport().height()
            )
            
        # top_view (冻结行)
        if self._frozen_rows > 0:
            self.top_view.setGeometry(
                v_header_width + frame_width,
                h_header_height + frame_width,
                self.viewport().width(),
                frozen_height
            )
            
        # corner_view (交集)
        if self._frozen_rows > 0 and self._frozen_cols > 0:
            self.corner_view.setGeometry(
                v_header_width + frame_width,
                h_header_height + frame_width,
                frozen_width,
                frozen_height
            )

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.update_frozen_geometry()

    def scrollTo(self, index, hint=QtWidgets.QAbstractItemView.EnsureVisible):
        # 如果目标在冻结区域内，不自动滚动主视图
        if index.row() < self._frozen_rows and index.column() < self._frozen_cols:
            return
        super().scrollTo(index, hint)
