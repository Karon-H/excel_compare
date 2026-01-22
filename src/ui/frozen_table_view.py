from PyQt5 import QtWidgets, QtCore, QtGui

class FrozenTableView(QtWidgets.QTableView):
    """
    支持冻结首列的 TableView。
    通过在主表格上方叠加一个只显示第一列的从属表格来实现。
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        
        # 创建冻结列的表格
        self.frozen_view = QtWidgets.QTableView(self)
        self._is_frozen = False
        
        self.init_frozen_view()
        
        # 连接信号：同步垂直滚动
        self.verticalScrollBar().valueChanged.connect(self.frozen_view.verticalScrollBar().setValue)
        self.frozen_view.verticalScrollBar().valueChanged.connect(self.verticalScrollBar().setValue)
        
        # 连接信号：同步垂直表头缩放
        self.verticalHeader().sectionResized.connect(self._on_v_section_resized)
        
        # 连接信号：同步表头缩放
        self.horizontalHeader().sectionResized.connect(self._on_section_resized)
        
    def init_frozen_view(self):
        self.frozen_view.setFocusPolicy(QtCore.Qt.NoFocus)
        self.frozen_view.verticalHeader().hide()
        self.frozen_view.horizontalScrollBar().hide()
        self.frozen_view.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.frozen_view.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        
        # 确保冻结层在最上层
        self.viewport().stackUnder(self.frozen_view)
        
        self.frozen_view.setStyleSheet("QTableView { border: none; background-color: white; selection-background-color: #0078d7; }")
        self.frozen_view.setSelectionModel(self.selectionModel())
        self.frozen_view.setAlternatingRowColors(True)
        self.frozen_view.hide() # 默认不冻结

    def set_frozen(self, frozen: bool):
        """启用或禁用冻结首列"""
        self._is_frozen = frozen
        if frozen:
            # 同步主表的行高设置到冻结表
            self.frozen_view.verticalHeader().setDefaultSectionSize(self.verticalHeader().defaultSectionSize())
            self.frozen_view.show()
            self.update_frozen_geometry()
            self._sync_frozen_columns()
        else:
            self.frozen_view.hide()

    def setModel(self, model):
        super().setModel(model)
        self.frozen_view.setModel(model)
        if self._is_frozen:
            self.frozen_view.verticalHeader().setDefaultSectionSize(self.verticalHeader().defaultSectionSize())
            self._sync_frozen_columns()

    def _sync_frozen_columns(self):
        """同步冻结视图的列状态：只显示第一列"""
        if not self.model():
            return
        
        # 在主视图中，第一列可以保持可见，也可以隐藏（如果冻结层覆盖了它）
        # 这里选择让冻结层完全覆盖主视图的第一列
        for col in range(self.model().columnCount()):
            if col == 0:
                self.frozen_view.setColumnHidden(col, False)
                self.frozen_view.setColumnWidth(col, self.columnWidth(col))
            else:
                self.frozen_view.setColumnHidden(col, True)

    def _on_v_section_resized(self, logicalIndex, oldSize, newSize):
        self.frozen_view.setRowHeight(logicalIndex, newSize)

    def _on_section_resized(self, logicalIndex, oldSize, newSize):
        if logicalIndex == 0:
            self.frozen_view.setColumnWidth(0, newSize)
            self.update_frozen_geometry()

    def update_frozen_geometry(self):
        if not self._is_frozen:
            return
        
        # 计算冻结区域的几何位置
        # 需要考虑表头高度和垂直表头宽度
        v_header_width = self.verticalHeader().width() if self.verticalHeader().isVisible() else 0
        h_header_height = self.horizontalHeader().height()
        
        col_width = self.columnWidth(0)
        
        # 冻结表格应该覆盖从主表格左侧开始的第一列
        self.frozen_view.setGeometry(
            v_header_width + self.frameWidth(),
            self.frameWidth(),
            col_width,
            self.viewport().height() + h_header_height
        )

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.update_frozen_geometry()

    def scrollTo(self, index, hint=QtWidgets.QAbstractItemView.EnsureVisible):
        # 如果滚动到第一列以外，正常执行
        if index.column() > 0:
            super().scrollTo(index, hint)
        # 如果是第一列，则不需要在主视图中水平滚动（因为它已被冻结层覆盖）
