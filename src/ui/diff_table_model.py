from PyQt5 import QtCore, QtGui

class DiffTableModel(QtCore.QAbstractTableModel):
    """
    支持差异高亮的自定义表格模型
    """
    def __init__(self, data=None, headers=None, diff_infos=None, role_type="left", parent=None):
        super().__init__(parent)
        self._data = data or []
        self._headers = headers or []
        self._diff_infos = diff_infos or []  # 存储每行的差异元数据
        self._role_type = role_type  # "left", "right" 或 "diff" (用于差异汇总表)

        # 定义颜色
        self.added_color = QtGui.QColor("#CCFFCC")      # 浅绿
        self.deleted_color = QtGui.QColor("#FFFFCC")    # 浅黄
        self.modified_color = QtGui.QColor("#E6F3FF")   # 浅蓝
        self.cell_diff_bg = QtGui.QColor("#FFCCCC")     # 浅红 (差异单元格背景)
        self.cell_diff_fg = QtGui.QColor("#CC0000")     # 深红 (差异单元格文字)

    def rowCount(self, parent=QtCore.QModelIndex()):
        return len(self._data)

    def columnCount(self, parent=QtCore.QModelIndex()):
        return len(self._headers)

    def data(self, index, role=QtCore.Qt.DisplayRole):
        if not index.isValid():
            return None

        row = index.row()
        col = index.column()

        if role == QtCore.Qt.DisplayRole:
            return str(self._data[row][col])

        if role == QtCore.Qt.BackgroundRole:
            diff_info = self._diff_infos[row]
            status = diff_info.get('status', 'equal')
            diff_cols = diff_info.get('diff_cols', [])

            # 1. 单元格级差异高亮 (最高优先级)
            if status == 'modified' and col in diff_cols:
                return self.cell_diff_bg
            
            # 2. 行级状态高亮
            if self._role_type == "left":
                if status == 'deleted':
                    return self.deleted_color
            elif self._role_type == "right":
                if status == 'added':
                    return self.added_color
            
            # 状态列 (第一列) 在修改时也高亮
            if col == 0:
                if status == 'modified':
                    return self.modified_color
                elif status == 'deleted':
                    return self.deleted_color
                elif status == 'added':
                    return self.added_color

        if role == QtCore.Qt.ForegroundRole:
            diff_info = self._diff_infos[row]
            status = diff_info.get('status', 'equal')
            diff_cols = diff_info.get('diff_cols', [])
            if status == 'modified' and col in diff_cols:
                return self.cell_diff_fg

        return None

    def headerData(self, section, orientation, role=QtCore.Qt.DisplayRole):
        if orientation == QtCore.Qt.Horizontal and role == QtCore.Qt.DisplayRole:
            if section < len(self._headers):
                return self._headers[section]
        return None
