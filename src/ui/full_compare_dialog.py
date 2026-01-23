from PyQt5 import QtWidgets, QtCore, QtGui

class FullCompareSummaryDialog(QtWidgets.QDialog):
    """
    全表比对结果摘要对话框
    """
    def __init__(self, summary, parent=None):
        super().__init__(parent)
        self.summary = summary
        self.setWindowTitle("全表比对结果摘要")
        self.resize(800, 500)
        self.setup_ui()

    def setup_ui(self):
        layout = QtWidgets.QVBoxLayout(self)

        # 1. 顶部统计信息
        total_sheets = len(self.summary)
        success_sheets = len([s for s in self.summary if s['status'] == 'success'])
        
        header_label = QtWidgets.QLabel(f"共检测到 {total_sheets} 个工作表，其中 {success_sheets} 个已完成比对。")
        header_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        layout.addWidget(header_label)

        # 2. 结果表格
        self.table = QtWidgets.QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(["文件 1 Sheet", "文件 2 Sheet", "状态", "新增", "删除", "修改"])
        self.table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setAlternatingRowColors(True)

        for i, item in enumerate(self.summary):
            self.table.insertRow(i)
            
            # Sheet 1
            s1 = item.get('sheet1', '-')
            self.table.setItem(i, 0, QtWidgets.QTableWidgetItem(str(s1) if s1 else "-"))
            
            # Sheet 2
            s2 = item.get('sheet2', '-')
            self.table.setItem(i, 1, QtWidgets.QTableWidgetItem(str(s2) if s2 else "-"))
            
            # Status
            status = item['status']
            status_text = "成功"
            color = "#FFFFFF"
            
            if status == 'success':
                stats = item.get('stats', {})
                if stats.get('added', 0) > 0 or stats.get('deleted', 0) > 0 or stats.get('modified', 0) > 0:
                    status_text = "发现差异"
                    color = "#FFF2CC" # 浅黄
                else:
                    status_text = "一致"
                    color = "#E2EFDA" # 浅绿
            elif status == 'error':
                status_text = f"错误: {item.get('error', '未知')}"
                color = "#F8CBAD" # 浅红
            elif status == 'only_in_file1':
                status_text = "仅在文件 1"
                color = "#D9E1F2" # 浅蓝
            elif status == 'only_in_file2':
                status_text = "仅在文件 2"
                color = "#D9E1F2" # 浅蓝
                
            status_item = QtWidgets.QTableWidgetItem(status_text)
            status_item.setBackground(QtGui.QColor(color))
            self.table.setItem(i, 2, status_item)
            
            # Stats columns
            if status == 'success':
                stats = item.get('stats', {})
                self.table.setItem(i, 3, QtWidgets.QTableWidgetItem(str(stats.get('added', 0))))
                self.table.setItem(i, 4, QtWidgets.QTableWidgetItem(str(stats.get('deleted', 0))))
                self.table.setItem(i, 5, QtWidgets.QTableWidgetItem(str(stats.get('modified', 0))))
            else:
                for col in range(3, 6):
                    self.table.setItem(i, col, QtWidgets.QTableWidgetItem("-"))

        layout.addWidget(self.table)

        # 3. 底部按钮
        btn_layout = QtWidgets.QHBoxLayout()
        close_btn = QtWidgets.QPushButton("关闭")
        close_btn.clicked.connect(self.accept)
        btn_layout.addStretch()
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)

        # 4. 双击表格跳转到特定 Sheet 比对 (TODO)
        # self.table.doubleClicked.connect(self.go_to_sheet_compare)
