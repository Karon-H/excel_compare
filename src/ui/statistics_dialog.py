from PyQt5 import QtWidgets, QtCore, QtGui

class StatisticsDialog(QtWidgets.QDialog):
    """
    比对结果统计仪表盘
    """
    def __init__(self, stats, parent=None):
        super().__init__(parent)
        self.setWindowTitle("比对统计仪表盘")
        self.resize(500, 450)
        self.stats = stats  # {'added': X, 'deleted': Y, 'modified_rows': Z, 'modified_cells': C, 'equal': E}
        self.setup_ui()

    def setup_ui(self):
        layout = QtWidgets.QVBoxLayout(self)
        
        # 1. 标题
        title_label = QtWidgets.QLabel("比对结果统计摘要")
        title_label.setAlignment(QtCore.Qt.AlignCenter)
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; margin-bottom: 10px;")
        layout.addWidget(title_label)

        # 2. 中间：图表与文字统计
        content_layout = QtWidgets.QHBoxLayout()
        
        # 左侧：饼图绘制区域
        self.chart_view = ChartWidget(self.stats)
        content_layout.addWidget(self.chart_view, 2)
        
        # 右侧：详细文字统计
        stats_text_layout = QtWidgets.QVBoxLayout()
        stats_text_layout.setSpacing(15)
        stats_text_layout.addStretch()
        
        # 定义颜色 (与图表一致)
        colors = {
            '新增': "#CCFFCC",
            '删除': "#FFFFCC",
            '修改': "#E6F3FF",
            '一致': "#F0F0F0"
        }

        self.add_stat_item(stats_text_layout, "新增行数", self.stats['added'], colors['新增'])
        self.add_stat_item(stats_text_layout, "删除行数", self.stats['deleted'], colors['删除'])
        self.add_stat_item(stats_text_layout, "修改行数", self.stats['modified_rows'], colors['修改'])
        self.add_stat_item(stats_text_layout, "修改单元格", self.stats['modified_cells'], "#FFCCCC")
        self.add_stat_item(stats_text_layout, "一致行数", self.stats['equal'], colors['一致'])
        
        stats_text_layout.addStretch()
        content_layout.addLayout(stats_text_layout, 1)
        
        layout.addLayout(content_layout)

        # 3. 底部：确认按钮
        buttons = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok)
        buttons.accepted.connect(self.accept)
        layout.addWidget(buttons)

    def add_stat_item(self, layout, label, value, color):
        item_layout = QtWidgets.QHBoxLayout()
        
        color_box = QtWidgets.QLabel()
        color_box.setFixedSize(16, 16)
        color_box.setStyleSheet(f"background-color: {color}; border: 1px solid #999;")
        
        text_label = QtWidgets.QLabel(f"{label}:")
        text_label.setStyleSheet("font-weight: bold;")
        
        value_label = QtWidgets.QLabel(str(value))
        value_label.setStyleSheet("color: #333;")
        
        item_layout.addWidget(color_box)
        item_layout.addWidget(text_label)
        item_layout.addWidget(value_label)
        item_layout.addStretch()
        
        layout.addLayout(item_layout)

class ChartWidget(QtWidgets.QWidget):
    """
    使用 QPainter 绘制简单的饼图
    """
    def __init__(self, stats, parent=None):
        super().__init__(parent)
        self.stats = stats
        self.setMinimumSize(250, 250)

    def paintEvent(self, event):
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.Antialiasing)
        
        rect = QtCore.QRect(20, 20, min(self.width(), self.height()) - 40, min(self.width(), self.height()) - 40)
        
        # 统计总数（不包括修改单元格，只统计行状态）
        total = self.stats['added'] + self.stats['deleted'] + self.stats['modified_rows'] + self.stats['equal']
        if total == 0:
            painter.drawText(rect, QtCore.Qt.AlignCenter, "暂无数据")
            return

        # 准备数据
        data = [
            (self.stats['added'], QtGui.QColor("#CCFFCC")),    # 新增
            (self.stats['deleted'], QtGui.QColor("#FFFFCC")),  # 删除
            (self.stats['modified_rows'], QtGui.QColor("#E6F3FF")), # 修改
            (self.stats['equal'], QtGui.QColor("#F0F0F0"))     # 一致
        ]
        
        current_angle = 90 * 16  # 从 12 点钟方向开始 (Qt 以 1/16 度为单位)
        
        for value, color in data:
            if value == 0:
                continue
            
            span_angle = int((value / total) * 360 * 16)
            
            painter.setBrush(QtGui.QBrush(color))
            painter.setPen(QtGui.QPen(QtGui.QColor("#999999"), 1))
            painter.drawPie(rect, current_angle, span_angle)
            
            current_angle += span_angle
