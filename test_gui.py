import sys
import os
from PyQt5 import QtWidgets, QtCore

# 设置插件路径
import PyQt5
os.environ["QT_QPA_PLATFORM_PLUGIN_PATH"] = os.path.join(os.path.dirname(PyQt5.__file__), "Qt5", "plugins")

app = QtWidgets.QApplication(sys.argv)
w = QtWidgets.QMainWindow()
w.setWindowTitle("Test Window")
w.resize(400, 300)
w.show()
print("Window shown, entering exec...")
sys.exit(app.exec_())
