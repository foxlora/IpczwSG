import sys
from PyQt5.QtGui import QColor,QIcon
from PyQt5.QtWidgets import QApplication,QMainWindow,QFileDialog,QTableWidgetItem,QTableWidget,QMessageBox
import os
import tools.common as common

from mainWindow import *


class MyMainWindow(QMainWindow,Ui_MainWindow):
    def __init__(self):
        super(MyMainWindow,self).__init__()
        self.setupUi(self)
        self.setGeometry(200,200,1024,600)
        self.setWindowTitle('IP承载网脚本自动化生成v1.0')
        self.setWindowIcon(QIcon("../image/timg.png"))
        palette = QtGui.QPalette()
        # palette.setBrush(self.backgroundRole(), QBrush(QPixmap("../image/bg.png").scaled(self.size(), QtCore.Qt.IgnoreAspectRatio, QtCore.Qt.SmoothTransformation)))
        self.setPalette(palette)
        self.setAutoFillBackground(True)  # 设置自动填充背景
        # self.setFixedSize(1024, 600)  # 禁止显示最大化按钮及调整窗体大小

    def excelFileClick(self):#单击浏览Excel源文件所触发方法
        file_path = QFileDialog.getOpenFileName(self,"请选择Excel调单文件")
        print(file_path)
        if file_path[0]:
            try:
                self.lineEdit_2.setText(file_path[0])
            except:
                self.lineEdit_2.setText("打开文件失败，可能是文件内yyy型错误")



def show_MainWindow():
    app = QApplication(sys.argv)
    main = MyMainWindow()

    main.pushButton_2.clicked.connect(main.excelFileClick)



    main.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    show_MainWindow()