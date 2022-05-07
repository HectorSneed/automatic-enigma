import sys
import os
from PyQt5 import QtWidgets, QtGui, QtCore
from MyGui import Ui_widget
from PyQt5.QtWidgets import *
from PyQt5.QtCore import QFileInfo
from Solution import analyFile


class MainGui(QtWidgets.QWidget, Ui_widget):
    def __init__(self):
        super(MainGui, self).__init__()
        self.setupUi(self)

    def SelectFile(self):
        fileName, fileType = QFileDialog.getOpenFileName(self, "选择文件", "/", "Excel Files (*.xls);;Excel Files (*.xlsx)")
        self.lineEdit.setText(fileName)
        #print(fileName, fileType)
        fileinfo = QFileInfo(fileName)
        file_path = fileinfo.absolutePath()
        file_name = fileinfo.fileName()
        file_absolute_path = file_path + '/' + file_name
        return file_absolute_path

    def SelectDir(self):
        directory = QFileDialog.getExistingDirectory(self, "选择文件夹", "/")
        self.lineEdit_2.setText(directory)
        #print(directory)
        return directory

    def Start(self):
        file_name = self.lineEdit.text()
        dir_path = self.lineEdit_2.text()
        begin_pos = self.lineEdit_3.text()
        end_pos = self.lineEdit_4.text()
        addr_input = self.lineEdit_5.text()
        addr_output = self.lineEdit_6.text()
        distence_output = self.lineEdit_7.text()
        if file_name == "":
            reply = QMessageBox.information(self, "ERROR", "请选择一个excel文件", QMessageBox.Ok)
            if reply == QMessageBox.Ok:
                return
        elif (not os.path.exists(file_name)) and (os.path.split(file_name)[1] != ".xls") and (os.path.split(file_name)[1] != ".xls"):
            reply = QMessageBox.information(self, "ERROR", "文件不存在或不为excel文件", QMessageBox.Ok)
            if reply == QMessageBox.Ok:
                return
        elif (not os.path.exists(dir_path)) or (not os.path.isdir(dir_path)):
            reply = QMessageBox.information(self, "ERROR", "请选择一个文件夹保存截图", QMessageBox.Ok)
            if reply == QMessageBox.Ok:
                return
        elif begin_pos == "" or end_pos == "" or addr_output == "" or distence_output == "" or addr_input == "":
            reply = QMessageBox.information(self, "ERROR", "请输入对应的列序号", QMessageBox.Ok)
            if reply == QMessageBox.Ok:
                return
        else:
            file_name = file_name.replace('\\', '/')
            #print(file_name)
            #print(addr_input)
            #print(addr_output)
            #print(distence_output)
            #print(dir_path)
            name_input = begin_pos[0:1]
            begin_pos_int = int(begin_pos[1:])
            end_pos_int = int(end_pos[1:])
            analyFile(file_name, dir_path, begin_pos_int, end_pos_int, name_input, addr_input, addr_output, distence_output)

    def clearText(self):
        self.ui.textEdit.clear()


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    myDlg = MainGui()
    myDlg.show()
    sys.exit(app.exec_())