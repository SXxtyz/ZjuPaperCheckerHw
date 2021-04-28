# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'E:\Pycharm\exam\UploadWidget.ui'
#
# Created by: PyQt5 UI code generator 5.9.2
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_MainWidget(object):
    def setupUi(self, MainWidget):
        MainWidget.setObjectName("MainWidget")
        MainWidget.resize(916, 619)
        self.load_file_button = QtWidgets.QPushButton(MainWidget)
        self.load_file_button.setGeometry(QtCore.QRect(350, 250, 141, 81))
        font = QtGui.QFont()
        font.setFamily("宋体")
        font.setPointSize(12)
        self.load_file_button.setFont(font)
        self.load_file_button.setObjectName("load_file_button")

        self.retranslateUi(MainWidget)
        QtCore.QMetaObject.connectSlotsByName(MainWidget)

    def retranslateUi(self, MainWidget):
        _translate = QtCore.QCoreApplication.translate
        MainWidget.setWindowTitle(_translate("MainWidget", "选择文件"))
        self.load_file_button.setText(_translate("MainWidget", "上传文件"))

