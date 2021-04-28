# coding=utf-8
import os
import sys
from UploadWidget import Ui_MainWidget
from paper_write_hw import PaperTextChecker
from PyQt5.QtWidgets import QWidget, QApplication, QMessageBox, QFileDialog
ui = Ui_MainWidget


class Achieve(QWidget, ui):

    def __init__(self):
        QWidget.__init__(self)
        ui.__init__(self)
        self.setupUi(self)
        self.load_file_button.clicked.connect(self.load_file_button_clicked)

    def load_file_button_clicked(self):
        file_dir = QFileDialog.getOpenFileName(self, '请选择要检测的文件', os.path.dirname(__file__))
        if file_dir[0] == "":
            return
        try:
            file_path = file_dir[0].replace('/', '\\')
            checker = PaperTextChecker(file_path=file_path)
            checker.check_all_paragraph()
        except ValueError as error:
            QMessageBox.warning(self, '错误提示', str(error))
        except Exception as error:
            QMessageBox.warning(self, '未知错误', str(error))
        else:
            QMessageBox.information(self, '完成提示', f'检测完成, 检测报告已保存在{checker.check_report_file_path}文件')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ui = Achieve()
    ui.show()
    sys.exit(app.exec_())
