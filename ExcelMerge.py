import os
import sys

import openpyxl
import xlrd
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox

import ExcelMergeUI


def input_file_read(input_file):
    row_list = []
    wb = xlrd.open_workbook(input_file)
    ws = wb.sheet_by_name('问题记录和解决')
    for i in range(ws.nrows):
        if i == 0:
            continue
        print(ws.row_values(i))
        row_list.append(ws.row_values(i))
    return row_list


def get_all_files(dir):
    files_ = []
    list = os.listdir(dir)
    for i in range(0, len(list)):
        path = os.path.join(dir, list[i])
        if os.path.isdir(path):
            files_.extend(get_all_files(path))
        if os.path.isfile(path):
            files_.append(path)
    return files_


def main(input_dir, output_file, file_type):
    output_wb = openpyxl.Workbook()
    output_ws = output_wb.active
    output_ws.title = "合并结果"
    files = get_all_files(input_dir)
    for f in files:
        if f[-(len(file_type) + 1):] == '.' + file_type:
            print(f)
            content = input_file_read(f)
            for c in content:
                output_ws.append(c)
    output_wb.save(output_file)


class ExcelMerge(QMainWindow):
    def __init__(self):
        super(ExcelMerge, self).__init__()
        self.ui = ExcelMergeUI.Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.runButton.clicked.connect(self.run_clicked)
        self.ui.inputFolderButton.clicked.connect(self.input_folder_button_clicked)
        self.ui.outputFileButton.clicked.connect(self.output_file_button_clicked)

    def run_clicked(self):
        input_dir = self.ui.inputFolderEdit.text()
        output_file = self.ui.outputFileEdit.text()
        file_type = self.ui.fileTypeComboBox.currentText()
        main(input_dir, output_file, file_type)
        QMessageBox.information(self, "Run", "Finished")

    def input_folder_button_clicked(self):
        self.ui.inputFolderEdit.setText(QFileDialog.getExistingDirectory(self, "Select Input Folder"))

    def output_file_button_clicked(self):
        fileName, _ = QFileDialog.getSaveFileName(self, "Select Output File", "output.xlsx",
                                                  "All Files (*);;Excel Files (*.xlsx)")
        self.ui.outputFileEdit.setText(fileName)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ExcelMerge()
    window.show()
    sys.exit(app.exec_())
