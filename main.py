import os
import sys
import tempfile
import uuid

import PyPDF2
import cups
from PyPDF2 import PdfFileReader
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QApplication, QDesktopWidget, QPushButton, QFileDialog, \
    QLineEdit, QWidget, QGridLayout, QLabel, QMessageBox, QCheckBox, QComboBox


class MainWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.conn = None
        self.print_pdf_btn = QPushButton('打印')
        self.select_pdf_btn = QPushButton('选择PDF')
        self.double_page_checkbox = QCheckBox()
        self.file_path_edit = QLineEdit()
        self.file_path_edit.setEnabled(False)
        self.file_name_edit = QLineEdit()
        self.file_name_edit.setEnabled(False)
        self.printer_combo_box = QComboBox()
        self._init_ui()
        self._find_printer()

    def _center(self):
        """
        将弹框移到屏幕中央
        :return: 无
        """
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def _init_ui(self):
        grid = QGridLayout()
        grid.setSpacing(20)
        cur_line = 1
        # printer
        printer_name = QLabel('选择打印机')
        grid.addWidget(printer_name, cur_line, 0)
        grid.addWidget(self.printer_combo_box, cur_line, 1)
        cur_line += 1

        # pdf file path
        file_name_label = QLabel('文件名称')
        grid.addWidget(file_name_label, cur_line, 0)
        grid.addWidget(self.file_name_edit, cur_line, 1)
        cur_line += 1
        file_path_label = QLabel('文件路径')
        grid.addWidget(file_path_label, cur_line, 0)
        grid.addWidget(self.file_path_edit, cur_line, 1)
        cur_line += 1
        # double page
        double_page_label = QLabel('双面打印')
        grid.addWidget(double_page_label, cur_line, 0)
        grid.addWidget(self.double_page_checkbox, cur_line, 1)
        cur_line += 1
        # buttons
        self.print_pdf_btn.clicked.connect(self.print_dialog)
        self.select_pdf_btn.clicked.connect(self.select_pdf_dialog)
        grid.addWidget(self.select_pdf_btn, cur_line, 0)
        grid.addWidget(self.print_pdf_btn, cur_line, 1)

        self.setLayout(grid)

        self.setGeometry(300, 300, 800, 500)
        self._center()
        self.setWindowTitle('如何双面打印我的PDF')
        self.setWindowIcon(QIcon('./favicon.ico'))
        self.show()

    def print_dialog(self):
        """
        打印的提示
        :return: 无
        """
        self.double_page_checkbox.setDisabled(True)
        self.printer_combo_box.setDisabled(True)
        self.select_pdf_btn.setDisabled(True)
        self.print_pdf_btn.setDisabled(True)
        if not os.path.isfile(self.file_path_edit.text()):
            QMessageBox.critical(self, '错误', '亲，您选择的PDF文件貌似出了点问题')
            return
        if not self.printer_combo_box.currentText():
            QMessageBox.critical(self, '错误', '亲，您选择的打印机貌似有些不对劲')
            return
        try:
            tmp_dir = tempfile.TemporaryDirectory().name
            os.makedirs(tmp_dir, exist_ok=True)
            tmp_path = os.path.join(tmp_dir, f'{uuid.uuid4()}.pdf')
            fin = open(self.file_path_edit.text(), 'rb')
            pdf = PdfFileReader(fin)
            page_num = pdf.getNumPages()
            temp_pdf = PyPDF2.PdfFileWriter()
            temp_pdf.appendPagesFromReader(pdf)
            if self.double_page_checkbox.isChecked() and page_num % 2 != 0:
                temp_pdf.addBlankPage()
            with open(tmp_path, 'wb') as fout:
                temp_pdf.write(fout)
            fin.close()
            tmp_pdf = PdfFileReader(tmp_path)
            page_num = tmp_pdf.getNumPages()
            if self.double_page_checkbox.isChecked():
                all_pages = [str(it + 1) for it in range(page_num)]
                odd_str = ','.join(all_pages[0::2])
                even_str = ','.join(all_pages[1::2])
                self.conn.printFile(self.printer_combo_box.currentText(), tmp_path, self.file_name_edit.text(),
                                    {'page-ranges': odd_str})
                QMessageBox.warning(self, '注意',
                                    "正面已经发送给打印机！请将纸重新放回打印机，然后点击确认按钮打印背面", QMessageBox.Yes, QMessageBox.Yes)
                self.conn.printFile(self.printer_combo_box.currentText(), tmp_path, self.file_name_edit.text(),
                                    {'outputorder': 'reverse', 'page-ranges': even_str})
                QMessageBox.information(self, '成功', '好耶！你的文档的背面也已经发送给打印机了')
            else:
                self.conn.printFile(self.printer_combo_box.currentText(), tmp_path, self.file_name_edit.text(), {})
                QMessageBox.information(self, '成功', '你的文档已经被发送到打印机')
        except:
            QMessageBox.critical(self, '错误', '哎呀，发生了一个错误，你的文件可能没法打印')
        finally:
            self.printer_combo_box.setEnabled(True)
            self.double_page_checkbox.setEnabled(True)
            self.select_pdf_btn.setEnabled(True)
            self.print_pdf_btn.setEnabled(True)

    def select_pdf_dialog(self):
        """
        选择打印的PDF的框
        :return: 无
        """
        f_name = QFileDialog.getOpenFileName(self, '选择需要打印的PDF文件', os.path.curdir, "PDF (*.pdf)")
        if f_name[0]:
            self.file_path_edit.setText(f_name[0])
            self.file_name_edit.setText(os.path.basename(f_name[0]))

    def _find_printer(self):
        """
        寻找打印机
        :return: 无
        """
        self.conn = cups.Connection()
        self.printers = self.conn.getPrinters()
        for k, _ in self.printers.items():
            self.printer_combo_box.addItem(k, k)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = MainWidget()
    sys.exit(app.exec_())
