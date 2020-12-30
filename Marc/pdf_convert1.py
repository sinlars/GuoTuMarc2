from PyPDF2 import PdfFileReader, PdfFileWriter
import os, sys
import fitz
from PyQt5.QtWidgets import QApplication, QProgressBar, QWidget, QPushButton, QFileDialog, QLabel, QFrame, QMessageBox, QLineEdit
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QIntValidator
from xml.dom.minidom import parse
import xml.dom.minidom

class MainUI(QWidget):

    def __init__(self):
        super().__init__()
        self.init_ui()

    # 窗体GUI部分
    def init_ui(self):
        # 设置窗体大小
        self.setGeometry(500, 200, 350, 200)
        # 设置固定大小
        self.setFixedSize(520, 370)
        self.setWindowTitle('pdf拆分')
        # 进度条
        self.pbar = QProgressBar(self)
        # 进图条位置及大小
        self.pbar.setGeometry(20, 120, 450, 20)

        # todo 按钮
        # 选择XML文件按钮
        self.btn_select_file1 = QPushButton('选择xml文件', self)
        self.btn_select_file1.move(20, 190)
        self.btn_select_file1.setFixedSize(125, 30)
        self.btn_select_file1.clicked.connect(self.file_dialog1)
        # 选择文件按钮
        self.btn_select_file = QPushButton('选择pdf文件', self)
        self.btn_select_file.move(20, 230)
        self.btn_select_file.setFixedSize(125, 30)
        self.btn_select_file.clicked.connect(self.file_dialog)
        # 输出文件按钮
        self.btn_output_path = QPushButton('选择输出文件夹', self)
        self.btn_output_path.move(20, 270)
        self.btn_output_path.setFixedSize(125, 30)
        self.btn_output_path.clicked.connect(self.output_dialog)
        # 开始按钮
        self.btn = QPushButton('开始', self)
        # 创建按钮并移动
        self.btn.move(150, 310)
        self.btn.setFixedSize(200, 40)
        # 点击按钮，连接事件函数
        self.btn.clicked.connect(self.btn_action)

        # todo 标签
        # xml文件路径标签
        self.lab_select_path1 = QLabel('xml文件路径', self)
        self.lab_select_path1.move(150, 190)
        self.lab_select_path1.setFixedSize(320, 30)
        self.lab_select_path1.setFrameShape(QFrame.Box)
        self.lab_select_path1.setFrameShadow(QFrame.Raised)
        # 文件路径标签
        self.lab_select_path = QLabel('文件路径', self)
        self.lab_select_path.move(150, 230)
        self.lab_select_path.setFixedSize(320, 30)
        self.lab_select_path.setFrameShape(QFrame.Box)
        self.lab_select_path.setFrameShadow(QFrame.Raised)
        # 输出标签
        self.lab_output_path = QLabel('文件路径', self)
        self.lab_output_path.move(150, 270)
        self.lab_output_path.setFixedSize(320, 30)
        self.lab_output_path.setFrameShape(QFrame.Box)
        self.lab_output_path.setFrameShadow(QFrame.Raised)
        # 说明标签
        content = '说明:\n    选择需要拆分的pdf，选择输出结果文件夹。\n    pdf拆分按照填写的开始页码和截止页码的区间来拆分，只拆分当前目录内容。'
        self.description_lab = QLabel(content, self)
        self.description_lab.move(20, 10)
        self.description_lab.setFixedSize(450, 100)
        self.description_lab.setAlignment(Qt.AlignTop)
        self.description_lab.setFrameShape(QFrame.Box)
        self.description_lab.setFrameShadow(QFrame.Raised)
        # 自动换行
        # self.description_lab.adjustSize()
        self.description_lab.setWordWrap(True)
        #
        # 提取开始页码
        self.btn_select_file2 = QPushButton('提取开始页码', self)
        self.btn_select_file2.move(20, 150)
        self.btn_select_file2.setFixedSize(100, 30)

        int_validato = QIntValidator(50, 100, self)  # 实例化整型验证器，并设置范围为50-100
        self.int_le = QLineEdit(self)  # 整型文本框
        self.int_le.setValidator(int_validato)  # 设置验证
        self.int_le.setFixedSize(100, 30)
        self.int_le.move(120, 150)

        # 截止页码
        self.btn_select_file1 = QPushButton('提取截止页码', self)
        self.btn_select_file1.move(260, 150)
        self.btn_select_file1.setFixedSize(100, 30)

        int_validato1 = QIntValidator(50, 100, self)  # 实例化整型验证器，并设置范围为50-100
        self.int_le1 = QLineEdit(self)  # 整型文本框
        self.int_le1.setValidator(int_validato1)  # 设置验证
        self.int_le1.setFixedSize(100, 30)
        self.int_le1.move(360, 150)

        self.step = 0

        # 显示
        self.show()

    # 按钮点击
    def btn_action(self):
        if self.btn.text() == '完成':
            self.close()
        else:
            file_path = '{}'.format(self.lab_select_path.text())
            xml_file_path = '{}'.format(self.lab_select_path1.text())
            output_path = '{}'.format(self.lab_output_path.text())
            start_pageno = self.int_le.text()
            end_pageno = self.int_le1.text()

            if self.btn.text() == '开始':
                if not os.path.exists(file_path):
                    QMessageBox.warning(self, '', '请选择路径', QMessageBox.Yes)
                elif not os.path.exists(output_path):
                    QMessageBox.warning(self, '', '请选择输出路径',
                                        QMessageBox.Yes)
                elif start_pageno == '':
                    QMessageBox.warning(self, '', '请输入开始页码',
                                        QMessageBox.Yes)
                elif end_pageno == '':
                    QMessageBox.warning(self, '', '请输入截止页码',
                                        QMessageBox.Yes)
                else:
                    self.btn.setText('程序进行中')
                    # self.run()
                    self.downloadThread = downloadThread(file_path,
                                                         output_path, start_pageno, end_pageno)
                    self.downloadThread.download_proess_signal.connect(
                        self.set_progerss_bar)
                    self.downloadThread.start()

    # 选择输入文件路径
    def file_dialog(self):
        # './'表示当前路径
        # path = QFileDialog.getExistingDirectory(self, '选取文件', './', 'pdf文件(*.pdf)')
        path = QFileDialog.getOpenFileName(self, '选取文件', './',
                                           'pdf文件(*.pdf)')
        print(path)
        # 标签框显示文本路径
        self.lab_select_path.setText(path[0])
        # 自动调整标签框大小
        self.lab_select_path.adjustSize()

    # 选择输入xml文件路径
    def file_dialog1(self):
        # './'表示当前路径
        # path = QFileDialog.getExistingDirectory(self, '选取文件', './', 'pdf文件(*.pdf)')
        path = QFileDialog.getOpenFileName(self, '选取文件', './',
                                           'xml文件(*.xml)')
        print(path)
        # 标签框显示文本路径
        self.lab_select_path1.setText(path[0])
        # 自动调整标签框大小
        self.lab_select_path1.adjustSize()

    # 选择输出路径
    def output_dialog(self):
        path = QFileDialog.getExistingDirectory(self, '选取文件夹', './')

        # 标签框显示文本路径
        self.lab_output_path.setText(path)
        # 自动调整标签框大小
        self.lab_output_path.adjustSize()

    def set_progerss_bar(self, num):
        '''
        设置进图条函数
        :param num: 进度条进度（整数）
        :return:
        '''
        self.step = num
        self.pbar.setValue(self.step)
        if num == 100:
            self.btn.setText('完成')
            QMessageBox.information(self, "提示", "提取完成！")
            return




    def pdf_splitter(path):
        """
            pdf按照页码拆分
        :param path:
        :return:
        """
        fname = os.path.splitext(os.path.basename(path))[0]
        pdf = PdfFileReader(path)
        for page in range(pdf.getNumPages()):
            pdf_writer = PdfFileWriter()
            pdf_writer.addPage(pdf.getPage(page))
            output_filename = '{}_page_{}.pdf'.format(fname, page+1)
            with open(output_filename, 'wb') as out:
                pdf_writer.write(out)
            print('created:{}'.format(output_filename))





class downloadThread(QThread):

    download_proess_signal = pyqtSignal(int)  # 创建信号

    def __init__(self, pdf_path, pdf_out, start_page_no, end_page_no):
        super(downloadThread, self).__init__()
        self.pdf_path = pdf_path
        self.pdf_out = pdf_out
        self.start_page_no = int(start_page_no)
        self.end_page_no = int(end_page_no)


    def pdf_split(self, pdf, start, end, out_pdf_name):
        pdf_writer = PdfFileWriter()
        for page in range(start-1, end):
            pdf_writer.addPage(pdf.getPage(page))
        with open(out_pdf_name, 'wb') as out:
            pdf_writer.write(out)

    def strQ2B(self, ustring):
        """全角转半角"""
        rstring = ""
        for uchar in ustring:
            inside_code = ord(uchar)
            if inside_code == 12288:  # 全角空格直接转换
                inside_code = 32
            elif (inside_code >= 65281 and inside_code <= 65374):  # 全角字符（除空格）根据关系转化
                inside_code -= 65248

            rstring += chr(inside_code)
        return rstring

    def strB2Q(self, ustring):
        """半角转全角"""
        rstring = ""
        for uchar in ustring:
            inside_code = ord(uchar)
            if inside_code == 32:  # 半角空格直接转化
                inside_code = 12288
            elif inside_code >= 32 and inside_code <= 126:  # 半角字符（除空格）根据关系转化
                inside_code += 65248

            rstring += chr(inside_code)
        return rstring

    def run(self):
        try:

            doc = fitz.open(self.pdf_path)
            pdf = PdfFileReader(self.pdf_path)
            tocs = doc.getToC()
            print(tocs)
            for i, toc in enumerate(tocs):
                print(i, toc)
                toc_name = self.strQ2B(toc[1]).replace('\r','').replace('\t', '').replace('\n', '')
                if toc[0] == 1:
                    continue
                else:
                    if i+1 != len(tocs):
                        if toc[2] >= self.start_page_no and tocs[i + 1][2] <= self.end_page_no and toc[2] <= self.end_page_no:
                            pdf_name = os.path.join(self.pdf_out, '%s_%s-%s.pdf' % (toc[2], tocs[i+1][2], toc_name))
                            self.pdf_split(pdf, toc[2], tocs[i+1][2], pdf_name)
                            print(pdf_name)

                    else:
                        if toc[2] >= self.start_page_no and tocs[i + 1][2] <= self.end_page_no and toc[2] <= self.end_page_no:
                            pdf_name = os.path.join(self.pdf_out, '%s_%s-%s.pdf' % (toc[2], tocs[i + 1][2], toc_name))
                            self.pdf_split(pdf, toc[2], tocs[i + 1][2], pdf_name)
                            print(pdf_name)

        except Exception as e:
            print(e)
            self.download_proess_signal.emit(int(i / len(tocs) * 100))
        self.download_proess_signal.emit(int(100))


if __name__ == '__main__':
    # pdf_path = '/Users/mac5318/Downloads/B8AAE79AD77014450BACA793B2C5F71F4000.pdf'
    # doc = fitz.open(pdf_path)
    # fname = os.path.splitext(os.path.basename(pdf_path))[0]
    # pdf_dir_name =  os.path.join(os.path.dirname(pdf_path), fname)
    # if os.path.exists(pdf_dir_name):
    #     print("{}文件夹已存在".format(pdf_dir_name))
    # else:
    #     os.mkdir(pdf_dir_name)
    #     print("{}文件夹已创建".format(pdf_dir_name))
    # #print(doc.pageCount)
    # tocs = doc.getToC()
    # #print(tocs)
    # # print(list(filter(lambda x: x[0] == 2 and x[2] != -1, tocs)))
    # # tocs.sort(key=lambda x: x[0])
    # # print(max([x[0] for x in tocs]))
    # # for i in range(1, max([x[0] for x in tocs]) + 1):
    # #     print(list(filter(lambda x: x[0] == i and x[2] != -1, tocs)))
    # pdf = PdfFileReader(pdf_path)
    # start_page_no = 32
    # end_page_no = 204
    # for i, toc in enumerate(tocs):
    #     #print(i, toc)
    #     if toc[0] == 1:
    #         continue
    #     if i+1 != len(tocs):
    #         if toc[2] >= start_page_no and tocs[i + 1][2] <= end_page_no and toc[2] <= end_page_no:
    #             #print(i, toc, '%s_%s-%s.pdf' % (toc[1], toc[2], tocs[i+1][2]))
    #             pdf_name = os.path.join(pdf_dir_name, '%s_%s-%s.pdf' % (toc[1], toc[2], tocs[i+1][2]))
    #             pdf_split(pdf, toc[2], tocs[i+1][2], pdf_name)
    #             print(pdf_name)
    #     else:
    #         if toc[2] >= start_page_no and tocs[i + 1][2] <= end_page_no and toc[2] <= end_page_no:
    #             #print(i, toc, '%s_%s-%s.pdf' % (toc[1], toc[2], doc.pageCount))
    #             pdf_name = os.path.join(pdf_dir_name, '%s_%s-%s.pdf' % (toc[1], toc[2], tocs[i + 1][2]))
    #             pdf_split(pdf, toc[2], tocs[i + 1][2], pdf_name)
    #             print(pdf_name)

    print(fitz.VersionBind)
    app = QApplication(sys.argv)
    pbar = MainUI()
    sys.exit(app.exec_())