import os, sys, traceback
import fitz
from PyQt5.QtWidgets import QApplication, QProgressBar, QWidget, QPushButton, QFileDialog, QLabel, QFrame, QMessageBox, QLineEdit
from PyQt5.QtCore import Qt, QThread, pyqtSignal


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
        self.setWindowTitle('pdf查错')
        # 进度条
        self.pbar = QProgressBar(self)
        # 进图条位置及大小
        self.pbar.setGeometry(20, 120, 470, 10)

        # todo 按钮
        # 选择XML文件按钮
        self.btn_select_file1 = QPushButton('选择词典文件', self)
        self.btn_select_file1.move(20, 190)
        self.btn_select_file1.setFixedSize(125, 30)
        self.btn_select_file1.clicked.connect(self.file_dialog1)
        # 选择文件按钮
        self.btn_select_file = QPushButton('选择pdf文件夹', self)
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
        self.lab_select_path1 = QLabel('词典文件路径', self)
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
        content = '说明:\n    选择需要查错的pdf文件夹，选择输出批注结果文件夹。\n    按照词典文件中包含的词查错每个pdf并生成对应页的pdf文件。\n    词典新建txt文件，每一行为一个要查错的词。'
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
        # # 提取开始页码
        # self.btn_select_file2 = QPushButton('提取开始页码', self)
        # self.btn_select_file2.move(20, 150)
        # self.btn_select_file2.setFixedSize(100, 30)
        #
        # int_validato = QIntValidator(50, 100, self)  # 实例化整型验证器，并设置范围为50-100
        # self.int_le = QLineEdit(self)  # 整型文本框
        # self.int_le.setValidator(int_validato)  # 设置验证
        # self.int_le.setFixedSize(100, 30)
        # self.int_le.move(120, 150)
        #
        # # 截止页码
        # self.btn_select_file1 = QPushButton('提取截止页码', self)
        # self.btn_select_file1.move(260, 150)
        # self.btn_select_file1.setFixedSize(100, 30)
        #
        # int_validato1 = QIntValidator(50, 100, self)  # 实例化整型验证器，并设置范围为50-100
        # self.int_le1 = QLineEdit(self)  # 整型文本框
        # self.int_le1.setValidator(int_validato1)  # 设置验证
        # self.int_le1.setFixedSize(100, 30)
        # self.int_le1.move(360, 150)

        self.step = 0

        # 显示
        self.show()

    # 按钮点击
    def btn_action(self):
        if self.btn.text() == '完成':
            self.close()
        else:
            pdf_dir_path = '{}'.format(self.lab_select_path.text())
            txt_file_path = '{}'.format(self.lab_select_path1.text())
            output_path = '{}'.format(self.lab_output_path.text())
            #start_pageno = self.int_le.text()
            #end_pageno = self.int_le1.text()

            if self.btn.text() == '开始':
                if not os.path.exists(pdf_dir_path):
                    QMessageBox.warning(self, '', '请选择pdf文件夹路径', QMessageBox.Yes)
                elif not os.path.exists(output_path):
                    QMessageBox.warning(self, '', '请选择输出路径', QMessageBox.Yes)
                elif not os.path.exists(txt_file_path):
                    QMessageBox.warning(self, '', '请选择词典文件路径', QMessageBox.Yes)
                else:
                    self.btn.setText('程序进行中')
                    self.downloadThread = downloadThread(pdf_dir_path, output_path, txt_file_path)
                    self.downloadThread.download_proess_signal.connect(self.set_progerss_bar)
                    self.downloadThread.start()

    # 选择输入文件路径
    def file_dialog(self):
        # './'表示当前路径
        # path = QFileDialog.getExistingDirectory(self, '选取文件', './', 'pdf文件(*.pdf)')
        path = QFileDialog.getExistingDirectory(self, '选取文件', './')
        print(path)
        # 标签框显示文本路径
        self.lab_select_path.setText(path)
        # 自动调整标签框大小
        self.lab_select_path.adjustSize()

    # 选择输入xml文件路径
    def file_dialog1(self):
        # './'表示当前路径
        # path = QFileDialog.getExistingDirectory(self, '选取文件', './', 'pdf文件(*.pdf)')
        path = QFileDialog.getOpenFileName(self, '选取文件', './',
                                           '词典文件(*.txt)')
        print(path)
        # 标签框显示文本路径
        self.lab_select_path1.setText(path[0])
        # 自动调整标签框大小
        self.lab_select_path1.adjustSize()

    # 选择输出路径
    def output_dialog(self):
        path = QFileDialog.getExistingDirectory(self, '选取文件夹', './')
        print(path)
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


class Name_Correcting():
    def __init__(self, names_path, pdf_path_dir, copy_path, download_proess_signal):
        self.names = self.get_names(names_path)
        self.copy_path1 = copy_path
        self.pdf_dir_path = pdf_path_dir
        self.download_proess_signal = download_proess_signal

    def get_names(self, txt_path):
        txt = open(txt_path, 'r', encoding='utf-8')
        line = txt.readlines()
        txt.close()
        print(line)
        return line

    def do_somthing(self):
        txt = os.path.join(self.copy_path1, 'result.txt')
        dirs = os.listdir(os.chdir(self.pdf_dir_path))
        num = 1
        total_size = len(dirs)
        #print('total_size:',total_size)
        with open(txt, 'a+', encoding='utf-8') as file:
            for files in dirs:
                process_num = int(num / total_size * 100)
                print('当前进度：{} / {}'.format(num, total_size))
                self.download_proess_signal.emit(process_num)
                if files.endswith('pdf') or files.endswith('PDF'):
                    pdf_path = os.path.join(self.pdf_dir_path, files)
                    print(pdf_path)
                    #try:
                    self.main_logic(pdf_path, file)
                    # except Exception as error:
                    #     print(traceback.format_exc())
                    #     continue
                num += 1

    def main_logic(self, pdf_path, file):
        try:
            doc = fitz.open(pdf_path)
            doc1 = fitz.open()
            errors = []
            for page_number in range(0, doc.pageCount):
                page = doc.loadPage(page_number)
                dl = page.getDisplayList()
                pt = dl.getTextPage()
                new_page = []
                for word in self.names:
                    try:
                        #areas = page.searchFor(word.strip(), quads=True)
                        # areas = doc.searchPageFor(page_number, word.strip(),
                        #                           quads=True)
                        areas = pt.search(word.strip(), quads=True)
                        if len(areas) > 0:
                            # page.addSquigglyAnnot(areas)
                            new_page.append(areas)
                            #print(areas, word.strip(), page.number)
                            errors.append(
                                word.strip() + '_' + str(page_number + 1))
                    except Exception as e:
                        raise Exception(sys.exc_info())
                        continue
                #print(page_number)
                if len(new_page) > 0:
                    # doc.deletePage(page.number)
                    # print('删除页码：{}'.format(page.number))
                    # if len(doc1)<= 0:
                    page1 = doc1.newPage()
                    page1.showPDFpage(page.rect, doc, page_number)
                    for area in new_page:
                        #print(page1.rect, page.rect, area, page.annots())
                        page1.addSquigglyAnnot(area)
                    #print(page1.annots())
                    # doc1.insertPDF(doc, from_page=page.number, to_page=page.number)
                    # doc1.insertPage(page.number)
                del pt
                del dl
                del page
            if len(errors) > 0:
                file.write(os.path.splitext(os.path.basename(pdf_path))[0] + ' ' + ';'.join(errors))
                file.write('\n')
                file.flush()
                # print(os.path.splitext(os.path.basename(para))[0] + ' ' + ';'.join(errors) +'\n')
            # print(len(doc1), os.path.join(self.copy_path1, os.path.splitext(os.path.basename(para))[0] +'_1.pdf'))
            if len(doc1) > 0:
                new_pdf_path = os.path.join(self.copy_path1,
                                       os.path.splitext(os.path.basename(pdf_path))[0] + '_1.pdf')
                if os.path.exists(new_pdf_path):
                    os.remove(new_pdf_path)

                doc1.save(new_pdf_path)
            doc.close()
            doc1.close()
            #shutil.move(pdf_path, self.copy_path)
        except Exception as error:
            #print(traceback.format_exc())
            raise Exception(sys.exc_info())
        finally:
            if not doc1.isClosed:
                doc1.close()
            if not doc.isClosed:
                doc.close()



class downloadThread(QThread):

    download_proess_signal = pyqtSignal(int)  # 创建信号

    def __init__(self, pdf_path, pdf_out, names_path):
        super(downloadThread, self).__init__()
        self.pdf_path = pdf_path
        self.pdf_out = pdf_out
        self.names_path = names_path

    def run(self):
        try:
            nc = Name_Correcting(self.names_path, self.pdf_path, self.pdf_out, self.download_proess_signal)
            nc.do_somthing()
        except Exception:
            print(sys.exc_info())


if __name__ == '__main__':
    print(fitz.VersionBind)
    app = QApplication(sys.argv)
    pbar = MainUI()
    sys.exit(app.exec_())