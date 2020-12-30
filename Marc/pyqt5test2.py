import sys
import os
import time
from PyQt5.QtWidgets import QApplication, QProgressBar, QWidget, QPushButton, QFileDialog, QLabel, QFrame, QMessageBox, \
    QInputDialog, QLineEdit
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QIcon
import shutil
import fitz
import openpyxl
from PIL import Image, ImageOps


class MainUI(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    # 窗体GUI部分
    def init_ui(self):
        # 设置窗体大小
        self.setGeometry(500, 200, 350, 200)
        # 设置固定大小
        self.setFixedSize(500, 370)
        self.setWindowTitle('pdf图片提取程序')
        #self.setWindowIcon(QIcon('icon/769160.png'))

        # 设置软件过期时间
        data = '2220-8-21 13:50:00'
        data_array = time.strptime(data, "%Y-%m-%d %H:%M:%S")
        timeStamp = int(time.mktime(data_array))
        # print(timeStamp)
        # 判断软件是否过期
        if time.time() > timeStamp:
            QMessageBox.warning(self, '', '软件已过期，请联系作者', QMessageBox.Yes)
        else:
            # 进度条
            self.pbar = QProgressBar(self)
            # 进图条位置及大小
            self.pbar.setGeometry(20, 200, 480, 6)

            # todo 按钮
            # 选择文件按钮
            self.btn_select_file = QPushButton('选择pdf文件', self)
            self.btn_select_file.move(20, 226)
            self.btn_select_file.setFixedSize(125, 30)
            self.btn_select_file.clicked.connect(self.file_dialog)
            # 输出文件按钮
            self.btn_output_path = QPushButton('选择输出文件夹', self)
            self.btn_output_path.move(20, 265)
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
            # 文件路径标签
            self.lab_select_path = QLabel('文件路径', self)
            self.lab_select_path.move(150, 225)
            self.lab_select_path.setFixedSize(320, 30)
            self.lab_select_path.setFrameShape(QFrame.Box)
            self.lab_select_path.setFrameShadow(QFrame.Raised)
            # 输出标签
            self.lab_output_path = QLabel('文件路径', self)
            self.lab_output_path.move(150, 265)
            self.lab_output_path.setFixedSize(320, 30)
            self.lab_output_path.setFrameShape(QFrame.Box)
            self.lab_output_path.setFrameShadow(QFrame.Raised)
            # 说明标签
            content = '说明:\n    选择需要提取图片的pdf；选择输出结果文件夹。\n'
            self.description_lab = QLabel(content, self)
            self.description_lab.move(20, 10)
            self.description_lab.setFixedSize(450, 100)
            self.description_lab.setAlignment(Qt.AlignTop)
            self.description_lab.setFrameShape(QFrame.Box)
            self.description_lab.setFrameShadow(QFrame.Raised)
            # 自动换行
            # self.description_lab.adjustSize()
            self.description_lab.setWordWrap(True)

            self.step = 0

            # 显示
            self.show()

    # 按钮点击
    def btn_action(self):
        if self.btn.text() == '完成':
            self.close()
        else:
            file_path = '{}'.format(self.lab_select_path.text())
            output_path = '{}'.format(self.lab_output_path.text())

            if self.btn.text() == '开始':
                if not os.path.exists(file_path):
                    QMessageBox.warning(self, '', '请选择路径', QMessageBox.Yes)
                elif not os.path.exists(output_path):
                    QMessageBox.warning(self, '', '请选择输出路径', QMessageBox.Yes)
                else:
                    self.btn.setText('程序进行中')
                    #self.run()
                    self.downloadThread = downloadThread(file_path, output_path)
                    self.downloadThread.download_proess_signal.connect(self.set_progerss_bar)
                    self.downloadThread.start()

    # 选择输入文件路径
    def file_dialog(self):
        # './'表示当前路径
        #path = QFileDialog.getExistingDirectory(self, '选取文件', './', 'pdf文件(*.pdf)')
        path = QFileDialog.getOpenFileName(self, '选取文件', './', 'pdf文件(*.pdf)')
        print(path)
        # 标签框显示文本路径
        self.lab_select_path.setText(path[0])
        # 自动调整标签框大小
        self.lab_select_path.adjustSize()

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


class downloadThread(QThread):
    download_proess_signal = pyqtSignal(int)                        #创建信号

    def __init__(self, pdf_path, pic_path):
        super(downloadThread, self).__init__()
        self.pdf_path = pdf_path
        self.pic_path = pic_path

    def lookup_matrix(self, page, imgname):
        """Return the transformation matrix for an image name.

        Args:
            :page: the PyMuPDF page object
            :imgname: the image reference name, must equal the name in the
                list doc.getPageImageList(page.number).
        Returns:
            concatenated matrices preceeding the image invocation.

        Notes:
            We are looking up "/imgname Do" in the concatenated /Contents of the
            page first. If not found, also look it up in the streams of any
            Form XObjects of the page. If still not found, return the zero matrix.
        """
        doc = page.parent  # get the PDF document
        if not doc.isPDF:
            raise ValueError("not PDF")

        page._cleanContents()  # sanitize image invocation matrices
        xref = page._getContents()[0]  # the (only) contents object
        cont = doc._getXrefStream(xref)  # the contents object
        cont = cont.replace(b"/", b" /")  # prepend slashes with a space
        # split this, ignoring white spaces
        cont = cont.split()

        imgnm = bytes("/" + imgname, "utf8")
        if imgnm in cont:
            idx = cont.index(imgnm)  # the image name is found here
        else:  # not in page /contents, so look in Form XObjects
            cont = None
            xreflist = doc._getPageInfo(page.number, 3)  # XObject xrefs
            for item in xreflist:
                cont = doc._getXrefStream(item[0]).split()
                if imgnm not in cont:
                    cont = None
                    continue
                idx = cont.index(imgnm)  # image name found here
                break

        if cont is None:  # safeguard against inconsistencies
            return fitz.Matrix()

        # list of matrices preceeding image invocation command.
        # not really required, because clean contents has concatenated those
        mat_list = []
        while idx >= 0:  # start value is "/Image Do" location
            if cont[idx] == b"q":  # finished at leading stacking command
                break
            if cont[idx] == b"cm":  # encountered a matrix command
                mat = cont[idx - 6: idx]  # list of the 6 matrix values
                l = list(map(float, mat))  # make them floats
                mat_list.append(fitz.Matrix(l))  # append fitz matrix
                idx -= 6  # step backwards 6 entries
            else:
                idx -= 1  # step backwards

        l = len(mat_list)
        if l == 0:  # safeguard against unusual situations
            return fitz.Matrix()  # the zero matrix

        mat = fitz.Matrix(1, 1)  # concatenate encountered matrices to this one
        for m in reversed(mat_list):
            mat *= m

        return mat

    def get_rot(self, matrix):
        """
        :param matrix: 图片矩阵
        :return: 返回图片翻转的角度
        """
        rot = None
        if min(matrix.a, matrix.d) > 0 and matrix.b == matrix.c == 0:
            rot = 0
        elif matrix.a == matrix.d == 0:
            if matrix.b > 0 and matrix.c < 0:
                rot = 90
            elif matrix.b < 0 and matrix.c > 0:
                rot = -90
            else:
                rot = None
        elif min(matrix.a, matrix.d) < 0 and matrix.b == matrix.c == 0:
            rot = 180
        else:
            rot = None
        return rot

    def run(self):
        try:
            print(self.pdf_path)
            doc = fitz.open(self.pdf_path)
            repetitive_xref = []
            for page in doc:
                self.download_proess_signal.emit(int((page.number + 1) / doc.pageCount * 100))
                try:
                    list_image = page.getImageList()
                except Exception as e:
                    print('Page ' + page.number + 1 + ': 发生未知错误！', e)
                    continue
                short = lambda x: round(x, 4)
                #i = 1
                # if page.number == 17 :
                #     print(page.getText('dict'))
                #     for block in page.getText('dict')['blocks']:
                #         if block['type'] == 1:
                #             print(block)
                #         if block['type'] == 0:
                #             lines = block['lines']
                #             for line in lines:
                #                 print(''.join([span['text'] for span in line['spans'] if int(span['size']) <= 9 ]))
                #                     #print(span['size'], span['text'])

                for i, item in enumerate(list_image):  # loop through all images on the page
                    #print(item)
                    img = item[7]  # we need the image reference name
                    matrix = self.lookup_matrix(page, img)  # find matrix for the image
                    if not bool(matrix):  # no display command  found for image
                        print("Image '%s' not found on page %i." % (img, page.number))
                        continue
                    matrix = fitz.Matrix(tuple(map(short, matrix)))
                    #print(matrix)
                    # print(page.getImageBbox(item))
                    img_obj = doc.extractImage(item[0])
                    #print(img_obj)
                    image_name = os.path.join(self.pic_path, "P{}_{}.{}".format(page.number + 1, i, img_obj['ext']))
                    pix = fitz.Pixmap(doc, item[0])
                    print(len(pix))
                    if len(pix) <= 2048: #如果图片小于2KB，忽略
                        continue
                    if min(pix.width, pix.height) < 100:
                        continue
                    if item[0] in repetitive_xref:
                        print('重复的图片对象：', item[0])
                        continue
                    else:
                        repetitive_xref.append(item[0])
                    rot = self.get_rot(matrix)
                    print("Page %i / %i, image '%s', rotation: %s" % (page.number + 1, doc.pageCount, img, rot))

                    if pix.n - pix.alpha < 4:
                        pix.writeImage(image_name)
                    else:  # 否则先转换CMYK
                        pix0 = fitz.Pixmap(fitz.csRGB, pix)
                        pix0.writeImage(image_name)
                        pix0 = None
                    if rot == 180:
                        image = Image.open(image_name)
                        out3 = image.transpose(Image.FLIP_TOP_BOTTOM)
                        out3.save(image_name)
                    image1 = Image.open(image_name)
                    colorslist = image1.getcolors()
                    print(colorslist)
                    if colorslist:
                        if len(colorslist) <= 2:
                            if len(colorslist) == 2:
                                if colorslist[0][0] > colorslist[1][0]:
                                    image2 = ImageOps.invert(image1)
                                    image2.save(image_name)


                    #i = i + 1
            shutil.copyfile(self.pdf_path, os.path.join(self.pic_path, os.path.splitext(os.path.basename(self.pdf_path))[0] + '.pdf'))
            self.download_proess_signal.emit(int(100))
            self.exit(0)
        except Exception as e:
             print(e)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    pbar = MainUI()
    sys.exit(app.exec_())
    # file_path = '/Users/mac5318/Downloads/27juan.pdf'
    # output_path = '/Users/mac5318/Downloads/27juan'
    # print(os.path.splitext(os.path.basename(file_path))[0] + '.pdf')
