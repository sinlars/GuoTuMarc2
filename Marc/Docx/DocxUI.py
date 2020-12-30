import os, sys
from PyQt5.QtWidgets import QApplication, QProgressBar, QWidget, QPushButton, QFileDialog, QLabel, QFrame, QMessageBox
from PyQt5.QtCore import Qt
from Marc.Docx.DocxThread import downloadThread



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
        self.setWindowTitle('docx中参考文献识别')
        # 进度条
        self.pbar = QProgressBar(self)
        # 进图条位置及大小
        self.pbar.setGeometry(20, 120, 470, 10)

        # todo 按钮
        # 选择文件按钮
        self.btn_select_file = QPushButton('选择docx文件', self)
        self.btn_select_file.move(20, 230)
        self.btn_select_file.setFixedSize(125, 30)
        self.btn_select_file.clicked.connect(self.file_dialog)
        # # 输出文件按钮
        # self.btn_output_path = QPushButton('选择输出文件夹', self)
        # self.btn_output_path.move(20, 270)
        # self.btn_output_path.setFixedSize(125, 30)
        # self.btn_output_path.clicked.connect(self.output_dialog)
        # 开始按钮
        self.btn = QPushButton('开始', self)
        # 创建按钮并移动
        self.btn.move(150, 310)
        self.btn.setFixedSize(200, 40)
        # 点击按钮，连接事件函数
        self.btn.clicked.connect(self.btn_action)

        # todo 标签
        # 文件路径标签
        self.lab_select_path = QLabel('', self)
        self.lab_select_path.move(150, 230)
        self.lab_select_path.setFixedSize(320, 30)
        self.lab_select_path.setFrameShape(QFrame.Box)
        self.lab_select_path.setFrameShadow(QFrame.Raised)

        # 说明标签
        content = '说明:\n    选择docx文件位置，程序将自动识别文中的参考文献，并以批注的方式查找参考文献缩写。'
        self.description_lab = QLabel(content, self)
        self.description_lab.move(20, 10)
        self.description_lab.setFixedSize(450, 100)
        self.description_lab.setAlignment(Qt.AlignTop)
        self.description_lab.setFrameShape(QFrame.Box)
        self.description_lab.setFrameShadow(QFrame.Raised)
        # 自动换行
        self.description_lab.setWordWrap(True)

        self.step = 0

        # 显示
        self.show()

    # 按钮点击
    def btn_action(self):
        if self.btn.text() == '完成':
            self.close()
        else:
            docx_path = '{}'.format(self.lab_select_path.text())

            if self.btn.text() == '开始':
                if not os.path.exists(docx_path):
                    QMessageBox.warning(self, '', '请选择docx路径', QMessageBox.Yes)
                else:
                    self.btn.setText('程序进行中')
                    self.downloadThread = downloadThread(docx_path)
                    self.downloadThread.download_proess_signal.connect(self.set_progerss_bar)
                    self.downloadThread.start()

    # 选择输入文件路径
    def file_dialog(self):
        # './'表示当前路径
        path = QFileDialog.getOpenFileName(self, '选取文件', './', 'docx文件(*.docx)')
        print(path)
        # 标签框显示文本路径
        self.lab_select_path.setText(path[0])
        # 自动调整标签框大小
        self.lab_select_path.adjustSize()

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
            QMessageBox.information(self, "提示", "批注完成！")
            return



if __name__ == '__main__':

    #print(fitz.VersionBind)
    app = QApplication(sys.argv)
    pbar = MainUI()
    sys.exit(app.exec_())