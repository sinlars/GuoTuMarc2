import os, sys
from PyQt5.QtWidgets import QApplication, QProgressBar, QWidget, QPushButton, QFileDialog, QLabel, QFrame, QMessageBox
from PyQt5.QtCore import Qt
#from Marc.Docx.DocxThread import downloadThread
from PyQt5.QtCore import QThread, pyqtSignal
import win32com
from win32com.client import Dispatch
import itertools
from docx import Document

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


class downloadThread(QThread):

    download_proess_signal = pyqtSignal(int)  # 创建信号

    def __init__(self, docx_path):
        super(downloadThread, self).__init__()
        self.docx_path = docx_path


    def emit(self, num):
        self.download_proess_signal.emit(num)

    def run(self):
        self.download_proess_signal.emit(int(1))
        try:
            #word = win32com.client.Dispatch('Word.Application')  # 打开word应用程序
            word = win32com.client.DispatchEx('word.Application') # 启动独立的进程
            word.Visible = 0  # 后台运行，不打开程序
            word.DisplayAlerts = 0  # 不警告
            doc = word.Documents.Open(FileName=self.docx_path, Encoding='gbk')
            dc = DocxContent()
            _ref_list = dc.read(self.docx_path)
            dr = DocxRef()
            _authors_msg = dr.deal_authors_by_ref(_ref_list, self.download_proess_signal)
            _authors_msg1 = dr.deal_authors_merge(_authors_msg, self.download_proess_signal)
            i = 1
            if len(_authors_msg) > 0:
                for author in _authors_msg:
                    print('========================================================')
                    #print(author)
                    names = author['names']
                    print(names, author['a_names'])
                    for author_name in names:
                        if author_name != '':
                            #print('author_name = ',author_name)
                            while word.Selection.Find.Execute(author_name, False, False, True, False):
                                doc.Comments.Add(Range=word.Selection.Range, Text=author['ref']) #给选中的文字添加批注
                                word.Selection.Range.HighlightColorIndex = 4
                                #print('word.Selection.Range = ', word.Selection.Range)

                        word.Selection.Start = 0
                        word.Selection.End = 0
                    # print(author['names1'])
                    # for author_name in author['names1']:
                    #     if author_name !='':
                    #         while word.Selection.Find.Execute(author_name):
                    #             doc.Comments.Add(Range=word.Selection.Range, Text=author['ref'])
                    #             word.Selection.Range.HighlightColorIndex = 4
                    if 'names_err' in author:
                        names_err = author['names_err']
                        for err_name in names_err:
                            while word.Selection.Find.Execute(err_name, False, False, True, False):
                                doc.Comments.Add(Range=word.Selection.Range, Text='此处疑似有错误，请核对：' + author['ref'])  # 给选中的文字添加批注
                                word.Selection.Range.HighlightColorIndex = 6
                                word.Selection.Range.Underline = 27
                            word.Selection.Start = 0
                            word.Selection.End = 0



                    num = int(i / (len(_authors_msg) + len(_authors_msg1)) * 90 + 10)
                    if num != 100:
                        self.download_proess_signal.emit(num)
                    i = i + 1
            if len(_authors_msg1) > 0:
                for author in _authors_msg1:
                    print('========================================================')
                    #print(author)
                    names = author['names']
                    print(names, author['a_names'])
                    for author_name in names:
                        if author_name != '':
                            while word.Selection.Find.Execute(author_name, False, False, True, False):
                                doc.Comments.Add(Range=word.Selection.Range, Text=author['ref']) #给选中的文字添加批注
                                word.Selection.Range.HighlightColorIndex = 4
                                #print('word.Selection.Range = ', word.Selection.Range)

                        word.Selection.Start = 0
                        word.Selection.End = 0
                    # names1 = author['names1']
                    # print(names1)
                    # for author_name in names1:
                    #     if author_name != '':
                    #         while word.Selection.Find.Execute(author_name):
                    #             doc.Comments.Add(Range=word.Selection.Range, Text=author['ref']) #给选中的文字添加批注
                    #             word.Selection.Range.HighlightColorIndex = 4
                    #             #print('word.Selection.Range = ', word.Selection.Range)
                    #
                    #     word.Selection.Start = 0
                    #     word.Selection.End = 0
                    # names_err = author['names_err']
                    # for err_name in names_err:
                    #     while word.Selection.Find.Execute(FindText=err_name, MatchWildcards=True):
                    #         doc.Comments.Add(Range=word.Selection.Range, Text=author['ref'])  # 给选中的文字添加批注
                    #         #word.Selection.Range.HighlightColorIndex = 2
                    #         word.Selection.Range.Underline = 27
                    #     word.Selection.Start = 0
                    #     word.Selection.End = 0
                    if 'names_err' in author:
                        names_err = author['names_err']
                        for err_name in names_err:
                            while word.Selection.Find.Execute(err_name, False, False, True, False):
                                doc.Comments.Add(Range=word.Selection.Range, Text='此处疑似有错误，请核对：\r\n' + author['ref'])  # 给选中的文字添加批注
                                word.Selection.Range.HighlightColorIndex = 6
                                word.Selection.Range.Underline = 27
                            word.Selection.Start = 0
                            word.Selection.End = 0



                    num = int(i / (len(_authors_msg1) + len(_authors_msg)) * 90 + 10)
                    if num != 100:
                        self.download_proess_signal.emit(num)
                    i = i + 1

            doc.Close()
            word.Quit()

            #self.download_proess_signal.emit(int(100))

        except Exception as e:
            doc.Close()
            word.Quit()
            print(e)

        self.download_proess_signal.emit(int(100))

class DocxRef:

    #docxThread  = downloadThread()

    def check_contain_chinese(self, check_str):  # 判断字符串是否含有中文
        for _char in check_str:
            if '\u4e00' <= _char <= '\u9fa5': return True
        return False

    def merge_author(self, merge_arr, ref_obj):
        '''
        :param merge_arr:
        :param ref_obj:
        :return:
        '''
        for obj in merge_arr:
            if obj['author'] == ref_obj['author']:
                obj['year'] = obj['year'].join(ref_obj['year'])
                obj['ref'] = obj['ref'].join(ref_obj['ref'])

        return merge_arr

    def search_equals_names(self, author, comments_msg, num_arr):
        '''
        :param authors: ['Minshull T A', ' White R S']
        :param ref_list: 所有的参考文献列表
        :return:
        '''
        merge_arr = []
        cursor_num = comments_msg.index(author)
        #print("当前对象的下标是：", cursor_num)
        for author_obj in comments_msg:
            if len(author['a_names']) <= 2 and len(author_obj['a_names']) <= 2: #作者数小于等于2，必须完全相同
                if author['a_names'] == author_obj['a_names'] :
                    if cursor_num != comments_msg.index(author_obj):
                        if comments_msg.index(author_obj) not in num_arr:
                            merge_arr.append(author_obj)
                            num_arr.append(comments_msg.index(author_obj))
            else:
                if author['a_names'][0] == author_obj['a_names'][0] and len(author_obj['a_names']) > 2 and len(author['a_names']) > 2:
                    if cursor_num != comments_msg.index(author_obj):
                        if comments_msg.index(author_obj) not in num_arr:
                            merge_arr.append(author_obj)
                            num_arr.append(comments_msg.index(author_obj))
        return merge_arr

    def get_firstname(self, name):
        """
        :param name: 参考文献作者全名
        :return: 返回作者的姓氏
        """
        name_arr = name.split()
        names = []
        for n in name_arr:
            if len(n.strip()) > 1:
                if self.check_contain_chinese(n.strip()):
                    names.append(n)
                else:
                    if n.strip().istitle():
                        names.append(n)
                    else:
                        if not n.strip().isupper():
                            names.append(n)
        return ' '.join(names)


    def strQ2B(self, str):
        """把字符串全角转半角"""
        ss = []
        for s in str:
            rstring = ""
            for uchar in s:
                inside_code = ord(uchar)
                if inside_code == 12288:  # 全角空格直接转换
                    inside_code = 32
                elif (inside_code >= 65281 and inside_code <= 65374):  # 全角字符（除空格）根据关系转化
                    inside_code -= 65248
                rstring += chr(inside_code)
            ss.append(rstring)
        return ''.join(ss)

    def create_ref_abbr(self, ref, err_reflist, _comments_mgs):
        """
        :param ref: 一条参考文献
        :return:
        """
        ref_arr = ref.split('.', 2)
        if len(ref_arr) == 3:
            authors = ref_arr[0].strip()
            authors = self.strQ2B(authors)
            authors_arr = authors.split(',')
            authors_arr_new = []
            for author in authors_arr:
                if author.strip().isupper(): continue
                if len(author.strip()) == 1: continue
                authors_arr_new.append(author)
            authors_arr = authors_arr_new
            year = ref_arr[1].strip()
            #print(ref_arr, authors, authors_arr)
            if len(authors_arr) == 1:  # 一位英文作者的情况
                _ref_authors = self.ref_abbr(ref, authors_arr, year)
                _comments_mgs.append(_ref_authors)
            elif len(authors_arr) == 2:  # 两位作者的情况
                _ref_authors = self.ref_abbr2(ref, authors_arr, year)
                _comments_mgs.append(_ref_authors)
            elif len(authors_arr) > 2:  # 多位作者的情况
                _ref_authors = self.ref_abbr3(ref, authors_arr, year)
                _comments_mgs.append(_ref_authors)
            else:
                print('--------------错误，无作者')

        else:
            err_reflist.append(ref)


    def deal_many_authors(self, author, similar_authors, year_split_str='([,，、 ]){1,3}'):
        similar_authors.insert(0, author)
        author_names = author['a_names']
        year = author['year']
        group_authors_all = []
        comments_msg = []
        for step in range(2, 5):  #len(similar_authors) + 1
            iter1 = itertools.combinations(similar_authors, step)
            group_authors = list(list(t1) for t1 in iter1)
            for au in group_authors:
                group_authors_all.append(au)
        #print(group_authors_all)
        if len(author_names) == 1:
            author_name = author_names[0].strip().split()[0]
            for author1 in  group_authors_all:
                years_str = year_split_str.join([auth["year"]for auth in author1])
                ref = '\r\n'.join([auth["ref"] for auth in author1])
                ref_author = self.ref_abbr(ref, author_names, years_str)
                comments_msg.append(ref_author)
        elif len(author_names) == 2:
            for author1 in  group_authors_all:
                years_str = year_split_str.join([auth["year"]for auth in author1])
                ref = '\r\n'.join([auth["ref"] for auth in author1])
                ref_author = self.ref_abbr2(ref, author_names, years_str)
                comments_msg.append(ref_author)
            #print(f'{author_names}')
        else:
            for author1 in  group_authors_all:
                years_str = year_split_str.join([auth["year"]for auth in author1])
                ref = '\r\n'.join([auth["ref"] for auth in author1])
                ref_author = self.ref_abbr3(ref, author_names, years_str)
                comments_msg.append(ref_author)
        #print(comments_msg)
        return comments_msg

    def deal_authors_merge(self, _comments_mgs, download_proess_signal):
        num_arr = []
        com_msg = []
        i = 1
        for _author in _comments_mgs:
            similar_author = self.search_equals_names(_author, _comments_mgs, num_arr)
            #print(similar_author)
            if len(similar_author) > 0:
                num_arr.append(_comments_mgs.index(_author))
                arr = self.deal_many_authors(_author, similar_author)
                #print('arr:==========',arr)
                for arr_obj in arr:
                    com_msg.append(arr_obj)

            num = int(i / len(_comments_mgs) * 5)
            if num != 100:
                download_proess_signal.emit(num)

            i = i + 1

        return com_msg

    def deal_authors_by_ref(self, ref_list, download_proess_signal):
        '''
        :param ref_list: 根据参考文献列表获取参考文献缩写
        :return:
        '''
        err_refList = []  # 记录错误符号的参考文献
        _comments_mgs = []
        i = 1
        for ref in ref_list:
            self.create_ref_abbr(ref, err_refList, _comments_mgs)

            num = int(i / len(ref_list) * 5)
            if num != 100:
                download_proess_signal.emit(num)
            i = i + 1
        return _comments_mgs

    #@staticmethod
    def ref_abbr(self, ref, authors_arr, year):
        """
        :param ref: 参考文献全文
        :param authors_arr: #['Brown M M'] 一位英文作者 或者 ['陈吉余'] 一位中文作者
        :param year: 参考文献年份
        :return:
        """
        _ref_authors = {}
        _name = []
        first_name = self.get_firstname(authors_arr[0].strip())
        _name.append(first_name + '([,， \(\（]){1,3}' + year)

        # _name.append(first_name + '（' + year + "）")
        # _name.append(first_name + '(' + year + ")")
        # _name.append(first_name + '' + year)
        # _name.append(first_name + ',' + year)
        # _name_err = []
        # _name_err.append(first_name +' {1,}' + '（' + year + "）")
        # _name_err.append(first_name + ' {1,}' + '（ {1,}' + year + "）")
        # _name_err.append(first_name + ' {1,}' + '（ {1,}' + year + " {1,}）")
        #
        # _name_err.append(first_name + ' {1,}' + '(' + year + ")")
        # _name_err.append(first_name + ' {1,}' + '( {1,}' + year + ")")
        # _name_err.append(first_name + ' {1,}' + '( {1,}' + year + " {1,})")
        #
        # _name_err.append(first_name + ' {1,}，' + year)
        # _name_err.append(first_name + ' {1,}， {1,}' + year)
        #
        # _name_err.append(first_name + ' {1,},' + year)
        # _name_err.append(first_name + ' {1,}, {1,}' + year)
        #
        # _name1 = []
        # _name1.append(first_name + ' ' + '（' + year + "）")
        # _name1.append(first_name + ' ' + '（ ' + year + "）")
        # _name1.append(first_name + ' ' + '（ ' + year + " ）")
        #
        # _name1.append(first_name + ' ' + '(' + year + ")")
        # _name1.append(first_name + ' ' + '( ' + year + ")")
        # _name1.append(first_name + ' ' + '( ' + year + " )")
        #
        # _name1.append(first_name + ' ，' + year)
        # _name1.append(first_name + ' ， ' + year)
        #
        # _name1.append(first_name + ' ,' + year)
        # _name1.append(first_name + ' , ' + year)



        _ref_authors['ref'] = ref
        _ref_authors['names'] = _name
        _ref_authors['a_names'] = authors_arr
        _ref_authors['year'] = year
        #_ref_authors['names_err'] = _name_err
        #_ref_authors['names1'] = _name1
        return _ref_authors

    # '''
    #     authors_arr : #['Brown M'] 一位中文作者
    #     ref_arr:
    # '''
    #
    # def _ref_abbr_cn(self, ref, authors_arr, year):  #
    #     print(authors_arr)
    #     if '，' in authors_arr[0]:
    #         return self._ref_abbr3_cn(ref, authors_arr[0].split('，'), year)
    #     else:
    #
    #         _ref_authors = {}
    #         _name = []
    #         first_name = authors_arr[0].strip().split(' ')[0].strip()
    #         # _name.append(authors_arr[0].strip().split(' ')[0].strip() + '（' + year + "）")
    #         # _name.append(authors_arr[0].strip().split(' ')[0].strip() + '(' + year + ")")
    #         # _name.append(authors_arr[0].strip().split(' ')[0].strip() + '，' + year)
    #         # _name.append(authors_arr[0].strip().split(' ')[0].strip() + ',' + year)
    #
    #         _name.append(first_name + '（' + year + "）")
    #         _name.append(first_name + '(' + year + ")")
    #         _name.append(first_name + '，' + year)
    #         _name.append(first_name + ',' + year)
    #         _name_err = []
    #         _name_err.append(first_name + ' {1,}' + '（' + year + "）")
    #         _name_err.append(first_name + ' {1,}' + '（ {1,}' + year + "）")
    #         _name_err.append(first_name + ' {1,}' + '（ {1,}' + year + " {1,}）")
    #
    #         _name_err.append(first_name + ' {1,}' + '(' + year + ")")
    #         _name_err.append(first_name + ' {1,}' + '( {1,}' + year + ")")
    #         _name_err.append(first_name + ' {1,}' + '( {1,}' + year + " {1,})")
    #
    #         _name_err.append(first_name + ' {1,}，' + year)
    #         _name_err.append(first_name + ' {1,}， {1,}' + year)
    #
    #         _name_err.append(first_name + ' {1,},' + year)
    #         _name_err.append(first_name + ' {1,}, {1,}' + year)
    #
    #         _ref_authors['ref'] = ref
    #         _ref_authors['names'] = _name
    #         _ref_authors['a_names'] = authors_arr
    #         _ref_authors['year'] = year
    #         _ref_authors['names_err'] = _name_err
    #
    #         #print(_ref_authors)
    #         return _ref_authors

    def ref_abbr2(self, ref, authors_arr, year):
        _ref_authors = {}
        _name = []
        #_name_err = []
        #_name1 = []
        first_name = self.get_firstname(authors_arr[0].strip())
        sec_name = self.get_firstname(authors_arr[1].strip())
        if not self.check_contain_chinese(ref):  # 纯英文文献
            _name.append(first_name + '([,，\(\（ ]){1,4}和([ 、]){1,3}' + sec_name + '([,，\(\（ ]){1,4}' + year)
            _name.append(first_name + '和' + sec_name + '([,，\(\（ ]){1,4}' + year)
            _name.append(first_name + '([,，\(\（ ]){1,4}和' + sec_name + '([,，\(\（ ]){1,4}' + year)
            _name.append(first_name + '和([ 、]){1,3}' + sec_name + '([,，\(\（ ]){1,4}' + year)

            _name.append(first_name + '([,，\(\（ ]){1,4}and([ 、]){1,3}' + sec_name + '([,，\(\（ ]){1,4}' + year)
            # _name.append(first_name + '和' + sec_name + '([,，\(\（ ]){1,4}' + year)
            # _name.append(first_name + '和' + sec_name + '（' + year + "）")
            # _name.append(first_name + '和' + sec_name + '(' + year + ")")
            # _name.append(first_name + ' and ' + sec_name + '，' + year)
            # _name.append(first_name + ' and ' + sec_name + ',' + year)
            # _name.append(first_name + ' and ' + sec_name + '(' + year + ')')
            # _name.append(first_name + ' and ' + sec_name + '（' + year + '）')
            #
            # _name_err.append(first_name + ' {1,}和' + sec_name + '（' + year + "）")
            # _name_err.append(first_name + ' {1,}和 {1,}' + sec_name + '（' + year + "）")
            # _name_err.append(first_name + ' {1,}和 {1,}' + sec_name + ' {1,}（' + year + "）")
            # #_name_err.append(first_name + ' {1,}和 {1,}' + sec_name + ' {1,}（' + year + "） {1,}")
            #
            # _name_err.append(first_name + ' {1,}和' + sec_name + '(' + year + ")")
            # _name_err.append(first_name + ' {1,}和 {1,}' + sec_name + '(' + year + ")")
            # _name_err.append(first_name + ' {1,}和 {1,}' + sec_name + ' {1,}(' + year + ")")
            # #_name_err.append(first_name + ' {1,}和 {1,}' + sec_name + ' {1,}(' + year + ") {1,}")
            #
            # _name_err.append(first_name + ' and ' + sec_name + ' {1,}，' + year)
            # _name_err.append(first_name + ' and ' + sec_name + ' {1,}， {1,}' + year)
            #
            # _name_err.append(first_name + ' and ' + sec_name + ' {1,},' + year)
            # _name_err.append(first_name + ' and ' + sec_name + ' {1,}, {1,}' + year)


            # _name1.append(first_name + ' 和' + sec_name + '（' + year + "）")
            # _name1.append(first_name + ' 和 ' + sec_name + '（' + year + "）")
            # _name1.append(first_name + ' 和 ' + sec_name + ' （' + year + "）")
            # _name1.append(first_name + ' 和 ' + sec_name + '（ ' + year + "）")
            # _name1.append(first_name + ' 和 ' + sec_name + '（ ' + year + " ）")
            # #_name1.append(first_name + ' 和 ' + sec_name + '（ ' + year + "）")
            # #_name1.append(first_name + ' 和 ' + sec_name + '（' + year + " ）")
            #
            # _name1.append(first_name + ' 和' + sec_name + '(' + year + ")")
            # _name1.append(first_name + ' 和 ' + sec_name + '(' + year + ")")
            # _name1.append(first_name + ' 和 ' + sec_name + ' (' + year + ")")
            # _name1.append(first_name + ' 和 ' + sec_name + '( ' + year + ")")
            # _name1.append(first_name + ' 和 ' + sec_name + '( ' + year + " )")
            #
            # _name1.append(first_name + ' and ' + sec_name + ' ，' + year)
            # _name1.append(first_name + ' and ' + sec_name + ' ， ' + year)
            #
            # _name1.append(first_name + ' and ' + sec_name + ' ,' + year)
            # _name1.append(first_name + ' and ' + sec_name + ' , ' + year)

        else:
            if not self.check_contain_chinese(authors_arr[0].strip()):
                first_name = self.get_firstname(authors_arr[0].strip())
            else:
                first_name = authors_arr[0].strip()

            if not self.check_contain_chinese(authors_arr[1].strip()):
                sec_name = self.get_firstname(authors_arr[1].strip())
            else:
                sec_name = authors_arr[1].strip()

            _name.append(first_name + '([和、,， ]){1,4}' + sec_name + '([,，\(\（ ]){1,4}' + year)

            # _name.append(first_name + '和' + sec_name + '（' + year + "）")
            # _name.append(first_name + '和' + sec_name + '(' + year + ")")
            # _name.append(first_name + '和' + sec_name + '，' + year)
            # _name.append(first_name + '和' + sec_name + ',' + year)
            # _name.append(first_name + '、' + sec_name + '，' + year)
            # _name.append(first_name + '、' + sec_name + ',' + year)
            #
            # _name_err.append(first_name + ' {1,}和' + sec_name + '（' + year + "）")
            # _name_err.append(first_name + ' {1,}和 {1,}' + sec_name + '（' + year + "）")
            # _name_err.append(first_name + ' {1,}和 {1,}' + sec_name + ' {1,}（' + year + "）")
            # #_name_err.append(first_name + ' {1,}和 {1,}' + sec_name + ' {1,}（' + year + "） {1,}")
            #
            # _name_err.append(first_name + ' {1,}和' + sec_name + '(' + year + ")")
            # _name_err.append(first_name + ' {1,}和 {1,}' + sec_name + '(' + year + ")")
            # _name_err.append(first_name + ' {1,}和 {1,}' + sec_name + ' {1,}(' + year + ")")
            # #_name_err.append(first_name + ' {1,}和 {1,}' + sec_name + ' {1,}(' + year + ") {1,}")
            #
            # _name_err.append(first_name + ' {1,}和' + sec_name + '，' + year)
            # _name_err.append(first_name + ' {1,}和 {1,}' + sec_name + '，' + year)
            # _name_err.append(first_name + ' {1,}和 {1,}' + sec_name + ' {1,}，' + year)
            #
            # _name_err.append(first_name + ' {1,}和' + sec_name + ',' + year)
            # _name_err.append(first_name + ' {1,}和 {1,}' + sec_name + ',' + year)
            # _name_err.append(first_name + ' {1,}和 {1,}' + sec_name + ' {1,},' + year)
            #
            # _name_err.append(first_name + ' {1,}、' + sec_name + '，' + year)
            # _name_err.append(first_name + ' {1,}、 {1,}' + sec_name + '，' + year)
            # _name_err.append(first_name + ' {1,}、 {1,}' + sec_name + ' {1,}，' + year)
            #
            # _name_err.append(first_name + ' {1,}、' + sec_name + ',' + year)
            # _name_err.append(first_name + ' {1,}、 {1,}' + sec_name + ',' + year)
            # _name_err.append(first_name + ' {1,}、 {1,}' + sec_name + ' {1,},' + year)
            #
            # _name1.append(first_name + ' 和' + sec_name + '（' + year + "）")
            # _name1.append(first_name + ' 和 ' + sec_name + '（' + year + "）")
            # _name1.append(first_name + ' 和 ' + sec_name + ' （' + year + "）")
            # _name1.append(first_name + ' 和 ' + sec_name + '（ ' + year + "）")
            # _name1.append(first_name + ' 和 ' + sec_name + '（ ' + year + " ）")
            #
            # _name1.append(first_name + ' 和' + sec_name + '(' + year + ")")
            # _name1.append(first_name + ' 和 ' + sec_name + '(' + year + ")")
            # _name1.append(first_name + ' 和 ' + sec_name + ' (' + year + ")")
            # _name1.append(first_name + ' 和 ' + sec_name + '( ' + year + ")")
            # _name1.append(first_name + ' 和 ' + sec_name + '( ' + year + " )")
            #
            # _name1.append(first_name + ' 和' + sec_name + '，' + year)
            # _name1.append(first_name + ' 和 ' + sec_name + '，' + year)
            # _name1.append(first_name + ' 和 ' + sec_name + ' ，' + year)
            # _name1.append(first_name + ' 和 ' + sec_name + ' ， ' + year)
            #
            # _name1.append(first_name + ' 和' + sec_name + ',' + year)
            # _name1.append(first_name + ' 和 ' + sec_name + ',' + year)
            # _name1.append(first_name + ' 和 ' + sec_name + ' ,' + year)
            # _name1.append(first_name + ' 和 ' + sec_name + ' , ' + year)
            #
            # _name1.append(first_name + ' 、' + sec_name + '，' + year)
            # _name1.append(first_name + ' 、 ' + sec_name + '，' + year)
            # _name1.append(first_name + ' 、 ' + sec_name + ' ，' + year)
            # _name1.append(first_name + ' 、 ' + sec_name + ' ， ' + year)
            #
            # _name1.append(first_name + ' 、' + sec_name + ',' + year)
            # _name1.append(first_name + ' 、 ' + sec_name + ',' + year)
            # _name1.append(first_name + ' 、 ' + sec_name + ' ,' + year)
            # _name1.append(first_name + ' 、 ' + sec_name + ' , ' + year)
        _ref_authors['ref'] = ref
        _ref_authors['names'] = _name
        _ref_authors['a_names'] = authors_arr
        _ref_authors['year'] = year
        #_ref_authors['names_err'] = _name_err
        #_ref_authors['names1'] = _name1
        #print(_ref_authors)
        return _ref_authors



    # def _ref_abbr2_cn(self, ref, authors_arr, year):
    #     '''
    #         authors_arr :  ['高原', ' 郑斯华'] 两位中文作者
    #         ref_arr:
    #     '''
    #     _ref_authors = {}
    #     _name = []
    #     first_name = ''
    #     sec_name = ''
    #     if not self.check_contain_chinese(authors_arr[0].strip()):
    #         first_name = authors_arr[0].strip().split(' ')[0].strip()
    #     else:
    #         first_name = authors_arr[0].strip()
    #
    #     if not self.check_contain_chinese(authors_arr[1].strip()):
    #         sec_name = authors_arr[1].strip().split(' ')[0].strip()
    #     else:
    #         sec_name = authors_arr[1].strip()
    #
    #     _name.append(first_name + '和' + sec_name + '（' + year + "）")
    #     _name.append(first_name + '和' + sec_name + '(' + year + ")")
    #     _name.append(first_name + '和' + sec_name + '，' + year)
    #     _name.append(first_name + '和' + sec_name + ',' + year)
    #     _name.append(first_name + '、' + sec_name + '，' + year)
    #     _name.append(first_name + '、' + sec_name + ',' + year)
    #
    #     _ref_authors['ref'] = ref
    #     _ref_authors['names'] = _name
    #     _ref_authors['a_names'] = authors_arr
    #     _ref_authors['year'] = year
    #
    #     #print(_ref_authors)
    #     return _ref_authors

    '''
        authors_arr : ['Minshull T A', ' Muller M R', ' White R S'], ['Minshull T A', ' Muller M R', ' Robinson C J', ' et al'] 多位作者  
        ref_arr: 
    '''

    def ref_abbr3(self, ref, authors_arr, year):  #
        _ref_authors = {}
        _name = []
        _name_err = []
        #_name1 = []
        first_name = self.get_firstname(authors_arr[0].strip())
        if not self.check_contain_chinese(ref):
            _name .append(first_name + '等([,，\(\（ ]){1,4}' + year)
            _name.append(first_name + '([,，\(\（ ]){1,4}等([,，\(\（ ]){1,4}' + year)
            _name.append(first_name + '([,，\(\（ ]){1,4}et[ ]{1,3}al([,，\(\（ \.]){1,4}([,，\(\（ .]){1,4}' + year)
            _name_err.append(first_name + '([,，\(\（ ]){1,4}' + year)
            _name_err.append(first_name + '等([一-龥]{1,2})([,，\(\（ ]){1,4}' + year)
            # _name.append(first_name + '等' + '（' + year + "）")
            # _name.append(first_name + '等' + '(' + year + ")")
            # _name.append(first_name + ' et al.' + '，' + year)
            # _name.append(first_name + ' et al.' + ',' + year)
            #
            # _name_err.append(first_name + ' {1,}等' + '（' + year + "）")
            # _name_err.append(first_name + ' {1,}等 {1,}' + '（' + year + "）")
            # _name_err.append(first_name + ' {1,}等 {1,}' + '（ {1,}' + year + "）")
            # _name_err.append(first_name + ' {1,}等 {1,}' + '（ {1,}' + year + " {1,}）")
            #
            # _name_err.append(first_name + ' {1,}等' + '(' + year + ")")
            # _name_err.append(first_name + ' {1,}等 {1,}' + '(' + year + ")")
            # _name_err.append(first_name + ' {1,}等 {1,}' + '( {1,}' + year + ")")
            # _name_err.append(first_name + ' {1,}等 {1,}' + '( {1,}' + year + " {1,})")
            #
            # _name_err.append(first_name + '  {1,}et al.' + '，' + year)
            # _name_err.append(first_name + '  {1,}et al {1,}.' + '，' + year)
            # _name_err.append(first_name + '  {1,}et al {1,}. {1,}' + '，' + year)
            # _name_err.append(first_name + '  {1,}et al {1,}. {1,}' + '， {1,}' + year)
            #
            # _name_err.append(first_name + '  {1,}et al.' + ',' + year)
            # _name_err.append(first_name + '  {1,}et al {1,}.' + ',' + year)
            # _name_err.append(first_name + '  {1,}et al {1,}. {1,}' + ',' + year)
            # _name_err.append(first_name + '  {1,}et al {1,}. {1,}' + ', {1,}' + year)
            #
            # _name1.append(first_name + ' 等' + '（' + year + "）")
            # _name1.append(first_name + ' 等 ' + '（' + year + "）")
            # _name1.append(first_name + ' 等 ' + '（ ' + year + "）")
            # _name1.append(first_name + ' 等 ' + '（ ' + year + " ）")
            #
            # _name1.append(first_name + ' 等' + '(' + year + ")")
            # _name1.append(first_name + ' 等 ' + '(' + year + ")")
            # _name1.append(first_name + ' 等 ' + '( ' + year + ")")
            # _name1.append(first_name + ' 等 ' + '( ' + year + " )")
            #
            # _name1.append(first_name + '  et al.' + '，' + year)
            # _name1.append(first_name + '  et al .' + '，' + year)
            # _name1.append(first_name + '  et al . ' + '，' + year)
            # _name1.append(first_name + '  et al . ' + '， ' + year)
            #
            # _name1.append(first_name + '  et al.' + ',' + year)
            # _name1.append(first_name + '  et al .' + ',' + year)
            # _name1.append(first_name + '  et al . ' + ',' + year)
            # _name1.append(first_name + '  et al . ' + ', ' + year)

        else:
            _name.append(first_name + '等([,，\(\（ ]){1,4}' + year)
            _name.append(first_name + '([,，\(\（ ]){1,4}等([,，\(\（ ]){1,4}' + year)
            _name_err.append(first_name + '([,，\(\（ ]){1,4}' + year)
            _name_err.append(first_name + '等([一-龥]{1,2})([,，\(\（ ]){1,4}' + year)
            # _name.append(first_name + '等' + '（' + year + "）")
            # _name.append(first_name + '等' + '(' + year + ")")
            # _name.append(first_name + '等' + '，' + year)
            # _name.append(first_name + '等' + ',' + year)
            #
            # _name_err.append(first_name + ' {1,}等' + '（' + year + "）")
            # _name_err.append(first_name + ' {1,}等 {1,}' + '（' + year + "）")
            # _name_err.append(first_name + ' {1,}等 {1,}' + '（ {1,}' + year + "）")
            # _name_err.append(first_name + ' {1,}等 {1,}' + '（ {1,}' + year + " {1,}）")
            #
            # _name_err.append(first_name + ' {1,}等' + '(' + year + ")")
            # _name_err.append(first_name + ' {1,}等 {1,}' + '(' + year + ")")
            # _name_err.append(first_name + ' {1,}等 {1,}' + '( {1,}' + year + ")")
            # _name_err.append(first_name + ' {1,}等 {1,}' + '( {1,}' + year + " {1,})")
            #
            # _name_err.append(first_name + ' {1,}等' + '，' + year)
            # _name_err.append(first_name + ' {1,}等 {1,}' + '，' + year)
            # _name_err.append(first_name + ' {1,}等 {1,}' + '， {1,}' + year)
            # #_name_err.append(first_name + ' {1,}等 {1,}' + '， {1,}' + year)
            #
            # _name_err.append(first_name + ' {1,}等' + ',' + year)
            # _name_err.append(first_name + ' {1,}等 {1,}' + ',' + year)
            # _name_err.append(first_name + ' {1,}等 {1,}' + ', {1,}' + year)
            # #_name_err.append(first_name + ' {1,}等 {1,}' + '（ {1,}' + year + " {1,}）")
            # _name1.append(first_name + ' 等' + '（' + year + "）")
            # _name1.append(first_name + ' 等 ' + '（' + year + "）")
            # _name1.append(first_name + ' 等 ' + '（ ' + year + "）")
            # _name1.append(first_name + ' 等 ' + '（ ' + year + " ）")
            #
            # _name1.append(first_name + ' 等' + '(' + year + ")")
            # _name1.append(first_name + ' 等 ' + '(' + year + ")")
            # _name1.append(first_name + ' 等 ' + '( ' + year + ")")
            # _name1.append(first_name + ' 等 ' + '( ' + year + " )")
            #
            # _name1.append(first_name + ' 等' + '，' + year)
            # _name1.append(first_name + ' 等 ' + '，' + year)
            # _name1.append(first_name + ' 等 ' + '， ' + year)
            # # _name_err.append(first_name + ' {1,}等 {1,}' + '， {1,}' + year)
            #
            # _name1.append(first_name + ' 等' + ',' + year)
            # _name1.append(first_name + ' 等 ' + ',' + year)
            # _name1.append(first_name + ' 等 ' + ', ' + year)
            # _name_err.append(first_name + ' {1,}等 {1,}' + '（ {1,}' + year + " {1,}）")


        _ref_authors['ref'] = ref
        _ref_authors['names'] = _name
        _ref_authors['a_names'] = authors_arr
        _ref_authors['year'] = year
        _ref_authors['names_err'] = _name_err
        #_ref_authors['names1'] = _name1
        return _ref_authors


class DocxContent():

    # def __init__(self, docx_path):
    #     #super(DocxContent, self).__init__()
    #     self.docx_path = docx_path

    def read(self, docx_path):
        # print(docx_path)
        document = Document(docx_path)
        ref_str = False
        i = 1
        ref_list = []
        for paragraph in document.paragraphs:
            ref = paragraph.text.strip()
            if ref_str is True:
                if ref != '':
                    print('[' + str(i) + ']：', ref)
                    ref_list.append(ref)
                i += 1

            if ref != '':
                ref = ref.replace(' ', '')
                if ref == '参考文献':
                    ref_str = True
            else:
                ref_str = False
        return ref_list


if __name__ == '__main__':

    #print(fitz.VersionBind)
    app = QApplication(sys.argv)
    pbar = MainUI()
    sys.exit(app.exec_())