# encoding: utf-8
import os, sys, win32com, itertools, difflib
from PyQt5.QtWidgets import QApplication, QProgressBar, QWidget, QPushButton, QFileDialog, QLabel, QFrame, QMessageBox
from PyQt5.QtCore import Qt
from PyQt5.QtCore import QThread, pyqtSignal
from win32com.client import Dispatch


class MainUI(QWidget):

    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setGeometry(500, 200, 350, 200)
        self.setFixedSize(520, 370)
        self.setWindowTitle('docx中参考文献识别')
        self.pbar = QProgressBar(self)
        self.pbar.setGeometry(20, 120, 470, 10)
        self.btn_select_file = QPushButton('选择docx文件', self)
        self.btn_select_file.move(20, 230)
        self.btn_select_file.setFixedSize(125, 30)
        self.btn_select_file.clicked.connect(self.file_dialog)
        self.btn = QPushButton('开始', self)
        self.btn.move(150, 310)
        self.btn.setFixedSize(200, 40)
        self.btn.clicked.connect(self.btn_action)
        self.lab_select_path = QLabel('', self)
        self.lab_select_path.move(150, 230)
        self.lab_select_path.setFixedSize(320, 30)
        self.lab_select_path.setFrameShape(QFrame.Box)
        self.lab_select_path.setFrameShadow(QFrame.Raised)
        content = '说明:\n    选择docx文件位置，程序将自动识别文中的参考文献，并以批注的方式查找参考文献缩写。'
        self.description_lab = QLabel(content, self)
        self.description_lab.move(20, 10)
        self.description_lab.setFixedSize(450, 100)
        self.description_lab.setAlignment(Qt.AlignTop)
        self.description_lab.setFrameShape(QFrame.Box)
        self.description_lab.setFrameShadow(QFrame.Raised)
        self.description_lab.setWordWrap(True)
        self.step = 0
        self.show()

    def btn_action(self):
        if self.btn.text() == '完成':
            self.close()
        else:
            docx_path = '{}'.format(self.lab_select_path.text())
            #print(docx_path)
            if not os.path.exists(docx_path):
                QMessageBox.warning(self, '', '请选择docx路径', QMessageBox.Yes)
            else:
                self.btn.setText('程序进行中')
                self.downloadThread = downloadThread(docx_path)
                self.downloadThread.download_proess_signal.connect(self.set_progerss_bar)
                self.downloadThread.start()

    def file_dialog(self):
        path = QFileDialog.getOpenFileName(self, '选取文件', './', 'docx文件(*.docx)')
        print(path)
        self.lab_select_path.setText(path[0])
        self.lab_select_path.adjustSize()

    def set_progerss_bar(self, num):
        """
        设置进图条函数
        :param num: 进度条进度（整数）
        :return:
        """
        self.step = num
        self.pbar.setValue(self.step)
        if num == 100:
            self.btn.setText('完成')
            QMessageBox.information(self, '提示', '批注完成！')
            return


class downloadThread(QThread):

    download_proess_signal = pyqtSignal(int)

    def __init__(self, docx_path):
        super(downloadThread, self).__init__()
        self.docx_path = docx_path

    def word(self):
        word = win32com.client.DispatchEx('word.Application')
        word.Visible = 0
        word.DisplayAlerts = 0
        document = word.Documents.Open(FileName=self.docx_path, Encoding='gbk')
        #temp = word.Selection.Find.Execute(findText, False, False, True, False)
        # while word.Selection.Find.Execute(findText, False, False, True, False):
        #     print(word.Selection.Range.HighlightColorIndex, word.Selection.Range)
        # doc.Close()
        #word.Quit()
        return document

    def getRefs(self, document):
        ref_str = False
        i = 1
        ref_list = []
        for paragraph in document.Paragraphs:
            ref = paragraph.Range.Text.strip()
            if ref_str is True:
                if ref != '':
                    ref_list.append(ref)
                i += 1
            else:
                if ref != '':
                    ref = ref.replace(' ', '')
                    if ref == '参考文献':
                        ref_str = True
                else:
                    ref_str = False
        return ref_list

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
                    elif not n.strip().isupper():
                        names.append(n)
            else:
                break

        return ' '.join(names)

    def check_contain_chinese(self, check_str):
        '''
        :param check_str: 检查是否含有中文字符
        :return:
        '''
        for _char in check_str:
            if '一' <= _char <= '龥':
                return True

        return False


    def create_ref_abbr(self, ref):
        """
        :param ref: 一条参考文献
        :return:
        """
        ref_arr = ref.split('.', 2)
        if len(ref_arr) == 3:
            authors = ref_arr[0].strip()
            authors = authors.replace('，', ',') #将全角的逗号替换半角的逗号
            authors_arr = authors.split(',')
            authors_arr_new = []
            for author in authors_arr:
                if author.strip().isupper(): #字母全是大写，则不是作者
                    pass
                else:
                    if len(author.strip()) == 1:# 只有一个字母，也不是作者
                        pass
                    else:
                        authors_arr_new.append(author)

            authors_arr = authors_arr_new
            year = ref_arr[1].strip()
            #year = filter(lambda en: en in '0123456789', year1)
            if len(authors_arr) == 1:
                ref_author = self.ref_abbr(ref, authors_arr, year)
                return ref_author
            elif len(authors_arr) == 2:
                ref_author = self.ref_abbr2(ref, authors_arr, year)
                return ref_author
            elif len(authors_arr) > 2:
                ref_author = self.ref_abbr3(ref, authors_arr, year)
                return ref_author
            else:
                print('--------------错误，无作者')
                return None
        else:
            return None

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
        _name.append(first_name + '[,， \(\（]{1,3}' + year)
        _ref_authors['ref'] = ref
        _ref_authors['names'] = _name
        _ref_authors['a_names'] = authors_arr
        _ref_authors['year'] = year
        return _ref_authors

    def ref_abbr2(self, ref, authors_arr, year):
        _ref_authors = {}
        _name = []
        first_name = self.get_firstname(authors_arr[0].strip())
        sec_name = self.get_firstname(authors_arr[1].strip())
        if not self.check_contain_chinese(ref):
            _name.append(first_name + '[,，\(\（ ]{1,4}和[ 、]{1,3}' + sec_name + '[,，\(\（ ]{1,4}' + year)
            _name.append(first_name + '和' + sec_name + '[,，\(\（ ]{1,4}' + year)
            _name.append(first_name + '[,，\(\（ ]{1,4}和' + sec_name + '[,，\(\（ ]{1,4}' + year)
            _name.append(first_name + '和[ 、]{1,3}' + sec_name + '[,，\(\（ ]{1,4}' + year)
            _name.append(first_name + '[,，\(\（ ]{1,4}and[ 、]{1,3}' + sec_name + '[,，\(\（ ]{1,4}' + year)
            _name.append(first_name + '[,，\(\（ ]{1,4}et[ 、]{1,3}' + sec_name + '[,，\(\（ ]{1,4}' + year)
            _name.append(first_name + '[,，\(\（ ]{1,4}&[ 、]{1,3}' + sec_name + '[,，\(\（ ]{1,4}' + year)
        else:
            if not self.check_contain_chinese(authors_arr[0].strip()):
                first_name = self.get_firstname(authors_arr[0].strip())
            else:
                first_name = authors_arr[0].strip()
            if not self.check_contain_chinese(authors_arr[1].strip()):
                sec_name = self.get_firstname(authors_arr[1].strip())
            else:
                sec_name = authors_arr[1].strip()
            _name.append(first_name + '[&和、,， ]{1,4}' + sec_name + '[,，\(\（ ]{1,4}' + year)
            #_name.append(first_name + '[和、,， ]{1,4}' + sec_name + '[,，\(\（ ]{1,4}' + year)
        _ref_authors['ref'] = ref
        _ref_authors['names'] = _name
        _ref_authors['a_names'] = authors_arr
        _ref_authors['year'] = year
        return _ref_authors

    def ref_abbr3(self, ref, authors_arr, year):
        _ref_authors = {}
        _name = []
        _name_err = []
        first_name = self.get_firstname(authors_arr[0].strip())
        if not self.check_contain_chinese(ref):
            _name.append(first_name + '等[,，\(\（ ]{1,4}' + year)
            _name.append(first_name + '[,，\(\（ ]{1,4}等[,，\(\（ ]{1,4}' + year)
            _name.append(first_name + '[,，\(\（ ]{1,4}et[ ]{1,3}al[,，\(\（ \.]{1,4}' + year)
            _name_err.append(first_name + '[,，\(\（ ]{1,4}' + year)
            _name_err.append(first_name + '等[一-龥]{1,2}[,，\(\（ ]{1,4}' + year)
        else:
            _name.append(first_name + '等[,，\(\（ ]{1,4}' + year)
            _name.append(first_name + '[,，\(\（ ]{1,4}等[,，\(\（ ]{1,4}' + year)
            _name_err.append(first_name + '[,，\(\（ ]{1,4}' + year)
            _name_err.append(first_name + '等[一-龥]{1,2}[,，\(\（ ]{1,4}' + year)
        _ref_authors['ref'] = ref
        _ref_authors['names'] = _name
        _ref_authors['a_names'] = authors_arr
        _ref_authors['year'] = year
        _ref_authors['names_err'] = _name_err
        return _ref_authors

    def deal_authors_merge(self, _comments_mgs):
        num_arr = []
        com_msg = []
        for _author in _comments_mgs:
            similar_author = self.search_equals_names(_author, _comments_mgs, num_arr)
            if len(similar_author) > 0:
                num_arr.append(_comments_mgs.index(_author))
                arr = self.deal_many_authors(_author, similar_author)
                for arr_obj in arr:
                    com_msg.append(arr_obj)

        return com_msg


    def search_equals_names(self, author, comments_msg, num_arr):
        """
        :param authors: ['Minshull T A', ' White R S']
        :param ref_list: 所有的参考文献列表
        :return:
        """
        merge_arr = []
        cursor_num = comments_msg.index(author)
        for author_obj in comments_msg:
            if len(author['a_names']) <= 2 and len(author_obj['a_names']) <= 2:
                if author['a_names'] == author_obj['a_names'] and cursor_num != comments_msg.index(author_obj) and comments_msg.index(author_obj) not in num_arr:
                    merge_arr.append(author_obj)
                    num_arr.append(comments_msg.index(author_obj))
            else:
                if author['a_names'][0] == author_obj['a_names'][0]:
                    if len(author_obj['a_names']) > 2:
                        if len(author['a_names']) > 2:
                            if cursor_num != comments_msg.index(author_obj):
                                if comments_msg.index(author_obj) not in num_arr:
                                    merge_arr.append(author_obj)
                                    num_arr.append(comments_msg.index(author_obj))

        return merge_arr

    def deal_many_authors(self, author, similar_authors, year_split_str='[,，、 ]{1,3}'):
        similar_authors.insert(0, author)
        author_names = author['a_names']
        year = author['year']
        group_authors_all = []
        comments_msg = []
        for step in range(2, 5):
            iter1 = itertools.combinations(similar_authors, step)
            group_authors = list(list(t1) for t1 in iter1)
            for au in group_authors:
                group_authors_all.append(au)

        if len(author_names) == 1:
            author_name = author_names[0].strip().split()[0]
            for author1 in group_authors_all:
                years_str = year_split_str.join([auth['year'] for auth in author1])
                ref = '\r\n'.join([auth['ref'] for auth in author1])
                ref_author = self.ref_abbr(ref, author_names, years_str)
                comments_msg.append(ref_author)

        else:
            if len(author_names) == 2:
                for author1 in group_authors_all:
                    years_str = year_split_str.join([auth['year'] for auth in author1])
                    ref = '\r\n'.join([auth['ref'] for auth in author1])
                    ref_author = self.ref_abbr2(ref, author_names, years_str)
                    comments_msg.append(ref_author)

            else:
                for author1 in group_authors_all:
                    years_str = year_split_str.join([auth['year'] for auth in author1])
                    ref = '\r\n'.join([auth['ref'] for auth in author1])
                    ref_author = self.ref_abbr3(ref, author_names, years_str)
                    comments_msg.append(ref_author)

        return comments_msg

    def searchNoComments(self, findText, word, document):
        while word.Selection.Find.Execute(findText, False, False, True, False):
            if word.Selection.Range.Comments.Count <= 0:
                document.Comments.Add(Range=word.Selection.Range, Text=('此处疑似有错误，请核对：' + findText))
                word.Selection.Range.HighlightColorIndex = 6
                word.Selection.Range.Underline = 27

    def run(self):
        print(self.docx_path)
        num = 1
        self.download_proess_signal.emit(int(num))
        word = win32com.client.DispatchEx('word.Application')
        word.Visible = 0
        word.DisplayAlerts = 0
        document = word.Documents.Open(FileName=self.docx_path, Encoding='gbk')
        try:
            refs = self.getRefs(document)
            i = 1
            comments_msg = []
            for ref in refs:
                author = self.create_ref_abbr(ref)
                if author is not None:
                    comments_msg.append(author)
            merge_authors = self.deal_authors_merge(comments_msg) #多条合并之后的参考文献
            j = 1
            if len(merge_authors) > 0:
                for author in merge_authors:
                    names = author['names']
                    print(names, author['a_names'])
                    for author_name in names:
                        if author_name != '':
                            while word.Selection.Find.Execute(author_name, False, False, True, False):
                                #if word.Selection.Range.Comments.Count <= 0:
                                print(word.Selection.Range)
                                document.Comments.Add(Range=word.Selection.Range, Text=(author['ref']))
                                word.Selection.Range.HighlightColorIndex = 4

                        word.Selection.Start = 0
                        word.Selection.End = 0

                    if 'names_err' in author:
                        names_err = author['names_err']
                        for err_name in names_err:
                            while word.Selection.Find.Execute(err_name, False, False, True, False):
                                if word.Selection.Range.Comments.Count <= 0:#没有批注的，再增加批注
                                    document.Comments.Add(Range=word.Selection.Range, Text=('此处疑似有错误，请核对：\r\n' + author['ref']))
                                    word.Selection.Range.HighlightColorIndex = 6
                                    word.Selection.Range.Underline = 27

                            word.Selection.Start = 0
                            word.Selection.End = 0

                    num = num + int(j/len(merge_authors) * 49)
                    if num < 100: self.download_proess_signal.emit(num)
                    j += 1
            for author in comments_msg:
                names = author['names']
                print(names, author['a_names'])
                for author_name in names:
                    if author_name != '':
                        while word.Selection.Find.Execute(author_name, False, False, True, False):
                            document.Comments.Add(Range=word.Selection.Range, Text=author['ref'])
                            word.Selection.Range.HighlightColorIndex = 4

                    word.Selection.Start = 0
                    word.Selection.End = 0

                if 'names_err' in author:
                    names_err = author['names_err']
                    for err_name in names_err:
                        while word.Selection.Find.Execute(err_name, False, False, True, False):
                            if word.Selection.Range.Comments.Count <= 0:  # 没有批注的，再增加批注
                                document.Comments.Add(Range=word.Selection.Range, Text=('此处疑似有错误，请核对：' + author['ref']))
                                word.Selection.Range.HighlightColorIndex = 6
                                word.Selection.Range.Underline = 27

                        word.Selection.Start = 0
                        word.Selection.End = 0
                num = i/len(refs) * 49
                if num < 100: self.download_proess_signal.emit(int(num))
                i += 1

            findText = '\([a-zA-Z]{2,}[!a-zA-Z0-9]{1,}[0-9]{3,}\)'
            while word.Selection.Find.Execute(findText, False, False, True, False):
                searchText = word.Selection.Range.Text
                print(searchText)
                if word.Selection.Range.Comments.Count <= 0:
                    dif = 0
                    str = ''
                    for author in comments_msg:
                        if len(author['a_names']) > 1: continue
                        similar_str = self.get_firstname(author['a_names'][0]) + author['year']
                        s = difflib.SequenceMatcher(None, searchText, similar_str).ratio()
                        if s > dif:
                            dif = s
                            str = author['ref']
                    document.Comments.Add(Range=word.Selection.Range, Text=('此处疑似有错误，请核对最近接近的参考文献：' + str))
                    word.Selection.Range.HighlightColorIndex = 6
                    word.Selection.Range.Underline = 27
            #word.Selection.Start = 0
            #word.Selection.End = 0
        except:
            pass
        finally:
            document.Close()
            self.download_proess_signal.emit(int(100))

if __name__ == '__main__':
    app = QApplication(sys.argv)
    pbar = MainUI()
    sys.exit(app.exec_())
    # findText = 'Chappell et al.,'
    # word = win32com.client.DispatchEx('word.Application')
    # word.Visible = 0
    # word.DisplayAlerts = 0
    # doc = word.Documents.Open(FileName='C:/Users/dell/Desktop/日地空间物理学81.docx', Encoding='gbk')
    # ref_str = False
    # i = 1
    # ref_list = []
    # for paragraph in doc.Paragraphs:
    #     ref = paragraph.Range.Text.strip()
    #     if ref_str is True:
    #         if ref != '':
    #             print('[' + str(i) + ']：', ref)
    #             ref_list.append(ref)
    #         i += 1
    #     else:
    #         if ref != '':
    #             #print(ref)
    #             ref = ref.replace(' ', '')
    #             if ref == '参考文献':
    #                 ref_str = True
    #         else:
    #             ref_str = False
    #
    # # temp = word.Selection.Find.Execute(findText, False, False, True, False)
    # # while word.Selection.Find.Execute(findText, False, False, True, False):
    # #     print(word.Selection.Range.HighlightColorIndex, word.Selection.Range, word.Selection.Range.Comments.Count)
    # #doc.Save(True, 1)
    # doc.Close()
    #word.Quit()