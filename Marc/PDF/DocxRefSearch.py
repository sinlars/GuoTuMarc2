# uncompyle6 version 3.7.4
# Python bytecode 3.6 (3379)
# Decompiled from: Python 3.6.2 (v3.6.2:5fd33b5, Jul  8 2017, 04:57:36) [MSC v.1900 64 bit (AMD64)]
# Embedded file name: DocxRefSearch.py
import os, sys
from PyQt5.QtWidgets import QApplication, QProgressBar, QWidget, QPushButton, QFileDialog, QLabel, QFrame, QMessageBox
from PyQt5.QtCore import Qt
from PyQt5.QtCore import QThread, pyqtSignal
import win32com
from win32com.client import Dispatch
import itertools
from docx import Document

class MainUI(QWidget):

    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setGeometry(500, 200, 350, 200)
        self.setFixedSize(520, 370)
        self.setWindowTitle('docx�вο�����ʶ��')
        self.pbar = QProgressBar(self)
        self.pbar.setGeometry(20, 120, 470, 10)
        self.btn_select_file = QPushButton('ѡ��docx�ļ�', self)
        self.btn_select_file.move(20, 230)
        self.btn_select_file.setFixedSize(125, 30)
        self.btn_select_file.clicked.connect(self.file_dialog)
        self.btn = QPushButton('��ʼ', self)
        self.btn.move(150, 310)
        self.btn.setFixedSize(200, 40)
        self.btn.clicked.connect(self.btn_action)
        self.lab_select_path = QLabel('', self)
        self.lab_select_path.move(150, 230)
        self.lab_select_path.setFixedSize(320, 30)
        self.lab_select_path.setFrameShape(QFrame.Box)
        self.lab_select_path.setFrameShadow(QFrame.Raised)
        content = '˵��:\n    ѡ��docx�ļ�λ�ã������Զ�ʶ�����еĲο����ף�������ע�ķ�ʽ���Ҳο�������д��'
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
        if self.btn.text() == '���':
            self.close()
        else:
            docx_path = '{}'.format(self.lab_select_path.text())
            if self.btn.text() == '��ʼ':
                os.path.exists(docx_path) or QMessageBox.warning(self, '', '��ѡ��docx·��', QMessageBox.Yes)
            else:
                self.btn.setText('���������')
                self.downloadThread = downloadThread(docx_path)
                self.downloadThread.download_proess_signal.connect(self.set_progerss_bar)
                self.downloadThread.start()

    def file_dialog(self):
        path = QFileDialog.getOpenFileName(self, 'ѡȡ�ļ�', './', 'docx�ļ�(*.docx)')
        print(path)
        self.lab_select_path.setText(path[0])
        self.lab_select_path.adjustSize()

    def set_progerss_bar(self, num):
        """
        ���ý�ͼ������
        :param num: ���������ȣ�������
        :return:
        """
        self.step = num
        self.pbar.setValue(self.step)
        if num == 100:
            self.btn.setText('���')
            QMessageBox.information(self, '��ʾ', '��ע��ɣ�')
            return


class downloadThread(QThread):
    download_proess_signal = pyqtSignal(int)

    def __init__(self, docx_path):
        super(downloadThread, self).__init__()
        self.docx_path = docx_path

    def emit(self, num):
        self.download_proess_signal.emit(num)

    def run(self):
        self.download_proess_signal.emit(int(1))
        try:
            word = win32com.client.DispatchEx('word.Application')
            word.Visible = 0
            word.DisplayAlerts = 0
            doc = word.Documents.Open(FileName=(self.docx_path), Encoding='gbk')
            dc = DocxContent()
            _ref_list = dc.read(self.docx_path)
            dr = DocxRef()
            _authors_msg = dr.deal_authors_by_ref(_ref_list, self.download_proess_signal)
            _authors_msg1 = dr.deal_authors_merge(_authors_msg, self.download_proess_signal)
            i = 1
            if len(_authors_msg) > 0:
                for author in _authors_msg:
                    print('========================================================')
                    names = author['names']
                    print(names, author['a_names'])
                    for author_name in names:
                        if author_name != '':
                            while word.Selection.Find.Execute(author_name, False, False, True, False):
                                doc.Comments.Add(Range=(word.Selection.Range), Text=(author['ref']))
                                word.Selection.Range.HighlightColorIndex = 4

                        word.Selection.Start = 0
                        word.Selection.End = 0

                    if 'names_err' in author:
                        names_err = author['names_err']
                        for err_name in names_err:
                            while word.Selection.Find.Execute(err_name, False, False, True, False):
                                doc.Comments.Add(Range=(word.Selection.Range), Text=('�˴������д�����˶ԣ�' + author['ref']))
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
                    names = author['names']
                    print(names, author['a_names'])
                    for author_name in names:
                        if author_name != '':
                            while word.Selection.Find.Execute(author_name, False, False, True, False):
                                doc.Comments.Add(Range=(word.Selection.Range), Text=(author['ref']))
                                word.Selection.Range.HighlightColorIndex = 4

                        word.Selection.Start = 0
                        word.Selection.End = 0

                    if 'names_err' in author:
                        names_err = author['names_err']
                        for err_name in names_err:
                            while word.Selection.Find.Execute(err_name, False, False, True, False):
                                doc.Comments.Add(Range=(word.Selection.Range), Text=('�˴������д�����˶ԣ�\r\n' + author['ref']))
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
        except Exception as e:
            doc.Close()
            word.Quit()
            print(e)

        self.download_proess_signal.emit(int(100))


class DocxRef:

    def check_contain_chinese(self, check_str):
        for _char in check_str:
            if 'һ' <= _char <= '��':
                return True

        return False

    def merge_author(self, merge_arr, ref_obj):
        """
        :param merge_arr:
        :param ref_obj:
        :return:
        """
        for obj in merge_arr:
            if obj['author'] == ref_obj['author']:
                obj['year'] = obj['year'].join(ref_obj['year'])
                obj['ref'] = obj['ref'].join(ref_obj['ref'])

        return merge_arr

    def search_equals_names(self, author, comments_msg, num_arr):
        """
        :param authors: ['Minshull T A', ' White R S']
        :param ref_list: ���еĲο������б�
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

    def get_firstname(self, name):
        """
        :param name: �ο���������ȫ��
        :return: �������ߵ�����
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

        return ' '.join(names)

    def strQ2B(self, str):
        """���ַ���ȫ��ת���"""
        ss = []
        for s in str:
            rstring = ''
            for uchar in s:
                inside_code = ord(uchar)
                if inside_code == 12288:
                    inside_code = 32
                else:
                    if inside_code >= 65281:
                        if inside_code <= 65374:
                            inside_code -= 65248
                rstring += chr(inside_code)

            ss.append(rstring)

        return ''.join(ss)

    def create_ref_abbr(self, ref, err_reflist, _comments_mgs):
        """
        :param ref: һ���ο�����
        :return:
        """
        ref_arr = ref.split('.', 2)
        if len(ref_arr) == 3:
            authors = ref_arr[0].strip()
            authors = self.strQ2B(authors)
            authors_arr = authors.split(',')
            authors_arr_new = []
            for author in authors_arr:
                if author.strip().isupper():
                    pass
                else:
                    if len(author.strip()) == 1:
                        pass
                    else:
                        authors_arr_new.append(author)

            authors_arr = authors_arr_new
            year = ref_arr[1].strip()
            if len(authors_arr) == 1:
                _ref_authors = self.ref_abbr(ref, authors_arr, year)
                _comments_mgs.append(_ref_authors)
            else:
                if len(authors_arr) == 2:
                    _ref_authors = self.ref_abbr2(ref, authors_arr, year)
                    _comments_mgs.append(_ref_authors)
                else:
                    if len(authors_arr) > 2:
                        _ref_authors = self.ref_abbr3(ref, authors_arr, year)
                        _comments_mgs.append(_ref_authors)
                    else:
                        print('--------------����������')
        else:
            err_reflist.append(ref)

    def deal_many_authors(self, author, similar_authors, year_split_str='([,���� ]){1,3}'):
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

    def deal_authors_merge(self, _comments_mgs, download_proess_signal):
        num_arr = []
        com_msg = []
        i = 1
        for _author in _comments_mgs:
            similar_author = self.search_equals_names(_author, _comments_mgs, num_arr)
            if len(similar_author) > 0:
                num_arr.append(_comments_mgs.index(_author))
                arr = self.deal_many_authors(_author, similar_author)
                for arr_obj in arr:
                    com_msg.append(arr_obj)

            num = int(i / len(_comments_mgs) * 5)
            if num != 100:
                download_proess_signal.emit(num)
            i = i + 1

        return com_msg

    def deal_authors_by_ref(self, ref_list, download_proess_signal):
        """
        :param ref_list: ���ݲο������б��ȡ�ο�������д
        :return:
        """
        err_refList = []
        _comments_mgs = []
        i = 1
        for ref in ref_list:
            self.create_ref_abbr(ref, err_refList, _comments_mgs)
            num = int(i / len(ref_list) * 5)
            if num != 100:
                download_proess_signal.emit(num)
            i = i + 1

        return _comments_mgs

    def ref_abbr(self, ref, authors_arr, year):
        """
        :param ref: �ο�����ȫ��
        :param authors_arr: #['Brown M M'] һλӢ������ ���� ['�¼���'] һλ��������
        :param year: �ο��������
        :return:
        """
        _ref_authors = {}
        _name = []
        first_name = self.get_firstname(authors_arr[0].strip())
        _name.append(first_name + '([,�� \\(\\��]){1,3}' + year)
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
            _name.append(first_name + '([,��\\(\\�� ]){1,4}��([ ��]){1,3}' + sec_name + '([,��\\(\\�� ]){1,4}' + year)
            _name.append(first_name + '��' + sec_name + '([,��\\(\\�� ]){1,4}' + year)
            _name.append(first_name + '([,��\\(\\�� ]){1,4}��' + sec_name + '([,��\\(\\�� ]){1,4}' + year)
            _name.append(first_name + '��([ ��]){1,3}' + sec_name + '([,��\\(\\�� ]){1,4}' + year)
            _name.append(first_name + '([,��\\(\\�� ]){1,4}and([ ��]){1,3}' + sec_name + '([,��\\(\\�� ]){1,4}' + year)
        else:
            if not self.check_contain_chinese(authors_arr[0].strip()):
                first_name = self.get_firstname(authors_arr[0].strip())
            else:
                first_name = authors_arr[0].strip()
            if not self.check_contain_chinese(authors_arr[1].strip()):
                sec_name = self.get_firstname(authors_arr[1].strip())
            else:
                sec_name = authors_arr[1].strip()
            _name.append(first_name + '([�͡�,�� ]){1,4}' + sec_name + '([,��\\(\\�� ]){1,4}' + year)
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
            _name.append(first_name + '��([,��\\(\\�� ]){1,4}' + year)
            _name.append(first_name + '([,��\\(\\�� ]){1,4}��([,��\\(\\�� ]){1,4}' + year)
            _name.append(first_name + '([,��\\(\\�� ]){1,4}et[ ]{1,3}al([,��\\(\\�� \\.]){1,4}([,��\\(\\�� .]){1,4}' + year)
            _name_err.append(first_name + '([,��\\(\\�� ]){1,4}' + year)
            _name_err.append(first_name + '��([һ-��]{1,2})([,��\\(\\�� ]){1,4}' + year)
        else:
            _name.append(first_name + '��([,��\\(\\�� ]){1,4}' + year)
            _name.append(first_name + '([,��\\(\\�� ]){1,4}��([,��\\(\\�� ]){1,4}' + year)
            _name_err.append(first_name + '([,��\\(\\�� ]){1,4}' + year)
            _name_err.append(first_name + '��([һ-��]{1,2})([,��\\(\\�� ]){1,4}' + year)
        _ref_authors['ref'] = ref
        _ref_authors['names'] = _name
        _ref_authors['a_names'] = authors_arr
        _ref_authors['year'] = year
        _ref_authors['names_err'] = _name_err
        return _ref_authors


class DocxContent:

    def read(self, docx_path):
        document = Document(docx_path)
        ref_str = False
        i = 1
        ref_list = []
        for paragraph in document.paragraphs:
            ref = paragraph.text.strip()
            if ref_str is True:
                if ref != '':
                    print('[' + str(i) + ']��', ref)
                    ref_list.append(ref)
                i += 1
            else:
                if ref != '':
                    ref = ref.replace(' ', '')
                    if ref == '�ο�����':
                        ref_str = True
                else:
                    ref_str = False

        return ref_list


if __name__ == '__main__':
    app = QApplication(sys.argv)
    pbar = MainUI()
    sys.exit(app.exec_())
# okay decompiling F:\Programs\Python\Python36\exe_extracted\DocxRefSearch.pyc
