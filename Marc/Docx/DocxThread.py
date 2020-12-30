from Marc.Docx.DocxContent import DocxContent
from Marc.Docx.DocxRef import DocxRef
from PyQt5.QtCore import QThread, pyqtSignal
import win32com
from win32com.client import Dispatch


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
            #print(_ref_list)
            dr = DocxRef()
            _authors_msg = dr.deal_authors_by_ref(_ref_list, self.download_proess_signal)
            _authors_msg1 = dr.deal_authors_merge(_authors_msg,self.download_proess_signal)
            # for author in _authors_msg:
            #     print(author['names'],author['a_names'])
            #     print(author['ref'])
            # for author in _authors_msg1:
            #     print(author['names'])
            #     print(author['names_err'])
            i = 1
            if len(_authors_msg) > 0:
                for author in _authors_msg:
                    print('========================================================')
                    #print(author)
                    names = author['names']
                    print(names)
                    for author_name in names:
                        if author_name != '':
                            while word.Selection.Find.Execute(author_name):
                                doc.Comments.Add(Range=word.Selection.Range, Text=author['ref']) #给选中的文字添加批注
                                word.Selection.Range.HighlightColorIndex = 4
                                #print('word.Selection.Range = ', word.Selection.Range)

                        word.Selection.Start = 0
                        word.Selection.End = 0
                    print(author['names1'])
                    for author_name in author['names1']:
                        if author_name !='':
                            while word.Selection.Find.Execute(author_name):
                                doc.Comments.Add(Range=word.Selection.Range, Text=author['ref'])
                                word.Selection.Range.HighlightColorIndex = 4
                    # names_err = author['names_err']
                    # for err_name in names_err:
                    #     while word.Selection.Find.Execute(FindText=err_name, MatchWildcards=True):
                    #         doc.Comments.Add(Range=word.Selection.Range, Text=author['ref'])  # 给选中的文字添加批注
                    #         #word.Selection.Range.HighlightColorIndex = 2
                    #         word.Selection.Range.Underline = 27
                    #     word.Selection.Start = 0
                    #     word.Selection.End = 0



                    num = int(i / (len(_authors_msg) + len(_authors_msg1)) * 90 + 10)
                    if num != 100:
                        self.download_proess_signal.emit(num)
                    i = i + 1
            if len(_authors_msg1) > 0:
                for author in _authors_msg1:
                    print('========================================================')
                    #print(author)
                    names = author['names']
                    print(names)
                    for author_name in names:
                        if author_name != '':
                            while word.Selection.Find.Execute(author_name):
                                doc.Comments.Add(Range=word.Selection.Range, Text=author['ref']) #给选中的文字添加批注
                                word.Selection.Range.HighlightColorIndex = 4
                                #print('word.Selection.Range = ', word.Selection.Range)

                        word.Selection.Start = 0
                        word.Selection.End = 0
                    names1 = author['names1']
                    print(names1)
                    for author_name in names1:
                        if author_name != '':
                            while word.Selection.Find.Execute(author_name):
                                doc.Comments.Add(Range=word.Selection.Range, Text=author['ref']) #给选中的文字添加批注
                                word.Selection.Range.HighlightColorIndex = 4
                                #print('word.Selection.Range = ', word.Selection.Range)

                        word.Selection.Start = 0
                        word.Selection.End = 0
                    # names_err = author['names_err']
                    # for err_name in names_err:
                    #     while word.Selection.Find.Execute(FindText=err_name, MatchWildcards=True):
                    #         doc.Comments.Add(Range=word.Selection.Range, Text=author['ref'])  # 给选中的文字添加批注
                    #         #word.Selection.Range.HighlightColorIndex = 2
                    #         word.Selection.Range.Underline = 27
                    #     word.Selection.Start = 0
                    #     word.Selection.End = 0



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

if __name__ == '__main__':

    word = win32com.client.DispatchEx('word.Application')  # 启动独立的进程
    word.Visible = 0  # 后台运行，不打开程序
    word.DisplayAlerts = 0  # 不警告
    doc = word.Documents.Open(FileName='C:/Users/dell/Desktop/河口潮波动力学-打印发排-副本.docx', Encoding='gbk')

    while word.Selection.Find.Execute(FindText='Toffolon 和 Savenije（2011）Toffolon和Savenije（2011）', MatchWildcards=True):
        doc.Comments.Add(Range=word.Selection.Range, Text='错误')  # 给选中的文字添加批注
        word.Selection.Range.Underline = 27
    word.Selection.Start = 0
    word.Selection.End = 0
    doc.Close()
    word.Quit()