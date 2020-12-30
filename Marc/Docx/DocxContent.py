from docx import Document


class DocxContent():
    
    # def __init__(self, docx_path):
    #     #super(DocxContent, self).__init__()
    #     self.docx_path = docx_path

    def read(self, docx_path):
        #print(docx_path)
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
