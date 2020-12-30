import os, shutil
import fitz


class Name_Correcting():
    def __init__(self):
        self.names = get_names('names.txt')
        # self.copy_path = get_copy_path('/data/02/')
        self.copy_path = get_copy_path('H:/0102/')
        # self.copy_path1 = get_create_pdf_path('/data/2020-11-09-01/')
        self.copy_path1 = get_create_pdf_path('H:/2020-11-11-02/')

        self.pdf_dir_path = 'H:/0101/'
        #self.file = open('result.txt', 'a', encoding='utf-8')

    def do_somthing(self):
        with open('result1.txt', 'a+', encoding='utf-8') as file:
            num = 1
            for files in os.listdir(os.chdir(self.pdf_dir_path)):
                if files.endswith('pdf') or files.endswith('PDF'):
                    pdf_path = os.path.join(self.pdf_dir_path, files)
                    try:
                        print('序号{}：{}'.format(num,pdf_path))
                        self.main_logic(pdf_path, file)
                        num += 1
                    except Exception as error:
                        print(error)
                        continue

    def main_logic(self, pdf_path, file):
        try:
            doc = fitz.open(pdf_path)
            doc1 = fitz.open()
            errors = []

            for page in doc:
                new_page = []
                #print('当前页码：{}'.format(page.number))
                for word in self.names:
                    # print(word)
                    try:
                        areas = page.searchFor(word.strip(), quads=True)
                        if len(areas) > 0:
                            # page.addSquigglyAnnot(areas)
                            new_page.append(areas)
                            #print(areas, word.strip(), page.number)
                            errors.append(
                                word.strip() + '_' + str(page.number + 1))
                    except Exception as e:
                        # print(e)
                        continue
                if len(new_page) > 0:
                    # doc.deletePage(page.number)
                    # print('删除页码：{}'.format(page.number))
                    # if len(doc1)<= 0:
                    page1 = doc1.newPage()
                    page1.showPDFpage(page.rect, doc, page.number)
                    for area in new_page:
                        #print(page1.rect, page.rect, area, page.annots())
                        page1.addSquigglyAnnot(area)
                    #print(page1.annots())
                    # doc1.insertPDF(doc, from_page=page.number, to_page=page.number)
                    # doc1.insertPage(page.number)
            if len(errors) > 0:
                file.write(os.path.splitext(os.path.basename(pdf_path))[0] + ' ' + ';'.join(errors))
                file.write('\n')
                file.flush()
                # print(os.path.splitext(os.path.basename(para))[0] + ' ' + ';'.join(errors) +'\n')
            # print(len(doc1), os.path.join(self.copy_path1, os.path.splitext(os.path.basename(para))[0] +'_1.pdf'))
            if len(doc1) > 0:
                doc1.save(os.path.join(self.copy_path1,
                                       os.path.splitext(os.path.basename(pdf_path))[0] + '_1.pdf'))
            doc1.close()
            doc.close()
            shutil.move(pdf_path, self.copy_path)

        except Exception as error:
            print(error)
        finally:
            if not doc1.isClosed:
                doc1.close()
            if not doc.isClosed:
                doc.close()


def get_copy_path(copy_path):
    if os.path.exists(copy_path):
        print('{}文件夹已经存在'.format(copy_path))
    else:
        os.mkdir(copy_path)
        print('{}文件夹已经存在'.format(copy_path))
    return copy_path


def get_create_pdf_path(copy_path1):
    if os.path.exists(copy_path1):
        print('{}文件夹已经存在'.format(copy_path1))
    else:
        os.mkdir(copy_path1)
        print('{}文件夹已经存在'.format(copy_path1))
    return copy_path1


def get_names(name_path):
    txt = open(name_path, 'r', encoding='utf-8')
    # with open('names.txt',r') as txt:
    line = txt.readlines()
    txt.close()
    return line


if __name__ =='__main__':
    # line = get_names('names.txt')
    # print(line)
    obj = Name_Correcting()
    obj.do_somthing()