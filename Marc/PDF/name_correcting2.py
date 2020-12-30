import os, shutil
import fitz


class Name_Correcting():
    def __init__(self):
        self.names = get_names('names.txt')
        print(self.names)
        # self.copy_path = get_copy_path('/data/02/')
        #self.copy_path = get_copy_path('/Volumes/v1//02/')
        # self.copy_path1 = get_create_pdf_path('/data/2020-11-09-01/')
        self.copy_path1 = get_create_pdf_path('/Volumes/v1/2020-11-12-01/')

        self.pdf_dir_path = '/Volumes/v1/2019/'
        #self.file = open('result.txt', 'a', encoding='utf-8')

    def do_somthing(self):
        with open('/Volumes/v1/result.txt', 'a+', encoding='utf-8') as file:
            for files in os.listdir(os.chdir(self.pdf_dir_path)):
                if files.endswith('pdf') or files.endswith('PDF'):
                    pdf_path = os.path.join(self.pdf_dir_path, files)
                    print(pdf_path)
                    try:
                        self.main_logic(pdf_path, file)
                    except Exception as error:
                        print(error)
                        continue

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
                        # print(e)
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
                doc1.save(os.path.join(self.copy_path1,
                                       os.path.splitext(os.path.basename(pdf_path))[0] + '_1.pdf'))
            doc.close()
            doc1.close()
            #shutil.move(pdf_path, self.copy_path)
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
    txt = open(name_path, 'r')
    # with open('names.txt',r') as txt:
    line = txt.readlines()
    txt.close()
    return line


if __name__ =='__main__':
    # line = get_names('/data/errors/names.txt')
    # print(line)
    obj = Name_Correcting()
    obj.do_somthing()
    # pdf_path = '/Volumes/Seagate/02/B0A2FE5DC74A746C7B39CCCD3F6A01C48000.pdf'
    # doc = fitz.open(pdf_path)
    # doc1 = fitz.open()
    # # for page in doc.pages():
    # #     print(page)
    # # for page in doc:
    # #     print(page)
    # line = get_names('names.txt')
    # print(line)
    # errors = []
    # names = get_names('names.txt')
    # for page_number in range(0,doc.pageCount):
    #     new_page = []
    #     page = doc.loadPage(page_number)
    #     dl = page.getDisplayList()
    #     pt = dl.getTextPage()
    #     for word in line:
    #         #areas = doc.searchPageFor(page_number, word.strip(), quads=True)
    #         areas = pt.search(word.strip(), quads=True)
    #         if len(areas) > 0:
    #             new_page.append(areas)
    #             errors.append(
    #                 word.strip() + '_' + str(page_number + 1))
    #     if len(new_page) > 0:
    #         page1 = doc1.newPage()
    #         page1.showPDFpage(page.rect, doc, page_number)
    #         for area in new_page:
    #             # print(page1.rect, page.rect, area, page.annots())
    #             page1.addSquigglyAnnot(area)
    #     del pt
    #     del dl
    #     del page
    #     print(page_number)
    # if len(errors) > 0:
    #     print(os.path.splitext(os.path.basename(pdf_path))[
    #                    0] + ' ' + ';'.join(errors))
    # if len(doc1) > 0:
    #     doc1.save(os.path.join('/Volumes/Seagate/',
    #                            os.path.splitext(os.path.basename(pdf_path))[
    #                                0] + '_1.pdf'))
    # doc.close()
    # doc1.close()
    #     # page = doc.loadPage(page_number)
    #     # for word in names:
    #     #     # print(word)
    #     #     try:
    #     #         areas = page.searchFor(word.strip(), quads=True)
    #     #         if len(areas) > 0:
    #     #             # page.addSquigglyAnnot(areas)
    #     #             print(areas)
    #     #     except Exception as e:
    #     #         print(e)
    # print(page_number)

    #print(doc.pageCount)