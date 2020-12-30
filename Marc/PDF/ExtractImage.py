# encoding: utf-8

import fitz, os

def get_page(pdf_path):
    doc = fitz.open(pdf_path)
    page_no = 67
    pic_path = 'F:\zsdx'
    page = doc.loadPage(page_no)
    #images = page.getImageList()
    print(doc.xrefObject(378))
    print(doc.getPageImageList(page_no, full=True))
    print(doc.pageXref(page_no))
    # i = 1
    # for obj in page.getText('dict')['blocks']:
    #     if obj['type'] == 0:
    #         print(obj['bbox'])
    #         print(obj['lines'])
    #     mat = fitz.Matrix(4, 4)
    #     pix = page.getPixmap(matrix=mat, clip = obj['bbox'])
    #     image_name = os.path.join(pic_path, "P{}_{}.{}".format(page.number + 1, i, 'png'))
    #     pix.writeImage(image_name)
    #     i += 1
    #
    # for item in images:
    #     print(item)
    #     img_obj = doc.extractImage(item[0])
    #     image_name = os.path.join(pic_path, "P{}_{}.{}".format(page.number + 1,'01', img_obj['ext']))
    #     pix = fitz.Pixmap(doc, item[0])
    #     #pix1 = fitz.Pixmap(doc, item[1])
    #     print(pix.n, pix.alpha)
    #     if pix.n - pix.alpha < 4:
    #         #pix2 = pix.Pixmap(pix)
    #         #pix2.setAlpha(pix1.samples)
    #         pix.writeImage(image_name)
    #         #pix2.writeImage(os.path.join(pic_path, "P{}_{}.{}".format(page.number + 1, 2, img_obj['ext'])))
    #     else:  # 否则先转换CMYK
    #         pix0 = fitz.Pixmap(fitz.csRGB, pix)
    #         pix0.writeImage(image_name)
    #         pix0 = None

if __name__ == '__main__':
    get_page('F:\空间分析建模与应用-(1-6).pdf')