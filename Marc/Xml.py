#coding:utf8

import xml.etree.cElementTree as et
import io
import xlwt
import sys
import re
import pymysql
import pinyin
import datetime

#改变标准输出的默认编码
#sys.stdout = io.TextIOWrapper(sys.stdout.buffer,encoding='gb18030')


tree = et.ElementTree(file='C:/Users/dell/Desktop//中国维管植物科属词典-内文-标引-20180824_审核.xml')

def subtext(elem):
    for el in elem.iter():
        print('tag: ',el.tag, 'tag attrib :', el.attrib,'text :',el.text)

# for book in tree.iter(tag='book-meta'):
#
#     authors = []
#     for elem in book.iter():
#         if elem.tag == 'book-title':
#             print('书名 ：',elem.text)
#         if elem.tag == 'contrib':
#             names = []
#             for el in elem.iter():
#                 name = ''
#                 if el.tag == 'surname':
#                     #print('姓：',el.text)
#                     name = name + el.text
#                 if el.tag == 'given-names':
#                     #print('名：',el.text)
#                     name = name + el.text
#                 if name != '' :
#                     names.append(name)
#             #print(names)
#             print(elem.attrib['contrib-type'],'：',''.join(names))
#         if elem.tag == 'publisher':
#             for el in elem.iter():
#                 if el.tag == 'publisher-name':
#                     print('出版社：',el.text)
#                 if el.tag == 'address':
#                     print('出版地：',el.text)


    # print('tag: ',elem.tag, 'tag attrib :', elem.attrib,'text :',elem.text)
    # for el in elem.iter():
    #     print('tag: ',el.tag, 'tag attrib :', el.attrib,'text :',el.text)

# for book in tree.iter(tag='book-front'):
#     for elem in book.iter():
#         if elem.tag == 'preface':
#             for el in elem.iter():
#                 if el.tag == 'title':
#                     print(el.text)
#                 if el.tag == 'p':
#                     p = el.text
#                     for e in el:
#                         if e.tag == 'italic':
#                             #print('<i>',e.text,'</i>',e.tail)
#                             p = p + '<i>' + e.text + '</i>' + e.tail
#                     print(p)
#                 if el.tag == 'contrib':
#                     names = []
#                     for e in el.iter():
#                         name = ''
#                         if e.tag == 'surname':
#                             #print('姓：',el.text)
#                             name = name + e.text
#                         if e.tag == 'given-names':
#                             #print('名：',el.text)
#                             name = name + e.text
#                         if name != '' :
#                             names.append(name)
#                         #print(names)
#                     print(el.attrib['contrib-type'],'：',','.join(names))
#                 if el.tag == 'date':
#                     print(el.text)


def get_text(eles,return_str):
    if(len(eles) > 0):
        if(len(eles) == 1):
            return get_text(eles[0], return_str)
        else:
            ret_s = ''
            for ele in eles:
                if(len(ele) > 0):
                    return get_text(ele[0], return_str)
                else:
                    if (ele.text is not None):
                        # return_str = eles.text
                        if (ret_s is not None):
                            ret_s += ele.text.strip()
                        else:
                            ret_s = ele.text.strip()
                    if (ele.tail is not None):
                        # print(return_str,eles.tail)
                        if (ret_s is not None):
                            ret_s += ele.tail.strip()
                        else:
                            ret_s = ele.tail.strip()
            return ret_s.strip()
    else:
        if(eles.text is not None):
            #return_str = eles.text
            if (return_str is not None):
                return_str += eles.text.strip()
            else:
                return_str = eles.text.strip()
        if(eles.tail is not None):
            #print(return_str,eles.tail)
            if(return_str is not None):
                return_str += eles.tail.strip()
            else:
                return_str = eles.tail.strip()
        return return_str.strip()
    return return_str.strip()

def parse_xml():
    wb = xlwt.Workbook()
    ws = wb.add_sheet('中国维管植物科属词典', cell_overwrite_ok=True)
    ws.write(0, 0, '0_科-拉丁名')
    ws.write(0, 1, '1_科-中文名')  # 科-中文名
    ws.write(0, 2, '2_科-定名人')

    ws.write(0, 3, '3_属-拉丁名')
    ws.write(0, 4, '4_属-中文名')
    ws.write(0, 5, '5_属-定名人')
    ws.write(0, 6, '6_目-中文名')
    ws.write(0, 7, '7_正文内容')
    ws.write(0, 8, '8_去掉染色体和茎内容')
    ws.write(0, 9, '9_染色体')
    ws.write(0, 10, '10_茎')
    ws.write(0, 11, '11_种数')
    ws.write(0, 12, '12_分布区')
    #ws.write(0, 13, '13_种数与分布区内容')
    i = 1
    for book in tree.iter(tag='book-body'):
        for part in book:
            if part.tag == 'part':
                for chapter in part:
                    for chap in chapter:
                        if chap.tag == 'chapter-title':
                            if(len(chap) > 0) :
                                if(chap[0].tag == 'bold'):
                                    chapter_title = chap[0].text;
                                    print(chapter_title)
                                    #get_text(chap[0])
                        if(chap.tag == 'sec'):
                            sec_first = '';
                            for named_contented in chap:
                                #if(len(named_contented.attrib) > 0):
                                if(named_contented.tag == 'named-content'):
                                    genus_latin = named_contented.text

                                    if(len(named_contented) > 0) :
                                        # if (named_contented[0].tag == 'bold'):
                                        #     genus_latin = named_contented[0].text
                                        # if(len(named_contented[0]) > 0):
                                        #     if (named_contented[0][0].tag == 'italic'):
                                        #         if(genus_latin != ''):
                                        #             genus_latin = genus_latin + ' ' +named_contented[0][0].text
                                        #         else:
                                        #             genus_latin = named_contented[0][0].text
                                        genus_latin = get_text(named_contented, genus_latin)
                                        #print(genus_latin)
                                    else:
                                        genus_latin = named_contented.text
                                    #ws.write(i, 3, genus_latin)
                                    if(genus_latin is None):
                                        genus_latin = ''
                                    if(genus_latin.strip() == '' ) :
                                        genus_latin = get_text(named_contented,genus_latin)
                                        print(genus_latin)
                                    else:
                                        print(genus_latin)

                                    if named_contented.attrib['content-type'] == 'genus-cn':
                                        ws.write(i, 4, genus_latin)
                                    if named_contented.attrib['content-type'] == 'family-cn':
                                        ws.write(i, 1, genus_latin)
                                    if named_contented.attrib['content-type'] == 'order-cn':
                                        ws.write(i, 6, genus_latin)
                                    if named_contented.attrib['content-type'] == 'genus-latin':
                                        ws.write(i, 3, genus_latin)
                                    if named_contented.attrib['content-type'] == 'family-latin':
                                        ws.write(i, 0, genus_latin)
                                    if(len(named_contented.attrib) > 0):

                                        if(named_contented.attrib['content-type'] == 'genus-latin' or named_contented.attrib['content-type'] == 'family-latin'):
                                            sec_first = named_contented.attrib['content-type']
                                        if (sec_first == 'genus-latin'):
                                            #ws.write(i, 3, genus_latin)
                                            if(named_contented.attrib['content-type'] == 'namer'):
                                                ws.write(i, 5, genus_latin)
                                        if (sec_first == 'family-latin'):
                                            #ws.write(i, 0, genus_latin)
                                            if (named_contented.attrib['content-type'] == 'namer'):
                                                ws.write(i, 2, genus_latin)
                                    else:
                                        print(named_contented.attrib, genus_latin)


                                else:
                                    #if named_contented.tag == 'p':
                                        if named_contented.text is not None:
                                            p = ''
                                            if (named_contented.text != '' or named_contented.text != '\r' or named_contented.text != '\n'):
                                                p = p + named_contented.text
                                        for e in named_contented:
                                            if e.tag == 'italic':
                                                # print('<i>',e.text,'</i>',e.tail)
                                                if (e.text != ''):
                                                    p = p + '<i>' + e.text.strip() + '</i>'
                                                if (e.tail != ''):
                                                    p = p + e.tail.strip()
                                            if e.tag == 'underline':
                                                # print('<i>',e.text,'</i>',e.tail)
                                                # p = p + '<a>' + e.text + '</a>' + e.tail
                                                if (e.text != ''):
                                                    p = p + '<a>' + e.text.strip() + '</a>'
                                                if (e.tail != ''):
                                                    p = p + e.tail.strip()
                                            if e.tag == 'bold':
                                                # print('<i>',e.text,'</i>',e.tail)
                                                # p = p + '<span>' + e.text + '</span>' + e.tail
                                                if (e.text != ''):
                                                    p = p + '<strong>' + e.text.strip() + '</strong>'
                                                    ws.write(i, 12, e.text.strip())
                                                if (e.tail != ''):
                                                    p = p + e.tail.strip()
                                            if e.tag == 'named-content':
                                                if(e.attrib['content-type'] == 'num'):
                                                    if(e.text.strip() != ''):
                                                        p = p + '<strong data-toggle="tooltip" data-placement="top" title="世界种数/中国种数">' + e.text.strip().replace('；','') + '</strong>'
                                                        ws.write(i, 11, e.text.strip().replace('；',''))
                                                    else:
                                                        num_str = get_text(e,e.text)
                                                        p = p + '<strong data-toggle="tooltip" data-placement="top" title="世界种数/中国种数">' + num_str.strip().replace('；','') + '</strong>'
                                                        ws.write(i, 11, num_str.strip().replace('；',''))
                                                    if(e.tail != ''):
                                                        p = p + e.tail.strip()
                                                if (e.attrib['content-type'] == 'type'):
                                                    if (e.text.strip() != '' and len(e) <= 0):
                                                        print('---------------',e.text,'---------------')
                                                        p = p + '<strong data-toggle="tooltip" data-placement="top" title="' + e.text.strip().replace('；','') + '">' + e.text.strip().replace('；','') + '</strong>'
                                                        ws.write(i, 12, e.text.strip().replace('；',''))
                                                    else:
                                                        #num_str = get_text(e, e.text)
                                                        num_str = e.text.strip()
                                                        for se in e:
                                                            num_str += se.text;
                                                            if(se.tail.strip() != ''):
                                                                num_str += se.tail
                                                            if(len(se) > 0):
                                                                break
                                                        p = p + '<strong data-toggle="tooltip" data-placement="top" title="' + num_str.strip().replace('；','') + '">' + num_str.strip().replace('；','') + '</strong>'
                                                        ws.write(i, 12, num_str.strip().replace('；',''))
                                                    if (e.tail != ''):
                                                        p = p + e.tail.strip()

                                        print(p.strip())
                                        if(p.strip() != ''):
                                            ws.write(i, 7, p.strip())

                                        # match_str = re.search(r'染色体(.*?)。',p.strip(),re.M|re.I)
                                        match_str = re.search(r'(.*?)。(.*)。(.*?染色体.*?)。(.*)', p.strip(),re.M | re.I)

                                        #if (match_str is not None):
                                        #     print(match_str.group())
                                        #     ws.write(i, 8, match_str.group(2) + '。' + match_str.group(5))
                                        #     ws.write(i, 9, match_str.group(3))
                                        #     ws.write(i, 11, match_str.group(4))
                                        if (p.strip() != '' and match_str is not None):
                                            print(match_str.group(1))
                                            print(match_str.group(2))
                                            print(match_str.group(3))
                                            print(match_str.group(4))
                                            ws.write(i, 8, match_str.group(2)+'。' + match_str.group(4) + '。')
                                            ws.write(i, 9, match_str.group(3))
                                            ws.write(i, 10, match_str.group(1))
                                            #ws.write(i, 13, match_str.group(4))
                                        else:
                                            match_str1 = re.search(r'(.*?)。(.*)', p.strip(), re.M | re.I)
                                            if (match_str1 is not None):

                                                ws.write(i, 8, match_str1.group(2))
                                                ws.write(i, 10, match_str1.group(1))
                                                # if(match_str1.group(3) != ''):
                                                #     ws.write(i, 13, '<strong ' + match_str1.group(3) + '。')

                                        # match_str1 = re.search(r'(.*)。(.*?)，<strong>(.*?)</strong>；(.*)',p.strip(), re.M | re.I)
                                        # if (match_str1 is not None):
                                        #     print(match_str1.group(1))
                                        #     print(match_str1.group(5))
                                        #     ws.write(i, 13, match_str1.group(2))
                            i = i + 1
    wb.save("C:/Users/dell/Desktop/科属词典-15.xls")

def xml():
    wb = xlwt.Workbook()
    ws = wb.add_sheet('中国维管植物科属词典',cell_overwrite_ok=True)
    ws.write(0, 0, '科-拉丁名')
    ws.write(0, 1, '定名人')
    ws.write(0, 2, '科-中文名') #科-中文名
    ws.write(0, 3, '属-拉丁名')
    ws.write(0, 4, '定名人')
    ws.write(0, 5, '属-中文名')
    ws.write(0, 6, '目-中文名')
    ws.write(0, 7, '正文内容')
    ws.write(0, 8, '去掉染色体内容和分布')
    ws.write(0, 9, '染色体')
    ws.write(0, 10, '茎')
    ws.write(0, 11, '分布')
    ws.write(0, 12, '型')
    ws.write(0, 13, '种分布')
    i = 1
    for book in tree.iter(tag='book-body'):
        for part in book:
            if part.tag == 'part':
                for chapter in part:
                    for chap in chapter:
                        if chap.tag == 'chapter-title':
                            print(chap.text)
                        if chap.tag == 'sec':
                            print('=======================',chap.attrib['subject'],'=======================')
                            sec_first = '';
                            for sec in chap:
                                if sec.tag == 'genus-latin':
                                    genus_latin = sec.text
                                    if(genus_latin == '' or genus_latin is None) :
                                        print('-----')
                                    sec_first = 'genus-latin'
                                    ws.write(i, 3, genus_latin)
                                if sec.tag == 'namer':
                                    print(sec.text)
                                    p = sec.text;
                                    for sec1 in sec :
                                        if(sec1.tag =='italic'):
                                            p = p + sec1.text
                                        if (sec1.tag == 'underline'):
                                            if(p != ''):
                                                p = p + ' ' + sec1.text
                                            else:
                                                p = p + sec1.text
                                        if(sec1.tail != ''):
                                            p = p + ' ' + sec1.tail.strip()
                                    if(sec_first == 'genus-latin'):
                                        ws.write(i, 4, p)
                                    if(sec_first == 'family-latin'):
                                        ws.write(i, 1, p)

                                if sec.tag == 'genus-cn':
                                    print(sec.text)
                                    ws.write(i, 5, sec.text)
                                if sec.tag == 'family-latin':
                                    print(sec.text)
                                    sec_first = 'family-latin'
                                    p = sec.text
                                    for sec1 in sec :
                                        if(sec1.tag =='italic'):
                                            if (p != ''):
                                                p = p + ' ' + sec1.text
                                            else:
                                                p = p + sec1.text
                                        if (sec1.tag == 'underline'):
                                            if(p != ''):
                                                p = p + ' ' + sec1.text
                                            else:
                                                p = p + sec1.text
                                        if (sec1.tag == 'bold'):
                                            if(p != ''):
                                                p = p + ' ' + sec1.text
                                            else:
                                                p = p + sec1.text
                                        if(sec1.tail != ''):
                                            p = p + ' ' + sec1.tail.strip()
                                    ws.write(i, 0, p)
                                if sec.tag == 'family-cn':
                                    print(sec.text)
                                    ws.write(i, 2, sec.text)
                                if sec.tag == 'order-cn':
                                    print(sec.text)
                                    ws.write(i, 6, sec.text)
                                if sec.tag == 'p' or sec.tag == 'bold':
                                    p = ''
                                    if(sec.text != '' or sec.text != '\r' or sec.text != '\n'):
                                        p = p + sec.text
                                    for e in sec:
                                        if e.tag == 'italic':
                                            #print('<i>',e.text,'</i>',e.tail)
                                            if(e.text != '') :
                                                p = p + '<i>' + e.text.strip() + '</i>'
                                            if(e.tail != '') :
                                                p = p + e.tail.strip()
                                        if e.tag == 'underline':
                                            #print('<i>',e.text,'</i>',e.tail)
                                            #p = p + '<a>' + e.text + '</a>' + e.tail
                                            if(e.text != '') :
                                                p = p + '<a>' + e.text.strip() + '</a>'
                                            if(e.tail != '') :
                                                p = p + e.tail.strip()
                                        if e.tag == 'bold':
                                            #print('<i>',e.text,'</i>',e.tail)
                                            #p = p + '<span>' + e.text + '</span>' + e.tail
                                            if(e.text != '') :
                                                p = p + '<strong>' + e.text.strip() + '</strong>'
                                                ws.write(i,12,e.text.strip())
                                            if(e.tail != '') :
                                                p = p + e.tail.strip()
                                    print(p.strip())
                                    ws.write(i, 7, p.strip())
                                    #match_str = re.search(r'染色体(.*?)。',p.strip(),re.M|re.I)
                                    match_str = re.search(r'(.*?)。(.*)染色体(.*?)。(.*?)。(.*)', p.strip(), re.M | re.I)

                                    if(match_str is not None):
                                        print(match_str.group())
                                        ws.write(i, 8, match_str.group(2) + '。' + match_str.group(5))
                                        ws.write(i, 9, match_str.group(3))
                                        ws.write(i, 11, match_str.group(4))
                                    if(p.strip() != '' and match_str is not None) :
                                        print(match_str.group(1))
                                        ws.write(i, 10, match_str.group(1))
                                    match_str1 = re.search(r'(.*)。(.*?)，<strong>(.*?)</strong>；*(.*)', p.strip(),re.M | re.I)
                                    if(match_str1 is not None):
                                        ws.write(i, 13, match_str1.group(2))
                            i = i + 1
    wb.save("C:/Users/dell/Desktop/科属词典-10.xls")

def pat() :
    str = '灌木、小乔木或攀援藤本，有刺或无刺。<a>二回羽状复叶或叶片退化而叶柄变为扁平的叶状柄</a>，<a>总叶柄和叶轴上常有腺体</a>。花小，两性或杂性；<a>常5基数</a>，有时3基数，<a>头状或穗状花序腋生或圆锥花序式排列</a>；<a>萼常钟状</a>，<a>具裂齿</a>；花瓣分离或合生，<a>有时缺</a>；<a>雄蕊多数</a>。<a>荚果两瓣开裂或逐节断裂</a>。花粉16合体。染色体2<i>n</i>=26。<strong data-toggle="tooltip" data-placement="top" title="世界种数/中国种数">约1450/18（3）种</strong>，<strong data-toggle="tooltip" data-placement="top" title="2（5，6）型">2（5，6）型</strong>；广布热带和亚热带地区，尤以澳大利亚种类最多；中国产西南、华南热带和亚热带地区。金合欢<i>A. farnesiana</i>(Linnaeus) Willdenow的根和荚可作黑色染料；花可制香水。'
    #str = '乔木或灌木。单叶，3小叶，<a>掌状叶或至少掌状叶脉</a>，<a>常具有长叶柄</a>。伞形花序或者伞房花序，有时总状或大型圆锥花序；萼片与花瓣（4）5，极少6，稀无花瓣；雄蕊（5）8（10，12），花丝等长或不等；子房多为2室，每心皮具2胚珠。<a>果实系2枚相连的小坚果</a>，具1枚种子，果实凸起或扁平，<a>侧面有长翅</a>，胚芽具有淀粉或油，胚根伸长。花粉粒3（拟）孔沟，条纹、网状或皱波状纹饰。染色体2<i>n</i>=26。126/99（61）种，<span>8型</span>；广布亚洲，欧洲及北美的热带、亚热带区域；中国各省区均产。'
    #match_str = re.search(r'(.*?)。(.*)染色体(.*?)。(.*?)。(.*)',str,re.M|re.I)
    #match_str = re.search(r'(.*)。(.*?)，<strong>(.*?)</strong>；*(.*)', str, re.M | re.I)
    match_str = re.search(r'(.*?)。(.*)', str,
                          re.M | re.I)
    if(match_str is not None):
        #print(match_str.groups().count(match_str))
        #print(match_str.group())
        print('1',match_str.group(1))
        print('2',match_str.group(2))
        #print('3',match_str.group(3))
        #print('4',match_str.group(4))
        #print('5',match_str.group(5))
    else:
        print('没有匹配到：',match_str)
def insertMysql(sql):
    #sql = pymysql.escape_string(sql)
    lastid = 0
    db = pymysql.connect(host='localhost',port= 3306,user = 'root',passwd='123456',db='test',charset='utf8')
    cursor = db.cursor()
    db.escape(sql)
    try:
        #print(sql)
        cursor.execute(sql)
        lastid = db.insert_id();
        db.commit()
    except Exception as e:
        print(e)
        db.rollback()
    cursor.close()
    db.close()
    return lastid
def get_name():
    db = pymysql.connect(host='localhost',port= 3306,user = 'root',passwd='123456',db='test',charset='utf8')
    cursor = db.cursor()
    sql = 'select * from keshu'
    cursor.execute(sql)
    data = cursor.fetchall()
    #print(data)
    cursor.close()
    return data
def get_kes():
    db = pymysql.connect(host='localhost',port= 3306,user = 'root',passwd='123456',db='test',charset='utf8')
    cursor = db.cursor()
    sql = 'select * from keshu WHERE shu_latin IS NULL and ke_latin=\'Rhachidosoraceae\''
    cursor.execute(sql)
    data = cursor.fetchall()
    print(data)
    cursor.close()
    return data

def get_shu():
    db = pymysql.connect(host='localhost',port= 3306,user = 'root',passwd='123456',db='test',charset='utf8')
    cursor = db.cursor()
    sql = 'select * from keshu WHERE shu_latin IS NOT NULL'
    cursor.execute(sql)
    data = cursor.fetchall()
    #print(data)
    cursor.close()
    return data

def get_name_existed(name):
    db = pymysql.connect(host='localhost',port= 3306,user = 'root',passwd='123456',db='test',charset='utf8')
    cursor = db.cursor()
    sql = 'select * from tb_classsys where classsys_latin=\'%s\''
    sql = sql % name
    cursor.execute(sql)
    data = cursor.fetchone()
    #print(data)
    cursor.close()
    return data
def get_name_existed_byname(name):
    db = pymysql.connect(host='localhost',port= 3306,user = 'root',passwd='123456',db='test',charset='utf8')
    cursor = db.cursor()
    sql = 'select * from tb_classsys where classsys_cname=\'%s\''
    sql = sql % name
    print(sql)
    cursor.execute(sql)
    data = cursor.fetchone()
    #print(data)
    cursor.close()
    return data
def get_fenbuqu(name):
    db = pymysql.connect(host='localhost',port= 3306,user = 'root',passwd='123456',db='test',charset='utf8')
    cursor = db.cursor()
    sql = 'select * from fenbuqu where num_str=\'%s\''
    sql = sql % name
    #print(sql)
    cursor.execute(sql)
    data = cursor.fetchone()
    #print(data)
    cursor.close()
    return data
def get_name_existed_sheet(name):
    db = pymysql.connect(host='localhost',port= 3306,user = 'root',passwd='123456',db='test',charset='utf8')
    cursor = db.cursor()
    sql = 'select * from sheet1 where kelatin=\'%s\' and kenames is not null'
    sql = sql % name
    cursor.execute(sql)
    data = cursor.fetchone()
    #print(data)
    cursor.close()
    return data

def insert_classsys(row) :
    sql = 'insert into sheet1 (`classsys_classid`,`classsys_class`,`classsys_par`,`classsys_latin`,`classsys_author`,`classsys_cname`,`classsys_latin2`,`classsys_pinyin`,`classsys_py`,`classsys_addtime`,`type1`,`create_date`) values (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)'

def datas() :
    is_existed_kes = []
    datas = get_name()
    for row in datas :
        #print(row,'------------')
        if(row[4] is not None) : #属的拉丁名
            data = get_name_existed(row[4]) #属的拉丁名在物种表中存在
            new_id = 0
            pid = 0
            if(data is not None) :
                new_id = data[0] #得到在物种表存在属的ID
            else :  #不存在物种表中，就要将新数据入库
                ke_data = None
                ke_latin = row[1].strip()
                ke_cname = row[3].strip()
                if(ke_latin is not None) :#科的拉丁名不是空的
                    ke_data = get_name_existed(ke_latin)
                else :
                    ke_data = get_name_existed_byname(ke_cname)
                if(ke_data is None):
                    ke_data_new = None
                    if(ke_latin != '' or ke_latin is not None):
                        ke_data_new = get_name_existed_sheet(ke_latin)
                    if(ke_data_new is not None) :
                        munames = ke_data_new[7]
                        m_pid = 0
                        if(munames is not None) :
                            mu_name_arr = munames.split('/')
                            mu_name = ''
                            if(len(mu_name_arr) > 1):
                                mu_name = mu_name_arr[1]
                            else :
                                mu_name = munames
                            mu_data = get_name_existed_byname(mu_name)
                            if(mu_data is not None) :
                                m_pid = mu_data[0]
                        if(ke_latin not in is_existed_kes):
                            sql = "insert into tb_classsys (`classsys_classid`,`classsys_class`,`classsys_par`,`classsys_latin`,`classsys_author`,`classsys_cname`,`classsys_latin2`,`classsys_pinyin`,`classsys_py`,`classsys_addtime`,`type1`,`create_date`,`del_flag`) values ('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')"
                            sql = sql %(13,'科',m_pid,ke_data_new[1],ke_data_new[2],ke_data_new[3],ke_data_new[1],get_pinyin(ke_data_new[3]),get_pinyin_prefix(ke_data_new[3]),datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),1,datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),0)
                            pid = insertMysql(sql)
                            is_existed_kes.append(ke_latin)
                            if(ke_data_new[8] is not None) :
                                sql1 = "insert into tb_spdesc (`spcid`,`splatin2`,`spdescid`,`spdesc`,`sporderid`,`spaddtime`,`create_date`,`del_flag`) values ('%s','%s','%s','%s','%s','%s','%s','%s')"
                                sql1 = sql1 % (pid,ke_data_new[1],7,ke_data_new[8],7,datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),0)
                                insertMysql(sql1)
                            if(ke_data_new[9] is not None) :
                                sql1 = "insert into tb_spdesc (`spcid`,`splatin2`,`spdescid`,`spdesc`,`sporderid`,`spaddtime`,`create_date`,`del_flag`) values ('%s','%s','%s','%s','%s','%s','%s','%s')"
                                sql1 = sql1 % (pid,ke_data_new[1],31,ke_data_new[9],31,datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),0)
                                insertMysql(sql1)
                            if(ke_data_new[11] is not None) :
                                sql1 = "insert into tb_spdesc (`spcid`,`splatin2`,`spdescid`,`spdesc`,`sporderid`,`spaddtime`,`create_date`,`del_flag`) values ('%s','%s','%s','%s','%s','%s','%s','%s')"
                                sql1 = sql1 % (pid,ke_data_new[1],30,ke_data_new[11],30,datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),0)
                                insertMysql(sql1)
                else :
                    pid = ke_data[0]

                sql = "insert into tb_classsys (`classsys_classid`,`classsys_class`,`classsys_par`,`classsys_latin`,`classsys_author`,`classsys_cname`,`classsys_latin2`,`classsys_pinyin`,`classsys_py`,`classsys_addtime`,`type1`,`create_date`,`del_flag`) values ('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')"
                sql = sql %(18,'属',pid,row[4],row[5],row[6],row[4],get_pinyin(row[6]),get_pinyin_prefix(row[6]),datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),1,datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),0)
                new_id = insertMysql(sql)
            if(row[8] is not None) :
                sql1 = "insert into tb_spdesc (`spcid`,`splatin2`,`spdescid`,`spdesc`,`sporderid`,`spaddtime`,`create_date`,`del_flag`) values ('%s','%s','%s','%s','%s','%s','%s','%s')"
                sql1 = sql1 % (new_id,row[4],7,row[8],7,datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),0)
                insertMysql(sql1)
            if(row[9] is not None) :
                sql1 = "insert into tb_spdesc (`spcid`,`splatin2`,`spdescid`,`spdesc`,`sporderid`,`spaddtime`,`create_date`,`del_flag`) values ('%s','%s','%s','%s','%s','%s','%s','%s')"
                sql1 = sql1 % (new_id,row[4],31,row[9],31,datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),0)
                insertMysql(sql1)
            if(row[11] is not None) :
                sql1 = "insert into tb_spdesc (`spcid`,`splatin2`,`spdescid`,`spdesc`,`sporderid`,`spaddtime`,`create_date`,`del_flag`) values ('%s','%s','%s','%s','%s','%s','%s','%s')"
                sql1 = sql1 % (new_id,row[4],30,row[11],30,datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),0)
                insertMysql(sql1)
        else :
            if(ke_latin not in is_existed_kes) :
                ke_latin = row[1];ke_cname = row[3]
                if(ke_latin is not None) :#科的拉丁名不是空的
                    ke_data = get_name_existed(ke_latin)
                else :
                    ke_data = get_name_existed_byname(ke_cname)
                if(ke_data is None) :
                    munames = row[7]
                    m_pid = 0
                    if(munames is not None) :
                        mu_name = munames.split('/')[1]
                        mu_data = get_name_existed_byname(mu_name)
                        if(mu_data is not None) :
                            m_pid = mu_data[0]
                    sql = "insert into tb_classsys (`classsys_classid`,`classsys_class`,`classsys_par`,`classsys_latin`,`classsys_author`,`classsys_cname`,`classsys_latin2`,`classsys_pinyin`,`classsys_py`,`classsys_addtime`,`type1`,`create_date`,`del_flag`) values ('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')"
                    sql = sql %(13,'科',m_pid,row[1],row[2],row[3],row[1],get_pinyin(row[3]),get_pinyin_prefix(row[3]),datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),1,datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),0)
                    new_id = insertMysql(sql)
                    is_existed_kes.append(ke_latin)
                if(row[8] is not None) :
                    sql1 = "insert into tb_spdesc (`spcid`,`splatin2`,`spdescid`,`spdesc`,`sporderid`,`spaddtime`,`create_date`,`del_flag`) values ('%s','%s','%s','%s','%s','%s','%s','%s')"
                    sql1 = sql1 % (new_id,row[1],7,row[8],7,datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),0)
                    insertMysql(sql1)
                if(row[9] is not None) :
                    sql1 = "insert into tb_spdesc (`spcid`,`splatin2`,`spdescid`,`spdesc`,`sporderid`,`spaddtime`,`create_date`,`del_flag`) values ('%s','%s','%s','%s','%s','%s','%s','%s')"
                    sql1 = sql1 % (new_id,row[1],31,row[9],31,datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),0)
                    insertMysql(sql1)
                if(row[11] is not None) :
                    sql1 = "insert into tb_spdesc (`spcid`,`splatin2`,`spdescid`,`spdesc`,`sporderid`,`spaddtime`,`create_date`,`del_flag`) values ('%s','%s','%s','%s','%s','%s','%s','%s')"
                    sql1 = sql1 % (new_id,row[1],30,row[11],30,datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),0)
                    insertMysql(sql1)

def get_pinyin(str):
    #print(pinyin.get_initial(str,delimiter=''))
    # if isinstance(var_str, str):
    #     if var_str == 'None':
    #         return ""
    #     else:
    #         return pinyin.get(var_str, format='strip', delimiter="")
    # else:
    #     return ''
    if(str is None) :
        return ''
    str = str.strip()
    if(str == 'None' or str == ''):
        return ''
    return pinyin.get(str,format='strip',delimiter=' ')

def get_pinyin_prefix(str):
    #return pinyin.get_initial(str,delimiter='').upper()
    # if isinstance(var_str, str):
    #     if var_str == 'None':
    #         return ""
    #     else:
    #         return pinyin.get_initial(var_str, format='strip', delimiter="").upper()
    # else:
    #     return ''
    if (str is None):
        return ''
    str = str.strip()
    if (str == 'None' or str == ''):
        return ''
    return pinyin.get_initial(str,delimiter='').upper()

def deal_mu():
    rows = get_kes();
    is_existed_kes = []
    for row in rows:
        gang_mu = row[6]
        gang_mu_arr = gang_mu.split('/')
        print(row[6],gang_mu_arr[0],gang_mu_arr[1])

        classid_0 = ''
        classsys_0 = ''
        classid_1 = ''
        classsys_1 = ''
        cname_0 = ''
        cname_1 = ''
        new_id = 0
        if('亚纲' in gang_mu_arr[0].strip()):
            classid_0 = 6
            classsys_0 = '亚纲'
        else:
            if ('纲' in gang_mu_arr[0].strip()):
                classid_0 = 5
                classsys_0 = '纲'
        if ('亚目' in gang_mu_arr[1].strip()):
            classid_1 = 10
            classsys_1 = '亚目'
        else:
            if ('目' in gang_mu_arr[1].strip()):
                classid_1 = 9
                classsys_1 = '目'
        if (gang_mu_arr[0] not in is_existed_kes):
            cname_0 = gang_mu_arr[0].strip()
            is_existed_kes.append(gang_mu_arr[0].strip())
        else:
            data = get_name_existed_byname(gang_mu_arr[0].strip())
            print(data)
            new_id = data[0]
        if (gang_mu_arr[1] not in is_existed_kes):
            cname_1 = gang_mu_arr[1].strip()
            is_existed_kes.append(gang_mu_arr[1].strip())
        list(set(is_existed_kes))

        if(cname_0 != ''):
            sql = "insert into tb_classsys (`classsys_classid`,`classsys_class`,`classsys_par`,`classsys_latin`,`classsys_author`,`classsys_cname`,`classsys_latin2`,`classsys_pinyin`,`classsys_py`,`classsys_addtime`,`type1`,`create_date`,`del_flag`) values ('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')"
            sql = sql % (classid_0, classsys_0,0, '', '', cname_0, '', get_pinyin(cname_0), get_pinyin_prefix(cname_0),
                     datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 1,
                     datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 0)
            new_id = insertMysql(sql)
        if (cname_1 != ''):
            sql = "insert into tb_classsys (`classsys_classid`,`classsys_class`,`classsys_par`,`classsys_latin`,`classsys_author`,`classsys_cname`,`classsys_latin2`,`classsys_pinyin`,`classsys_py`,`classsys_addtime`,`type1`,`create_date`,`del_flag`) values ('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')"
            sql = sql % (classid_1, classsys_1, new_id, '', '', cname_1, '', get_pinyin(cname_1), get_pinyin_prefix(cname_1),
            datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 1,
            datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 0)
            insertMysql(sql)



    print('总数=',len(list(set(is_existed_kes))))
def deal_ke():
    rows = get_kes();
    is_existed_kes = []
    for row in rows:
        gang_mu = row[6]
        gang_mu_arr = gang_mu.split('/')
        print(row[6],gang_mu_arr[0],gang_mu_arr[1])
        data = get_name_existed_byname(gang_mu_arr[1].strip())
        par_id = 0
        new_id = 0
        if(data is not None):
            par_id = data[0]
        if(par_id != 0):
            sql = "insert into tb_classsys (`classsys_classid`,`classsys_class`,`classsys_par`,`classsys_latin`,`classsys_author`,`classsys_cname`,`classsys_latin2`,`classsys_pinyin`,`classsys_py`,`classsys_addtime`,`type1`,`create_date`,`del_flag`) values ('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')"
            sql = sql % (13, '科', par_id, row[0], row[2], row[1], row[0], get_pinyin(row[1]), get_pinyin_prefix(row[1]),
                         datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 1,
                         datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 0)
            new_id = insertMysql(sql)
            if (row[8] is not None):
                spdesc = row[8]
                if(row[12] is not None):
                    data1 = get_fenbuqu(row[12].strip())
                    if(data1 is not None):
                        spdesc = row[8].replace('title="'+row[12].strip()+'"','title="'+data1[1].strip()+'"')
                sql1 = "insert into tb_spdesc (`spcid`,`splatin2`,`spdescid`,`spdesc`,`sporderid`,`spaddtime`,`create_date`,`del_flag`) values ('%s','%s','%s','%s','%s','%s','%s','%s')"
                sql1 = sql1 % (new_id, row[0], 7, spdesc, 7, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                               datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 0)
                insertMysql(sql1)
            if (row[9] is not None):
                sql1 = "insert into tb_spdesc (`spcid`,`splatin2`,`spdescid`,`spdesc`,`sporderid`,`spaddtime`,`create_date`,`del_flag`) values ('%s','%s','%s','%s','%s','%s','%s','%s')"
                sql1 = sql1 % (new_id, row[0], 31, row[9].replace('染色体',''), 31, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                               datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 0)
                insertMysql(sql1)
            if (row[10] is not None):
                sql1 = "insert into tb_spdesc (`spcid`,`splatin2`,`spdescid`,`spdesc`,`sporderid`,`spaddtime`,`create_date`,`del_flag`) values ('%s','%s','%s','%s','%s','%s','%s','%s')"
                sql1 = sql1 % (new_id, row[0], 30, row[10], 30, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                               datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 0)
                insertMysql(sql1)

def deal_shu():
    rows = get_shu();
    is_existed_kes = []
    for row in rows:
        ke_latin = row[0]
        data = get_name_existed(ke_latin.strip())
        par_id = 0
        new_id = 0
        if(data is not None):
            par_id = data[0]
        else:
            print(ke_latin.strip())
            sql = "insert into tb_classsys (`classsys_classid`,`classsys_class`,`classsys_par`,`classsys_latin`,`classsys_author`,`classsys_cname`,`classsys_latin2`,`classsys_pinyin`,`classsys_py`,`classsys_addtime`,`type1`,`create_date`,`del_flag`) values ('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')"
            sql = sql % (13, '科', 0, ke_latin.strip(), '', row[1], ke_latin.strip(), get_pinyin(row[1]), get_pinyin_prefix(row[1]),
                         datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 1,
                         datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 0)
            par_id = insertMysql(sql)
        if(par_id != 0):
            sql = "insert into tb_classsys (`classsys_classid`,`classsys_class`,`classsys_par`,`classsys_latin`,`classsys_author`,`classsys_cname`,`classsys_latin2`,`classsys_pinyin`,`classsys_py`,`classsys_addtime`,`type1`,`create_date`,`del_flag`) values ('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')"
            sql = sql % (18, '属', par_id, row[3], row[5], row[4], row[3], get_pinyin(row[4]), get_pinyin_prefix(row[4]),
                         datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 1,
                         datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 0)
            new_id = insertMysql(sql)
            if (row[8] is not None):
                spdesc = row[8]
                if(row[12] is not None):
                    data1 = get_fenbuqu(row[12].strip())
                    if(data1 is not None):
                        spdesc = row[8].replace('title="'+row[12].strip()+'"','title="'+data1[1].strip()+'"')
                sql1 = "insert into tb_spdesc (`spcid`,`splatin2`,`spdescid`,`spdesc`,`sporderid`,`spaddtime`,`create_date`,`del_flag`) values ('%s','%s','%s','%s','%s','%s','%s','%s')"
                sql1 = sql1 % (new_id, row[0], 7, spdesc, 7, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                               datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 0)
                insertMysql(sql1)
            if (row[9] is not None):
                sql1 = "insert into tb_spdesc (`spcid`,`splatin2`,`spdescid`,`spdesc`,`sporderid`,`spaddtime`,`create_date`,`del_flag`) values ('%s','%s','%s','%s','%s','%s','%s','%s')"
                sql1 = sql1 % (new_id, row[0], 31, row[9].replace('染色体',''), 31, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                               datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 0)
                insertMysql(sql1)
            if (row[10] is not None):
                sql1 = "insert into tb_spdesc (`spcid`,`splatin2`,`spdescid`,`spdesc`,`sporderid`,`spaddtime`,`create_date`,`del_flag`) values ('%s','%s','%s','%s','%s','%s','%s','%s')"
                sql1 = sql1 % (new_id, row[0], 30, row[10], 30, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                               datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 0)
                insertMysql(sql1)


parse_xml()
#pat()
#get_name()
#get_name_existed('Aconitum')
#get_pinyin('西番莲科')
#print('木兰亚纲/川续断目'.split('/')[1])
#datas()
#deal_mu()
#deal_ke()

#deal_shu()