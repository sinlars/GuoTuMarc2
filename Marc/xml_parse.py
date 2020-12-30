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

tree = et.ElementTree(file='C:/Users/dell/Desktop/中国生物物种名录被子植物IX分册-内文标引.xml')

for book in tree.iter(tag='book-meta'):
    authors = []
    for elem in book.iter():
        if elem.tag == 'grant-sponsor':
            for child in elem.iter():
                if child.tag == 'image':
                    print(child.attrib['href'])
        if elem.tag == 'seriesname':
            print('seriesname ：',elem.text)
        if elem.tag == 'book-title':
            print('书名 ：',elem.text)
        if elem.tag == 'p':
            print(elem.attrib,elem.text,elem.tail)
            for p_child in elem.iter():
                if p_child.tag == 'phylum-cn':
                    print(p_child.text,p_child.tail)
                if p_child.tag == 'phylum-latin':
                    for p_child_child in p_child :
                        if(p_child_child.tag == 'bold'):
                            print(p_child_child.text, p_child_child.tail)
                if p_child.tag == 'family-cn':
                    print(p_child.text,p_child.tail)
                if p_child.tag == 'family-latin':
                    print(p_child.text,p_child.tail)
                if p_child.tag == 'trans-title':
                    print(p_child.text, p_child.tail)
                if p_child.tag == 'trans-subtitle':
                    print(p_child.text, p_child.tail)
                    for p_child_child in p_child :
                        if(p_child_child.tag == 'bold'):
                            print(p_child_child.text, p_child_child.tail)
        if elem.tag == 'contrib':
            names = []
            for el in elem.iter():
                name = ''
                if el.tag == 'surname':
                    #print('姓：',el.text)
                    name = name + el.text
                if el.tag == 'given-names':
                    #print('名：',el.text)
                    name = name + el.text
                if name != '' :
                    names.append(name)
            #print(names)
            print(elem.attrib['contrib-type'],'：',''.join(names))
        if elem.tag == 'publisher':
            for el in elem.iter():
                if el.tag == 'publisher-name':
                    print('出版社：',el.text)
                if el.tag == 'address':
                    print('出版地：',el.text)

for book in tree.iter(tag='book-front'):
    for elem in book.iter():
        if elem.tag == 'preface':
            for el in elem.iter():
                if el.tag == 'title':
                    print(el.text)
                if el.tag == 'p':
                    p = el.text
                    for e in el:
                        if e.tag == 'italic':
                            #print('<i>',e.text,'</i>',e.tail)
                            p = p + '<i>' + e.text + '</i>' + e.tail
                    print(p)
                if el.tag == 'contrib':
                    names = []
                    for e in el.iter():
                        name = ''
                        if e.tag == 'surname':
                            #print('姓：',el.text)
                            name = name + e.text
                        if e.tag == 'given-names':
                            #print('名：',el.text)
                            name = name + e.text
                        if name != '' :
                            names.append(name)
                        #print(names)
                    print(el.attrib['contrib-type'],'：',','.join(names))
                if el.tag == 'date':
                    print(el.text)
        if elem.tag == 'foreword':
            for el in elem.iter():
                if el.tag == 'title':
                    print(el.text)
                if el.tag == 'p':
                    p = el.text
                    for e in el:
                        if e.tag == 'italic':
                            #print('<i>',e.text,'</i>',e.tail)
                            p = p + '<i>' + e.text + '</i>' + e.tail
                    print(p)
                if el.tag == 'contrib':
                    names = []
                    for e in el.iter():
                        name = ''
                        if e.tag == 'surname':
                            #print('姓：',el.text)
                            name = name + e.text
                        if e.tag == 'given-names':
                            #print('名：',el.text)
                            name = name + e.text
                        if name != '' :
                            names.append(name)
                        #print(names)
                    print(el.attrib['contrib-type'],'：',','.join(names))
                if el.tag == 'date':
                    print(el.text)
for book in tree.iter(tag='book-body'):
    for part in book:
        if part.tag == 'part':
            for chapter in part:
                if chapter.tag == 'chapter':
                    for i,chapter_child in enumerate(chapter):
                        print("i = ",i,chapter_child.tag)
                        # if(i == 2):
                        #     break
                        if(chapter_child.tag == 'chapter-title'):
                            for ch_child in chapter_child :
                                if(ch_child.tag == 'phylum-cn'):
                                    print(ch_child.attrib,ch_child.text,ch_child.tail)
                                if(ch_child.tag == 'phylum-latin'):
                                    for ch in ch_child:
                                        if(ch.tag == 'bold'):
                                            print(ch.attrib,ch.text,ch.tail)
                        if(chapter_child.tag == 'sec') :
                            for ch in chapter_child:
                                if (ch.tag == 'title'):
                                    title = ''
                                    for c in ch :
                                        if(c.tag == 'family-no'):
                                            #print(c.attrib,c.text,c.tail)
                                            if(c.text != '') :
                                                title = title + c.text
                                        if (c.tag == 'family-cn'):
                                            #print(c.attrib, c.text, c.tail)
                                            if (c.text != ''):
                                                title = title + c.text
                                        if (c.tag == 'family-latin'):
                                            #print(c.attrib, c.text, c.tail)
                                            if (c.text != ''):
                                                title = title + c.text
                                    print('sec_title :',title)
                                if(ch.tag == 'p'):
                                    #print(ch.attrib,ch.text,ch.tail)
                                    p = ''
                                    if(ch.text != ''):
                                        p = p + ch.text
                                    for c in ch:
                                        if(c.tag == 'genus-num'):
                                            #print(c.attrib,c.text,c.tail)
                                            if (c.text != ''):
                                                p = p + c.text + c.tail
                                        if (c.tag == 'species-num'):
                                            #print(c.attrib, c.text, c.tail)
                                            if (c.text != ''):
                                                p = p + c.text + c.tail
                                    print('p = ',p)
                                if(ch.tag == 'sec1'):
                                    for c in ch :
                                        if(c.tag == 'title'):
                                            sec1_title = ''
                                            for cc in c:
                                                if(cc.tag == 'genus-cn'):
                                                    #print(cc.attrib,cc.text,cc.tail)
                                                    if(cc.text != '') :
                                                        sec1_title = sec1_title + cc.text.strip() + ' ' + cc.tail.strip()
                                                if(cc.tag == 'genus-latin'):
                                                    #print(cc.attrib, cc.text, cc.tail)
                                                    if (cc.text != ''):
                                                        sec1_title = sec1_title + cc.text.strip() + ' ' + cc.tail.strip()
                                                    for ccc in cc:
                                                        if(ccc.tag == 'bold'):
                                                            #print(ccc.attrib, ccc.text, ccc.tail)
                                                            if (ccc.text != ''):
                                                                sec1_title = sec1_title + '<b>' + ccc.text.strip() + '</b>'+ ' ' + cc.tail.strip()
                                                if(cc.tag == 'namer'):
                                                    #print(cc.attrib, cc.text, cc.tail)
                                                    if (cc.text != ''):
                                                        sec1_title = sec1_title + cc.text.strip() + ' ' + cc.tail.strip()
                                            print('sec1_title = ',sec1_title.strip())
                                        if(c.tag == 'sec2'):
                                            for cc in c:
                                                if(cc.tag == 'title'):
                                                    sec2_title = ''
                                                    for ccc in cc:
                                                        if(ccc.tag == 'special'):
                                                            sec2_title = sec2_title + ccc.text.strip() + ' ' + ccc.tail.strip()
                                                        if(ccc.tag == 'species-cn'):
                                                            #print(ccc.attrib, ccc.text, ccc.tail)
                                                            sec2_title = sec2_title + ccc.text.strip() + ' ' + ccc.tail.strip()
                                                    sec2_title = sec2_title + ' ' + cc.text.strip()
                                                    print('sec2_title = ',sec2_title)
                                                if(cc.tag == 'p'):
                                                    #print(cc.attrib, cc.text, cc.tail)
                                                    sec2_text = cc.text.strip() + ' ' + cc.tail.strip()
                                                    for ccc in cc:
                                                        if(ccc.tag == 'species-latin' or ccc.tag == 'synonym-latin'):
                                                            #print(ccc.attrib, ccc.text, ccc.tail)
                                                            for cccc in ccc:
                                                                if(cccc.tag == 'bold'):
                                                                    #print(cccc.attrib,cccc.text,cccc.tail)
                                                                    sec2_text = sec2_text + cccc.text.strip() + ' ' + cccc.tail.strip()
                                                                if (cccc.tag == 'italic'):
                                                                    #print(cccc.attrib, cccc.text, cccc.tail)
                                                                    sec2_text = sec2_text + cccc.text.strip() + ' ' + cccc.tail.strip()
                                                            sec2_text = sec2_text + ccc.tail.strip()
                                                        if(ccc.tag == 'namer'):
                                                            #print(ccc.attrib, ccc.text, ccc.tail)
                                                            sec2_text = sec2_text + ccc.text.strip() + ' ' + ccc.tail.strip()
                                                        if(ccc.tag == 'distribution'):
                                                            #print(ccc.attrib, ccc.text, ccc.tail)
                                                            sec2_text = sec2_text + ccc.text.strip() + ' ' + ccc.tail.strip()
                                                        #sec2_text = sec2_text + ccc.text.strip() + ccc.tail.strip()
                                                    print('sec2_text =', sec2_text)


for book in tree.iter(tag='book-back'):
    for ref_list in book:
        if(ref_list.tag == 'ref-list'):
            for ref in ref_list:
                if(ref.tag == 'title'):
                    print('title = ',ref.text)
                if(ref.tag == 'ref'):
                    for note in ref:
                        if(note.tag == 'note'):
                            note_text = ''
                            if(len(note) > 0) :
                                note_text = note_text + note.text.strip()
                                for italic in note:
                                    note_text = note_text + italic.text.strip() + ' ' + italic.tail.strip()
                                note_text = note_text + note.tail.strip()
                            else:
                                note_text = note.text.strip() + ' ' + note.tail.strip()
                            print('note :',note_text)