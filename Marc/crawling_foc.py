#coding:utf-8

import urllib.request
import lxml.html
from pymarc import Record, Field
from pymarc import MARCReader
import re
import xlwt
import sys,io
import openpyxl
from bs4 import BeautifulSoup
import gzip
import docx
from docx import Document
from io import BytesIO
import pymysql
import pinyin
import datetime
import requests

#改变标准输出的默认编码
#sys.stdout = io.TextIOWrapper(sys.stdout.buffer,encoding='gb18030')

headers = {
    'Accept':'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
    #'Accept-Encoding':'gzip, deflate, sdch',
    'Accept-Encoding':'gzip, deflate',
    'Accept-Language':'zh-CN,zh;q=0.9',
    #'Connection':'keep-alive',
    #'Cookie':'_gscu_413729954=00942062efyg0418; Hm_lvt_2cb70313e397e478740d394884fb0b8a=1500942062',
    #'Host':'opac.nlc.cn',
    'Cookie':'PHPSESSID=0f94e40864d4e71b5dfeb2a8cf392922; Hm_lvt_668f5751b331d2a1eec31f2dc0253443=1542012452,1542068702,1542164499,1542244740; Hm_lpvt_668f5751b331d2a1eec31f2dc0253443=1542246351',
    'Upgrade-Insecure-Requests':'1',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3141.7 Safari/537.36 Core/1.53.3226.400 QQBrowser/9.6.11682.400'}


def getHtml(url,num_retries = 5):
    print('Crawling url:',url)
    try:
        request = urllib.request.Request(url,headers=headers)
        response = urllib.request.urlopen(request,timeout=30)
        info = response.info();
        page_html = ''
        page_html = response.read()
        if info.get('Content-Encoding') == 'gzip':
            buff = BytesIO(page_html)  # 把content转为文件对象
            f = gzip.GzipFile(fileobj=buff)
            page_html = f.read().decode('utf-8')
        else:
            page_html = page_html.decode('utf-8','ignore')
            print(page_html)
    except Exception as e:
        print('Downloading error:',str(e))
        print('重试次数：', num_retries)
        page_html = None;
        if (num_retries > 0):
            if(hasattr(e, 'code') and 500 <= e.code < 600) :
                return getHtml(url,num_retries - 1)
            else:
                return getHtml(url, num_retries - 1)
        else :
            print('重试次数完毕：',num_retries)
            return page_html
    return page_html

def insertMysql(sql):
    #sql = pymysql.escape_string(sql)
    lastid = 0
    db = pymysql.connect(host='localhost',port= 3306,user = 'root',passwd='123456',db='zhiwu',charset='utf8')
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

def get_pinyin(str):
    if(str is None) :
        return ''
    str = str.strip()
    if(str == 'None' or str == ''):
        return ''
    return pinyin.get(str,format='strip',delimiter=' ')
def get_pinyin_prefix(str):
    if (str is None):
        return ''
    str = str.strip()
    if (str == 'None' or str == ''):
        return ''
    return pinyin.get_initial(str,delimiter='').upper()

def get_name_existed(name):
    db = pymysql.connect(host='localhost',port= 3306,user = 'root',passwd='123456',db='zhiwu',charset='utf8')
    cursor = db.cursor()
    sql = 'select * from tb_classsys where classsys_latin=\'%s\''
    sql = sql % name
    cursor.execute(sql)
    data = cursor.fetchone()
    #print(data)
    cursor.close()
    return data
def get_foc():
    db = pymysql.connect(host='localhost', port=3306, user='root', passwd='123456', db='zhiwu', charset='utf8')
    cursor = db.cursor()
    sql = 'select * from zhiwu2'
    #sql = sql % name
    cursor.execute(sql)
    data = cursor.fetchall()
    print(data)
    cursor.close()
    return data


def get_text_docx():
    file = docx.Document("C:\\Users\\dell\\Desktop\\高等九卷.docx")
    i = 1
    j = 0
    wb = xlwt.Workbook()
    ws = wb.add_sheet('中国高等植物彩色图鉴正文内容-第九卷', cell_overwrite_ok=True)
    ws.write(0, 0, '物种中文名')
    ws.write(0, 1, '物种拉丁名')  # 科-中文名
    ws.write(0, 2, '正文内容')
    ws.write(0, 3, '正文英文内容')
    ke = False
    for p in file.paragraphs:
        #if i > 20 :break
        #print('--------------------')d
        #if p.text.strip() == '':break

        if p.text.strip() == '' :
            continue
        if j%4 == 0:
            j = 0
            i = i + 1


        print('----------',i, j)
        ws.write(i, j, p.text.strip())

        if ke is True:
            j = 0
            #i = i + 1
            ke = False
        else :
            j = j + 1
        print(p.text,'---',p.style.name)
        #print(run.bold for run in p.runs)
        #if p.style.name == '种-英文' :
        for run in p.runs:
            if run.bold :
                print(run.text,run.bold)
            #print(run.bold)
        #print('--------------------')
        #j = j + 1
        if p.text.strip().endswith('科'):
            ke = True

    wb.save("C:/Users/dell/Desktop/高等九卷.xls")



def get_content(url='http://www.efloras.org/',cralw_url='http://www.efloras.org/browse.aspx?flora_id=2&page=%s',pages=2):
    for i in range(1,pages+1):
        cralw_url_i = cralw_url % (str(i))
        info = getHtml(cralw_url_i)
        #print(info)
        page_context = BeautifulSoup(info, "html.parser")
        divs = page_context.find_all(id='ucFloraTaxonList_panelTaxonList')
        #print(divs)
        if len(divs) > 0:
            div = divs[0]
            table = div.find_all('table')[0]
            #print(table)
            trs = table.find_all('tr')
            #print(trs)
            for tr in trs:
                tds = tr.find_all('td')
                if len(tds) == 5:
                    print(tds[0].text,tds[1].text,tds[2].text,tds[3].text,tds[4].text)
                #if tds[1].fina_all('a') is not None:
                    ke_urls = tds[1].select('a[href]')
                    print(ke_urls)
                    if len(ke_urls) > 0:
                        ke_url = ke_urls[0].get('href');
                        print('ke_url :',ke_url)
                        ke_context = getHtml(url)
                        #print(ke_context)
                        ke_context_soup = BeautifulSoup(ke_context, "html.parser")
                        table_ke = ke_context_soup.find_all('table',id='footerTable')
                        print(table_ke)
                    shu_urls = tds[3].select('a[href]')
                    print(shu_urls)
                    if len(shu_urls) > 0:
                        print('shu_url :',shu_urls[0].get('href'))

def get_ke_context(url):
    volume_content = {};
    ke_context = getHtml(url)
    volume_content['url'] = url
    volume_content['taxon_id'] = get_max_number(url)

    ke_context_soup = BeautifulSoup(ke_context, "html.parser")
    table_ke = ke_context_soup.find_all('table', id='footerTable')
    tds = table_ke[0].select('td[style]')
    #print(tds[0].text)# 科所在的卷册、页码等
    volume_content['volume_title'] = tds[0].text
    div_context = ke_context_soup.find_all('div', id='panelTaxonTreatment')
    #print(div_context[0].find_all(re.compile("^image")))
    #print('正文内容：',div_context[0].prettify())
    foc_taxon_chain = ke_context_soup.select_one('span[id="lblTaxonChain"]')
    #print(foc_taxon_chain)
    parent_links = foc_taxon_chain.find_all('a')
    if parent_links:
        parent_link = parent_links[len(parent_links)-1].get('href')
        volume_content['parent_taxon_id'] = get_max_number(parent_link)
    volume_list = foc_taxon_chain.find_all('a', href=re.compile("volume_id"), recursive=False)
    if len(volume_list) == 1:
        volume_content['volume_id'] = get_max_number(volume_list[0].get('href'))
        volume_content['volume'] = volume_list[0].text
    span = div_context[0].find_all('span',id='lblTaxonDesc')[0]

    #print('正文内容：', span.prettify())
    #print(span.prettify())

    #####################获取有image图片信息的部分内容################
    image_table = span.select_one('table')
    if image_table:
        image_table_tr_list = image_table.find_all('tr')
        for image_table_tr in image_table_tr_list:
            image_table_td_list = image_table_tr.find_all('td')
            for image_table_td in image_table_td_list:
                if image_table_td.a:
                    #print('图片连接：',image_table_td.select_one('a').img.get('src'))          ##获取图片的链接\
                    image_link = image_table_td.a.img.get('src')
                    #print('图片连接：', image_link)

                    #download_file(image_link,'F:\FloraData\images\\' + str(get_max_number(image_link)) + '.jpg')
                if image_table_td.a.next_sibling :
                    print('当前物种的拉丁名及链接等：',image_table_td.a.next_sibling.get('href'),image_table_td.a.next_sibling.text)
                if image_table_td.a.next_sibling.next_sibling:
                    print('Credit:',image_table_td.a.next_sibling.next_sibling.small.text)
        image_table.extract()
    ###############################################################
    #print(span.b.next_siblings)
    latin_name_object = []
    for wuzh in span.next_element.next_siblings:
        if wuzh.name == 'p':
            continue
        if wuzh.name == 'a': #表示直接跳转下个物种，类似 See Isoëtaceae # http://www.efloras.org/florataxon.aspx?flora_id=2&taxon_id=20790
            latin_name_object = []
            latin_name_object.append(wuzh)
            break
        if wuzh.name == 'small' :
            volume_content['small'] = wuzh.string.strip('\n\r ')
            continue
        if wuzh.string is not None and wuzh.string.strip('\n\r '):
            latin_name_object.append(wuzh)
        #else:
        #    print(repr(wuzh).strip(['\n', ' ', '\r\n']))
    print(latin_name_object)
    if len(latin_name_object) > 1:
        if latin_name_object[0].name is None: #如果第一个字符串是类似1.，7a,... 则表示序号
            volume_content['xuhao'] = latin_name_object[0].string.strip('\n\r ')
        else:
            volume_content['xuhao'] = ''
        if latin_name_object[len(latin_name_object)-1].name is None : #如果最后一个字符串是类似(Blume) Tagawa, Acta Phytotax. Geobot. 7: 83. 1938.则表示文献
            volume_content['latin_name'] = ' '.join(list(latin.string.strip('\n\r ') for latin in latin_name_object[1:len(latin_name_object)-1] ))
        else:
            volume_content['latin_name'] = ' '.join(list(latin.string.strip('\n\r ') for latin in latin_name_object[1:]))
    else:
        volume_content['xuhao'] = ''
        volume_content['latin_name'] = ' '.join(list(latin.string.strip('\n\r ') for latin in latin_name_object))
    #volume_content['xuhao'] = latin_name[0]
    #print(span.b.next_sibling) #当前物种信息的物种拉丁名
    #print(span.b.find_next_sibling("p").contents[0].strip())
    volume_content['latin_name_full'] = span.b.next_sibling.strip()
    #print(span.b.find_next_sibling("p"))
    #print('-----------------------')
    #print(span.b.find_next_sibling("p").contents[0])
    zh_name_and_pinyin = span.b.find_next_sibling("p").contents[0]
    if is_all_zh(zh_name_and_pinyin):   #含有中文
        print('#######################')
        print(zh_name_and_pinyin.split(' ')[0].strip())
        print(' '.join(zh_name_and_pinyin.split(' ')[1:]))
    #print(re.sub('[A-Za-z0-9\!\%\[\]\,\。\(\)]', '', zh_name_and_pinyin))
    #print(' '.join(re.findall(r'[A-Za-z\(\)]+', zh_name_and_pinyin)))
        volume_content['zh_name'] = zh_name_and_pinyin.split(' ')[0].strip()
        volume_content['zh_name_pinyin'] = ' '.join(zh_name_and_pinyin.split(' ')[1:]).strip()
    else:
        volume_content['zh_name'] = ''
        volume_content['zh_name_pinyin'] = zh_name_and_pinyin.strip()
    #authors = span.b.find_next_sibling("p").p.next_element #获取下面一个直接字符串
    spdesc_p_list = span.b.find_next_sibling("p").p
    #print('##############################################')
    #print(spdesc_p_list)
    #print('##############################################')
    #print(spdesc_p_list.find_all('a',recursive=False))
    authors_list = []
    authors_id_list = []
    for author in spdesc_p_list.find_all('a',recursive=False):
        #print(author.text,author.get('href'))
        authors_list.append(author.text)
        authors_id_list.append(str(get_max_number(author.get('href'))))
    volume_content['authors'] = ';'.join(authors_list)
    volume_content['authors_id'] = ';'.join(authors_id_list)

    #print(authors.find_all('p',recursive=False)[0].prettify())
    spdescs = spdesc_p_list.find_all('p',recursive=False)
    #print(spdescs)
    print('##############################################')
    if len(spdescs) > 0:
        specs_context = ''
        table = spdescs[0].select_one('table')
        if table is not None:
            #print(table.find_all('a'))

            for s in table.next_sibling.next_sibling.strings:
                #print(repr(s),type(s),s.parent.name=='i')
                if s.parent.name == 'i':
                    specs_context = specs_context + '<i>' + s.strip('\n') + '</i>'
                else:
                    specs_context = specs_context + s.strip('\n')
        else :
            #print(spdescs[0].strings)
            for s in spdescs[0].strings:
                #print(s)
                if s.parent.name == 'i':
                    specs_context = specs_context + '<i>' + s.strip('\n') + '</i>'
                else:
                    if s.parent.name == 'b':
                        specs_context = specs_context + '<b>' + s.strip('\n') + '</b>'
                    specs_context = specs_context + s.strip('\n')
            #print(specs_context.strip())
        #print('##############################################')
        #print(specs_context.strip())#获取正文内容
        volume_content['content'] = specs_context.strip()
    #volume_content['create_date'] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        # sql = "insert into volume_content (`content`,`create_date`,`del_flag`) values ('%s','%s','%s')"
        # sql = sql % (specs_context.strip(), datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 0)
        # pid = insertMysql(sql)
    #print('##############################################')
    print(volume_content)
    wuzhong_detail_sql(volume_content)
    table_jiansuobiao = div_context[0].find_all('table',id='tableKey') #获取检索表的内容
    if len(table_jiansuobiao) > 0:
        trs_jiansuobiao = table_jiansuobiao[0].find_all('tr')
        table_jsb = trs_jiansuobiao[1].find_all('table')
        if len(table_jsb) > 0:
            trs_jsb = table_jsb[0].find_all('tr')
            for tr in trs_jsb:
                tds_jsb = tr.find_all('td')
                tds_jxb_cs = tds_jsb[3].contents;
                goto_no = ''
                goto_id = ''
                for tds_jxb_c in tds_jxb_cs:
                    #print(tds_jxb_c.name)
                    if tds_jxb_c.name == 'a':
                        tds_jxb_c_href = tds_jxb_c.get('href')
                        tds_jxb_c_s = tds_jxb_c.string;
                        if tds_jxb_c_s is not None:
                            #print(tds_jxb_c)
                            goto_id = tds_jxb_c_href + '='+ tds_jxb_c_s
                        else:
                            goto_id = tds_jxb_c_href
                    else :
                        goto_no = tds_jxb_c
                #print(tds_jsb[0].text,tds_jsb[1].text,tds_jsb[2].text,goto_no,goto_id)
    ############################################################################################
    ###lower_taxa_ul = div_context[0].select_one('ul')#获取当前物种的下级物种信息
    ###print(lower_taxa_ul)
    # if lower_taxa_ul is not None:
    #     for li in lower_taxa_ul.find_all('li'):
    #         lower_taxa_a = li.select_one('a')
    #         #print(lower_taxa_a.get('href'),lower_taxa_a.b.string,lower_taxa_a.b.next_sibling)
    ############################################################################################
    related_objects = div_context[0].select_one('span[id="lblObjectList"]')
    #print(related_objects)
    if related_objects is not None:
        related_objects_trs = related_objects.find_all('tr')
        #print(related_objects_trs)
        for related_objects_tr in related_objects_trs:
            related_objects_tds = related_objects_tr.find_all('td')
            if len(related_objects_tds) == 2:
                related_objects_td_li = related_objects_tds[0].li
                if related_objects_td_li is not None:
                    li_a = related_objects_td_li.a
                    print(li_a.text,li_a.get('href'))
                else:
                    print(related_objects_tds[0].text)
                print(related_objects_tds[1].text)
            else:
                print('采集错误')

def get_foc_vol_list(url='http://www.efloras.org/index.aspx'): #从foc主页上获取foc卷册列表
    context = getHtml(url)
    context_soup = BeautifulSoup(context, "html.parser")
    span = context_soup.find_all('span',id='lblFloraList')
    url_list = []
    #print(span)
    if len(span) > 0:
        ul_list = span[0].find_all('ul')
        li_list = ul_list[2].find_all('li') #FOC在ul_list的第三个位置
        a_list = li_list[1].find_all('a')
        print(a_list)
        for a in a_list[1:]:  #a_list[1:]:
            a_href = a.get('href')
            print(' Volume :',a.text)
            #sql = "insert into volume (`url`,`volume_id`,`volume_no`,`create_date`,`create_by`,`del_flag`) values ('%s','%s','%s','%s','%s','%s')"
            if a_href is not None:
                url_list.append('http://www.efloras.org/' + a_href)
                volume_id = get_max_number(a_href)
                print('volume_id',str(volume_id))
            else:
                print('获取不到volume信息')

    else:
        print('未找到FOC卷册列表')

    return url_list

def get_foc_volume_list(volumes,index_url = 'http://www.efloras.org/',level = 0): # 根据卷册信息的地址找到科、属、种下属列表页，采集相关信息
    #url_list = []

    level = level + 1 #level = 1从科开始
    for vol in volumes:
        context = getHtml(vol)
        if context is None:
            continue
        context_soup = BeautifulSoup(context, "html.parser")
        div = context_soup.find_all('div', id='ucFloraTaxonList_panelTaxonList')
        volumeInfo = context_soup.select_one('span[id="ucVolumeInfo_lblVolumeInfo"]')
        volume_map = []
        if volumeInfo is not None:
            volumeInfo_table_trs = volumeInfo.table.find_all('tr')
            if len(volumeInfo_table_trs) > 0:
                for volumeInfo_table_tr in volumeInfo_table_trs:
                    volumeInfo_table_tds = volumeInfo_table_tr.find_all('td')
                    if len(volumeInfo_table_tds) == 2:
                        volume_map.append(volumeInfo_table_tds[1].text)
                    else:
                        volume_map.append('')
        if len(volume_map) != 5:
            for i in range(5-len(volume_map)): volume_map.append('')
        #print(volume_map)

        foc_taxon_chain = context_soup.select_one('span[id="ucFloraTaxonList_lblTaxonChain"]')
        parent_links = foc_taxon_chain.find_all('a')
        volume_list = foc_taxon_chain.find_all('a', href=re.compile("volume_id"), recursive=False)
        print(volume_list)

        if len(div) > 0:
            tr_list = div[0].find_all('tr',class_='underline')
            for tr in tr_list[2:]:
                td_list = tr.find_all('td') #科为四列，其他为五列，每一个都是一个物种信息
                wuzhong_list = {}
                wuzhong_list['parent_taxon_id'] = get_max_number(vol)
                wuzhong_list['type'] = str(level)
                wuzhong_list['type_name'] = ''
                wuzhong_list['taxon_name'] = ''
                wuzhong_list['title'] = volume_map[0]
                wuzhong_list['families'] =  volume_map[1]
                wuzhong_list['genera'] = volume_map[2]
                wuzhong_list['speces'] = volume_map[3]
                wuzhong_list['online_date'] = volume_map[4]
                wuzhong_list['taxon_id'] = td_list[0].text.strip()
                wuzhong_list['accepted_name'] = td_list[1].text.strip()
                wuzhong_detail_link_a = td_list[1].select_one('a')
                if wuzhong_detail_link_a:
                    wuzhong_list['accepted_name_url'] = index_url + wuzhong_detail_link_a.get('href')
                else:
                    wuzhong_list['accepted_name_url'] = ''
                wuzhong_list['accepted_name_cn'] = td_list[2].text.strip()
                wuzhong_list['lower_taxa'] = td_list[3].text.strip()
                lower_taxa_link_a = td_list[3].select_one('a')
                if lower_taxa_link_a:
                    wuzhong_list['lower_taxa_url'] = index_url + lower_taxa_link_a.get('href')
                else:
                    wuzhong_list['lower_taxa_url'] = ''
                if len(td_list) == 4:
                    if len(volume_list) == 1:
                        wuzhong_list['volume_no'] = get_max_number(volume_list[0].get('href'))
                        wuzhong_list['volume_name'] = volume_list[0].text
                    else:
                        wuzhong_list['volume_no'] = 0
                        wuzhong_list['volume_name'] = 0
                if len(td_list) == 5:
                    volume_link_a = td_list[4].select_one('a')
                    if volume_link_a:
                        wuzhong_list['volume_no'] = get_max_number(volume_link_a.get('href'))
                        wuzhong_list['volume_name'] = volume_link_a.text
                    else:
                        wuzhong_list['volume_no'] = 0
                        wuzhong_list['volume_name'] = 0

                print(wuzhong_list)

                wuzhong_list_sql(wuzhong_list)
                if wuzhong_list['accepted_name_url'] :
                    print('开始采集详细内容：',wuzhong_list['accepted_name_url'])
                    get_ke_context(wuzhong_list['accepted_name_url'])
                if wuzhong_list['lower_taxa_url'] :
                    print('开始采集：',wuzhong_list['accepted_name_cn'],'  的下级内容', wuzhong_list['accepted_name_url'])
                    url_list = []
                    url_list.append(wuzhong_list['lower_taxa_url'])
                    get_foc_volume_list(url_list,index_url,level)

        else:
            print('无法找到')
        volume_related_links_table = context_soup.find_all('table', id='ucVolumeResourceList_dataListResource')
        #print(volume_related_links_table)
        if len(volume_related_links_table) > 0:
            #print(volume_related_links_table[0])
            volumes_relateds = volume_related_links_table[0].find_all('tr',recursive=False) #搜索当前节点的直接子节点
            if len(volumes_relateds) > 0:
                #print(volumes_relateds)
                for volume in volumes_relateds[1:]:
                    trs=volume.find_all('tr')
                    if len(trs) > 0:
                        tds = trs[0].find_all('td')
                        if len(tds) > 1:
                            a = tds[0].select_one('a')
                            href = a.get('href')
                            print('--------',a.text,' ',href)
                            print('=====',tds[1].text)
                            sql1 = "insert into volume_related_links (`taxid`,`type`,`url`,`title`,`resource_type`,`files`,`create_date`,`create_by`,`del_flag`) values ('%s','%s','%s','%s','%s','%s','%s','%s','%s')"
                            if tds[1].text.strip() == 'PDF':
                                paths = href.split('/')
                                print(paths)
                                #download_file(href,'f://FloraData//' + paths[len(paths)-1])
                                #sql1 = sql1 % (vol.split('&')[0]),tds[1].text,href,a.text,paths[len(paths)-1],datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),'luoxuan',0)
                                sql1 = sql1 % (re.sub("\D", "", vol.split('&')[0]),tds[1].text,href,a.text,'PDF',paths[len(paths)-1],datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),'luoxuan',0)
                            else: #tds[1].text == 'Treatment'
                                #get_ke_context(href)
                                sql1 = sql1 % (re.sub("\D", "", vol.split('&')[0]),tds[1].text,href,a.text,'',tds[1].text,
                                               datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 'luoxuan', 0)
                            #print(insertMysql(sql1))
        else:
            print('无法找到volume_related_links')

def get_max_number(str): #获得连接中最大的数字
    return max(list(map(int,re.findall(r"\d+\.?\d*",str))))

def is_all_zh(s): #是否含有中文
    for ch in s:
        if u'\u4e00' <= ch <= u'\u9fff':
            return True
    return False

def insert_related_objects(related_objects):#插入相关内容到表中，返回当前的id
    sql = "insert into volume_related_links (`taxon_id`,`parent_taxon_id`,`type`,`url`,`parent_title`,`title`,`content`,`resource_type`,`files`,`create_date`,`create_by`,`del_flag`) " \
          "values ('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')"
    sql = sql % (related_objects['taxon_id'],related_objects['parent_taxon_id'],related_objects['type'],related_objects['url'],
                 related_objects['parent_title'],related_objects['title'],related_objects['content'],related_objects['resource_type'],
                 related_objects['files'],datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),'luoxuan',0)
    pid = insertMysql(sql)
    return pid

def insert_jiansuobiao(jiansuobiao): #插入检索表内容到表中，返回当前的id
    sql = "insert into volume_related_links (`taxon_id`,`first_no`,`first_no2`,`content`,`no_name`,`second_no`,`latin_name`,`goto_taxon_id`,`goto_taxon_url`,`create_date`,`create_by`,`del_flag`) " \
          "values ('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')"
    sql = sql % (jiansuobiao['taxon_id'],jiansuobiao['first_no'],jiansuobiao['first_no2'],jiansuobiao['content'],
                 jiansuobiao['no_name'],jiansuobiao['second_no'],jiansuobiao['latin_name'],jiansuobiao['goto_taxon_id'],
                 jiansuobiao['goto_taxon_url'],datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),'luoxuan',0)
    pid = insertMysql(sql)
    return pid

def wuzhong_detail_sql(volume_content): #插入详细内容到表中，返回当前插入的id
    small = ''
    if 'small' in volume_content  : small = volume_content['small']
    sql = "insert into volume_content (`url`,`content`,`taxon_id`,`parent_taxon_id`,`xuhao`,`latin_name`,`latin_name_full`,`zh_name`,`zh_name_pinyin`,`authors`,`authors_id`,`volume_id`,`volume`,`volume_title`,`create_date`,`del_flag`,`small`) values ('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')"
    sql = sql % (volume_content['url'],volume_content['content'],volume_content['taxon_id'],volume_content['parent_taxon_id'],volume_content['xuhao'],volume_content['latin_name'],
                 volume_content['latin_name_full'],volume_content['zh_name'],volume_content['zh_name_pinyin'],volume_content['authors'],volume_content['authors_id'],
                 volume_content['volume_id'],volume_content['volume'],volume_content['volume_title'],datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 0,small)
    pid = insertMysql(sql)
    return pid
def wuzhong_list_sql(wuzhong_list):
    sql = "insert into volume_ke (`parent_taxon_id`,`type`,`type_name`,`taxon_id`,`taxon_name`,`accepted_name`,`accepted_name_url`,`accepted_name_cn`,`lower_taxa`,`lower_taxa_url`,`volume_no`,`volume_name`,`title`,`families`,`genera`,`speces`,`online_date`,`create_date`,`create_by`,`del_flag`) values ('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')"
    sql = sql % (wuzhong_list['parent_taxon_id'], wuzhong_list['type'],
                 wuzhong_list['type_name'], wuzhong_list['taxon_id'],
                 wuzhong_list['taxon_name'],wuzhong_list['accepted_name'],
                 wuzhong_list['accepted_name_url'], wuzhong_list['accepted_name_cn'],
                 wuzhong_list['lower_taxa'], wuzhong_list['lower_taxa_url'],
                 wuzhong_list['volume_no'], wuzhong_list['volume_name'],
                 wuzhong_list['title'], wuzhong_list['families'],
                 wuzhong_list['genera'], wuzhong_list['speces'],
                 wuzhong_list['online_date'], datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                 'luoxuan', 0)
    pid = insertMysql(sql)
    return pid

def download_file(url,path): #下载文件
    print('Download file:',url,path)
    request = requests.get(url)
    with open(path, "wb") as code:
        code.write(request.content)

if __name__ == '__main__':
    #search_isbn()
    #print(html)
    #read07Excel('C:/Users/dell/Desktop/书单：PDA_全库（2015）_20180621 科学文库书单第二版2.xlsx')

    #get_page_html()

    #get_ke_context('http://www.efloras.org/florataxon.aspx?flora_id=2&taxon_id=250098342')
    #get_ke_context('http://www.efloras.org/florataxon.aspx?flora_id=2&taxon_id=20790')
    #get_text_docx()
    #read07_excel('C:/Users/dell/Desktop/高等二卷.xlsx')
    #mings = ['f','fsdf','fsdf1','fsdfs','fsdfs']
    #print(mings[2:len(mings)])
    # i = 0
    # datas = get_foc();
    # for data in datas:
    #     i = i + 1
    #     print(data)
    #     if i >= 10:break
    lists = get_foc_vol_list()
    ##print(lists)
    get_foc_volume_list(lists)
    #print(getHtml('http://flora.huh.harvard.edu/FloraData/002/Vol11/foc11-Preface.htm'))
    #print(get_page_html())
    #vol = 'http://www.efloras.org/browse.aspx?flora_id=2&start_taxon_id=103074,volume_page.aspx?volume_id=2002&flora_id=2'

    #print(is_all_zh('剑叶铁角蕨 jian ye tie jiao jue'))
    #print(is_all_zh('jian ye tie jiao jue'))
    #print(re.findall(r"\d+\.?\d*",vol),get_max_number(vol))