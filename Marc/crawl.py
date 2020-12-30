#coding:utf-8
from urllib import request
import re
import socket
import sys


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

def pick_charset(html):
    """
    从文本中提取 meta charset
    :param html:
    :return:
    """
    charset = None
    m = re.compile('<meta .*(http-equiv="?Content-Type"?.*)?charset="?([a-zA-Z0-9_-]+)"?', re.I).search(html)
    if m and m.lastindex == 2:
        charset = m.group(2).lower()
    return charset


def download_page(url,timeout=10):
    response = request.urlopen(url,timeout=100)
    code = response.getcode()
    info = response.info()
    charset = None
    try:
        if info :
            m = re.findall(r'charset=([a-zA-Z0-9_-]+)', re.I)
            if m:
                charset = str(m[0]).lower()
        if code == 200:
            html = response.read()
            if not charset and html:
                charset = pick_charset(html)

            # 如果完全采不到 charset,默认使用 gbk 反正都是乱码
            if not charset or charset == "gb2312":
                charset = 'gbk'

            if charset and charset != 'utf-8':
                try:
                    html = html.decode(charset).encode('utf-8')
                except:
                    pass
        else:
            html = ''
        print(html,charset)
        return (code,response.geturl(),charset,html)

    except request.URLError as e:
        return str("%r" % e)
    except socket.timeout as e:
        return str("%r" % e)
    except:
        return str(sys.exc_info())


if __name__ == '__main__':
    info = download_page('http://flora.huh.harvard.edu/FloraData/002/Vol11/foc11-Preface.htm')
    print(info)