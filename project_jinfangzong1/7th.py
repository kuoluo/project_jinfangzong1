import urllib.request
import xlwt
import time
from lxml import etree
import json

ht = 'http://www.okooo.com/jingcai/shuju/zhishu/'

def get_html(url):
    headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                             'Chrome/80.0.3987.149 Safari/537.36'}
    page1 = urllib.request.Request(url, headers=headers)
    page = urllib.request.urlopen(page1)
    html = page.read()
    return html


def get_js(url):
    headers = {
        'Referer': 'http: // www.okooo.com / soccer / match / 1097704 / okoooexponent /',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.149 Safari/537.36'}
    page1 = urllib.request.Request(url=url, headers=headers, method='POST')
    page = urllib.request.urlopen(page1)
    html = page.read()
    return html


def create_html(ht):
    # file = urllib.request.urlopen(ht)
    # wb_data = file.read().decode('GBK')
    file = get_js(ht)
    wb_data = file.decode('GBK')
    html = etree.HTML(wb_data)
    temp_list = html.xpath('/html/body[@class="magazine"]/div[@class="chartboxbg"]/div[@class="card"]/div['
                           '@class="Clear container_wrapper magazineDate"]/table['
                           '@class="magazine_table"]/tr/td/span/a/@href')
    final_list = []
    for i in temp_list:
        tep = 'http://www.okooo.com' + i
        final_list.append(tep)
    return temp_list, final_list


def get_name(url):
    file = get_js(url)
    wb_data = file.decode('GBK')
    html = etree.HTML(wb_data)
    name = html.xpath('/html/head/meta[@name="keywords"]/@content')
    return name


def get_js_data(url):
    tep_list = []
    file = get_js(url)
    wb_data = file.decode('GBK')
    wb_data = json.loads(wb_data)
    t1 = wb_data["odds"]["start"]
    t2 = wb_data["odds"]["end"]
    for key in t1:
        tep_list.append(t1[key])
    for key in t2:
        tep_list.append(t1[key])
    return tep_list


def get_all_data(num_list, html_list):
    # 首先处理ajax数据
    all_data = []
    for i, j in zip(num_list, html_list):
        aoke_html = 'http://www.okooo.com' + i + 'xmlData/?type=okooo'
        kaili_html = 'http://www.okooo.com' + i + 'xmlData/?type=okoooexponent'
        name = get_name(j)
        # print(name)
        # print(type(name))
        aoke_list = get_js_data(aoke_html)
        kaili_list = get_js_data(kaili_html)
        new_list = name + aoke_list + kaili_list
        # all_data.append([name, aoke_list, kaili_list])
        all_data.append(new_list)
    return all_data


def write_xls(all_data):
    now = time.strftime('%Y-%m-%d %H%M%S', time.localtime())
    # 以ASCII码的形式创建excel
    wb = xlwt.Workbook(encoding='ascii')
    # 追加列表，后面的是可以重复写入
    ws = wb.add_sheet(now, cell_overwrite_ok=True)
    ws.col(0).width = 256 * 25
    for i in range(len(all_data)):
        for j in range(len(all_data[i])):
            ws.write(i, j, all_data[i][j])
    wb.save('aoke'+now + '.xls')


numlist, htmllist = create_html(ht)
data = get_all_data(numlist, htmllist)
write_xls(data)
# http://www.okooo.com/soccer/match/1097704/okoooexponent/xmlData/?type=okooo#澳客
# http://www.okooo.com/soccer/match/1097704/okoooexponent/xmlData/?type=okoooexponent#凯莉离散度
