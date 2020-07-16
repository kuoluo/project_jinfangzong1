
import urllib.request
import re
import xlwt
import time
from lxml import etree

#htm = 'http://bf.win007.com/football/Next_20200715.htm'
#need_name_list = ['英超', '英冠', '英甲', '德甲', '德乙', '意甲', '法甲', '法乙', '荷兰超', '荷乙', '日职联', '日职乙', '韩K联', '美职', '巴甲', '欧罗杯',
#                  '欧冠', '欧洲杯']

htm = 'http://bf.win007.com/football/Next_20200715.htm'
need_name_list = ['英超', '英冠', '英甲', '德甲', '德乙', '意甲', '法甲', '法乙', '荷兰超', '荷乙', '日职联', '日职乙', '韩K联', '美职', '巴甲', '欧罗杯',
                  '欧冠', '欧洲杯','葡超','挪超', '俄超','日职乙','瑞典超']
def create_html(ht):
    html_list = []
    file = urllib.request.urlopen(ht)
    wb_data = file.read().decode('GBK')
    html = etree.HTML(wb_data)
    name_list = html.xpath('/html/body/div[@class="resultBox"]/div[@class="content"]/table/tr/td/font/text()')
    num_list = html.xpath('/html/body/div[@class="resultBox"]/div[@class="content"]/table/tr/td/a/@onclick')
    new_num_list = []
    for i in num_list:
        if 'EuropeOdds' in i:
            new_num_list.append(re.sub(r'\D', "", i))
    for i in new_num_list:
        the_url = 'http://op1.win007.com/oddslist/' + i + '.htm'
        html_list.append(the_url)
    return html_list, name_list


nlist = []
sum = []


# 需要数据的下标
countlist = [3, 4, 5, 6, 7, 8, 9, 17, 18, 19]

# 一定要写成凯利指数123，因为默认不重复
secondrow = ['国家', '主队', '客队', '赛果', '胜', '平', '负', '主胜率', '和率', '客胜率', '返还率', '凯利指数1', '凯利指数2', '凯利指数3']

# 公司一 写入的列
countcol = [4, 5, 6, 13, 14, 15, 16, 17, 18, 19]


def reurl(url):
    """
    这个函数是从url网址利用正则提取里面js的网址
    :param url: 传入网址的url
    :return: 返回js的url
    """
    pointnum = re.search(r'oddslist/(.*).htm', url).group(1)
    urljs = 'http://1x2d.win007.com/' + pointnum + '.js'
    return urljs


def rematch(temdata):
    """
    读取到的数据正则表达式提取联赛名字、主场队伍名字、客场队伍名字
    :param temdata:爬取的js数据
    :return:是一个列表，里面的参数是联赛名字、主场队伍名字、客场队伍名字
    """
    matchname_cn = re.search(r'matchname_cn="(.*)"', temdata).group(1)
    hometeam_cn = re.search(r'hometeam_cn="(.*)"', temdata).group(1)
    guestteam_cn = re.search(r'guestteam_cn="(.*)"', temdata).group(1)
    defsum = [matchname_cn, hometeam_cn, guestteam_cn]
    return defsum


def renumber(temdata):
    """
    这个函数是处理的具体nlist
    :param temdata:
    :return:
    """
    threedata = []
    num = re.search(r'Array\((.*)\)', temdata).group(1)
    new = re.split('","', num)
    for i in new:
        if '365(英国)' in i:
            threedata.append(i)
        elif '威廉希尔' in i:
            threedata.append(i)
        elif '利记' in i:
            threedata.append(i)
    bet_jia = '"9999|95630973|Bet 365|99999|99999|99999|99999|99999|99999|99999|99999|99999|99999|99999|99999|99999|99999|99999|99999|99999|2020,07-1,13,14,26,00|bet 365(英国)|1|0'
    weilianxier_jia = '"9999|95630973|Bet 365|99999|99999|99999|99999|99999|99999|99999|99999|99999|99999|99999|99999|99999|99999|99999|99999|99999|2020,07-1,13,14,26,00|威廉希尔(英国)|1|0'
    liji_jia = '"9999|95630973|Bet 365|99999|99999|99999|99999|99999|99999|99999|99999|99999|99999|99999|99999|99999|99999|99999|99999|99999|2020,07-1,13,14,26,00|利记sbobet(英国)|1|0'
    tep_list = [0, 0, 0]
    for i in threedata:
        if '365(英国)' in i:
            tep_list[0] = 1
        if '威廉希尔' in i:
            tep_list[1] = 1
        if '利记' in i:
            tep_list[2] = 1
    if tep_list[0] == 0:
        threedata.append(bet_jia)
    if tep_list[1] == 0:
        threedata.append(weilianxier_jia)
    if tep_list[2] == 0:
        threedata.append(liji_jia)
    #threedata = new[0:3]
    print(threedata)
    for i in threedata:
        nlist.append(i.split('|'))


def allwritexls(nl):
    """
    这个函数是将输入的全部url爬下来的数据写入excel
    :param sg:是赛果数据
    :param nl: 传入的数据
    :return:
    """
    now = time.strftime('%Y-%m-%d %H%M%S', time.localtime())
    # 以ASCII码的形式创建excel
    wb = xlwt.Workbook(encoding='ascii')
    # 追加列表，后面的是可以重复写入
    ws = wb.add_sheet(now, cell_overwrite_ok=True)
    # 设置列宽
    ws.col(1).width = 256 * 16
    ws.col(2).width = 256 * 16

    # for index, el in enumerate(sg):
    #     ws.write(index + 2, 3, el)

    # 写入第0行
    # dzh = [4, 7, 10, 13, 20, 27]
    dzh = [7, 4, 10, 13, 20, 27]  # 调换两列

    for iindex, i in enumerate(dzh):
        ws.write(0, i, nl[iindex % 3][21])
    # 写入第一行
    d1 = secondrow[0:4]
    d2 = secondrow[4:7]
    d3 = secondrow[7:]
    for iindex, i in enumerate(d1):
        ws.write(1, iindex, i)
    tc = 4
    while tc <= 12:
        ws.write(1, tc, d2[(tc - 4) % 3])
        tc = tc + 1
    td = 13
    while td <= 33:
        ws.write(1, td, d3[(td - 13) % 7])
        td = td + 1

    # 写入剩下行
    for iindex, i in enumerate(sum):
        for jindex, j in enumerate(i):
            ws.write(iindex + 2, jindex, label=j)  # 写入国家、主、客
    for cindex, c in enumerate(nl):
        t = int(cindex / 3)
        if (cindex + 1) % 3 == 2:
            for cl, k in zip(countcol, countlist):
                ws.write(t + 2, cl, label=c[k])
        elif (cindex + 1) % 3 == 1:
            for cl, k in zip(countcol, countlist):
                if cl < 10:
                    ws.write(t + 2, cl + 3, label=c[k])
                else:
                    ws.write(t + 2, cl + 7, label=c[k])
        elif (cindex + 1) % 3 == 0:
            for cl, k in zip(countcol, countlist):
                if cl < 10:
                    ws.write(t + 2, cl + 6, label=c[k])
                else:
                    ws.write(t + 2, cl + 14, label=c[k])

    wb.save(now + '.xls')


def getdata(ht, name):
    """
    这个函数是用来获取js的数据的
    :param ht: 是输入的一组原始页面的url（html）
    :return: 返回的爬取到的数据
    """
    for i_index, i in enumerate(ht):
        url_js = reurl(i)
        # 爬取urljs页面，保存到file
        print(url_js)
        try:
            file = urllib.request.urlopen(url_js)
        except:
            ht.remove(i)
            name.remove(name[i_index])
        # file读取的数据解码 保存在data中
        else:
            data = file.read().decode()
            sum.append(rematch(data))
            renumber(data)

    allwritexls(nlist)


def choose_name(need_name_list, now_name_list, now_html_list):
    tep_name_list = []
    tep_html_list = []
    for i in range(len(now_name_list)):
        if now_name_list[i] in need_name_list:
            tep_name_list.append(now_name_list[i])
            tep_html_list.append(now_html_list[i])
    return tep_name_list, tep_html_list


html_list, name_list = create_html(htm)
# 特定队伍
name_list, html_list = choose_name(need_name_list, name_list, html_list)
getdata(html_list, name_list)
