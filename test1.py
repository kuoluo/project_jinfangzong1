# @Time : 2020/7/3 16:19 

# @Author : 于川清

# @File : test1.py 

# @Software: PyCharm
import xlrd
import matplotlib
import pyecharts

# import NetworkX
"户口簿户号"
"户主姓名"
"成员姓名"
"与户主关系"
"身份证号码"
"出生日期"
"性别"
"拟测劳龄"
"存在状态"
"取得该存在状态原因"
"户籍地址"
"是否取得承包地"

diccontent = {"户口簿户号": 0, "户主姓名": 0, "成员": 0}


def readxls(filename):
    """
    这是读取biao.xls的一个函数
    :param filename: xls的名字
    :return: 返回一个列表加字典的嵌套数据
    """
    data = xlrd.open_workbook(filename).sheets()[0]
    nrows = data.nrows
    for i in range(0, nrows):
        print(data.cell(i, 0).value)
    print(data)
    print(type(data))


readxls('biao.xls')
