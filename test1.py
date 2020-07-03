# @Time : 2020/7/3 16:19 

# @Author : 于川清

# @File : test1.py 

# @Software: PyCharm
import xlrd
import matplotlib
import pyecharts


# import NetworkX

def readxls(filename):
    data = xlrd.open_workbook(filename).sheets()[0]

    print(data)
    print(type(data))


readxls('biao.xls')
