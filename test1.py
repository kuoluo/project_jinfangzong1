# @Time : 2020/7/3 16:19 

# @Author : 于川清

# @File : test1.py 

# @Software: PyCharm


import pyecharts
import networkx
import matplotlib.pyplot as plt
import seaborn as sns
import xlrd
import pandas as pd
from pyecharts import options as opts
from pyecharts.charts import Graph
from pyecharts.charts import Pie


def read_xls(filename):
    """
    这个函数是用来读取excel里面的数据
    :param filename: excel的文件名称
    :return:
    """
    # 读取总的数据表
    data = xlrd.open_workbook(filename).sheets()[0]
    num_rows = data.nrows  # 获取该sheet中的有效行数
    dic_xls = {'户口簿户号': 0, '户主姓名': 0, '成员': 0}
    xls_list = []
    for i in range(num_rows):
        t = data.row_values(i)  # 返回由该行中所有的单元格对象组成的列表
        xls_list.append(t)

    return xls_list


print(read_xls('biao.xls')[0][7])

print(read_xls('biao.xls')[0][6][:4])


def cal_edge(xls_list):
    male_edge_list = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    female_edge_list = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    edge_list = []
    print(xls_list)
    for i in xls_list:
        print(i)
        t = int((2020 - int(i[6][:4])) / 10)
        print(t)
        if i[7] == '男':
            male_edge_list[t] += 1
        elif i[7] == '女':
            female_edge_list[t] += 1
    for i, j in zip(male_edge_list, female_edge_list):
        edge_list.append(i + j)

    return male_edge_list, female_edge_list, edge_list


def plot_people(people_data):
    index_list = ['0-9', '10-19', '20-29', '30-39', '40-49', '50-59', '60-69', '70-79', '80-89', '90-99']
    pyramid_data = {'male_edge_list': people_data[0], 'female_edge_list': [-i for i in people_data[1]],
                    'index': index_list, }
    order_list = index_list.reverse()
    pyramid_data = pd.DataFrame(pyramid_data)
    pyramid_data.columns
    bar_plot = sns.barplot(y='index', x="male_edge_list", color="blue", data=pyramid_data,
                           order=order_list, )
    bar_plot = sns.barplot(y='index', x="female_edge_list", color="red", data=pyramid_data,
                           order=order_list, )
    plt.xticks([-150, -120, -90, -70, -50, -30, 0, 30, 50, 70, 90, 120, 150],
               [150, 120, 90, 70, 50, 30, 0, 30, 50, 70, 90, 120, 150])
    # sns is seaborn alias
    plt.rcParams['font.sans-serif'] = ['SimHei']
    plt.rcParams['axes.unicode_minus'] = False
    bar_plot.set(xlabel="人口数量", ylabel="年龄层", title="某村落人口年龄结构金字塔")
    plt.show()


plot_people(cal_edge(read_xls('biao.xls')))
