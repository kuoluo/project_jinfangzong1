# @Time : 2020/7/3 16:19 

# @Author : 于川清

# @File : test1.py 

# @Software: PyCharm

import json
import networkx as nx
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


def relation_graph():
    with open("weibo.json", "r", encoding="utf-8") as f:
        j = json.load(f)
        nodes, links, categories, cont, mid, userl = j
    c = (
        Graph()
            .add(
            "",
            nodes,
            links,
            categories,
            repulsion=50,
            linestyle_opts=opts.LineStyleOpts(curve=0.2),
            label_opts=opts.LabelOpts(is_show=False),
        )
            .set_global_opts(
            legend_opts=opts.LegendOpts(is_show=False),
            title_opts=opts.TitleOpts(title="Graph-微博转发关系图"),
        )
            .render("graph_weibo.html")
    )


def write_data(xls_list):
    temp_dict = {}
    data_list = []
    for i in xls_list:
        temp_dict['name'] = i[3]+i[6][:]
        temp_dict['des'] = i[4]
        data_list.append(temp_dict.copy())
    return data_list


def temp_source_target(xls_list):
    temp_relation_list = []
    tmp = {}
    xls_len = len(xls_list)
    for i_index in range(xls_len):
        if xls_list[i_index][4] == '户主':
            j = 1
            tmp['source'] = xls_list[i_index][3] + xls_list[i_index][6]
            while i_index + j < xls_len and xls_list[i_index + j][4] != '户主':
                tmp['target'] = xls_list[i_index + j][3] + xls_list[i_index+j][6]
                tmp['name'] = xls_list[i_index + j][4]
                tmp['des'] = xls_list[i_index + j][3] + '是' + xls_list[i_index][3] + '的' + tmp['name']
                temp_relation_list.append(tmp.copy())
                j += 1
    return temp_relation_list

def hujian_source_target(xls_list):
    xls_len = len(xls_list)
    temp_relation_list = []
    tmp = {}
    for i in range(xls_len):
        for j in range(xls_len):
            if i!=j:
                if xls_list[i][3][0]==xls_list[j][3][0] and xls_list[i][3][1]==xls_list[j][3][1]:
                    tmp['source'] = xls_list[i][3] + xls_list[i][6]
                    tmp['target'] = xls_list[j][3] + xls_list[j][6]
                    tmp['name'] = '兄弟'
                    tmp['des'] = xls_list[i][3] + '是' + xls_list[j][3] + '的' + tmp['name']
                    temp_relation_list.append(tmp.copy())
                if len(xls_list[j][3])==3 and len(xls_list[i][3])==3 and xls_list[i][3][2] is not None and xls_list[j][3][2] is not None:
                    if xls_list[i][3][0] == xls_list[j][3][0] and xls_list[i][3][2] == xls_list[j][3][2]:
                        tmp['source'] = xls_list[i][3] + xls_list[i][6]
                        tmp['target'] = xls_list[j][3] + xls_list[j][6]
                        tmp['name'] = '兄弟'
                        tmp['des'] = xls_list[i][3] + '是' + xls_list[j][3] + '的' + tmp['name']
                        temp_relation_list.append(tmp.copy())
    return temp_relation_list
# tmp = {'source': '高育良', 'target': '侯亮平', 'name': '师生', 'des': '侯亮平是高育良的得意门生'}


relation_list = ['户主', '配偶', '子', '女', '孙', '父母', '祖父母', '兄弟', '姐妹', '儿媳', '非亲属']
xls = read_xls('biao.xls')
print(write_data(xls))
print(len(write_data(xls)))
print(temp_source_target(xls).append(hujian_source_target(xls)))
