#!/usr/bin/env python
# encoding: utf-8
'''
@Author: Joven Chu
@Email: jovenchu@163.com
@File: Model_analysis.py
@Time: 2020-01-08 17:38
@Project: 模型数据分析
@About: 联合多个表格进行数据的统计分析，保存为表格形式
'''

import pandas as pd
import numpy as np
from datetime import datetime
import re,os
import openpyxl
import matplotlib.pyplot as plt


# Excel表格的读取路径
a_path = "audit_report.xlsx"
b_path = "audit_report.xlsx"
# 分析数据结果保存的路径
output_path = "audit_result.xlsx"

def Excel2df(path):
    """
    读取表格将其转为pandas的DataFrame对象格式
    :param path: 表格路径
    :return: Dataframe对象格式
    """
    excel_df = pd.DataFrame(pd.read_excel(path)) # 默认获取表格的第一个表单（sheet0）
    return excel_df

# 获取每个表的关键字段
a_df = Excel2df(a_path)
b_df = Excel2df(b_path)

# 数据概览
alldata_list = []
# 定义结果表的title
alldata_title = ["批次","推送时间","数据量总和","平均时效"]
alldata_list.append(alldata_title)

# 功能1 获取一个字段的不同值集合
model_name = '费用'
df_model = a_df.loc(a_df['模型名称']==model_name)
batchall = df_model[['批次']].values.T.tolist()[:][0]
batch = list(set(batchall))


# 功能2 获取一个字段的第一个值
push_time = b_df[['推送时间']].values[0][0]

# 功能3 获取一个字段的值总和
number_all = a_df['数据总和'].sum()
number1 = a_df['数据'].sum()

# 功能4 求两值相除的百分比
number_efficiency = '%.2f%%' % ((number1 / number_all) * 100)

# 功能4 某两个时间字段的逐一相减，求平均值
# （1）timedelta64[ns]格式的时间列表
onsite_time = (pd.to_datetime(a_df['时间1']) - pd.to_datetime(a_df['时间2'])).astype('timedelta64[D]')
onsite_average_time = np.mean(onsite_time.values.T.tolist()[:][0])

# (2)Timedelta格式，先求字段时间的最大值
onsite_date_max = pd.to_datetime(max(a_df['1时间'].values.T.tolist()[:]))
apply_date_max = pd.to_datetime(max(a_df['2时间'].values.T.tolist()[:]))
onsite_batch_date = (onsite_date_max - apply_date_max).day # 转换成天数（int）

# 功能5 消除列表中的空值
incoms = []
incoms = [incom for incom in incoms if str(incom) != 'nan']

# 功能6 保存数据到表格中
summaryDataFrame = pd.DataFrame(alldata_list)
summaryDataFrame.to_excel(output_path, encoding='utf-8', index=False, header=False)