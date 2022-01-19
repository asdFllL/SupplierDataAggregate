import pandas as pd
import os
import re
import numpy as np

# 导入汇总名单并整理
data_list = pd.read_excel(r'/Users/ky_c/PycharmProjects/sgd/checking_list/产品信用户汇总1227.xlsx')
data_list['账号'] = data_list['账号'].astype(str)
use_list = data_list['账号'].tolist()
data_list = data_list.set_index('账号')


# 设定整理方法


path = r'/Users/ky_c/PycharmProjects/sgd/checking_list/11/'

# TODO 缺少其他券商方法
# 广发证券
def guang_fa(filename):
    data = pd.read_excel(root + '/' + filename, index_col=0)
    data = data.loc['融券负债合约明细':'其他负债合约明细',
                    ['Unnamed: 1', 'Unnamed: 3', 'Unnamed: 5', 
                     'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 8']]
    data = data.set_axis(data.iloc[1], axis=1, inplace=False)
    data = data.drop(['其他负债合约明细'])
    data = data.drop(data.index[[0, 1]])
    index_list = range(0, len(data.index))
    data.loc[:, 'index'] = index_list
    data = data.set_index('index')
    # 添加判断是否为空表
    if data.iloc[0, 0] is np.nan:
        print('空表')
    else:
        return data


# 遍历文件夹
data_sum_list = []
for root, dirs, filenames in os.walk(path):
    for filename in filenames:
        if re.search('([(]\d+[)]?)', filename): # 分解文件名，匹配列表判断
            a = re.search('([(]\d+[)]?)', filename)
            a = a.group()
            a = a.replace('(', '').replace(')', '')
            if a in use_list:
                print('需要方法')
        elif re.search('(\d{8,})', filename):
            b = re.search('(\d{8,})', filename)
            b = b.group()
            if b in use_list:
                company_name = data_list.loc[b, '经纪公司']
                if company_name == '广发证券':
                    data = guang_fa(filename)
                    data.loc[:, '产品'] = data_list.loc[b, '产品名称']
                    data.loc[:, '团队'] = data_list.loc[b, '团队名称']
                    data_sum_list.append(data)
data = pd.concat(data_sum_list, join='outer')
data = data.replace('/', np.nan)
data = data.dropna()
data = data.drop_duplicates(subset = '合约编号')
data['归还截至日期'] = pd.to_datetime(data['归还截至日期'])
index_list = range(0, len(data.index))
data.loc[:, 'index'] = index_list
data = data.set_index('index')
data.to_csv('demo.csv')




