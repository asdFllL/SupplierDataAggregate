import os
import pandas as pd
import re


# 载入信用户账号列表
data_list = pd.read_excel('产品信用户汇总1227.xlsx')
data_list['账号'] = data_list['账号'].astype(str)
use_list = data_list['账号'].tolist()

# 读取所有文件并保留需求信用户账号文件
path = r'/Users/ky_c/PycharmProjects/sgd/checking_list/11'

for root, dirs, filenames in os.walk(path):
    a = 1
    b = 1
    c = 1
    for filename in filenames:
        if filename.startswith('030680180261-信用'):
            NewName = '81205121' + '_' + str(a) + '.xlsx'
            os.rename(
                os.path.join(root, filename),
                os.path.join(root, NewName))
            print('文件%s成功重命名为%s' % (filename, NewName))
            a += 1
        elif filename.startswith('030680181168-信用'):
            NewName = '81205155' + '_' + str(b) + '.xlsx'
            os.rename(
                os.path.join(root, filename),
                os.path.join(root, NewName))
            print('文件%s成功重命名为%s' % (filename, NewName))
            b += 1
        elif filename.startswith('121200388388-信用'):
            NewName = '69501821' + '_' + str(c) + '.xlsx'
            os.rename(
                os.path.join(root, filename),
                os.path.join(root, NewName))
            print('文件%s成功重命名为%s' % (filename, NewName))
            c += 1
        elif re.search('([(]\d+[)]?)', filename):
            account = re.search('([(]\d+[)]?)', filename)
            account = account.group()
            account = account.replace('(', '').replace(')', '')
            if account not in use_list:
                dir_filename = os.path.join(root, filename)
                os.remove(dir_filename)
        elif re.search('(\d{7,})', filename):
            account = re.search('(\d{7,})', filename)
            account = account.group()
            if account not in use_list:
                dir_filename = os.path.join(root, filename)
                os.remove(dir_filename)
        elif filename.endswith('.RAR', ):
            os.remove(os.path.join(root, filename))
        elif filename.endswith('.TXT',):
            os.remove(os.path.join(root, filename))
        elif filename.endswith('.txt', ):
            os.remove(os.path.join(root, filename))
        elif filename.endswith('.rar',):
            os.remove(os.path.join(root, filename))
        elif filename.endswith('.ZIP',):
            os.remove(os.path.join(root, filename))
        else:
            continue

# 清洗完毕 清除空文件夹
dir_list = []
for root, dirs, filenames in os.walk(path):
    dir_list.append(root)
for root in dir_list[::-1]:
    if not os.listdir(root):
        os.rmdir(root)


