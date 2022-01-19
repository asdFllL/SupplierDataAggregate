import pandas as pd

'''
将归还日期提成列表，设置成索引，然后遍历列表并和今天日期组成series后求时间差，
不处于0-7期间则删除后，成立新的表格并发送
'''
today = pd.to_datetime('2022-06-10')
basic_data = pd.read_csv(
    r'/Users/ky_c/PycharmProjects/sgd/checking_list/11/demo.csv')
check_list = basic_data['归还截至日期'].tolist()
basic_data = basic_data.set_index('归还截至日期')

for i in check_list:
    j = pd.to_datetime(i)
    k = j - today
    k = k.days
    try:
        if k == 3:
            continue
        else:
            basic_data = basic_data.drop(i)
    except Exception as e:
        pass
    continue

basic_data.to_csv('basic.csv')