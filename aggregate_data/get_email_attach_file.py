import poplib
import pandas as pd
from email.parser import Parser
from email.header import decode_header, Header
from email.utils import parseaddr

email = str(input('plz enter account:'))
password = str(input('plz enter account`s password:'))


# 设置券商券池读取
def ZhongJing(fileName):
    zj = pd.read_excel(fileName, usecols=['证券代码', '证券名称', '合约到期日', '费率', '数量'])
    zj.loc[:, 'Broker ID'] = '中金'
    zj = zj.rename(columns={'费率': '利率'})
    zj['利率'] = zj['利率'].map(lambda x: x * 0.01)
    zj['利率'] = zj['利率'].map(lambda x: format(x, '.2%'))
    zj = zj.rename(columns={'数量': '可融数量'})
    zj['可融数量'] = zj['可融数量'].astype(int)
    return zj


def ZhongXin(fileName):
    zx = pd.read_excel(fileName, usecols=['证券代码', '证券名称', '期限', '数量上限'])
    zx.loc[:, 'Broker ID'] = '中信'
    zx = zx.rename(columns={'数量上限': '可融数量'})
    zx['可融数量'] = zx['可融数量'].astype(int)
    return zx


def YinHe(fileName):
    import pandas as pd

    sheet = pd.read_excel(fileName, sheet_name=None)
    sheet_list = []
    for i in list(sheet.keys()):
        if i == '非公募券单':
            yinhe1 = sheet[i].iloc[:, 1:5]
            yinhe1.columns = ['证券代码', '证券名称', '期限', '可融数量']
            yinhe1 = yinhe1.dropna()
            sheet_list.append(yinhe1)

        elif i == '公募券单':
            yinhe2 = sheet[i].iloc[:, 1:5]
            yinhe2.columns = ['证券代码', '证券名称', '期限', '可融数量']
            yinhe2 = yinhe2.dropna()
            sheet_list.append(yinhe2)

        elif i == '公募券单（可按篮子或个股出借）':
            df_list = []
            for x in range(3, 47, 4):
                y = x + 2
                df = sheet[i].iloc[:, x:y]
                df = df.set_axis(df.iloc[0], axis=1, inplace=False)
                df = df.drop(index=0)
                df = df.dropna()
                df_list.append(df)
            yinhe3 = pd.concat(df_list, join='outer')
            sheet_list.append(yinhe3)

        elif i == '其他公募券单':
            df_list = []
            for i in range(3, 50, 4):
                j = i + 3
                df = sheet[i].iloc[:, i:j]
                df = df.set_axis(df.iloc[0], axis=1, inplace=False)
                df.columns = ['证券代码', '证券名称', '可融数量']
                df = df.drop(index=0)
                df = df.dropna()
                df_list.append(df)
            yinhe4 = pd.concat(df_list, join='outer')
            sheet_list.append(yinhe4)

        else:
            continue
    yh = pd.concat(sheet_list, join='outer')
    yh.loc[:, 'Broker ID'] = '银河'
    return yh

def DongWu(fileName):
    dw = pd.read_excel(fileName, usecols=['证券代码', '证券名称', '到期日期', '可用数量'])
    dw.loc[:, 'Broker ID'] = '东吴'
    dw = dw.rename(columns={'到期日期': '合约到期日'})
    dw = dw.rename(columns={'可用数量': '可融数量'})
    dw['可融数量'] = dw['可融数量'].astype(int)
    return dw


def GuangFa(fileName):
    gf = pd.read_excel(fileName, usecols=['证券代码', '证券名称', '期限', '利率', '数量(股)'])
    gf.loc[:, 'Broker ID'] = '广发'
    gf['利率'] = gf['利率'].replace('待定', 0)
    gf['利率'] = gf['利率'].astype(float)
    gf['利率'] = gf['利率'].map(lambda x: format(x, '.2%'))
    gf = gf.rename(columns={'数量(股)': '可融数量'})
    gf['可融数量'] = gf['可融数量'].astype(int)
    return gf


def DongFang(fileName):
    df = pd.read_excel(fileName, usecols=['证券代码', '证券名称', '期限', '费率'])
    df.loc[:, 'Broker ID'] = '东方'
    df = df.rename(columns={'费率': '利率'})
    return df


def HuaTai(fileName):
    sheet1 = pd.read_excel(fileName, sheet_name='T0券单-库存标的',
                           usecols=['证券代码', '证券名称', '期限(天)', '利率', '证券数量(万股)'])

    sheet2 = pd.read_excel(fileName, sheet_name='T0券单-战略配售标的',
                           usecols=['证券代码', '证券名称', '期限(天)', '利率', '证券数量(万股)'])

    sheet3 = pd.read_excel(fileName, sheet_name='T1券单-转融通来源组借入标的',
                           usecols=['证券代码', '证券名称', '期限(天)', '利率'])

    sheet4 = pd.read_excel(fileName, sheet_name='T1券单-客户持仓标的',
                           usecols=['证券代码', '证券名称', '期限(天)', '利率', '证券数量(万股)'])

    sheet5 = pd.read_excel(fileName, sheet_name='T1券单-转融通市场行情标的',
                           usecols=['证券代码', '证券名称', '期限(天)', '利率', '证券数量(万股)'])
    sheet1_5 = pd.concat([sheet1, sheet2, sheet3, sheet4, sheet5], axis=0, join='outer')
    sheet1_5['证券数量(万股)'] = sheet1_5['证券数量(万股)'].str.replace(',', '')
    sheet1_5['证券数量(万股)'] = sheet1_5['证券数量(万股)'].astype(float)
    sheet1_5 = sheet1_5.rename(columns={'证券数量(万股)': '证券数量(份)'})
    sheet1_5 = sheet1_5.fillna(0)
    sheet1_5['证券数量(份)'] = sheet1_5['证券数量(份)'].map(lambda x: x * 10000)
    sheet1_5['证券数量(份)'] = sheet1_5['证券数量(份)'].astype(int)
    sheet6 = pd.read_excel(fileName, sheet_name='库存ETF标的',
                           usecols=['证券代码', '证券名称', '期限(天)', '利率', '证券数量(份)'])
    sheet6['证券数量(份)'] = sheet6['证券数量(份)'].str.replace(',', '')
    sheet6['证券数量(份)'] = sheet6['证券数量(份)'].astype(float)
    sheet6['证券数量(份)'] = sheet6['证券数量(份)'].astype(int)
    ht = pd.concat([sheet1_5, sheet6], axis=0, join='outer')
    ht.loc[:, 'Broker ID'] = '华泰'
    ht = ht.rename(columns={'证券数量(份)': '可融数量'})
    ht = ht.rename(columns={'期限(天)': '期限'})
    ht['利率'] = ht['利率'].map(lambda x: format(x, '.2%'))
    ht['可融数量'] = ht['可融数量'].astype(int)
    return ht


def DongfangCaifu(fileName):
    df = pd.read_excel(fileName, usecols=['证券代码', '证券名称', '可用股数', '到期日'])
    df = df.rename(columns={'到期日': '合约到期日', '可用股数': '可融股数'})
    df.loc[:, 'Broker ID'] = '东方财富'
    return df


def DecodeStr(s):
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    return value


def print_info(msg):
    # 输出发件人，收件人，邮件主题信息
    for header in ['From', 'To', 'Subject']:
        value = msg.get(header, '')
        if value:
            if header == 'Subject':
                value = DecodeStr(value)  # 将主题名称解密
            else:
                hdr, addr = parseaddr(value)
                name = DecodeStr(hdr)
                value = u'%s <%s>' % (name, addr)
        print('%s: %s' % (header, value))
        return addr


# 获取邮件主体信息
def AttachmentFiles(msg):
    attachment_files = []
    for part in msg.walk():
        file_name = part.get_filename()  # 获取附件名称类型
        contentType = part.get_content_type()  # 获取数据类型
        mycode = part.get_content_charset()  # 获取编码格式
        if file_name:
            h = Header(file_name)
            dh = decode_header(h)  # 对附件名称进行解码
            filename = dh[0][0]
            if dh[0][1]:
                filename = DecodeStr(str(filename, dh[0][1]))  # 将附件名称可读化
            attachment_files.append(filename)
            data = part.get_payload(decode=True)  # 下载附件
            with open(filename, 'wb') as f:  # 在当前目录下创建文件，注意二进制文件需要用wb模式打开
                # with open('指定目录路径'+filename, 'wb') as f: 也可以指定下载目录
                f.write(data)  # 保存附件
            print(f'附件 {filename} 已下载完成')
            return filename


        elif contentType == 'text/plain':  # or contentType == 'text/html':
            # 输出正文 也可以写入文件
            data = part.get_payload(decode=True)
            content = data.decode(mycode)
            print('正文：', content)
    print('附件文件名列表', attachment_files)


# 获取邮件主体信息
def AttachmentFiles1(msg):
    attachment_files = []
    for part in msg.walk():
        file_name = part.get_filename()  # 获取附件名称类型
        contentType = part.get_content_type()  # 获取数据类型

        if file_name:
            h = Header(file_name)
            dh = decode_header(h)  # 对附件名称进行解码
            filename = dh[0][0]
            if dh[0][1]:
                filename = DecodeStr(str(filename, dh[0][1]))  # 将附件名称可读化
            attachment_files.append(filename)
            data = part.get_payload(decode=True)  # 下载附件
            with open(filename, 'wb') as f:  # 在当前目录下创建文件，注意二进制文件需要用wb模式打开
                # with open('指定目录路径'+filename, 'wb') as f: 也可以指定下载目录
                f.write(data)  # 保存附件
            print(f'附件 {filename} 已下载完成')
            return filename


        elif contentType == 'text/plain':  # or contentType == 'text/html':
            # 输出正文 也可以写入文件
            data = part.get_payload(decode=True)
            content = data.decode('gbk')
            print('正文：', content)
    print('附件文件名列表', attachment_files)


server = poplib.POP3('pop.exmail.qq.com', 110)
server.user(email)
server.pass_(password)

# 可选:打印POP3服务器的欢迎文字:
print(server.getwelcome().decode('utf-8'))

# stat()返回邮件数量和占用空间:
print('Messages: %s. Size: %s' % server.stat())
# list()返回所有邮件的编号:
resp, mails, octets = server.list()
# 可以查看返回的列表类似[b'1 82923', b'2 2184', ...]
print(mails)
# 由于pop3协议不支持对已读未读邮件的标记，因此，要判断一封pop邮箱中的邮件是否是新邮件必须与邮件客户端联合起来才能做到。
index = len(mails)
print('未读邮件的数量', index)

ht_addr = 'chenzixuan@htsc.com'
dw_addr = 'sz_zhus@dwzq.com.cn'
gf_addr = '3353989@gf.com.cn'
zx_addr = 'pbsblbj@citics.com'
yh_addr = 'duandeyi@chinastock.com.cn'
df_addr = 'caojia@orientsec.com.cn'
zj_addr = 'Yatong.Zheng@cicc.com.cn'
dfcf_addr = 'wangqifan@18.cn'

security_list = []

for i in range(index, 1, -1):
    resp, lines, octets = server.retr(i)
    msg_content = b'\r\n'.join(lines).decode('utf-8')
    msg = Parser().parsestr(msg_content)
    # 获取邮件内容
    addr = print_info(msg)
    if addr not in security_list:
        if addr == ht_addr:
            ht = HuaTai(AttachmentFiles(msg))
            security_list.append(addr)
        elif addr == dw_addr:
            dw = DongWu(AttachmentFiles(msg))
            security_list.append(addr)
        elif addr == yh_addr:
            yh = YinHe(AttachmentFiles(msg))
            security_list.append(addr)
        elif addr == gf_addr:
            gf = GuangFa(AttachmentFiles(msg))
            security_list.append(addr)
        elif addr == zj_addr:
            zj = ZhongJing(AttachmentFiles1(msg))
            security_list.append(addr)
        elif addr == zx_addr:
            zx = ZhongXin(AttachmentFiles(msg))
            security_list.append(addr)
        elif addr == df_addr:
            df = DongFang(AttachmentFiles(msg))
            security_list.append(addr)
        elif addr == dfcf_addr:
            dfcf = DongfangCaifu(AttachmentFiles(msg))
            security_list.append(addr)
    elif len(security_list) == 8:
        break
server.quit()

# 合并
combineData = pd.concat([zj, zx, yh, gf, dw, df, ht, dfcf], axis=0, join='outer')
combineData['证券代码'] = combineData['证券代码'].astype(str)
combineData['证券代码'] = combineData['证券代码'].str.zfill(6)
combineData = combineData.dropna(subset=['证券代码'])
combineData = combineData.sort_values(by='证券代码', ascending=True)
combineData = combineData.fillna(0)
# combineData['合约到期日'] = combineData['合约到期日'].astype(int)
combineData['合约到期日'] = combineData['合约到期日'].astype(str)
combineData['证券代码'] = combineData['证券代码'].str.extract('([0-9]+)')
# TODO 将表格上传到数据库，包括需求，最后再做一个脚本去对比发送邮件
# 读取CTA需求并匹配
# TODO 将需求设置为自动读取后处理上传或者直接本地匹配
xuqiu1 = pd.read_excel('盛冠达融券需求.xlsx', usecols=['证券代码', '证券名称'])
cta = pd.merge(xuqiu1, combineData, how='left', on='证券名称')
cta.to_csv('CTA.csv')
# 读取T0需求并匹配
etf = pd.read_excel('盛冠达投资团队融券需求(11.17)(1).xls', usecols=[
    'ETF标的', 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4'])
eft = etf.set_axis(etf.iloc[1], axis=1, inplace=True)
etf = etf.shift(axis=0, periods=-2)
etf['证券代码'] = etf['证券代码'].str.extract('([0-9]+)')
etf_result = pd.merge(etf, combineData, how='left', on='证券代码')
etf_result.insert(0, 'ETF', 'ETF', allow_duplicates=True)

fifty_p = pd.read_excel('盛冠达投资团队融券需求(11.17)(1).xls', usecols=[
    '50成分股', 'Unnamed: 8', 'Unnamed: 9'])
fifty_p = fifty_p.set_axis(fifty_p.iloc[2], axis=1, inplace=False)
fifty_p = fifty_p.shift(axis=0, periods=-3)
fifty_p['证券代码'] = fifty_p['证券代码'].astype(str)
fifty_p_result = pd.merge(fifty_p, combineData, how='left', on='证券代码')
fifty_p_result.insert(0, '50成分股', '50成分股', allow_duplicates=True)

three_h_p = pd.read_excel('盛冠达投资团队融券需求(11.17)(1).xls', usecols=[
    '300成分股', 'Unnamed: 11', 'Unnamed: 12'])
three_h_p = three_h_p.set_axis(three_h_p.iloc[2], axis=1, inplace=False)
three_h_p = three_h_p.shift(axis=0, periods=-3)
three_h_p['证券代码'] = three_h_p['证券代码'].astype(str)
three_h_p['证券代码'] = three_h_p['证券代码'].str.zfill(6)
three_h_p_result = pd.merge(three_h_p, combineData, how='left', on='证券代码')
three_h_p_result.insert(0, '300成分股', '300成分股', allow_duplicates=True)

five_h_p = pd.read_excel('盛冠达投资团队融券需求(11.17)(1).xls', usecols=[
    '500成分股', 'Unnamed: 14', 'Unnamed: 15'])
five_h_p = five_h_p.set_axis(five_h_p.iloc[2], axis=1, inplace=False)
five_h_p = five_h_p.shift(axis=0, periods=-3)
five_h_p['证券代码'] = five_h_p['证券代码'].astype(str)
five_h_p['证券代码'] = five_h_p['证券代码'].str.zfill(6)
five_h_p_result = pd.merge(five_h_p, combineData, how='left', on='证券代码')
five_h_p_result.insert(0, '500成分股', '500成分股', allow_duplicates=True)

T_0 = pd.concat([etf_result, fifty_p_result, three_h_p_result, five_h_p_result],
                axis=1, join='outer')
T_0.to_csv('T0.csv')
