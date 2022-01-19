import pandas as pd
import email
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email.mime.application import MIMEApplication

host_server = 'smtp.exmail.qq.com'
sender_qq = str(input('plz enter email:'))
psw = str(input('plz enter password:'))
receiver = '1565013912@qq.com'
mail_title = 'testing'

mail_content = 'testing mail'
msg = MIMEMultipart()
msg['Subject'] = Header(mail_title, 'utf-8')
msg['From'] = sender_qq
msg['To'] = receiver

msg.attach(MIMEText(mail_content, 'html'))
attachment = MIMEApplication(open(
    r'/Users/ky_c/PycharmProjects/sgd/checking_list/11/basic.csv', 'rb').read())
attachment['Content-Type'] = 'application/octet-stream'
attachment['Content-Disposition'] = 'attachement; filename = "basic.csv"'
msg.attach(attachment)

try:
    smtp = smtplib.SMTP_SSL(host_server, 465)
    smtp.set_debuglevel(1)
    smtp.ehlo(host_server)
    smtp.login(sender_qq, psw)
    smtp.sendmail(sender_qq, receiver, msg.as_string())
    smtp.quit()
    print('succes')
except smtplib.SMTPException:
    print('failed')
