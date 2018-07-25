from email import encoders
from email.header import Header
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import parseaddr, formataddr
from email.mime.base import MIMEBase
import smtplib
import datetime
import csv
import os
import random

def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr((Header(name, 'utf-8').encode(), addr))
    
def _writeToEmail(convertedContent):
    res = ''
    for info in convertedContent:
        temp = '<tr><td>'+info[0]+'</td><td>' + info[1] + '</td><td>' + info[2] + '</td><td>' + info[3] + '</td></tr>'
        res += temp
    res += '<td>每日总结</td><td colspan="3">' + conclusionToday +'</td></tr></tbody></table>'
    return res
def _writeToCsv(convertedContent):
    if os.path.exists('data/'+strToday +'.csv'): 
        os.remove('data/'+strToday +'.csv')
    with open('data/'+strToday +'.csv', 'w', newline='') as csvfile:
        spamwriter = csv.writer(csvfile,dialect='excel')
        title = str(today.year)+'年'+ strToday + '尹凯4Y日结果报表'
        spamwriter.writerow(["",title, "",""])
        spamwriter.writerow(("日期","工作地点","结果定义","完成情况汇报"))
        for info in convertedContent:
            spamwriter.writerow(info)
        spamwriter.writerow(("每日总结", conclusionToday ,"",""))
def _contentHelper(dirpath = 'input/source.txt'):
    with open(dirpath, 'r') as f:
        lines = f.readlines()
        lines = list(map(lambda x:x.strip(), lines))
    try:
        start = lines.index(strToday)
    except ValueError:
        print('没有找到当天的数据')
    else: 
        end = lines.index('end', start, start + 20)
        content = lines[start + 1: end]
    res = []
    for info in content:
        info = info.split(',')
        if(info[0] == ''):info[0] = '公司'
        if(info[2] == ''):info.insert(0,strTomorrow)
        else: info.insert(0,strToday)
        res.append(info)
    return res
# content = [['公司', '编写报告', '完成今日编写计划'], ['公司', '内部oa系统使用培训', '完成oa系统培训'], ['公司', '编写报告','']]
def _conclusionHelper(dirpath = 'input/conclusion.txt'):
    with open(dirpath, 'r') as f:
        lines = f.readlines()
        conclusion = list(map(lambda x:x.strip(), lines))
        conclusionToday = random.sample(conclusion, 1)[0]
    return conclusionToday
conclusionToday = _conclusionHelper()

#from_addr = 'yink@relialab.com'
#password = 'FuckU0'
#to_addr = 'beihangink@sina.com'

from_addr = 'beihangink@sina.com'
password = ''
#to_addr = 'yink_bise@163.com'
to_addr = 'beijing4y@relialab.com'
# smtp_server = raw_input('SMTP server: ')


today = datetime.datetime.now().date()
tomorrow = today + datetime.timedelta(days=1)
strToday = str(today.month)+'月'+str(today.day)+'日'
strTomorrow = str(tomorrow.month)+'月'+str(tomorrow.day)+'日'

convertedContent = _contentHelper()
title = '<table border="6"><caption>'+str(today.year)+'年'+ strToday + '尹凯4Y日结果报表</caption><tbody><tr><td>日期</td><td>工作地点</td><td>结果定义</td><td>完成情况汇报</td></tr>'
infomation = _writeToEmail(convertedContent)
text = title + infomation

msg = MIMEMultipart()
msg.attach(MIMEText(text, 'html', 'utf-8'))
msg['From'] = _format_addr(u'尹凯 <%s>' % from_addr)
msg['To'] = _format_addr(u'4Y管理员 <%s>' % to_addr)
msg['Subject'] = Header(strToday + '尹凯4Y日结果', 'utf-8').encode()

_writeToCsv(convertedContent)

with open('data/'+strToday +'.csv', 'rb') as f:
    # 设置附件的MIME和文件名，这里是png类型:
    mime = MIMEBase('data', 'csv', filename=strToday + '.csv')
    # 加上必要的头信息:
    mime.add_header('Content-Disposition', 'attachment', filename=strToday + '-北京分部-尹凯''.csv')
    mime.add_header('Content-ID', '<0>')
    mime.add_header('X-Attachment-Id', '0')
    # 把附件的内容读进来:
    mime.set_payload(f.read())
    # 用Base64编码:
    encoders.encode_base64(mime)
    # 添加到MIMEMultipart:
    msg.attach(mime)

server = smtplib.SMTP("smtp.sina.com")
#server = smtplib.SMTP_SSL('smtp.exmail.qq.com', port = 465)
server.set_debuglevel(1)
server.login(from_addr, password)
server.sendmail(from_addr, [to_addr], msg.as_string())
server.quit()