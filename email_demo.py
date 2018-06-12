# -*- coding:utf-8 -*-
import xlwt
import pymysql
import datetime
import time
import smtplib
from email import encoders
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email.utils import parseaddr,formataddr
from email.mime.base import MIMEBase

#create a Excel
excelTable = xlwt.Workbook()
sheet1 = excelTable.add_sheet('vip',cell_overwrite_ok=True)
sheet1.write(0,0,'vip_name')
sheet1.write(0,1,'vip_phone')
sheet1.write(0,2,'invite_name')
sheet1.write(0,3,'invite_phone')
sheet1.write(0,4,'time')
sheet1.write(0,5,'cck_commission')

#获取数据库连接
conn = pymysql.connect(
    host='10.211.55.22',
    port=3306,
    user='root',
    password='root',
    db='database',
    charset='utf8'
)

EXE_SQL="""
    select * from demo where create_time>={yesterday_s} and create_time<{today_s};
"""

#获取今天和昨天0点的时间戳
today = datetime.date.today()
yesterday = today - datetime.timedelta(days=1)
yesterday_stamp = time.mktime(time.strptime(str(yesterday),'%Y-%m-%d'))
today_stamp = time.mktime(time.strptime(str(today),'%Y-%m-%d'))

#执行sql获取数据
cur = conn.cursor()
num = cur.execute(EXE_SQL.format(yesterday_s=yesterday_stamp,today_s=today_stamp))
i=1
while(True):
    record = cur.fetchone()
    if(record is not None):
        for j in range(6):
            sheet1.write(i,j,str(record[j]))
        i=i+1
    else:
        break

#将数据库返回结果写进excel
excelTable.save('/home/wangfan/pyfile/'+str(yesterday)+'_new_vip.xls')
cur.close()
conn.close()

def _format_addr(s):
    name,addr = parseaddr(s)
    return formataddr((Header(name,'utf-8').encode(),addr))

from_addr = "demo@163.com"
password = "123456"
smtp_server = "smtp.163.com"
to_addr = ["xxxxx@qq.com","xxxxxx@163.com"]

msg=MIMEMultipart()
msg['From'] = _format_addr(u'wangfan<%s>' % from_addr)
msg['To'] = _format_addr(u'myfriend<%s>' % to_addr)
msg['Subject'] = Header(u'new vip list','utf-8').encode()

#这里是邮件正文
msg.attach(MIMEText("today new vip",'plain','utf-8'))

#添加一个附件就是加上一个MIMEBase,读取excel文件
with open("/home/wangfan/pyfile/"+str(yesterday)+'_new_vip.xls',"rb") as f:
    mime = MIMEBase('file','xls',filename=str(yesterday)+'.xls')
    mime.add_header('Content-Disposition','attachment',filename=str(yesterday)+"_new_vip.xls")
    mime.add_header('Content-ID','<0>')
    mime.add_header('X-Attachment-Id','0')
    mime.set_payload(f.read())
    encoders.encode_base64(mime)
    msg.attach(mime)
    
#将邮件发送出去
server = smtplib.SMTP(smtp_server,25)
server.set_debuglevel(1)
server.login(from_addr,password)
server.sendmail(from_addr,to_addr,msg.as_string())
server.quit()
