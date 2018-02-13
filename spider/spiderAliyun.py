# -*- coding:utf-8 -*-
__author__ = 'TanXing'
__time__ = '2018.02.13'

import os,time
import requests
from bs4 import BeautifulSoup

from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from cStringIO import StringIO
from io import open
from pdfminer.pdfpage import PDFPage

import smtplib
from email.mime.text import MIMEText
from email.header import Header

'''
工作原理：将脚本添加进计划任务，固定时间去爬取指定页面的内容，并在当前目录生成txt文件保存通知标题，供下一次
爬取作比对，如果有更新，将更新的通知已邮件形式发送，并将新标题添加进txt。如此反复。

注：使用前先确认邮件服务器配置正确！！
'''
url1 = "https://help.aliyun.com"    #供拼接的主域名
url2 = "https://help.aliyun.com/noticelist/9004748.html?spm=a2c4g.789213612.n2.1.5kjNIE"    #阿里云所有公告的爬取页面


def email(title,body):
    '''邮件服务器设置，按需自行设置'''
    mail_host = "xxxx.com"   #邮件服务器地址
    sender = "xxxx.com"    #发件人名称
    receivers = ['xxxx.com'] #收件人地址，自行设置

    mail_user = "xxxx.com" #邮件服务器用户名
    mail_pass = "passwd"  #邮件服务用户密码

    message = MIMEText(body, 'plain', 'utf-8')  #邮件正文设置
    message['From'] = Header("zabbix", 'utf-8') #发件人别名
    message['To'] = Header("xxxx", 'utf-8') #收件人别名

    subject = title
    message['Subject'] = Header(subject, 'utf-8')

    smtpObj = smtplib.SMTP()
    smtpObj.connect(mail_host, 25)
    smtpObj.login(mail_user, mail_pass)
    smtpObj.sendmail(sender, receivers, message.as_string())

def firstRunAliyun():
    '''初次运行，先爬取阿里云公告页面一次，生成列表，供后面比对'''
    html = requests.get(url2)
    soup = BeautifulSoup(html.text, 'lxml')
    old_list = soup.find_all('li', limit=7)  #limit为爬取列表的条数，此处表示只抓取符合条件的7条内容，可自行设置
    tmp_list = []
    for i in range(len(old_list)):
        name = old_list[i].select('a')[0].get_text()
        if name.rstrip():
            tmp = name.rstrip() + '\n'
            tmp_list.append(tmp)
    with open('aliyun.txt', "w",encoding='utf8') as f:
        f.writelines(tmp_list)

def _pdf2txt(url):
    '''将pdf转换成txt，迅游之家的行政文件是pdf格式，想提取其内容，需将其进行格式转换'''
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    f = requests.get(url).content
    fp = StringIO(f)
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos = set()
    for page in PDFPage.get_pages(fp,
                                  pagenos,
                                  maxpages=maxpages,
                                  password=password,
                                  caching=caching,
                                  check_extractable=True):
        interpreter.process_page(page)
    fp.close()
    device.close()
    tmp_str = retstr.getvalue()
    retstr.close()
    str = tmp_str.splitlines()
    return str

def readHistory(filename):
    '''读list文件'''
    with open(filename, "r",encoding='utf8') as f:
        com_list = f.readlines()
        return com_list

def aliyunPage():
    '''爬取阿里云的<所有公告>页面'''
    html = requests.get(url2)
    soup = BeautifulSoup(html.text, 'lxml')
    lilist = soup.find_all('li',limit=13)
    for li in lilist:
        title = li.select('a')[0].get_text()
        # messg_time = li.select('span')[0].get_text()
        tmp_title = title + '\n'
        com_list = readHistory('aliyun.txt')
        if tmp_title not in com_list:
            str1 = li.select('a')[0].attrs.get('href')
            suburl = url1 + str1
            subhtml = requests.get(suburl)
            subsoup = BeautifulSoup(subhtml.text, 'lxml')
            plist = subsoup.find_all("p")
            h3list = subsoup.find_all("h3")
            messg_str = ''
            for p in plist:
                line = p.get_text()
                messg_str = messg_str + line +'\n'
            if messg_str.strip() == '':
                for h3 in h3list:
                    line = h3.get_text()
                    messg_str = messg_str + line + '\n'
            email_title = '【阿里云公告】'+ title.encode('utf8')
            email(email_title,messg_str)
            # print email_title,messg_str
            time.sleep(1)
            with open("aliyun.txt", "a", encoding='utf8') as f:
                f.write(tmp_title)


if __name__ == "__main__":
    try:
        if not os.path.exists("aliyun.txt"):
            firstRunAliyun()
        else:
            aliyunPage()
    except Exception as e:
        pass
