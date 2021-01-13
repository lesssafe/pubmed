#!/usr/bin/python3
#-*- coding:utf-8 -*-
# author:Less is More
# from:https://github.com/lesssafe
import requests
from bs4 import BeautifulSoup
import re
import xlsxwriter
import tkinter
import urllib.parse
import threading
import random
workbook = xlsxwriter.Workbook('pubmed.xlsx',{'constant_memory':True,'strings_to_numbers':True})
worksheet = workbook.add_worksheet("pubmed爬取数据")
worksheet.write(0, 0, "网址链接")
worksheet.write(0, 1, "标题")  # 写入行，列，内容
worksheet.write(0, 2, "摘要")
headers={
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:84.0) Gecko/20100101 Firefox/84.0',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Language': 'zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2',
        'Accept-Encoding': 'gzip, deflate',
        'Connection': 'close',
        'Upgrade-Insecure-Requests': '1',
        'Cache-Control': 'max-age=0'
    }
def pages_amount1(url):
    pubmed_url=url
    geturl = requests.get(pubmed_url, headers=headers)
    soupurl = BeautifulSoup(geturl.text, "html.parser")
    amount = soupurl.find('div', attrs={'class': 'search-results-chunk results-chunk'})
    pages_amount = amount['data-pages-amount']
    return pages_amount
def pubmed_search(pageid,pageurl):
    number=int(pageid)*10-10
    pubmedurl=pageurl+'&page='+pageid
    getres=requests.get(pubmedurl,headers=headers)
    soup=BeautifulSoup(getres.text,"html.parser")
    lis=soup.find_all('a',attrs={'class': 'docsum-title'})
    for li in lis:
        number=number+1
        url='https://pubmed.ncbi.nlm.nih.gov'+li['href']
        geturl=requests.get(url,headers=headers)
        soupurl=BeautifulSoup(geturl.text,"html.parser")
        titlesoup=soupurl.find('h1',attrs={'class': 'heading-title'})
        title = re.findall(r'<h1 class="heading-title">\n  \n    \n    \n    \n    \n      \n  ([\S\s]*)\n\n\n    \n  \n</h1>',str(titlesoup))
        abstractsoup = soupurl.find('div', attrs={'class': 'abstract-content selected'})
        abstract=BeautifulSoup(str(abstractsoup)).get_text(separator=" ")
        write_xls(url,str(title),abstract,number)
def write_xls(url,title,abstract,row):
    worksheet.write(row, 0, url)
    worksheet.write(row, 1, title)
    worksheet.write(row, 2, abstract)
def mainrun():
    keyword = entry.get()
    years=entry1.get()
    years=2021-int(years)
    keyword = urllib.parse.quote(keyword)
    listbox.insert(tkinter.END, '已经开始爬取请耐心等待')
    url = 'https://pubmed.ncbi.nlm.nih.gov/?term=' + keyword + '&filter=years.'+str(years)+'-2021'
    pages_amount = pages_amount1(url)
    for pageid in range(1, int(pages_amount) + 1):
        pubmed_search(str(pageid), url)
        notice="已爬取" + str(pageid) + "页数据"
        listbox.insert(tkinter.END,notice)
    workbook.close()
    listbox.insert(tkinter.END, '哈哈哈哈，爬取完成了，请享用！')
def thread_it(func):
    t = threading.Thread(target=func)
    # 守护 !!!
    t.setDaemon(True)
    # 启动
    t.start()
if __name__ == '__main__':
    top = tkinter.Tk()
    top.title("pubmed文章下载")
    top.geometry('560x460')
    label = tkinter.Label(top, text='Keyword', font=('微软雅黑', 13))
    label.grid(row=0, column=0)
    entry = tkinter.Entry(top, font=('微软雅黑', 13), width=33)
    entry.grid(row=0, column=1)
    label1 = tkinter.Label(top, text='最近几年数据', font=('微软雅黑', 13))
    label1.grid(row=1, column=0)
    entry1 = tkinter.Entry(top, font=('微软雅黑', 13), width=33)
    entry1.grid(row=1, column=1)
    button = tkinter.Button(top, text='GO', font=('微软雅黑', 10),command=lambda: thread_it(mainrun))
    button.grid(row=1, column=2)
    listbox = tkinter.Listbox(top, font=('微软雅黑', 10), width=69, height=18)
    listbox.grid(row=2, columnspan=3)
    randomtxt=['(^_^)∠※','~(@^_^@)~','(☆＿☆)']
    randomnum=random.randint(0,2)
    print(randomnum)
    label2=tkinter.Label(top,text=randomtxt[randomnum],font=('微软雅黑',10))
    label2.grid(row=3,column=0)
    top.mainloop()