# This is a Spider coded by Python @OMGwq

from bs4 import BeautifulSoup
from lxml import html
from lxml import etree
import requests
import xlwings
import re
import time
import tkinter
import tkinter.messagebox
import random
import pywintypes

def get_agent():
    '''
    模拟header的user-agent字段，
    返回一个随机的user-agent字典类型的键值对
    '''
    agents = ['Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/5.0;',
              'Opera/9.80 (Macintosh; Intel Mac OS X 10.6.8; U; en) Presto/2.8.131 Version/11.11',
              'Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1; 360SE)',
              'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.116 Safari/537.36',
              'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.6; rv,2.0.1) Gecko/20100101 Firefox/4.0.1']
    fakeheader = {}
    fakeheader['User-agent'] = agents[random.randint(0, len(agents)-1)]
    return fakeheader

def request_construct(url):
    header = get_agent()

    try:
        #response = requests.get(url, proxies=proxies)
        response = requests.get(url=url, headers=header).text
    except requests.exceptions.ConnectionError:
        print('警告：遭遇反爬虫机制，休息30秒')
        time.sleep(30)
        try:
            response = requests.get(url=url, headers=header).text
        except requests.exceptions.ConnectionError:
            print('警告：休息30秒无法解决反爬机制，爬取结束')

    terr = etree.HTML(response)
    return terr

def request_best_books_url(best_books_url,sht):
    terr = request_construct(best_books_url)
    book_lists = terr.xpath('//ul[@class="bang_list clearfix bang_list_mode"]/li')
    for i in book_lists:
        book_url = i.xpath('./div[3]/a/@href')[0]
        request_each_book_Url(book_url,sht)

def request_each_book_Url(each_book_url,sht):
    global book_name_colunm
    global book_author_colunm
    global book_publishing_colunm
    global book_publishing_time_colunm
    global book_price_colunm
    global book_ISBN_colunm

    terr = request_construct(each_book_url)

    try:
        book_name = terr.xpath('//*[@id ="product_info"]/div[1]/h1/@title')[0]
    except IndexError:
        print('警告：遭遇反爬虫机制，休息30秒')
        time.sleep(30)
        try:
            terr = request_construct(each_book_url)
            book_name = terr.xpath('//*[@id ="product_info"]/div[1]/h1/@title')[0]
        except IndexError:
            book_name =''
            print('警告：休息30秒无法解决反爬机制')

    try:
        book_author = terr.xpath('//*[@id="author"]/a/text()')[0]
    except IndexError:
        book_author =''

    try:
        book_publishing = terr.xpath('//*[@id="product_info"]/div[2]/span[2]/a/text()')[0]
    except IndexError:
        book_publishing = ''

    try:
        book_publishing_time = terr.xpath('//*[@id="product_info"]/div[2]/span[3]/text()')[0]
    except IndexError:
        try:
            book_publishing_time = terr.xpath('//*[@id="product_info"]/div[2]/span[2]/text()')[0]
        except IndexError:
            book_publishing_time = ''

    book_price = terr.xpath('//*[@id="original-price"]/text()')[1]
    book_ISBN = terr.xpath('//*[@id="detail_describe"]/ul/li[5]/text()')[0]
    book_ISBN = re.search('\d+',str(book_ISBN)).group()

    book_name_colunm = book_name_colunm + 1
    book_name_xls = "A" + str(book_name_colunm)
    sht.range(book_name_xls).value = book_name
    sht.range(book_name_xls).columns.autofit()

    book_author_colunm = book_author_colunm + 1
    book_author_xls = "B" + str(book_author_colunm)
    sht.range(book_author_xls).value = book_author
    sht.range(book_author_xls).columns.autofit()

    book_publishing_colunm = book_publishing_colunm + 1
    book_publishing_xls = "C" + str(book_publishing_colunm)
    sht.range(book_publishing_xls).value = book_publishing
    sht.range(book_publishing_xls).columns.autofit()

    book_publishing_time_colunm = book_publishing_time_colunm + 1
    book_publishing_time_xls = "D" + str(book_publishing_time_colunm)
    sht.range(book_publishing_time_xls).value = book_publishing_time
    sht.range(book_publishing_time_xls).columns.autofit()

    book_price_colunm = book_price_colunm + 1
    book_price_xls = "E" + str(book_price_colunm)
    sht.range(book_price_xls).value = book_price
    sht.range(book_price_xls).columns.autofit()

    book_ISBN_colunm = book_ISBN_colunm + 1
    book_ISBN_xls = "F" + str(book_ISBN_colunm)
    sht.range(book_ISBN_xls).value = book_ISBN
    sht.range(book_ISBN_xls).columns.autofit()
    print("正在写入----name:" + book_name + "----author:" + book_author + "---- publish:" + book_publishing + "----time:" + book_publishing_time + "----price:" + book_price + "----ISBN:" + book_ISBN)
    time.sleep(random.random()*4+1)

def site_scan():
    global book_name_colunm
    global book_author_colunm
    global book_publishing_colunm
    global book_publishing_time_colunm
    global book_price_colunm
    global book_ISBN_colunm
    book_name_colunm = 0
    book_author_colunm = 0
    book_publishing_colunm = 0
    book_publishing_time_colunm = 0
    book_price_colunm = 0
    book_ISBN_colunm = 0

    input_path=str.strip(entry_path.get())
    url_site = str.strip(entry_site.get())

    url_site = url_site[:-1]
    start = 1
    app = xlwings.App(visible=True, add_book=False)
    wb = app.books.add()

    try:
        if re.match(r'^[A-Z]:(/.*?)*/$',input_path)==None:
            tkinter.messagebox.showinfo('警告', '格式错误，请检查结尾是否存在/，或/的方向')
    except SyntaxError:
        tkinter.messagebox.showinfo('警告', '格式错误，请检查/方向')

    url_path = input_path + "desired_book" + ".xlsx"

    try:
        wb.save(url_path)
    except pywintypes.com_error:
        tkinter.messagebox.showinfo('警告', '请关闭desired_book.xlsx')

    sht = wb.sheets["sheet1"]
    while start <= 25:
        request_best_books_url(url_site + str(start), sht)
        start = start + 1
    wb.save()
    wb.close()
    app.quit()
    tkinter.messagebox.showinfo('警告', '输入完成！！！')

if __name__ == "__main__":
    win=tkinter.Tk()
    win.title("Top500一键输出 v1.03")
    win.geometry("600x200+200+50")

    label_path=tkinter.Label(win, text="请输入书单的存储地址")
    label_path.pack()
    label_format=tkinter.Label(win, text="格式（注意“/”方向）,比如：C:/Users/XXX/Desktop/")
    label_format.pack()

    entry_path=tkinter.Entry(win)
    entry_path.pack()

    label_site = tkinter.Label(win, text="请输入书单的网页地址")
    label_site.pack()
    label_site_format = tkinter.Label(win, text="格式,比如：http://bang.dangdang.com/books/newhotsales/01.00.00.00.00.00-24hours-0-0-1-1")
    label_site_format.pack()

    entry_site=tkinter.Entry(win)
    entry_site.pack()

    button=tkinter.Button(win, text="输出", command=site_scan)
    button.pack()

    label_author=tkinter.Label(win, text="Author:OMGw~ Best wishes~~~")
    label_author.pack()

    win.mainloop()
    #pyinstaller -F -w dangdang_famous_book.py




