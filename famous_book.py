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

def request_construct(url):
    header = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36"
    }

    try:
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
    time.sleep(random.random()*3+1)

def site_scan():
    global book_name_colunm
    global book_author_colunm
    global book_publishing_colunm
    global book_publishing_time_colunm
    global book_price_colunm
    global book_ISBN_colunm
    urls={"new_books":"http://bang.dangdang.com/books/newhotsales/01.00.00.00.00.00-24hours-0-0-1-","best_books":"http://bang.dangdang.com/books/bestsellers/01.00.00.00.00.00-24hours-0-0-1-","best_books_army":"http://bang.dangdang.com/books/bestsellers/01.27.00.00.00.00-24hours-0-0-1-"}
    for key,value in urls.items():
        start = 1
        book_name_colunm = 0
        book_author_colunm = 0
        book_publishing_colunm = 0
        book_publishing_time_colunm = 0
        book_price_colunm = 0
        book_ISBN_colunm = 0

        app = xlwings.App(visible=True, add_book=False)
        wb = app.books.add()
        url=entry.get() + key + ".xlsx"
        wb.save(url)
        sht = wb.sheets["sheet1"]
        while start <= 25:
            request_best_books_url(value + str(start),sht)
            start = start + 1
        wb.save()
        wb.close()
        app.quit()
    tkinter.messagebox.showinfo('警告','输入完成！！！')

if __name__ == "__main__":
    win=tkinter.Tk()
    win.title("Top500一键输出 v1.02")
    win.geometry("400x200+200+50")

    label_path=tkinter.Label(win, text="请输入书单的地址")
    label_path.pack()
    label_format=tkinter.Label(win, text="格式（注意“/”方向）,比如：C:/Users/XXX/Desktop/")
    label_format.pack()

    entry=tkinter.Entry(win)
    entry.pack()

    button=tkinter.Button(win, text="输出", command=site_scan)
    button.pack()

    label_author=tkinter.Label(win, text="Author:wuqi~ Best wishes~~~")
    label_author.pack()

    win.mainloop()




