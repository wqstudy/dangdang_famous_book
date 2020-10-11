# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
from bs4 import BeautifulSoup
from lxml import html
from lxml import etree
import xml
import requests
import xlwings
import re

book_name_colunm = 0
book_score_colunm = 0
book_evaluate_colunm = 0


def requestUrl(start):
    global book_name_colunm
    global book_score_colunm
    global book_evaluate_colunm

    url = "https://book.douban.com/top250"
    header = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36"
    }
    params = {
        "start": start
    }
    response = requests.get(url=url, params=params, headers=header).text
    terr = etree.HTML(response)
    terr_lis = terr.xpath('//div[@class="indent"]/table')

    for i in terr_lis:
        book_name = i.xpath('./tr/td[2]/div[1]/a/text()')[0]
        book_score = i.xpath('./tr/td[2]/div[2]/span[2]/text()')[0]
        book_evaluate = i.xpath('./tr/td[2]/div[2]/span[3]/text()')[0]
        book_name = re.sub('\n | \( | \) | \xa0*', '', book_name)
        book_name=book_name.strip()
        book_score = re.sub('\n | \( | \) | \xa0*', '', book_score)
        book_score = book_score.strip()
        book_evaluate = re.sub('\n | \( | \) | \xa0*', '', book_evaluate)
        book_evaluate = book_evaluate.strip()
        # print(text + ' 评分:' + number + "\n")

        book_name_colunm = book_name_colunm + 1
        book_name_xls = "A" + str(book_name_colunm)
        sht.range(book_name_xls).value = book_name

        book_score_colunm = book_score_colunm + 1
        book_score_xls = "B" + str(book_score_colunm)
        sht.range(book_score_xls).value = book_score

        book_evaluate_colunm = book_evaluate_colunm + 1
        book_evaluate_xls = "C" + str(book_evaluate_colunm)
        sht.range(book_evaluate_xls).value = book_evaluate
        sht.range(book_evaluate_xls).columns.autofit()
        print("正在写入----" + book_name)


if __name__ == "__main__":

    wb = xlwings.Book("C:/Users/wqstu/Desktop/test/new6/book_top250.xlsx")
    sht = wb.sheets["sheet1"]
    start = 0
    while start <= 225:
        requestUrl(start)
        start = start + 25
