# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
from bs4 import BeautifulSoup
from lxml import html
from lxml import etree
import xml
import requests
import xlwings

video_name_colunm = 0
video_score_colunm = 0
video_evaluate_colunm = 0


def requestUrl(start):
    global video_name_colunm
    global video_score_colunm
    global video_evaluate_colunm

    url = "https://movie.douban.com/top250"
    header = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36"
    }
    params = {
        "start": start,
        "filter": ""
    }
    response = requests.get(url=url, params=params, headers=header).text
    terr = etree.HTML(response)
    terr_lis = terr.xpath('//ol[@class="grid_view"]/li')

    for i in terr_lis:
        video_name = i.xpath('./div/div/div/a/span/text()')[0]
        video_score = i.xpath('./div/div/div/div/span[2]/text()')[0]
        video_evaluate = i.xpath('./div/div/div/div/span[4]/text()')[0]
        # print(text + ' 评分:' + number + "\n")

        video_name_colunm = video_name_colunm + 1
        video_name_xls = "A" + str(video_name_colunm)
        sht.range(video_name_xls).value = video_name

        video_score_colunm = video_score_colunm + 1
        video_score_xls = "B" + str(video_score_colunm)
        sht.range(video_score_xls).value = video_score

        video_evaluate_colunm = video_evaluate_colunm + 1
        video_evaluate_xls = "C" + str(video_evaluate_colunm)
        sht.range(video_evaluate_xls).value = video_evaluate
        sht.range(video_evaluate_xls).columns.autofit()
        print("正在写入----" + video_name)


if __name__ == "__main__":

    wb = xlwings.Book("C:/Users/wqstu/Desktop/test/new6/video_top250.xlsx")
    sht = wb.sheets["sheet1"]
    start = 0
    while start <= 225:
        requestUrl(start)
        start = start + 25
