from bs4 import BeautifulSoup
from lxml import html
import xml
import requests
import xlwings

rank = 1
video_name_colunm = 0
video_score_colunm = 0
video_year_colunm = 0
video_country_colunm = 0
video_type_colunm = 0

def write_one_page(soup):
    global rank
    global video_name_colunm
    global video_score_colunm
    global video_year_colunm
    global video_country_colunm
    global video_type_colunm
    for k in soup.find('div',class_='article').find_all('div',class_='info'):
        name = k.find('div',class_='hd').find_all('span')#电影名字
        score = k.find('div',class_='star').find_all('span')#分数
        #inq = k.find('p',class_='quote').find('span')#一句话简介
        #抓取年份、国家
        actor_infos_html = k.find(class_='bd')
        #strip() 方法用于移除字符串头尾指定的字符（默认为空格）
        actor_infos = actor_infos_html.find('p').get_text().strip().split('\n')
        actor_infos1 = actor_infos[0].split('\xa0\xa0\xa0')
        director = actor_infos1[0][3:]
        role = actor_infos[1]
        year_area = actor_infos[1].lstrip().split('\xa0/\xa0')
        year = year_area[0]
        country = year_area[1]
        type = year_area[2]

        video_name_colunm = video_name_colunm + 1
        video_name_xls = "A" + str(video_name_colunm)
        sht.range(video_name_xls).value = name[0].string

        video_score_colunm = video_score_colunm + 1
        video_score_xls = "B" + str(video_score_colunm)
        sht.range(video_score_xls).value = score[1].string

        video_year_colunm = video_year_colunm + 1
        video_year_xls = "C" + str(video_year_colunm)
        sht.range(video_year_xls).value = year

        video_country_colunm = video_country_colunm + 1
        video_country_xls = "D" + str(video_country_colunm)
        sht.range(video_country_xls).value = country
        sht.range(video_country_xls).columns.autofit()

        video_type_colunm = video_type_colunm + 1
        video_type_xls = "E" + str(video_type_colunm)
        sht.range(video_type_xls).value = type

        #print(rank,name[0].string,score[1].string,inq.string,year,country,type)
        print(rank, name[0].string, score[1].string, year, country, type)
        rank = rank + 1
        #写txt
    #write_to_file(rank,name[0].string,score[1].string,year,country,type,inq.string)
    #write_to_file(rank, name[0].string, score[1].string, year, country, type)


# #def write_to_file(rank,name,score,year,country,type,quote):
# def write_to_file(rank, name, score, year, country, type):
#     with open('C:/Users/wqstu/Desktop/test/new6/Top_250_movie.txt', 'a', encoding='utf-8') as f:
#         #f.write(str(rank)+';'+str(name)+';'+str(score)+';'+str(year)+';'+str(country)+';'+str(type)+';'+str(quote)+'\n')
#         f.write(str(rank) + ';' + str(name) + ';' + str(score) + ';' + str(year) + ';' + str(country) + ';' + str(
#             type) + ';'  + '\n')
#         f.close()

if __name__ == '__main__':
    wb = xlwings.Book("C:/Users/wqstu/Desktop/test/new6/video_top250_method2.xlsx")
    sht = wb.sheets["sheet1"]
    for i in range(10):
        a = i*25
        url = "https://movie.douban.com/top250?start="+str(a)+"&filter="
        header = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36"
        }
        f = requests.get(url=url, headers=header)
        soup = BeautifulSoup(f.content, "lxml")
        write_one_page(soup)