from bs4 import BeautifulSoup
import requests
from datetime import datetime
from openpyexcel import load_workbook
from openpyexcel import Workbook

dateList = ["20220107"]
sectorList = ["105"]         #   105 : IT/과학,  101 : 경제
headers = {"user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.71 Safari/537.36"}
idx = 2
h8 = datetime.strptime('8:0', '%H:%M')
h12 = datetime.strptime('12:0', '%H:%M')
h3 = datetime.strptime('3:0', '%H:%M')


load_excel = load_workbook(filename='뉴스기사모음\예제.xlsx', data_only=True)
read_sheet = load_excel["data"]


for date in dateList:
    print(date + " 진행중")
    for sector in sectorList:
        print(date + " " + sector + " 진행중")
        pageNum = 1
        tmp = ""
        if sector == "105":
            tmp = "IT/과학"
        else:
            tmp = "경제"

        while pageNum <= 300:
            url = "https://news.naver.com/main/list.naver?mode=LSD&sid1=" + \
                    sector + "&mid=sec&listType=summary&date=" + date + "&page=" + str(pageNum)

            #print(tmp + ", " + str(pageNum) + "번째 생성된 url: ",url)

            original_html = requests.get(url, headers = headers)
            html = BeautifulSoup(original_html.text, "html.parser")

            headlines = html.find_all('dt', class_=False)
            headlines_date = html.find_all('span', 'date')
            #headlines1 = html.select("div#main_content > div.list_body.newsflash_body > ul.type06_headline > li > dl > dt > a.nclicks\(fls\.list\)")

            for headline, headline_date in zip(headlines, headlines_date):
                title = headline.get_text(strip = True)
                title = title.replace("\t","")
                if (title != "검색상위") and (title != "기준"):
                    read_sheet.cell(idx, 1, value=tmp)
                    read_sheet.cell(idx, 2, value=title)
                    timedata = (headline_date.get_text()).split(' ')
                    print(timedata)
                    det = datetime.strptime(timedata[2], '%H:%M')
                    if (timedata[1] == '오전' and det >= h12) or (timedata[1] == '오전' and det < h8) or \
                        ((timedata[1] == '오후' and det > h3) and (timedata[1] == '오후' and det < h12)):                #목표시간 설정
                        continue
                    col_time = 3
                    for time in timedata:
                        if col_time == 3:
                            YYYYMMDD = (time.split('.'))[:3]
                            for i in YYYYMMDD:
                                read_sheet.cell(idx, col_time, value=i)
                                col_time += 1
                        else:
                            read_sheet.cell(idx, col_time, value=time)
                            col_time += 1
                    idx += 1

            #print("status code : ", original_html.status_code)

            pageNum += 1

load_excel.save(filename="뉴스기사모음\예제.xlsx")


print("end")
