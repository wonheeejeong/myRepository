from bs4 import BeautifulSoup
import requests
from openpyexcel import load_workbook
from openpyexcel import Workbook
from datetime import datetime

today = datetime.today().strftime('%Y%m%d')

load_excel = load_workbook(filename='결과값_' + today + '.xlsm', data_only=True)
read_sheet = load_excel["뉴스정보_코스피"]

list_KOSPI = []
list_KOSPI.append(read_sheet.cell(2,1).value)
list_KOSPI.append(read_sheet.cell(12,1).value)
list_KOSPI.append(read_sheet.cell(22,1).value)
list_KOSPI.append(read_sheet.cell(32,1).value)
list_KOSPI.append(read_sheet.cell(42,1).value)
list_KOSPI.append(read_sheet.cell(52,1).value)
list_KOSPI.append(read_sheet.cell(62,1).value)
list_KOSPI.append(read_sheet.cell(72,1).value)
list_KOSPI.append(read_sheet.cell(82,1).value)
list_KOSPI.append(read_sheet.cell(92,1).value)
print(list_KOSPI)



idx = 2

for i in list_KOSPI:
    search = i
    url = "https://search.naver.com/search.naver?where=news&sm=tab_pge&query=" + search
    print("생성url: ",url)

    original_html = requests.get(url)
    html = BeautifulSoup(original_html.text, "html.parser")

    articles = html.select("div.news_area > a")
    print(articles)
    # 검색된 기사의 갯수
    print(len(articles),"개의 기사가 검색됌.")


    for j in articles:
        read_sheet.cell(idx,2, value = j.attrs['title'])
        read_sheet.cell(idx,3, value = j.attrs['href'])
        idx += 1






read_sheet = load_excel["뉴스정보_코스닥"]

list_KOSDAQ = []
list_KOSDAQ.append(read_sheet.cell(2,1).value)
list_KOSDAQ.append(read_sheet.cell(12,1).value)
list_KOSDAQ.append(read_sheet.cell(22,1).value)
list_KOSDAQ.append(read_sheet.cell(32,1).value)
list_KOSDAQ.append(read_sheet.cell(42,1).value)
list_KOSDAQ.append(read_sheet.cell(52,1).value)
list_KOSDAQ.append(read_sheet.cell(62,1).value)
list_KOSDAQ.append(read_sheet.cell(72,1).value)
list_KOSDAQ.append(read_sheet.cell(82,1).value)
list_KOSDAQ.append(read_sheet.cell(92,1).value)

print(list_KOSDAQ)

idx = 2

for i in list_KOSDAQ:
    search = i
    url = "https://search.naver.com/search.naver?where=news&sm=tab_pge&query=" + search
    print("생성url: ",url)

    original_html = requests.get(url)
    html = BeautifulSoup(original_html.text, "html.parser")

    articles = html.select("div.news_area > a")
    print(articles)
    # 검색된 기사의 갯수
    print(len(articles),"개의 기사가 검색됌.")



    for j in articles:
        read_sheet.cell(idx,2, value = j.attrs['title'])
        read_sheet.cell(idx,3, value = j.attrs['href'])
        idx += 1


load_excel.save(filename=today +'.xlsx')
