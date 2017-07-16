#
# 이 코드는 장이 완전히 종료된 저녁 6시 이후에 실행시켜야 하도록 작성되었음.
# 장 종료된 당일 거래량이 100만 이상되는 종목만을 추리는 코드임.

import urllib
import time
import win32com.client # to deal with excel

from urllib.request import urlopen
from bs4 import BeautifulSoup
from pandas import Series, DataFrame

print("코스피 종목코드 압축 실행중... 약 1분30초~ 3분30초 가량이 소요됩니다.")

code_df = DataFrame(columns=("ItemName","Code"))

excel_1 = win32com.client.Dispatch("Excel.Application")
excel_1.Visible = False
wb = excel_1.Workbooks.Open('C:\\Users\\DG\\PycharmProjects\\Stock\\kospi.xls')
ws = wb.ActiveSheet

ITEM = 772

for i in range(2,ITEM):
    rows = [str(ws.Cells(i,1).Value), str(ws.Cells(i,2).Value)]
    code_df.loc[len(code_df)] = rows
excel_1.Quit()


excel_2 = win32com.client.Dispatch("Excel.Application")
excel_2.Visible = False
wb = excel_2.Workbooks.Add()
ws = wb.Worksheets("Sheet1")

k = 2
ws.Cells(1, 1).Value = '회사명'
ws.Cells(1, 2).Value = '종목코드'
ws.Cells(1, 3).Value = '거래량'

for i in range(ITEM-2):
    stockName = code_df.ix[i,0]
    stockCode = code_df.ix[i,1]

    url = 'http://finance.naver.com/item/sise_day.nhn?code=' + stockCode
    html = urlopen(url)
    source = BeautifulSoup(html.read(), "html.parser")
    srlists = source.find_all("tr")

    for j in range(1, len(srlists) - 1):
        # 데이터를 읽어오는 부분
        if (srlists[j].span != None):
            day = (srlists[j].find_all("td", align="center")[0].text).replace('.', '')
            volume = (srlists[j].find_all("td", class_="num")[5].text).replace(',', '')
            break;

    if(int(volume) >= 100000):
        ws.Cells(k, 1).Value = stockName
        ws.Cells(k, 2).Value = '\'' + stockCode
        ws.Cells(k, 3).Value = volume
        k += 1

wb.SaveAs('C:\\Users\\DG\\PycharmProjects\\Stock\\zipKospi.xls')
excel_2.Quit()