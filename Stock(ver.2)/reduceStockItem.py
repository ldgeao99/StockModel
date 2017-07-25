# 이 코드는 장이 완전히 종료된 오후 6시에 장이 마감된 다음 실행해야 오늘의 거래량도 포함된다.
# 시가총액은 아마 전날 값이 들어갈텐데 크게 상관 없을거 같다.

import urllib
import time
import win32com.client

from urllib.request import urlopen
from bs4 import BeautifulSoup
from pandas import DataFrame

class ReduceStockItem:
    def __init__(self, TOTAL_ITEM, SOURCE_EXCEL_PATH, TARGET_EXCEL_PATH, MINMARKET_CAPITALIZATION, MAXMARKET_CAPITALIZATION, MINPRICE, MAXPRICE, DAYS, MIN_NDAYS_MEAN_VOLUME):
        print("__init__()호출됨")
        self.nameAndCode_df = DataFrame()  #엑셀의 정보를 옮겨담을 데이터프레임
        self.TOTAL_ITEM = TOTAL_ITEM
        self.SOURCE_EXCEL_PATH = SOURCE_EXCEL_PATH
        self.TARGET_EXCEL_PATH = TARGET_EXCEL_PATH
        self.MINMARKET_CAPITALIZATION = MINMARKET_CAPITALIZATION
        self.MAXMARKET_CAPITALIZATION = MAXMARKET_CAPITALIZATION
        self.MINPRICE = MINPRICE
        self.MAXPRICE = MAXPRICE
        self.DAYS = DAYS
        self.MIN_NDAYS_MEAN_VOLUME = MIN_NDAYS_MEAN_VOLUME

    #소스엑셀에서 데이터프레임으로 종목이름과 코드를 옮겨온다.
    def load_StockName_StockCode_FromExcel(self):
        print("load_StockName_StockCode_FromExcel()호출됨")
        self.nameAndCode_df = DataFrame(columns=("ItemName", "Code"))
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(self.SOURCE_EXCEL_PATH)
        ws = wb.ActiveSheet
        for i in range(2, self.TOTAL_ITEM):
            rows = [str(ws.Cells(i, 1).Value), str(ws.Cells(i, 2).Value)]
            self.nameAndCode_df.loc[len(self.nameAndCode_df)] = rows
        excel.Application.Quit()

    #코스피 또는 코스닥의 종목코드를 이용해 시가총액을 가져와 엑셀로 저장한다. => 회사명 / 종목코드 / 시가총액 / 가격 / N일 평균거래량
    def saveMarketCapitalization_Price_NdayMean(self):
        print("saveMarketCapitalization()호출됨")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Add()
        ws = wb.Worksheets("Sheet1")

        ws.Cells(1, 1).Value = '회사명'
        ws.Cells(1, 2).Value = '종목코드'
        ws.Cells(1, 3).Value = '시가총액'
        ws.Cells(1, 4).Value = '가격'
        ws.Cells(1, 5).Value = str(self.DAYS) + '일 평균거래량'

        k = 2    # 정상 작동했을때만 엑셀에 입력되게 하기 위함.
        for i in range(self.TOTAL_ITEM - 2):
            try:
                stockName = self.nameAndCode_df.ix[i, 0]
                stockCode = self.nameAndCode_df.ix[i, 1]

                #시가총액, 가격관련부분
                url_1 = 'http://companyinfo.stock.naver.com/v1/company/c1010001.aspx?cmp_cd=' + stockCode
                html_1 = urlopen(url_1)
                source = BeautifulSoup(html_1.read(), "html.parser")
                source_divFind = source.find("table",summary="기업의 기본적인 시세정보(주가/전일대비/수익률,52주최고/최저,액면가,거래량/거래대금,시가총액,유통주식비율,외국인지분율,52주베타,수익률(1M/3M/6M/1Y))를 제공합니다.")
                source_divFind_tdFindAll = source_divFind.find_all("td", class_="num")

                marketCapitalization = (source_divFind_tdFindAll[4].text).replace(",", "").replace("억원", "").strip()
                price = (source_divFind_tdFindAll[0].text).split('/')[0].replace(",", "").replace("원", "").strip()

                if(int(marketCapitalization) >= self.MINMARKET_CAPITALIZATION and int(marketCapitalization) <= self.MAXMARKET_CAPITALIZATION and int(price) >= self.MINPRICE and int(price) <= self.MAXPRICE):
                    #거래량 관련부분
                    url_2 = 'http://finance.naver.com/item/frgn.nhn?code=' + stockCode
                    html_2 = urlopen(url_2)
                    source = BeautifulSoup(html_2.read(), "html.parser")
                    dealInfo_tableFind = source.find("table", summary="외국인 기관 순매매 거래량에 관한표이며 날짜별로 정보를 제공합니다.")
                    dealInfo_tableFind_trFindAll = dealInfo_tableFind.find_all("tr")

                    volumeLst = []  # 최근 20개의 거래량을 가져와 저장할 리스트.

                    for j in range(3, 8):
                        volumeLst.append(int(dealInfo_tableFind_trFindAll[j].find_all("td")[4].text.replace(',', '')))
                    for j in range(11, 16):
                        volumeLst.append(int(dealInfo_tableFind_trFindAll[j].find_all("td")[4].text.replace(',', '')))
                    for j in range(19, 24):
                        volumeLst.append(int(dealInfo_tableFind_trFindAll[j].find_all("td")[4].text.replace(',', '')))
                    for j in range(27, 32):
                        volumeLst.append(int(dealInfo_tableFind_trFindAll[j].find_all("td")[4].text.replace(',', '')))

                    sum = 0

                    for t in range(self.DAYS):
                        sum += volumeLst[t]

                    nDayMean = str(int(sum / self.DAYS)) # 소수점 아래 버림

                    if(int(nDayMean) >= self.MIN_NDAYS_MEAN_VOLUME):
                        ws.Cells(k, 1).Value = stockName
                        ws.Cells(k, 2).Value = '\'' + stockCode
                        ws.Cells(k, 3).Value = marketCapitalization
                        ws.Cells(k, 4).Value = price
                        ws.Cells(k, 5).Value = nDayMean
                        print(i)
                        k += 1

            except Exception as error:
                print(stockName, stockCode, "에서 문제발생해서 패스")
                print(error)

        wb.SaveAs(self.TARGET_EXCEL_PATH)
        excel.Application.Quit()

if __name__ == '__main__':

    SELECT_MODE = 'KOSDAQ'

    if(SELECT_MODE == 'KOSPI'):
        TOTAL_ITEM = 772   # SOURCE 코스피 엑셀 맨 마지막인덱스 + 1
        SOURCE_EXCEL_PATH = "C:\\Users\\DG\\PycharmProjects\\Stock(ver.2)\\kospi.xls"
        TARGET_EXCEL_PATH = "C:\\Users\\DG\\PycharmProjects\\Stock(ver.2)\\zipKospi.xls"
        MAXMARKET_CAPITALIZATION = 50000  # 필터할 최대시총
    elif(SELECT_MODE == 'KOSDAQ'):
        TOTAL_ITEM = 1232  # SOURCE 코스닥 엑셀 맨 마지막인덱스 + 1
        SOURCE_EXCEL_PATH = "C:\\Users\\DG\\PycharmProjects\\Stock(ver.2)\\kosdaq.xls"
        TARGET_EXCEL_PATH = "C:\\Users\\DG\\PycharmProjects\\Stock(ver.2)\\zipKosdaq.xls"
        MAXMARKET_CAPITALIZATION = 150000  # 필터할 최대시총
    else:
        print("모드를 올바르게 입력하세요!!")
        exit()

    MINMARKET_CAPITALIZATION = 300  # 필터할 최소시총

    MINPRICE = 1000     #필터할 최소가격
    MAXPRICE = 400000   #필터할 최대가격

    DAYS = 10   #최대 20일의 평균까지 가능
    MIN_NDAYS_MEAN_VOLUME = 100000

    reduceStockItem = ReduceStockItem(TOTAL_ITEM, SOURCE_EXCEL_PATH, TARGET_EXCEL_PATH, MINMARKET_CAPITALIZATION, MAXMARKET_CAPITALIZATION, MINPRICE, MAXPRICE, DAYS, MIN_NDAYS_MEAN_VOLUME)
    reduceStockItem.load_StockName_StockCode_FromExcel()
    reduceStockItem.saveMarketCapitalization_Price_NdayMean()