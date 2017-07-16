#
#주의!!!! 이 코드를 실행하기 전에 ZipKospiCode를 실행해서 코드를 선별하셔야 합니다.
# 이 코드는 순간적으로 메릴린치가 거래상위에 들어왔을때를 보여준다.

import urllib
import time
import win32com.client # to deal with excel

from sqlalchemy import create_engine #The flavor 'mysql' is deprecated in pandas version 0.19. You have to use the engine from sqlalchemy to create the connection with the database.
from urllib.request import urlopen
from bs4 import BeautifulSoup
from pandas import Series, DataFrame

print("전일 10만주 이상의 거래량, 오늘 메릴린치가 1만주 이상 매수, 외인매도 상위 수량 집계가 0인 종목 탐색중...")

code_df = DataFrame(columns=("ItemName","Code"))

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
wb = excel.Workbooks.Open('C:\\Users\\DG\\PycharmProjects\\Stock\\zipKospi.xls')
ws = wb.ActiveSheet

ITEM = 342 #엑셀 맨끝 왼쪽에있는 인덱스 + 1

for i in range(2,ITEM):
    rows = [str(ws.Cells(i,1).Value), str(ws.Cells(i,2).Value)]
    code_df.loc[len(code_df)] = rows

excel.Quit()

######################################################################################################
count = 0
dic = {}
while(1):
    for i in range(ITEM-2):
        stockName = code_df.ix[i,0]
        stockCode = code_df.ix[i,1]
        try:
            url = 'http://finance.naver.com/item/frgn.nhn?code=' + stockCode
            html = urlopen(url)
            source = BeautifulSoup(html.read(), "html.parser")
            dataSection = source.find("table", summary="거래원정보에 관한표이며 일자별 누적 정보를 제공합니다.")
            filter1 = dataSection.find_all("tr")

            sellForeign_volume = filter1[11].find_all("td", class_="num bg01")[0].text.replace(',', '')
            sellForeign_volume = sellForeign_volume.replace('\n', '')

            if(sellForeign_volume != "" and int(sellForeign_volume) == 0):
                for j in range(4, 9):
                    # buyFirm_name   : 매수한 증권사
                    # buyFirm_volume : 매수한 증권사의 물량
                    buyFirm_name = filter1[j].find_all("td", class_="title bg02")[0].text
                    if (buyFirm_name == "메릴린치"):
                        buyFirm_volume = filter1[j].find_all("td", class_="num bg02")[0].text.replace(',', '')
                        if (int(buyFirm_volume) >= 2000):
                            now = time.localtime()
                            nowTime = "%02d:%02d:%02d" % (now.tm_hour, now.tm_min, now.tm_sec)
                            if(count == 0):
                                print(nowTime, stockCode, stockName, buyFirm_name, buyFirm_volume)
                                dic[stockName] = buyFirm_volume
                            else:
                                if(int(dic[stockName]) < int(buyFirm_volume)):
                                    print(nowTime, stockCode, stockName, buyFirm_name, dic[stockName], '->', buyFirm_volume, '(+', int(buyFirm_volume) - int(dic[stockName]), ')')
                                    dic[stockName] = buyFirm_volume
        except:
            print(stockCode, "부분에서 에러발생!!")



    print("-----------------------------------------")
    count += 1