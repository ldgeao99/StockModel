#
#주의!!!! 이 코드를 실행하기 전에 ZipKospiCode를 실행해서 코드를 선별하셔야 합니다.
# 이 코드는 순간적으로 메릴린치가 거래상위에 들어왔을때를 보여준다.

import urllib
import time
import win32com.client # to deal with excel
import slackweb # pip install slackweb, 참고 - http://qiita.com/satoshi03/items/14495bf431b1932cb90b

from sqlalchemy import create_engine #The flavor 'mysql' is deprecated in pandas version 0.19. You have to use the engine from sqlalchemy to create the connection with the database.
from urllib.request import urlopen
from bs4 import BeautifulSoup
from pandas import Series, DataFrame

slack = slackweb . Slack ( url = "https://hooks.slack.com/services/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" )


print("전일 10만주 이상의 거래량, 오늘 메릴린치가 1만주 이상 매수, 외인매도 상위 수량 집계가 0인 종목 탐색중...")


#<엑셀에 있는 종목이름과 종목코드를 Dataframe으로 읽어오는 작업>#
code_df = DataFrame(columns=("ItemName","Code"))

ITEM = 587 # zipKosdaq.xls 엑셀 맨끝 왼쪽에있는 인덱스 + 1

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
wb = excel.Workbooks.Open('C:\\Users\\DG\\PycharmProjects\\Stock(ver.1)\\zipKosdaq.xls')
ws = wb.ActiveSheet

for i in range(2,ITEM):
    rows = [str(ws.Cells(i,1).Value), str(ws.Cells(i,2).Value)] #종목이름, 종목코드
    code_df.loc[len(code_df)] = rows

excel.Quit()
###########################################################

dic_BuyVolume = {} # 종목이름 - 순매수물량
dic_BuyCount = {} # 종목이름 - 매수횟수
count = 0

while(1):
    if(count == 0):
        print("<스타트 첫 종목 정리>")
        slack.notify(text="<스타트 첫 종목 정리>")
    for i in range(ITEM-2):
        message = ""
        stockName = code_df.ix[i,0]
        stockCode = code_df.ix[i,1]
        try:
            url = 'http://finance.naver.com/item/frgn.nhn?code=' + stockCode
            html = urlopen(url)
            source = BeautifulSoup(html.read(), "html.parser")

            dealFirmSection = source.find("table", summary="거래원정보에 관한표이며 일자별 누적 정보를 제공합니다.")
            filter1 = dealFirmSection.find_all("tr")

            currentPriceSection = source.find("div", class_="new_totalinfo")
            filter2 = currentPriceSection.find_all("dl", class_="blind")[0]

            lst = filter2.find_all("dd")[3].text.split(' ') # ex) 현재가 39,100 전일대비 상승 1,200 플러스 3.17 퍼센트
            #현재가 1,130 전일대비 보합 0  0.00 퍼센트
            currentPrice = lst[1].replace(',', '')

            if(lst[3] == '상승'):
                state = '+'
            elif(lst[3] == '하락'):
                state = '-'
            else:
                state = '보합'

            fluctuationRate = lst[6] #등락률

            sellForeign_volume = filter1[11].find_all("td", class_="num bg01")[0].text.replace(',', '')
            sellForeign_volume = sellForeign_volume.replace('\n', '')

            if(sellForeign_volume != "" and int(sellForeign_volume) == 0):
                for j in range(4, 9):
                    # buyFirm_name   : 매수한 증권사
                    # buyFirm_volume : 매수한 증권사의 물량
                    buyFirm_name = filter1[j].find_all("td", class_="title bg02")[0].text
                    if (buyFirm_name == "메릴린치"):
                        #메릴린치가 매수 상위로 존재한다면.
                        buyFirm_volume = filter1[j].find_all("td", class_="num bg02")[0].text.replace(',', '')
                        if (int(buyFirm_volume) >= 1000):
                            # 메릴린치가 매수 상위이고 2000주 이상을 순매수 했다면.

                            now = time.localtime()
                            nowTime = "%02d:%02d:%02d" % (now.tm_hour, now.tm_min, now.tm_sec)

                            if(stockName in dic_BuyVolume and stockName in dic_BuyCount):
                                if(int(dic_BuyVolume[stockName]) < int(buyFirm_volume)): #메릴 좀전 매수물량, 지금 매수물량에 변동이 있다면.
                                    dic_BuyCount[stockName] = dic_BuyCount[stockName] + 1
                                    print(nowTime, stockCode, stockName, currentPrice + '원', state + fluctuationRate + '%', buyFirm_name, str(dic_BuyCount[stockName])+'번째포착', dic_BuyVolume[stockName],'->', buyFirm_volume, '(+', int(buyFirm_volume) - int(dic_BuyVolume[stockName]), ')')
                                    message = nowTime + ' ' + stockCode + ' ' + stockName + currentPrice + '원' + '\n' + state + ' ' + fluctuationRate + '%' + ' ' + buyFirm_name + ' ' + str(dic_BuyCount[stockName]) + '번째포착' + '\n' + dic_BuyVolume[stockName] + '->' + buyFirm_volume + '(+' + str(int(buyFirm_volume) - int(dic_BuyVolume[stockName])) + ')'
                                    slack.notify(text=message)
                                    dic_BuyVolume[stockName] = buyFirm_volume
                            else:
                                dic_BuyVolume[stockName] = buyFirm_volume
                                dic_BuyCount[stockName] = 0
                                if (count != 0):
                                    print("**처음으로 매수상위, 2000주 이상 순매수에 포착된 종목**")
                                print(nowTime, stockCode, stockName, currentPrice + '원', state + fluctuationRate + '%', buyFirm_name, str(dic_BuyCount[stockName])+'번째포착', buyFirm_volume)
                                message = nowTime + ' ' + stockCode + ' ' + stockName + currentPrice + '원' + '\n' + state + ' ' + fluctuationRate + '%' + ' ' + buyFirm_name + ' ' + str(dic_BuyCount[stockName]) + '번째포착' + ' ' + buyFirm_volume
                                if(count != 0):
                                    slack.notify(text="**처음으로 매수상위, 2000주 이상 순매수에 포착된 종목**")
                                slack.notify(text=message)
        except Exception as e:
            print(stockCode, "부분에서 에러발생!!")
            print(e)

    print("-----------------------------------------")
    slack.notify(text="-----------------------------------------")
    count = 1