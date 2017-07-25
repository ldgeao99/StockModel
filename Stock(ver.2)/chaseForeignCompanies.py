import urllib
import time
import win32com.client # to deal with excel
import slackweb # pip install slackweb, 참고 - http://qiita.com/satoshi03/items/14495bf431b1932cb90b

from urllib.request import urlopen
from bs4 import BeautifulSoup
from pandas import DataFrame
from multiprocessing import Process

class ChaseForeignCompanies:
    def __init__(self, TOTAL_ITEM, EXCEL_PATH, DESTINATION_URL, chaseTarget, MINIMUM_TOTAL_FOREIGN_VOLUME, MINIMUM_VARIANCE):
        print("__init__()호출됨")
        self.TOTAL_ITEM = TOTAL_ITEM
        self.EXCEL_PATH = EXCEL_PATH
        self.DESTINATION_URL = DESTINATION_URL
        self.MINIMUM_TOTAL_FOREIGN_VOLUME = MINIMUM_TOTAL_FOREIGN_VOLUME
        self.MINIMUM_VARIANCE = MINIMUM_VARIANCE
        self.nameAndCode_df = DataFrame()  # 엑셀에서 읽어온 종목이름과 코드 저장할 변수
        self.repeatCount_traceBuy_SlackAlarm = 0
        self.slack = slackweb.Slack(url=self.DESTINATION_URL)
        self.chaseTarget = chaseTarget
        self.catchedStock = {} #ex {"카카오" : {"메릴린치":20000, "CS증권":30000}, ....}

    def load_StockName_StockCode_FromExcel(self):
        print("load_StockName_StockCode_FromExcel()호출됨")
        self.nameAndCode_df = DataFrame(columns=("ItemName", "Code"))       #엑셀의 정보를 옮겨담을 데이터프레임
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(self.EXCEL_PATH)
        ws = wb.ActiveSheet
        for i in range(2, self.TOTAL_ITEM):
            rows = [str(ws.Cells(i, 1).Value), str(ws.Cells(i, 2).Value)]
            self.nameAndCode_df.loc[len(self.nameAndCode_df)] = rows
        excel.Application.Quit()

    def sendToSlack(self, message):
        self.slack.notify(text=message)

    def tradingTrends_CurrnetPrice_fluctuationRate_Crawler(self, stockCode): # 매수, 매도 상위 5위의 증권사 및 현재가격 및 변동량을 읽어온다.
        # 이 변수에 장중 거래원정보 및 현재가격을 가져와서 반환해준다.
        tradingTrends_df = DataFrame(columns=('sellFirmName', 'sellFirmVolume', 'buyFirmName', 'buyFirmVolume'))
        currentPrice = ''
        fluctuationRate = ''

        try:
            url = 'http://finance.naver.com/item/frgn.nhn?code=' + stockCode
            html = urlopen(url)
            source = BeautifulSoup(html.read(), "html.parser")

            #거래원정보 크롤링
            dealFirmInfo_tableFind = source.find("table", summary="거래원정보에 관한표이며 일자별 누적 정보를 제공합니다.")
            dealFirmInfo_tableFind_trFindAll_lst = dealFirmInfo_tableFind.find_all("tr")

            if(len(dealFirmInfo_tableFind_trFindAll_lst[11].find_all("span")) == 1):
                sellTotalForeignVolume = '0'
                buyTotalForeignVolume = '0'
            else:
                sellTotalForeignVolume = dealFirmInfo_tableFind_trFindAll_lst[11].find_all("span")[1].text.replace(',', '')
                buyTotalForeignVolume = dealFirmInfo_tableFind_trFindAll_lst[11].find_all("span")[3].text.replace(',', '')

            for i in range(4, 9): #4~8
                sellFirmName = dealFirmInfo_tableFind_trFindAll_lst[i].find_all("td")[0].text
                sellFirmVolume = dealFirmInfo_tableFind_trFindAll_lst[i].find_all("td")[1].text.replace(',', '')
                buyFirmName = dealFirmInfo_tableFind_trFindAll_lst[i].find_all("td")[2].text
                buyFirmVolume = dealFirmInfo_tableFind_trFindAll_lst[i].find_all("td")[3].text.replace(',', '')

                row = [sellFirmName, sellFirmVolume, buyFirmName, buyFirmVolume]
                tradingTrends_df.loc[len(tradingTrends_df)] = row


            row = [None, sellTotalForeignVolume, None, buyTotalForeignVolume]
            tradingTrends_df.loc[len(tradingTrends_df)] = row

            #현재가격 크롤링
            currentPrice_divFind = source.find("div", class_="new_totalinfo")
            currentPrice_divFind_dlFind = currentPrice_divFind.find("dl", class_="blind")
            currentPrice_divFind_dlFindAll_ddFindAll_lst = currentPrice_divFind_dlFind.find_all("dd")

            infoLst = currentPrice_divFind_dlFindAll_ddFindAll_lst[3].text.split(' ')  # ex) 현재가 39,100 전일대비 상승 1,200 플러스 3.17 퍼센트

            currentPrice = infoLst[1].replace(',', '')

            if (infoLst[3] == '상승' or infoLst[3] == '보합'):
                state = '+'
            elif (infoLst[3] == '하락'):
                state = '-'

            fluctuationRate = state + infoLst[6]

        except Exception as error:
            print(stockCode, "부분에서 에러발생!!")
            print(error)


        return tradingTrends_df, currentPrice, fluctuationRate

    def traceBuy_SlackAlarm(self, startIndex, endIndex): #멀티프로세싱을 해주고자 한다면 for문에 들어가는 인자의 범위를 달리하면 될 거 같다.
        while (1):
            try:
                if (self.repeatCount_traceBuy_SlackAlarm == 0):
                    print("<첫탐색>")
                    self.sendToSlack("<첫탐색>")

                for i in range(startIndex, endIndex): #for i in range(self.TOTAL_ITEM - 2):
                    stockName = self.nameAndCode_df.ix[i, 0]
                    stockCode = self.nameAndCode_df.ix[i, 1]

                    tradingTrends_df, currentPrice, fluctuationRate = self.tradingTrends_CurrnetPrice_fluctuationRate_Crawler(stockCode)

                    for j in range(5):
                        if (tradingTrends_df.ix[j, 2] in self.chaseTarget):  # 추적하고자하는 증권사가 있는 경우만 필터
                            if (int(tradingTrends_df.ix[5, 3]) >= self.MINIMUM_TOTAL_FOREIGN_VOLUME):  # 외인 매수 총 합이 MINIMUM_TOTAL_FOREIGN_VOLUME이상인 종목 필터
                                if (int(tradingTrends_df.ix[5, 1]) == 0):  # 외인 매도가 없는 종목 필터
                                    buyFirmName = tradingTrends_df.ix[j, 2]
                                    buyVolume = tradingTrends_df.ix[j, 3]  # 특정 증권사가 매수한  총 물량
                                    now = time.localtime()
                                    nowTime = "%02d:%02d:%02d" % (now.tm_hour, now.tm_min, now.tm_sec)
                                    if (stockName in self.catchedStock and buyFirmName in self.catchedStock[
                                        stockName]):  # 최종으로 찾고자하는 증권사가 있는 곳 필터
                                        previousBuyCount = str(self.catchedStock[stockName][buyFirmName][0])
                                        previousVolume = str(self.catchedStock[stockName][buyFirmName][1])
                                        currentVolume = buyVolume
                                        variationVolume = int(currentVolume) - int(previousVolume)  # 매수 변동량
                                        if (variationVolume > self.MINIMUM_VARIANCE):
                                            self.catchedStock[stockName][buyFirmName] = [int(previousBuyCount) + 1, int(buyVolume)]
                                            message = (nowTime + ' ' + stockCode + ' ' + stockName + ' ' + buyFirmName + ' ' + str(int(previousBuyCount) + 1) + '번매수' + ' ' + fluctuationRate + '%' + ' ' + currentPrice + '원' + ' '+ str(round((int(currentPrice)*int(buyVolume))/10000000, 1)) + '천만원' +' ' + previousVolume + '->' + currentVolume + '주' + '(+' + str(variationVolume) + ')')
                                            if (len(self.catchedStock[stockName]) >= 2):
                                                message = message + '!!!2개 이상의 외국증권계가 매수중인 종목!!!'
                                            print(message)
                                            self.sendToSlack(message)

                                    else:
                                        if(stockName in self.catchedStock):
                                            self.catchedStock[stockName][buyFirmName] = [1, int(buyVolume)]
                                        else:
                                            self.catchedStock[stockName] = {buyFirmName: [0, int(buyVolume)]}
                                        message = (nowTime + ' ' + stockCode + ' ' + stockName + ' ' + buyFirmName + ' ' + '0번매수' + ' ' + fluctuationRate + '%' + ' ' + currentPrice + '원' + ' ' + str(round((int(currentPrice)*int(buyVolume))/10000000, 1)) + '천만원')
                                        if(self.repeatCount_traceBuy_SlackAlarm > 0):
                                            message = message + "!!!처음으로 포착된 종목!!!"
                                        print(message)
                                        self.sendToSlack(message)
            except Exception as error:
                print("traceBuy_SlackAlarm() 에서 에러가 발생하였음.")
                print(stockName + "이 문제일까?")
                print(error)
            self.sendToSlack("------------------------------")
            print("------------------------------")
            self.repeatCount_traceBuy_SlackAlarm = 1

if __name__ == '__main__':

    START_TIME = "09:02:00"

    #코스피, 코스닥 공통
    chaseTarget = ['메릴린치', 'CS증권', '모간서울'] #추적하고 싶은 증권사 이름을 추가하면 됨.
    MINIMUM_TOTAL_FOREIGN_VOLUME = 1000  # 최소 이만큼의 외인 매수가 있을때 추적할 종목으로 포함.
    MINIMUM_VARIANCE = 50  # 최소 이만큼의 변화가 있을때 화면에 출력하게 되는 것.

    #코스피 프로세스 관련
    TOTAL_ITEM = 280
    EXCEL_PATH = "C:\\Users\\DG\\PycharmProjects\\Stock(ver.2)\\zipKospi.xls"
    DESTINATION_URL = "https://hooks.slack.com/services/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"
    ob1 = ChaseForeignCompanies(TOTAL_ITEM, EXCEL_PATH, DESTINATION_URL, chaseTarget, MINIMUM_TOTAL_FOREIGN_VOLUME, MINIMUM_VARIANCE)
    ob1.load_StockName_StockCode_FromExcel()

    #코스닥 프로세스 관련
    TOTAL_ITEM = 592
    EXCEL_PATH = "C:\\Users\\DG\\PycharmProjects\\Stock(ver.2)\\zipKosdaq.xls"
    DESTINATION_URL = "https://hooks.slack.com/services/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"
    ob2 = ChaseForeignCompanies(TOTAL_ITEM, EXCEL_PATH, DESTINATION_URL, chaseTarget, MINIMUM_TOTAL_FOREIGN_VOLUME, MINIMUM_VARIANCE)
    ob2.load_StockName_StockCode_FromExcel()

    while(1):
        now = time.localtime()
        nowTime = "%02d:%02d:%02d" % (now.tm_hour, now.tm_min, now.tm_sec)
        print(nowTime)
        if(nowTime == START_TIME): #09:02:00
            print(nowTime, "시작합니다.")
            break
        time.sleep(0.5)

    pro_kp1 = Process(target=ob1.traceBuy_SlackAlarm, args=(0, 140))
    pro_kp2 = Process(target=ob1.traceBuy_SlackAlarm, args=(140, 278))

    pro_kd1 = Process(target=ob2.traceBuy_SlackAlarm, args=(0, 280))
    pro_kd2 = Process(target=ob2.traceBuy_SlackAlarm, args=(280, 590))

    pro_kp1.start()
    pro_kp2.start()
    pro_kd1.start()
    pro_kd2.start()