import urllib
import time
import win32com.client # to deal with excel
import slackweb # pip install slackweb, 참고 - http://qiita.com/satoshi03/items/14495bf431b1932cb90b

from urllib.request import urlopen
from bs4 import BeautifulSoup
from pandas import DataFrame
from multiprocessing import Process

class SearchBreakthrough:
    def __init__(self, TOTAL_ITEM, EXCEL_PATH, DESTINATION_URL, TARGET_PERCENT):
        print("__init__()호출됨")
        self.TOTAL_ITEM = TOTAL_ITEM
        self.EXCEL_PATH = EXCEL_PATH
        self.nameAndCode_df = DataFrame()  # 엑셀에서 읽어온 종목이름과 코드 저장할 변수
        self.slack = slackweb.Slack(url=DESTINATION_URL)
        self.TARGET_PERCENT = TARGET_PERCENT

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

    #해당종목의 collectionDays일 동안의 최고가를 반환함
    def getHighestPriceFor20Days(self, stockCode):
        highestPrice = 0
        highPriceListfor20Days = []

        for page in range(2):
            url = "http://finance.naver.com/item/sise_day.nhn?code=" + stockCode + "&page=" + str(page+1)
            html = urlopen(url)
            source = BeautifulSoup(html.read(), "html.parser")

            priceInfo_FindAll = source.find_all('tr')

            for j in range(2,7):
                highPrice = priceInfo_FindAll[j].find_all('td')[4].text.replace(',', '')
                highPriceListfor20Days.append(highPrice)

            for j in range(10, 15):
                highPrice = priceInfo_FindAll[j].find_all('td')[4].text.replace(',', '')
                highPriceListfor20Days.append(highPrice)

        highPriceListfor20Days = map(int, highPriceListfor20Days) # 문자열리스트 였던 것을 정수형 리스트로 변환
        highestPrice = str(max(highPriceListfor20Days))
        return highestPrice

    #해당종목의 현재가를 반환함
    def getCurrentPrice(self, stockCode):
        url = "http://finance.naver.com/item/main.nhn?code=" + stockCode
        html = urlopen(url)
        source = BeautifulSoup(html.read(), "html.parser")
        totalInfo_dlFind = source.find("dl")
        totalInfo_dlFind_FindAll = totalInfo_dlFind.find_all('dd')
        currentPrice = totalInfo_dlFind_FindAll[3].text.split(' ')[1].replace(',', '')
        return currentPrice

    def searchBreakthroughLoop(self):
        print("탐색중...")
        for i in range(self.TOTAL_ITEM - 2):  # for i in range(self.TOTAL_ITEM - 2):
            try:
                stockName = self.nameAndCode_df.ix[i, 0]
                stockCode = self.nameAndCode_df.ix[i, 1]

                currentPrice = self.getCurrentPrice(stockCode)
                highestPriceFor20Days = self.getHighestPriceFor20Days(stockCode)

                percent = str(int((float(currentPrice) / float(highestPriceFor20Days)) * 100))

                if(int(percent) >= self.TARGET_PERCENT):
                    now = time.localtime()
                    nowTime = "%02d:%02d:%02d" % (now.tm_hour, now.tm_min, now.tm_sec)
                    print(stockName, percent, '현재: ' + currentPrice, '전고: ' + highestPriceFor20Days)
                    message = nowTime + ' ' +stockName + ' ' + percent + '% ' + '현재: ' +currentPrice + ' ' + '전고: ' +highestPriceFor20Days
                    self.sendToSlack(message)

            except Exception as error:
                print(stockCode, "부분에서 에러발생!!")
                print(error)

if __name__ == '__main__':
    TARGET_PERCENT = 100
    PROJECT_PLACE = "C:\\Users\\NEPS\\PycharmProjects\\SBModel(ver0.5)\\"

    #코스피
    TOTAL_ITEM = 280
    EXCEL_PATH = PROJECT_PLACE + "zipKospi.xls"
    DESTINATION_URL = "https://hooks.slack.com/services/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"
    sb1 = SearchBreakthrough(TOTAL_ITEM, EXCEL_PATH, DESTINATION_URL, TARGET_PERCENT)
    sb1.load_StockName_StockCode_FromExcel()

    #코스닥
    TOTAL_ITEM = 280
    EXCEL_PATH = PROJECT_PLACE + "zipKosdaq.xls"
    DESTINATION_URL = "https://hooks.slack.com/services/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"
    sb2 = SearchBreakthrough(TOTAL_ITEM, EXCEL_PATH, DESTINATION_URL, TARGET_PERCENT)
    sb2.load_StockName_StockCode_FromExcel()

    pro_kp1 = Process(target=sb1.searchBreakthroughLoop)
    pro_kp2 = Process(target=sb2.searchBreakthroughLoop)

    pro_kp1.start()
    pro_kp2.start()