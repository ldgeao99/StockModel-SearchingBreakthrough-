import urllib
import time
import win32com.client # to deal with excel
import slackweb # pip install slackweb, 참고 - http://qiita.com/satoshi03/items/14495bf431b1932cb90b
import os

from urllib.request import urlopen
from bs4 import BeautifulSoup
from pandas import DataFrame
from multiprocessing import Process

class SearchBreakthrough:
    def __init__(self, TOTAL_ITEM, EXCEL_PATH, DESTINATION_URL, TARGET_PERCENT, NDAYS, MARKET):
        print("__init__()호출됨")
        self.TOTAL_ITEM = TOTAL_ITEM
        self.EXCEL_PATH = EXCEL_PATH
        self.nameAndCode_df = DataFrame()  # 엑셀에서 읽어온 종목이름과 코드 저장할 변수
        self.slack = slackweb.Slack(url=DESTINATION_URL)
        self.TARGET_PERCENT = TARGET_PERCENT
        self.NDAYS = NDAYS
        self.MARKET = MARKET

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

    #해당종목의 수집Days일 동안의 최고가를 반환함
    def getHighestPriceForNDays(self, stockCode):
        highestPrice = 0
        highPriceListforNDays = []
        isFinish = False


        for page in range(10):
            url = "http://finance.naver.com/item/sise_day.nhn?code=" + stockCode + "&page=" + str(page+1)
            html = urlopen(url)
            source = BeautifulSoup(html.read(), "html.parser")
            priceInfo_FindAll = source.find_all('tr')

            for j in range(2,7):
                if page == 0 and j ==2: #오늘것은 수집 하지 않기 위한 처리
                    continue
                else:
                    highPrice = priceInfo_FindAll[j].find_all('td')[4].text.replace(',', '')
                    highPriceListforNDays.append(highPrice)
                    if len(highPriceListforNDays) == self.NDAYS:
                        isFinish = True
                        break

            if(isFinish):
                break

            for j in range(10, 15):
                highPrice = priceInfo_FindAll[j].find_all('td')[4].text.replace(',', '')
                highPriceListforNDays.append(highPrice)
                if len(highPriceListforNDays) == self.NDAYS:
                    isFinish = True
                    break
            if (isFinish):
                break

        highPriceListforNDays = map(int, highPriceListforNDays) # 문자열리스트 였던 것을 정수형 리스트로 변환
        highestPrice = str(max(highPriceListforNDays))
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
        storageItems = {}
        isFirtstWhileLoop = True
        
        while(1):

            for i in range(self.TOTAL_ITEM - 2):  # for i in range(self.TOTAL_ITEM - 2):
                try:
                    stockName = self.nameAndCode_df.ix[i, 0]
                    stockCode = self.nameAndCode_df.ix[i, 1]
    
                    currentPrice = self.getCurrentPrice(stockCode)
                    highestPriceForNDays = self.getHighestPriceForNDays(stockCode)
    
                    percent = str(round((float(currentPrice) / float(highestPriceForNDays) * 100), 1))
    
                    if(float(percent) >= self.TARGET_PERCENT):
                        #첫 While루프이고 TARGET_PERCENT를 만족하는 경우
                        if (isFirtstWhileLoop):
                            storageItems[stockName] = percent  # 딕셔너리에 종목 추가
                            now = time.localtime()
                            nowTime = "%02d:%02d:%02d" % (now.tm_hour, now.tm_min, now.tm_sec)
                            message = nowTime + ' ' + self.MARKET + ' ' + stockName + ' ' + percent + '% ' + '현재: ' + currentPrice + ' ' + '전고: ' + highestPriceForNDays
                            print(message)
                            self.sendToSlack(message)
                        #두번째 이상의 While루프이고 TARGET_PERCENT를 만족하는 경우
                        else:
                            isExistItem = False
                            isExistItem = stockName in storageItems

                            if(isExistItem):
                                beforePercent = storageItems[stockName]
                                currentPercent = percent
                                if(beforePercent < currentPercent):
                                    now = time.localtime()
                                    nowTime = "%02d:%02d:%02d" % (now.tm_hour, now.tm_min, now.tm_sec)
                                    message = nowTime + ' ' + self.MARKET + ' ' + stockName + ' ' + percent + '% ' + '현재: ' + currentPrice + ' ' + '전고: ' + highestPriceForNDays + " !percent가 이전보다 증가하였음!"
                                    print(message)
                                    self.sendToSlack(message)
                                storageItems[stockName] = currentPercent
                            else:
                                storageItems[stockName] = percent #딕셔너리에 종목 추가
                                now = time.localtime()
                                nowTime = "%02d:%02d:%02d" % (now.tm_hour, now.tm_min, now.tm_sec)
                                message = nowTime + ' ' + self.MARKET + ' ' + stockName + ' ' + percent + '% ' + '현재: ' + currentPrice + ' ' + '전고: ' + highestPriceForNDays + " !신규포착종목!"
                                print(message)
                                self.sendToSlack(message)
                    else:
                        isExistItem = False
                        isExistItem = stockName in storageItems

                        if(isExistItem):
                            del storageItems[stockName]
                            message = stockName + ' ' + self.MARKET + " 종목이 TARGET_PERCENT에서 이탈하여 딕셔너리에서 제거됩니다."
                            print(message)
                            self.sendToSlack(message)

                        #맨처음 루프때 명칭, %를 딕셔너리 형태로 저장해서 %가 상승하거나 새로 등록된 종목은 알람을 보내준다.
                except Exception as error:
                    print(stockCode, "부분에서 에러발생!!")
                    print(error)


            if(isFirtstWhileLoop):
                print(self.MARKET + " 첫번째 While루프 종료")
                isFirtstWhileLoop = False


if __name__ == '__main__':
    TARGET_PERCENT = 101
    PROJECT_PLACE = os.getcwd() + '\\'
    NDAYS = 20          #전고점을 어느 날짜까지 탐색할 것인가

    #코스피
    TOTAL_ITEM = 358
    EXCEL_PATH = PROJECT_PLACE + "zipKospi.xls"
    DESTINATION_URL = "https://hooks.slack.com/services/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"
    MARKET = 'KOSPI'
    sb1 = SearchBreakthrough(TOTAL_ITEM, EXCEL_PATH, DESTINATION_URL, TARGET_PERCENT, NDAYS, MARKET)
    sb1.load_StockName_StockCode_FromExcel()


    #코스닥
    TOTAL_ITEM = 461
    EXCEL_PATH = PROJECT_PLACE + "zipKosdaq.xls"
    DESTINATION_URL = "https://hooks.slack.com/services/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"
    MARKET = 'KOSDAQ'
    sb2 = SearchBreakthrough(TOTAL_ITEM, EXCEL_PATH, DESTINATION_URL, TARGET_PERCENT, NDAYS, MARKET)
    sb2.load_StockName_StockCode_FromExcel()

    pro_kp1 = Process(target=sb1.searchBreakthroughLoop)
    pro_kp2 = Process(target=sb2.searchBreakthroughLoop)

    pro_kp1.start()
    pro_kp2.start()

#아이디어 : 실시간으로 돌파종목을 감시하는데 감시선상에 올라온 종목에 대해서 최근 6개월 데이터의 차트를 그려주는 사이트 및 서버개발
#내가 가지고 있는 검색기를 가지고 매수, 매도를 자동으로 장중에 해보고 그에대한 성과를 보여주는 것.
#우상향인 차트인지 구분하는 코드