# 이 코드는 장이 완전히 종료된 오후 6시에 장이 마감된 다음 실행해야 오늘의 거래량도 포함된다.
# 시가총액은 아마 전날 값이 들어갈텐데 크게 상관 없을거 같다.

'''
1. 시가총액 만족
2. 가격 만족
3. 평균거래량 만족
4. 기관 외인 거래 활발 만족

위 4가지를 모두 만족하는 조건을 선별해냄
'''

import urllib
import time
import win32com.client

from urllib.request import urlopen
from bs4 import BeautifulSoup
from pandas import DataFrame

class ReduceStockItem:
    def __init__(self, TOTAL_ITEM, SOURCE_EXCEL_PATH, TARGET_EXCEL_PATH, MINMARKET_CAPITALIZATION,
                 MAXMARKET_CAPITALIZATION, MINPRICE, MAXPRICE, DAYS, MIN_NDAYS_MEAN_VOLUME):
        print("__init__()호출됨")
        self.nameAndCode_df = DataFrame()  # load_StockName_StockCode_FromExcel()를 통해 종목이름과 코드를 옮겨담을 데이터프레임
        self.TOTAL_ITEM = TOTAL_ITEM
        self.SOURCE_EXCEL_PATH = SOURCE_EXCEL_PATH
        self.TARGET_EXCEL_PATH = TARGET_EXCEL_PATH
        self.MINMARKET_CAPITALIZATION = MINMARKET_CAPITALIZATION
        self.MAXMARKET_CAPITALIZATION = MAXMARKET_CAPITALIZATION
        self.MINPRICE = MINPRICE
        self.MAXPRICE = MAXPRICE
        self.DAYS = DAYS
        self.MIN_NDAYS_MEAN_VOLUME = MIN_NDAYS_MEAN_VOLUME


    # 소스엑셀에서 데이터프레임으로 종목이름과 코드를 옮겨온다.
    def load_StockName_StockCode_FromExcel(self):
        print("load_StockName_StockCode_FromExcel()호출됨")
        self.nameAndCode_df = DataFrame(columns=("ItemName", "Code"))
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(self.SOURCE_EXCEL_PATH)
        ws = wb.ActiveSheet
        for i in range(2, self.TOTAL_ITEM+2):
            rows = [str(ws.Cells(i, 1).Value), str(ws.Cells(i, 2).Value)]
            self.nameAndCode_df.loc[len(self.nameAndCode_df)] = rows
        excel.Application.Quit()

    # 시가총액과 요구하는 시가총액의 범위에 맞는지 확인한 상태를 반환한다.
    def checkMarketCapitalization(self, stockCode):
        #print("checkMarketCapitalization()호출됨")
        url_1 = 'http://companyinfo.stock.naver.com/v1/company/c1010001.aspx?cmp_cd=' + stockCode
        html_1 = urlopen(url_1)
        source = BeautifulSoup(html_1.read(), "html.parser")
        source_divFind = source.find("table", summary="기업의 기본적인 시세정보(주가/전일대비/수익률,52주최고/최저,액면가,거래량/거래대금,시가총액,유통주식비율,외국인지분율,52주베타,수익률(1M/3M/6M/1Y))를 제공합니다.")
        source_divFind_tdFindAll = source_divFind.find_all("td", class_="num")
        marketCapitalization = (source_divFind_tdFindAll[4].text).replace(",", "").replace("억원", "").strip()

        isSatisfyCondition = False
        if(int(marketCapitalization) >= self.MINMARKET_CAPITALIZATION and int(marketCapitalization) <= self.MAXMARKET_CAPITALIZATION):
            isSatisfyCondition = True
        else:
            isSatisfyCondition = False
        return marketCapitalization, isSatisfyCondition

    # 가격과 가격의 범위에 맞는지 확인한 상태를 반환한다.
    def checkPrice(self, stockCode):
        #print("checkPrice()호출됨")
        url_1 = 'http://finance.naver.com/item/sise_day.nhn?code=' + stockCode
        html_1 = urlopen(url_1)
        source = BeautifulSoup(html_1.read(), "html.parser")
        source_divFind = source.find("table")
        source_divFind_trFindAll = source_divFind.find_all("tr")
        source_divFind_trFindAll_tdFindAll = source_divFind_trFindAll[2].find_all("td")
        price = source_divFind_trFindAll_tdFindAll[1].text.replace(",", "")

        isSatisfyCondition = False
        if(int(price) >= self.MINPRICE and int(price) <= self.MAXPRICE):
            isSatisfyCondition = True
        else:
            isSatisfyCondition = False
        return price, isSatisfyCondition

    # NDay평균거래량과 NDay평균 거래량 범위에 맞는지 확인한 상태를 반환한다.
    def checkNDaysVoumeMean(self, stockCode):
        #print("checkNDaysVoumeMean()호출됨")
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

        nDayMean = str(int(sum / self.DAYS))  # 소수점 아래 버림

        isSatisfyCondition = False
        if(int(nDayMean) >= MIN_NDAYS_MEAN_VOLUME):
            isSatisfyCondition = True
        else:
            isSatisfyCondition = False
        return nDayMean, isSatisfyCondition

    # 해당 주식을 기관과 외국인이 순매수한 거래량과 각각이 모두 0이 아닌지 확인한 상태를 반환한다.
    def checkBuyStateInstitutionAndForeign(self, stockCode):
        #print("checkBuyStateInstitutionAndForeign()호출됨")

        url_2 = 'http://finance.naver.com/item/frgn.nhn?code=' + stockCode
        html_2 = urlopen(url_2)
        source = BeautifulSoup(html_2.read(), "html.parser")
        dealInfo_tableFind = source.find("table", summary="외국인 기관 순매매 거래량에 관한표이며 날짜별로 정보를 제공합니다.")
        dealInfo_tableFind_trFindAll = dealInfo_tableFind.find_all("tr")

        institutionBuyVolume = dealInfo_tableFind_trFindAll[3].find_all("td")[5].text.replace(',', '')
        foreignBuyVolume = dealInfo_tableFind_trFindAll[3].find_all("td")[6].text.replace(',', '')

        isSatisfyCondition = False

        if(abs(int(institutionBuyVolume)) > 1000 and abs(int(foreignBuyVolume)) > 1000):
            isSatisfyCondition = True
        else:
            isSatisfyCondition = False
        return institutionBuyVolume, foreignBuyVolume, isSatisfyCondition

    # 위의 함수들을 종합적으로 사용하여 결과를 만들어 내는 함수.
    def totalCheckAndMakeResultExcelFile(self):
        print("totalCheckAndMakeResultExcelFile()호출됨")

        self.load_StockName_StockCode_FromExcel()

        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Add()
        ws = wb.Worksheets("Sheet1")

        ws.Cells(1, 1).Value = '회사명'
        ws.Cells(1, 2).Value = '종목코드'
        ws.Cells(1, 3).Value = '시가총액'
        ws.Cells(1, 4).Value = '가격'
        ws.Cells(1, 5).Value = str(self.DAYS) + '일 평균거래량'
        ws.Cells(1, 6).Value = '기관순매수'
        ws.Cells(1, 7).Value = '외인순매수'

        k = 2  # 정상 작동했을때만 엑셀에 입력되게 하기 위함.

        for i in range(self.TOTAL_ITEM):
            try:
                stockName = self.nameAndCode_df.ix[i, 0]
                stockCode = self.nameAndCode_df.ix[i, 1]

                print(str(i) + "번째 " + stockName + " " +stockCode + "대해 진행중...")

                marketCapitalization, condition1 = self.checkMarketCapitalization(stockCode)
                price, condition2 = self.checkPrice(stockCode)
                nDayVoluemMean, condition3 = self.checkNDaysVoumeMean(stockCode)
                institutionBuyVolume, foreignBuyVolume, condition4 = self.checkBuyStateInstitutionAndForeign(stockCode)

                if (condition1 and condition2 and condition3 and condition4):
                    ws.Cells(k, 1).Value = stockName
                    ws.Cells(k, 2).Value = '\'' + stockCode
                    ws.Cells(k, 3).Value = marketCapitalization
                    ws.Cells(k, 4).Value = price
                    ws.Cells(k, 5).Value = nDayVoluemMean
                    ws.Cells(k, 6).Value = institutionBuyVolume
                    ws.Cells(k, 7).Value = foreignBuyVolume
                    k += 1
                else:
                    print(stockName + " 조건미충족으로 인해 엑셀에서 제외")
            except Exception as error:
                print(stockName, stockCode, "에서 문제발생해서 패스")
                i += 1
                print(error)

        wb.SaveAs(self.TARGET_EXCEL_PATH)
        excel.Application.Quit()


if __name__ == '__main__':

    #Setting 변수들
    SELECT_MODE = 'KOSDAQ'                                              # KOSDAQ과 KOSPI 둘 중에 하나를 선택할 수 있다.
    PROJECT_PLACE = "C:\\Users\\NEPS\\PycharmProjects\\SBModel(ver0.8)\\" # 프로젝트의 경로
    MINMARKET_CAPITALIZATION = 300                                      # 필터할 최소시총
    MAXMARKET_CAPITALIZATION = 9999999
    MINPRICE = 1000                                                     # 필터할 최소가격
    MAXPRICE = 9999999                                                   # 필터할 최대가격
    DAYS = 10                                                           # 최대 20일의 평균까지 가능
    MIN_NDAYS_MEAN_VOLUME = 30000                                      # DAYS 동안의 평균중 얼마를 최저치로 잡을 것인가

    if(SELECT_MODE == 'KOSPI'):
        TOTAL_ITEM = 772   # SOURCE 코스피 엑셀 맨 마지막인덱스 + 1
        SOURCE_EXCEL_PATH = PROJECT_PLACE + "kospi.xls"
        TARGET_EXCEL_PATH = PROJECT_PLACE + "zipKospi.xls"
    elif(SELECT_MODE == 'KOSDAQ'):
        TOTAL_ITEM = 1232  # SOURCE 코스닥 엑셀 맨 마지막인덱스 + 1
        SOURCE_EXCEL_PATH = PROJECT_PLACE + "kosdaq.xls"
        TARGET_EXCEL_PATH = PROJECT_PLACE + "zipKosdaq.xls"
    else:
        print("모드를 올바르게 입력하세요!!")
        exit()

    reduceStockItem = ReduceStockItem(TOTAL_ITEM, SOURCE_EXCEL_PATH, TARGET_EXCEL_PATH, MINMARKET_CAPITALIZATION,
                                      MAXMARKET_CAPITALIZATION, MINPRICE, MAXPRICE, DAYS, MIN_NDAYS_MEAN_VOLUME)
    reduceStockItem.totalCheckAndMakeResultExcelFile()