import win32com.client
import GetChart
import pandas as pd
import time

def ReqeustData(obj):
    # 데이터 요청
    obj.BlockRequest()

    # 통신 결과 확인
    rqStatus = obj.GetDibStatus()
    rqRet = obj.GetDibMsg1()
    # print("통신상태", rqStatus, rqRet)
    if rqStatus != 0:
        return False

    # 일자별 정보 데이터 처리
    # count = obj.GetHeaderValue(1)  # 데이터 개수
    date = []
    open = []
    high = []
    low = []
    close = []
    vol = []
    for i in range(20):
        date.append(obj.GetDataValue(0, i))  # 일자
        open.append(obj.GetDataValue(1, i))  # 시가
        high.append(obj.GetDataValue(2, i))  # 고가
        low.append(obj.GetDataValue(3, i))  # 저가
        close.append(obj.GetDataValue(4, i))  # 종가
        vol.append(obj.GetDataValue(6, i))  # 거래량

    df = pd.DataFrame({'일자':date,
                      '시가':open,
                      '고가':high,
                      '저가':low,
                      '종가':close,
                      '거래량':vol})

    return df

# 연결 여부 체크
objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
bConnect = objCpCybos.IsConnect
if (bConnect == 0):
    print("PLUS가 정상적으로 연결되지 않음. ")
    exit()

# 종목코드 리스트 구하기
objCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
codeList = objCpCodeMgr.GetStockListByMarket(2)  # 코스닥
ChartData = GetChart.CpStockChart()

stockCode = []
for i, code in enumerate(codeList):
    print(i)
    time.sleep(0.3)
    secondCode = objCpCodeMgr.GetStockSectionKind(code)
    name = objCpCodeMgr.CodeToName(code)
    stdPrice = objCpCodeMgr.GetStockStdPrice(code)
    objStockWeek = win32com.client.Dispatch("DsCbo1.StockWeek")
    objStockWeek.SetInputValue(0, code)
    objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
    objStockMst.SetInputValue(0, code)

    # 최초 데이터 요청
    stock_data = ReqeustData(objStockWeek)
    if stock_data['거래량'][1] > 1000000: # 날짜 확인 0:오늘 1:어제 2:그저께
        MA20 = stock_data['종가'].mean()
        if stock_data['고가'][2] < MA20 and stock_data['고가'][3] < MA20 and stock_data['고가'][4] < MA20: # 날짜 확인 0:오늘 1:어제 2:그저께
            if stock_data['저가'][1] < MA20 and stock_data['고가'][1] > MA20:
                stockCode.append(code)

    if len(stockCode) >= 10:
        break
    else:
        pass

print(stockCode)

Numstock = len(stockCode)
# 연결 여부 체크
objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
bConnect = objCpCybos.IsConnect
if (bConnect == 0):
    print("PLUS가 정상적으로 연결되지 않음. ")
    exit()

# 주문 초기화
objTrade = win32com.client.Dispatch("CpTrade.CpTdUtil")
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')
objCpTradeCpTdUtil = win32com.client.Dispatch('CpTrade.CpTdUtil')
objCpTradeCpTd6033 = win32com.client.Dispatch('CpTrade.CpTd6033')
objCpTradeCpTd0311 = win32com.client.Dispatch('CpTrade.CpTd0311')
objCpTradeCpTdNew5331A = win32com.client.Dispatch('CpTrade.CpTdNew5331A')
# logging.logger.debug('CpTrade API 연결 완료')

initCheck = objTrade.TradeInit(0)

if (initCheck != 0):
    print("주문 초기화 실패")
    exit()

for code in stockCode:
    # 주식 매수 주문
    acc = objTrade.AccountNumber[0] #계좌번호
    accFlag = objTrade.GoodsList(acc, 1)  # 주식상품 구분

    objCpTradeCpTdNew5331A.SetInputValue(0, acc)
    objCpTradeCpTdNew5331A.SetInputValue(1, accFlag[0])
    objCpTradeCpTdNew5331A.BlockRequest()

    print(acc, accFlag[0])
    objStockOrder = win32com.client.Dispatch("CpTrade.CpTd0311")
    objStockOrder.SetInputValue(0, "2")   # 2: 매수
    objStockOrder.SetInputValue(1, acc)   #  계좌번호
    objStockOrder.SetInputValue(2, accFlag[0])   # 상품구분 - 주식 상품 중 첫번째
    objStockOrder.SetInputValue(3, code)   # 종목코드 - 필요한 종목으로 변경 필요
    objStockOrder.SetInputValue(4, 10)   # 매수수량 - 요청 수량으로 변경 필요
    # objStockOrder.SetInputValue(5, 10000)   # 주문단가 - 필요한 가격으로 변경 필요
    objStockOrder.SetInputValue(7, "0")   # 주문 조건 구분 코드, 0: 기본 1: IOC 2:FOK
    # IOC (Immediate-Or-Cancel Order) : 주문즉시 체결 그리고 잔량 자동취소
    # FOK (Fill-Or-kill Order) : 주문즉시 전부체결 또는 전부자동취소
    objStockOrder.SetInputValue(8, "03")   # 주문호가 구분코드 - 03: 시장가

    # 매수 주문 요청
    nRet = objStockOrder.BlockRequest()
    if (nRet != 0) :
        print("주문요청 오류", nRet)
        # 0: 정상,  그 외 오류, 4: 주문요청제한 개수 초과
        exit()

    rqStatus = objStockOrder.GetDibStatus()
    errMsg = objStockOrder.GetDibMsg1()
    if rqStatus != 0:
        print("주문 실패: ", rqStatus, errMsg)
        exit()
