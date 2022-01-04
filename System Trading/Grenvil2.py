import win32com.client
import pandas as pd

# 연결 여부 체크
objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
bConnect = objCpCybos.IsConnect
if (bConnect == 0):
    print("PLUS가 정상적으로 연결되지 않음. ")
    exit()

# 주문 초기화
objTrade = win32com.client.Dispatch("CpTrade.CpTdUtil")
initCheck = objTrade.TradeInit(0)
if (initCheck != 0):
    print("주문 초기화 실패")
    exit()

return_list = []
objStockRead = win32com.client.Dispatch("CpTrade.CpTd6033")
cnt = objStockRead.GetHeaderValue(7)
for i in range(cnt):
    evalPerc = objStockRead.GetDataValue(11, i)  # 평가손익
    return_list.append(evalPerc)

print(return_list)

table = pd.read_csv('./평가손익.csv')

# 주식 매도 주문
acc = objTrade.AccountNumber[0]  # 계좌번호
accFlag = objTrade.GoodsList(acc, 1)  # 주식상품 구분
print(acc, accFlag[0])
# objStockOrder = win32com.client.Dispatch("CpTrade.CpTd0311")
# objStockOrder.SetInputValue(0, "1")  # 1: 매도
# objStockOrder.SetInputValue(1, acc)  # 계좌번호
# objStockOrder.SetInputValue(2, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
# objStockOrder.SetInputValue(3, "A003540")  # 종목코드 - A003540 - 대신증권 종목
# objStockOrder.SetInputValue(4, 10)  # 매도수량 10주
# objStockOrder.SetInputValue(5, 14100)  # 주문단가  - 14,100원
# objStockOrder.SetInputValue(7, "0")  # 주문 조건 구분 코드, 0: 기본 1: IOC 2:FOK
# objStockOrder.SetInputValue(8, "01")  # 주문호가 구분코드 - 01: 보통

# 매도 주문 요청
objStockOrder.BlockRequest()

rqStatus = objStockOrder.GetDibStatus()
rqRet = objStockOrder.GetDibMsg1()
print("통신상태", rqStatus, rqRet)
if rqStatus != 0:
    exit()

    self.objBuySell.BlockRequest()

    rqStatus = self.objBuySell.GetDibStatus()
    rqRet = self.objBuySell.GetDibMsg1()
    print('통신상태', rqStatus, rqRet)
    if rqStatus != 0:

objBuySell.BlockRequest()

rqStatus = objBuySell.GetDibStatus()
rqRet = objBuySell.GetDibMsg1()
print('통신상태', rqStatus, rqRet)
if rqStatus != 0:
    exit()