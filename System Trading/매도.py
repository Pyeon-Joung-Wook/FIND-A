import sys
from PyQt5.QtWidgets import *
import win32com.client
import pandas as pd


# 설명: 주식 계좌잔고 종목(최대 200개)을 가져와 현재가  실시간 조회하는 샘플
# CpEvent: 실시간 현재가 수신 클래스
# CpStockCur : 현재가 실시간 통신 클래스
# Cp6033 : 주식 잔고 조회
# CpMarketEye: 복수 종목 조회 서비스 - 200 종목 현재가를 조회 함.

# CpEvent: 실시간 이벤트 수신 클래스
class CpEvent:
    def set_params(self, client):
        self.client = client

    def OnReceived(self):
        code = self.client.GetHeaderValue(0)  # 초
        name = self.client.GetHeaderValue(1)  # 초
        timess = self.client.GetHeaderValue(18)  # 초
        exFlag = self.client.GetHeaderValue(19)  # 예상체결 플래그
        cprice = self.client.GetHeaderValue(13)  # 현재가
        diff = self.client.GetHeaderValue(2)  # 대비
        cVol = self.client.GetHeaderValue(17)  # 순간체결수량
        vol = self.client.GetHeaderValue(9)  # 거래량

        # if (exFlag == ord('1')):  # 동시호가 시간 (예상체결)
        #     print("실시간(예상체결)", name, timess, "*", cprice, "대비", diff, "체결량", cVol, "거래량", vol)
        # elif (exFlag == ord('2')):  # 장중(체결)
        #     print("실시간(장중 체결)", name, timess, cprice, "대비", diff, "체결량", cVol, "거래량", vol)


# CpStockCur: 실시간 현재가 요청 클래스
class CpStockCur:
    def Subscribe(self, code):
        self.objStockCur = win32com.client.Dispatch("DsCbo1.StockCur")
        handler = win32com.client.WithEvents(self.objStockCur, CpEvent)
        self.objStockCur.SetInputValue(0, code)
        handler.set_params(self.objStockCur)
        self.objStockCur.Subscribe()

    def Unsubscribe(self):
        self.objStockCur.Unsubscribe()


# Cp6033 : 주식 잔고 조회
class Cp6033:
    def __init__(self):
        # 통신 OBJECT 기본 세팅
        self.objTrade = win32com.client.Dispatch("CpTrade.CpTdUtil")
        initCheck = self.objTrade.TradeInit(0)
        if (initCheck != 0):
            print("주문 초기화 실패")
            return

        acc = self.objTrade.AccountNumber[0]  # 계좌번호
        accFlag = self.objTrade.GoodsList(acc, 1)  # 주식상품 구분
        print(acc, accFlag[0])

        self.objRq = win32com.client.Dispatch("CpTrade.CpTd6033")
        self.objRq.SetInputValue(0, acc)  # 계좌번호
        self.objRq.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        self.objRq.SetInputValue(2, 50)  # 요청 건수(최대 50)

    # 실제적인 6033 통신 처리
    def rq6033(self, retcode):
        self.objRq.BlockRequest()

        # 통신 및 통신 에러 처리
        rqStatus = self.objRq.GetDibStatus()
        rqRet = self.objRq.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False

        cnt = self.objRq.GetHeaderValue(7)

        code_list =[]
        amount_list = []
        return_list = []

        print("종목개수: ", cnt)

        for i in range(cnt):
            code = self.objRq.GetDataValue(12, i)  # 종목코드
            retcode.append(code)
            if len(retcode) >= 200:  # 최대 200 종목만,
                break

            amount = self.objRq.GetDataValue(7, i)  # 체결잔고수량
            evalPerc = self.objRq.GetDataValue(11, i)  # 평가손익

            code_list.append(code)
            return_list.append(evalPerc)
            amount_list.append(amount)

        table = pd.DataFrame({'종목코드':code_list,
                              '수량':amount_list,
                              '평가손익':return_list})

        past_table = pd.read_csv('./평가손익.csv')

        stop_loss = table['평가손익'] - past_table['평가손익']

        # 주식 매도 주문
        for num in range(len(stop_loss)):
            if stop_loss[num] <-0:
                print(num)
                objBuySell = win32com.client.Dispatch('CpTrade.CpTd0311')  # 매수
                objBuySell.SetInputValue(0, '1')  # 1 매도 2 매수
                objBuySell.SetInputValue(1, acc)  # 계좌번호
                objBuySell.SetInputValue(2, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
                objBuySell.SetInputValue(3, past_table['종목코트'][num])  # 종목코드
                objBuySell.SetInputValue(4, past_table['수량'][num])  # 수량
                # objBuySell.SetInputValue(5, price)  # 주문단가
                objBuySell.SetInputValue(7, 0)  # 주문 조건 구분 코드, 0: 기본 1: IOC 2:FOK
                objBuySell.SetInputValue(8, '01')  # 주문호가 구분코드 - 01: 보통 03 시장가 05 조건부지정가

                # 주문 요청
                self.objBuySell.BlockRequest()

                rqStatus = self.objBuySell.GetDibStatus()
                rqRet = self.objBuySell.GetDibMsg1()
                print('통신상태', rqStatus, rqRet)
                if rqStatus != 0:
                    exit()

            else :
                pass

    def Request(self, retCode):
        self.rq6033(retCode)

        # 연속 데이터 조회 - 200 개까지만.
        while self.objRq.Continue:
            self.rq6033(retCode)
            print(len(retCode))
            if len(retCode) >= 200:
                break
        # for debug
        size = len(retCode)
        # for i in range(size):
        #     print(retCode[i])
        return True


class MyWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("PLUS API TEST")
        self.setGeometry(300, 300, 300, 180)
        self.isSB = False
        self.objCur = []

        btnStart = QPushButton("요청 시작", self)
        btnStart.move(20, 20)
        btnStart.clicked.connect(self.btnStart_clicked)

        btnStop = QPushButton("요청 종료", self)
        btnStop.move(20, 70)
        btnStop.clicked.connect(self.btnStop_clicked)

        btnExit = QPushButton("종료", self)
        btnExit.move(20, 120)
        btnExit.clicked.connect(self.btnExit_clicked)

    def StopSubscribe(self):
        if self.isSB:
            cnt = len(self.objCur)
            for i in range(cnt):
                self.objCur[i].Unsubscribe()
            print(cnt, "종목 실시간 해지되었음")
        self.isSB = False

        self.objCur = []

    def btnStart_clicked(self):
        self.StopSubscribe();
        codes = []
        obj6033 = Cp6033()
        if obj6033.Request(codes) == False:
            return

    def btnStop_clicked(self):
        self.StopSubscribe()

    def btnExit_clicked(self):
        self.StopSubscribe()
        exit()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
