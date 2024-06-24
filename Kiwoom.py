import pandas as pd
import re
import time
from datetime import datetime
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()

        ## 계좌 비번
        self.cert_password = ""

        # Kiwoom 인스터스 생성/초기화
        self._create_kiwoom_instance()

        # 종목 마스터 데이터프레임
        self.stock_master_df = pd.DataFrame()

        # 매수 대기 종목 관련 정보 리스트 초기화
        self.stock_master_df_list = []
        self.bought_stock_df = pd.DataFrame(columns=['s_code', 's_name', 's_bought_price', 's_bought_num']) 
        self.stock_waitlist = []
        self.stock_bought_list = self.bought_stock_df.s_code.tolist()
        self.cancel_list = []

        # 매수 희망 종목 손절가 정보 초기화
        self.loss_cut_df = pd.DataFrame(columns = ['s_code', 'loss_cut_price', 'target_profit_rate']) 

        # 매수 희망 종목 수 초기화
        self.stock_nums_to_buy = 1

        # 목표 수익률 클릭 수 초기화
        self.target_profit_click_cnt = 0

        # 실시간 감시 종목 등록용 데이터 프레임
        self.real_stock_df = pd.DataFrame(columns = ['item_code'])
        
        # 비동기 API 요청 응답 핸들러 on
        self.OnEventConnect.connect(self._event_connect)
        self.OnReceiveTrData.connect(self.tr_slot)
        #self.OnReceiveChejanData.connect(self.chejan_slot)
        
    #################### 초기 로그인 동작 #######################
    def _create_kiwoom_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def comm_connect(self):
        self.dynamicCall("CommConnect()")
        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()

    def _event_connect(self, err_code):
        if err_code == 0:
            print("Kiwoom OpenAPI+ connected")
        else:
            print("Kiwoom OpenAPI+ connection error")

        self.login_event_loop.exit()
    ############################################################

    #################### 접속 서버 파악  ########################
    def rq_connected_server(self):
        server_type = self.dynamicCall("GetLoginInfo(QString)", "GetServerGubun")
        self.server_name = ""
        
        if server_type == "1":
            self.server_name = "모의 서버"
        else:
            self.server_name = "실거래 서버"
    ############################################################

    #################### 계좌 기본 정보 조회 동작 ###############
    def rq_account_info(self, account_number_input):
        account_list = self.dynamicCall("GetLoginInfo(QString)", "ACCNO")
        account_number_list = account_list.split(';')
        account_number_list = [an.strip() for an in account_number_list]
        account_number = [a for a in account_number_list if a == account_number_input][0]
        self.account_number = account_number
    ############################################################

    ####################### 자금 현황 조회 ######################
    def rq_money_status(self):
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_number)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", self.cert_password)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "2")
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "계좌평가잔고내역요청", "opw00018", "0", "0")

        self.get_status_loop = QEventLoop()
        self.get_status_loop.exec_()
    ############################################################

    ####################### 예수금 조회 #########################
    def rq_takeaway(self):
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_number)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", self.cert_password)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "2")
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "계좌평가현황요청", "opw00004", "0", "0")

        self.get_takeaway_loop = QEventLoop()
        self.get_takeaway_loop.exec_()
    ############################################################

    ####################### 매수 대기 종목 등록 ##################
    def register_waitlist(self, stock_code):
        self.dynamicCall("SetInputValue(QString, QString)", "종목코드", stock_code)
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "주식기본정보요청", "opt10001", "0", "0")

        self.waitlist_loop = QEventLoop()
        self.waitlist_loop.exec_()
    ############################################################

    ############ 매수 대기 종목 정보 테이블 등록 ##################
    def register_master_df(self, s_code, s_name, s_current_price):
        # 현재 시간 및 기타 변수 정의
        s_time = datetime.now().strftime('%Y-%m-%d %H:%M')

        # 빈 stock_master_df 생성
        stock_master_df = pd.DataFrame(columns= ['s_time','s_code', 's_name', 's_current_price'])

        # 새로 추가할 데이터 딕셔너리
        stock_dic_obj = {
            's_time': s_time,
            's_code': s_code,
            's_name': s_name,
            's_current_price': s_current_price,
        }
        # 데이터 프레임에 추가
        stock_master_df.loc[len(stock_master_df)] = stock_dic_obj
        self.stock_master_df_list.append({'s_code': s_code ,'stock_master_df': stock_master_df, 'buy_status': 'not_bought'})
        
        # 실시간 시세 정보에 등록
        self.set_real_reg(s_code)
    ############################################################

    ################# 매수 종목 테이블 업데이트 ##################
    def update_bought_stock_df(self, s_code, s_name, s_bought_price, s_bought_num):

        bought_stock_dic = {
            's_code': s_code,
            's_name': s_name,
            's_bought_price': s_bought_price,
            's_bought_num': s_bought_num
        }

        self.bought_stock_df.index = range(self.bought_stock_df.shape[0])
        self.bought_stock_df.loc[len(self.bought_stock_df)] = bought_stock_dic

        for obj in self.stock_waitlist:
            if obj[0] == s_code:
                obj[2] = 'bought'
    ############################################################

    ####################### 종목 등록 ###########################
    def set_real_reg(self, item_code):
        real_type = ''
        
        if self.real_stock_df.shape[0] == 0:
            real_type = 0
            self.dynamicCall("SetRealReg(QString, QString, QString, QString)", ['0101', item_code, '10', real_type])
        else:
            real_type = 1
            self.dynamicCall("SetRealReg(QString, QString, QString, QString)", ['0101', item_code, '10', real_type])

        stock_dic = pd.DataFrame({'item_code': [item_code]})
        self.real_stock_df = pd.concat([self.real_stock_df, stock_dic], ignore_index=True)
    ############################################################

    ####################### 실시간 종목 삭제 #####################
    def remove_real_reg(self, item_code):
        self.dynamicCall("SetRealRemove(QString, QString)", ['0101', item_code])
    ############################################################

    ####################### 실시간 현재가 업데이트 ###############
    def update_current_price(self, s_code, s_current_price):
        s_time = datetime.now().strftime('%Y-%m-%d %H:%M')

        # 종목코드에 해당하는 데이터 프레임 리스트
        stock_master_search = [obj['stock_master_df'] for obj in self.stock_master_df_list if obj['s_code'] == s_code]
        s_name = stock_master_search[0].s_name.iloc[0]
        
        # 새로 추가할 데이터 딕셔너리
        stock_dic_obj = {
            's_time': s_time,
            's_code': s_code,
            's_name': s_name,
            's_current_price': s_current_price,
        }

        for stock_master_df in stock_master_search:
            s_time = datetime.now().strftime('%Y-%m-%d %H:%M')
            
            # 중복 검색
            duplicate_search = stock_master_df.loc[
                (stock_master_df.s_time == s_time) &
                (stock_master_df.s_code == s_code)
            ]

            # 중복이 없으면 새 행 추가
            if duplicate_search.shape[0] == 0:
                stock_master_df.loc[len(stock_master_df)] = stock_dic_obj
            else:
                # 중복이 있으면 현재 가격 업데이트
                stock_master_df.loc[duplicate_search.index, 's_current_price'] = s_current_price
    ############################################################

    ####################### 2분 이평선 계산 #####################
    def calculate_smoothing_line(self, s_code):
        # 3분 이동평균 값 추가 & 이동평균선 차분값 추가
        stock_master_search = [obj['stock_master_df'] for obj in self.stock_master_df_list if obj['s_code'] == s_code]

        for stock_master_df in stock_master_search:
            stock_master_df['min2_smoothing'] = stock_master_df['s_current_price'].rolling(window=2).mean().tolist()
            stock_master_df['min5_smoothing'] = stock_master_df['s_current_price'].rolling(window=5).mean().tolist()
            stock_master_df['min15_smoothing'] = stock_master_df['s_current_price'].rolling(window=15).mean().tolist()
            stock_master_df['min60_smoothing'] = stock_master_df['s_current_price'].rolling(window=60).mean().tolist()
            stock_master_df['min90_smoothing'] = stock_master_df['s_current_price'].rolling(window=90).mean().tolist()
            stock_master_df['min120_smoothing'] = stock_master_df['s_current_price'].rolling(window=120).mean().tolist()
            stock_master_df['min2_smoothing_diff'] = stock_master_df.min2_smoothing - stock_master_df.min2_smoothing.shift(1)
            stock_master_df['min5_smoothing_diff'] = stock_master_df.min5_smoothing - stock_master_df.min5_smoothing.shift(1)
            stock_master_df['min15_smoothing_diff'] = stock_master_df.min15_smoothing - stock_master_df.min15_smoothing.shift(1)
            stock_master_df['min60_smoothing_diff'] = stock_master_df.min60_smoothing - stock_master_df.min60_smoothing.shift(1)
    ############################################################

    ##################### 이전 종가 데이터 수신 ##################
    def get_past_price_data(self, stock_code, rq_cnt=180):
        self.dynamicCall("SetInputValue(QString, QString)", "종목코드", stock_code)
        self.dynamicCall("SetInputValue(QString, QString)", "틱범위", "1")
        self.dynamicCall("SetInputValue(QString, QString)", "수정주가구분", "1")
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "주식분봉차트조회요청", "OPT10080", rq_cnt, "0")
        
        self.past_data_loop = QEventLoop()
        self.past_data_loop.exec_()
    ############################################################

    ####################### 매수/매도 주문 처리 ######################    
    def rq_order(self, rqname, order_type, code, quantity, screen = "0101", acc_no = "", price=0, hoga_gb="03", org_order_no=""):

        acc_no = self.account_number              
        
        self.dynamicCall(
                "SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)", 
                [rqname, screen, acc_no, order_type, code, quantity, price ,hoga_gb, org_order_no])    
    ############################################################

    #################### 이벤트 핸들러 슬롯 #####################
    def tr_slot(self, sScrNo, sRQName, sTrCode, srecordName, sPrevNext):
        ## 오버 나잇 종목 정보 수신
        if sRQName == "계좌평가잔고내역요청":            
            ## 종목 관련 정보 요청 (각 종목별 개별 데이터)
            self.num_of_bought_stocks = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName,0, "조회건수")
            self.num_of_bought_stocks = int(self.num_of_bought_stocks)

            tr_items = ["종목번호", "종목명", "보유수량", "매입가", "현재가", "수익률(%)"]
            self.info_list_by_stocks = []

            for idx in range(self.num_of_bought_stocks):
                ## 개별 종목 별 (종목번호, 종목명, 보유수량, ...) 임시 리스트 생성
                info_packet = []
                for item in tr_items:
                    result = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, idx, item)
                    if result == "":
                        result = "0"
                    
                    info_packet.append(result)
                
                ## 종목 별 리스트들을 담은 info_list_by_stocks에 개별 종목별 정보 리스트 저장 
                self.info_list_by_stocks.append(info_packet)
                
            ## 종목 별 개별 데이터 전역 변수에 저장
            for info in self.info_list_by_stocks:
                stock_code = info[0].strip()[1:]
                stock_name = info[1].strip()
                stock_num = int(info[2].strip())
                stock_buy_price = int(info[3].strip())
                stock_current_price = int(info[4].strip())
                stock_profit_rate = round(float((int(info[4].strip()) - int(info[3].strip())) / (int(info[3].strip())) * 100),3)

                ## stock_master_df에 등록
                self.register_master_df(s_code=stock_code, s_name=stock_name, s_current_price=stock_current_price)
                self.calculate_smoothing_line(s_code=stock_code)

                ## bought_stock_df에 등록
                self.update_bought_stock_df(s_code=stock_code, s_name=stock_name, s_bought_price=stock_buy_price, s_bought_num=stock_num)

                ## add stock_code to stock_bought_list 
                self.stock_bought_list = self.bought_stock_df.s_code.tolist()

                ## add to waitlist
                self.stock_waitlist.append([stock_code, stock_name, 'bought'])

            ## 디폴트 수익률 지정
            if self.target_profit_click_cnt == 0:
                self.target_profit = 2.0
            
            self.get_status_loop.exit()

        ## 예수금(D+2) 정보 수신
        if sRQName == "계좌평가현황요청":
            takeaway_money = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName,0, "D+2추정예수금")
            self.takeaway_money = int(takeaway_money)
            self.trading_money = self.takeaway_money
            self.get_takeaway_loop.exit()

        ## 매수 대기 종목 정보 수신
        if sRQName == "주식기본정보요청":
            stock_code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName,0, "종목코드")
            stock_name = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName,0, "종목명")
            stock_current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName,0, "현재가")

            self.stock_code_waiting = stock_code.strip()
            self.stock_name_waiting = stock_name.strip()
            try:
                self.stock_current_price = abs(int(stock_current_price))
            except ValueError:
                self.stock_current_price = 0
            
            bought_status = 'not_bought'

            self.is_new_waitlist = len([wl['s_code'] for wl in self.stock_master_df_list if wl['s_code'] == self.stock_code_waiting]) == 0
            
            if (self.stock_name_waiting != "") & (self.is_new_waitlist):
                self.stock_waitlist.append([self.stock_code_waiting, self.stock_name_waiting, bought_status])
                self.register_master_df(s_code=self.stock_code_waiting, s_name=self.stock_name_waiting, s_current_price=self.stock_current_price)
            
            self.waitlist_loop.exit()
        if sRQName == "주식분봉차트조회요청":
            stock_code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "종목코드")
            stock_code = stock_code.strip()

            master_df = [obj['stock_master_df'] for obj in self.stock_master_df_list if obj['s_code'] == stock_code][0]
            s_name_data = master_df.iloc[0].s_name

            s_time_list = []
            s_current_price_list = []

            for data_idx in range(500):
                time_stamp = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, data_idx, "체결시간")
                time_stamp = datetime.strptime(time_stamp.strip(), '%Y%m%d%H%M%S').strftime('%Y-%m-%d %H:%M')

                price_data = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, data_idx, "현재가")
                price_data = abs(int(price_data.strip()))

                s_time_list.append(time_stamp)
                s_current_price_list.append(price_data)

            s_code_list = [stock_code] * len(s_time_list)
            s_name_list = [s_name_data] * len(s_time_list)

            past_price_dic = {'s_time': s_time_list, 's_code': s_code_list, 's_name': s_name_list, 's_current_price': s_current_price_list}
            master_df_replace = pd.DataFrame(past_price_dic)
            master_df_replace.sort_values(by="s_time", ascending=True, inplace = True)
            master_df_replace.index = range(master_df_replace.shape[0])

            [obj for obj in self.stock_master_df_list if obj['s_code'] == stock_code][0]['stock_master_df'] = master_df_replace
            self.calculate_smoothing_line(s_code=stock_code)

            self.past_data_loop.exit()
    ############################################################