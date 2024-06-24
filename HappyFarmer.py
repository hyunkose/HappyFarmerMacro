import time
import sys
import re
import pandas as pd
import pickle
import locale
from datetime import datetime, timedelta
from Kiwoom import *

from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5 import uic
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap 

form_class = uic.loadUiType("assets/stock_bot.ui")[0]
locale.setlocale(locale.LC_ALL, '')

class BotWindow(QMainWindow, form_class):
    def __init__(self):
        ## StockBot 프로그램 실행
        super().__init__()
        self.setupUi(self)
        self.setWindowTitle("Happy Farmer")

        pixmap = QPixmap("assets/HappyFarmerLogo.png")
        self.logo_image.resize(50,50)
        self.logo_image.setPixmap(pixmap.scaled(self.logo_image.width(), self.logo_image.height(), Qt.KeepAspectRatio))
    
        ## 키움 OpenAPI 로그인 처리
        self.kiwoom = Kiwoom()
        self.kiwoom.comm_connect()

        ## 공인 인증서 비밀번호 입력 받기
        cert_password, is_ok = QInputDialog.getText(
                    self,
                    "공인인증서 비밀번호 입력창",
                    "\n\n  1) 비밀번호가 가려지지 않습니다! 공공장소에서 이용하지마세요.\n\n  2) 2~6자리 공인인증서 비밀번호 입력\n\n",
                    QLineEdit.PasswordEchoOnEdit, 
                    "",)
        
        self.kiwoom.cert_password = cert_password
        
        ## Stock Bot Trader 접속 서버 조회
        self.kiwoom.rq_connected_server()
        self.server_label.setText(self.kiwoom.server_name)

        ## 계좌 번호 입력, 해당 계좌 정보 조회
        while True:
            try:
                account_number_input, is_ok =  QInputDialog.getText(self, "계좌 번호 입력창","\n\n 접속할 계좌의 계좌 번호를 입력하세요. ", QLineEdit.PasswordEchoOnEdit, "",)
                account_number_input = account_number_input.strip()
                self.kiwoom.rq_account_info(account_number_input)
                account_num = self.kiwoom.dynamicCall("GetLoginInfo(QString)", ["ACCNO"])
                self.account_label.setText(self.kiwoom.account_number)
                break
            except IndexError:
                QMessageBox.information(self, "계좌 번호 입력 오류", "정확한 계좌번호를 입력하세요")

        ## 예수금 (D+2) 조회
        self.kiwoom.rq_takeaway()
        takeaway_money_label = locale.currency(self.kiwoom.takeaway_money, grouping=True)
        self.takeaway_label.setText(takeaway_money_label)

        ## 예비금
        self.reserve_money_rate = 0.0
        self.kiwoom.trading_money = self.kiwoom.takeaway_money * (100-self.reserve_money_rate) / 100
        self.takeaway_info_label.setText('※ 전체 예수금의 {0}%만 트레이딩에 쓰입니다'.format(100-self.reserve_money_rate))
        self.reserver_money_save.clicked.connect(self.set_reserve_money)

        ## 매수 대기 종목
        #self.show_waitlist()
        self.buy_stock_save.clicked.connect(self.add_waitlist)
        self.buy_stock_delete.clicked.connect(self.delete_waitlist)

        ## 매수 대기 종목 타이틀 변경
        self.waitlist_group.setTitle('매수 대기 종목 (구매 희망 종목 수 {0}개)'.format(self.kiwoom.stock_nums_to_buy))

        ## 긴급 손절 모드 초기화 (off)
        self.emergency_sell_mode = False
        self.emergency_sell_off.setChecked(True)
        self.emergency_sell_on.clicked.connect(self.set_emergency_sell_mode)
        self.emergency_sell_off.clicked.connect(self.set_emergency_sell_mode)

        ## 예수금 정보 수신
        self.kiwoom.rq_money_status()
  
        ## 매수 희망 종목 수 지정
        self.desired_num_stocks_save.clicked.connect(self.set_stock_nums_to_buy)

        ## 손절가 변경
        self.loss_cut_save.clicked.connect(self.change_loss_cut)

        ## 이미 매수한 종목 종류 수 정보 저장
        self.already_bought_stocks = []

        ## 목표 수익률 값 지정 or 수정
        self.sell_option_save.clicked.connect(self.set_target_profit)

        ## 매수 종목 감시 실행
        self.kiwoom.OnReceiveRealData.connect(self.real_slot)

        ## 종목 매수 완료시 데이터 최신화 & 화면 refresh
        self.kiwoom.OnReceiveChejanData.connect(self.chejan_slot)

        ## 종목 매수 증거금 부족 handling
        self.kiwoom.OnReceiveMsg.connect(self.order_slot)

        ## 프로그램 매도가 아닌 사용자의 영웅문 앱을 통한 매도 여부 판단
        self.is_program_sell = False

        ## 오버나잇 종목 데이터 불러오기
        # 이전 분봉 데이터 수신
        overnight_stocks = set(self.kiwoom.bought_stock_df.s_code.tolist())
        for stock_code in overnight_stocks:
            self.kiwoom.get_past_price_data(stock_code, "0")
        
        # bought_stock_df 최신화 (전일 분할 매수 된 정보를 반영하게끔 파일 형태의 저장 정보 불러오기)
        try:
            past_bought_stocks = set(self.kiwoom.bought_stock_df.s_code) ## 기존 키움 서버에서의 데이터 내 매수 종목 코드 정보

            with open("./assets/bought_stock_df", "rb") as f:
                bought_stock_df_past = pickle.load(f)

            # bought_stock_df 최신화 (키움 서버의 데이터와 비교하여, 이미 사용자가 장외에 매도한 종목은 bought_stock_df에서 제외)   
            bought_stock_df_past = bought_stock_df_past.loc[bought_stock_df_past.s_code.isin(past_bought_stocks)]
            
            self.kiwoom.bought_stock_df = bought_stock_df_past
            self.kiwoom.stock_bought_list = self.kiwoom.bought_stock_df.s_code.tolist()
            self.kiwoom.num_of_bought_stocks = self.kiwoom.bought_stock_df.shape[0]
            
            # stock_master_df_list 최신화 (사용자에 의해 장외 매도된 종목은 제외 + 이미 구매된 종목의 buy_status 'bought'으로 변경) 
            # stock_master_df_list_past = [obj for obj in stock_master_df_list_past if obj['s_code'] in past_bought_stocks]
            for obj in self.kiwoom.stock_master_df_list:
                if obj['s_code'] in past_bought_stocks:
                    obj['buy_status'] = 'bought'
            
            # loss_cut_df 최신화
            with open("./assets/loss_cut_df", "rb") as f:
                self.kiwoom.loss_cut_df = pickle.load(f)

        except FileNotFoundError:
            ## 프로그램을 처음 실행하여, 이전 저장 데이터가 없는 경우 키움 서버에서 종목 정보 수신
            # 처음 실행 하여 loss_cut_df가 없거나 비어있을 시, 자동 손절가 0원 적용
            for stock_code in past_bought_stocks:
                loss_cut_dic = {'s_code': stock_code, 'loss_cut_price': 0, 'target_profit_rate': 2}
                self.kiwoom.loss_cut_df.loc[len(self.kiwoom.loss_cut_df)] = loss_cut_dic

        self.show_waitlist()
        self.show_bought_status()
    #################### 예비금 비중 control ####################
    def set_reserve_money(self):
        self.reserve_money_rate = self.reserve_money_select.value()
        self.kiwoom.trading_money = self.kiwoom.takeaway_money * (100-self.reserve_money_rate) / 100
        text = '※ 전체 예수금의 {0}%만 트레이딩에 쓰입니다'.format(100-self.reserve_money_rate)
        self.takeaway_info_label.setText(text)
        
        QMessageBox.information(self, "예비금 비중 변경", "예비금 비중 변경 완료")
    ############################################################

    #################### 매수 대기 종목 추가 ####################
    def add_waitlist(self):
        stock_code = self.buy_stock_input.text().strip()
        self.kiwoom.register_waitlist(stock_code)
        stock_idx_list = [idx for idx, c in enumerate(self.kiwoom.stock_waitlist) if c[0] == stock_code]

        not_cancel_before = len([s for s in self.kiwoom.cancel_list if s == stock_code]) == 0

        loss_cut_price = self.loss_cut_price.text()

        if self.kiwoom.stock_name_waiting == "":
            QMessageBox.information(self, "종목코드 오류", "올바른 종목코드를 입력해주세요")
        elif (not self.kiwoom.is_new_waitlist) & (not_cancel_before):
            QMessageBox.information(self, "종목코드 오류", "이미 추가한 종목입니다")
        else:
            try:
                ## 종목 손절가 정보 추가
                self.kiwoom.register_waitlist(stock_code)
                loss_cut_dic = {'s_code': stock_code, 'loss_cut_price': int(loss_cut_price), 'target_profit_rate': 2}
                self.kiwoom.loss_cut_df.loc[len(self.kiwoom.loss_cut_df)] = loss_cut_dic
                self.kiwoom.calculate_smoothing_line(s_code=stock_code)
                self.kiwoom.get_past_price_data(stock_code, "0")
                QMessageBox.information(self, "종목 매수 예약창", "종목 매수 예약 완료")
            except ValueError:
                ## 유효하지 않은 손절가 입력시 
                remove_idx =  [idx for idx, obj in enumerate(self.kiwoom.stock_master_df_list) if obj['s_code'] == stock_code][0]
                del self.kiwoom.stock_master_df_list[remove_idx]
                del self.kiwoom.stock_waitlist[stock_idx_list[0]]
                self.kiwoom.remove_real_reg(stock_code)
                QMessageBox.information(self, "손절가 오류", "정확한 손절가를 입력해주세요")

        self.show_waitlist()
    ############################################################

    #################### 매수 대기 종목에서 제거 #################
    def delete_waitlist(self):
        stock_code = self.buy_stock_input.text().strip()
        stock_idx_list = [idx for idx, c in enumerate(self.kiwoom.stock_waitlist) if c[0] == stock_code]

        if len(stock_idx_list) == 0:
            QMessageBox.information(self, "종목 매수 예약 취소창", "해당 종목이 예약 리스트에 없습니다")
        else:
            remove_idx =  [idx for idx, obj in enumerate(self.kiwoom.stock_master_df_list) if obj['s_code'] == stock_code][0]
            del self.kiwoom.stock_master_df_list[remove_idx]
            del self.kiwoom.stock_waitlist[stock_idx_list[0]]
            self.kiwoom.remove_real_reg(stock_code)
            self.kiwoom.cancel_list.append(stock_code)

            ## loss_cut_df 내 종목 정보 삭제
            self.kiwoom.loss_cut_df = self.kiwoom.loss_cut_df.loc[self.kiwoom.loss_cut_df.s_code != stock_code]
            self.kiwoom.loss_cut_df.index = range(self.kiwoom.loss_cut_df.shape[0])

            self.show_waitlist()

            QMessageBox.information(self, "종목 매수 예약 취소창", "매수 예약 취소 완료")
    ############################################################

    ####################### 손절가 변경 #########################
    def change_loss_cut(self):
        stock_code = self.loss_cut_stock_code.text().strip()
        loss_cut_price = self.loss_cut_change.text().strip()

        stock_code_search = [obj[0] for obj in self.kiwoom.stock_waitlist if obj[0] == stock_code]
        try:
            if len(stock_code_search) == 0:
                QMessageBox.information(self, "종목 코드 오류", "정확한 종목 코드를 입력해주세요")
            else:
                search_index = self.kiwoom.loss_cut_df.loc[self.kiwoom.loss_cut_df.s_code == stock_code_search[0]].index[0]
                target_profit_rate = self.kiwoom.loss_cut_df.iloc[search_index].target_profit_rate
                self.kiwoom.loss_cut_df.iloc[search_index] = {'s_code': stock_code_search[0], 'loss_cut_price': int(loss_cut_price), 'target_profit_rate': target_profit_rate}
                QMessageBox.information(self, "손절가 변경창", "손절가 변경완료")
        except ValueError:
            QMessageBox.information(self, "손절가 오류", "정확한 손절가를 입력하세요")

        self.show_waitlist()
        self.show_bought_status()
    ############################################################

    ################# 긴급 손절 모드 변경 ########################
    def set_emergency_sell_mode(self):
        
        mode_on = self.emergency_sell_on.isChecked()
        mode_off = self.emergency_sell_off.isChecked()

        if mode_on:
            self.emergency_sell_mode = True
            QMessageBox.information(self, "긴급 손절모드 변경창", "긴급 자동 손절모드 On")
        elif mode_off:
            self.emergency_sell_mode = False
            QMessageBox.information(self, "긴급 손절모드 변경창", "긴급 자동 손절모드 Off")
    ############################################################


    #################### 매수 대기 종목 조회  ###################
    def show_waitlist(self):
        table_column = ['종목번호', '종목명', '매수상태', '최소 수익률(%)','손절가']

        self.waitlitst_table.setColumnCount(len(table_column))
        self.waitlitst_table.setRowCount(len(self.kiwoom.stock_waitlist))
        self.waitlitst_table.setHorizontalHeaderLabels(table_column)
                
        ## 매수 대기 종목 테이블 구성
        for row_idx in range(len(self.kiwoom.stock_waitlist)):
            for col_idx, info in enumerate(self.kiwoom.stock_waitlist[row_idx]):
                text = ""
                if (col_idx == 2) & (info == "bought"):
                    text = "매수 완료"
                elif (col_idx == 2) & (info == "not_bought"):
                    text = "매수 대기중"
                else:
                    text = info
                ## 종목 기본 정보 표시
                text_obj = QTableWidgetItem(text)
                text_obj.setTextAlignment(Qt.AlignCenter)
                self.waitlitst_table.setItem(row_idx,col_idx,text_obj)

                ## 종목 목표 수익률 표시
                if col_idx == 0:
                    target_profit = self.kiwoom.loss_cut_df.loc[self.kiwoom.loss_cut_df.s_code == info, "target_profit_rate"].iloc[0] 

                    text = str(target_profit)+"%"
                    text_obj = QTableWidgetItem(text)
                    text_obj.setTextAlignment(Qt.AlignCenter)
                    self.waitlitst_table.setItem(row_idx,3,text_obj)


                ## 종목 손절가 정보 추가 표시
                try:
                    loss_cut_price = self.kiwoom.loss_cut_df.loc[self.kiwoom.loss_cut_df.s_code == self.kiwoom.stock_waitlist[row_idx][0]].loss_cut_price.iloc[0]
                except:
                    loss_cut_price = 0

                text = locale.currency(int(loss_cut_price), grouping=True)
                text_obj = QTableWidgetItem(text)
                text_obj.setTextAlignment(Qt.AlignCenter)
                self.waitlitst_table.setItem(row_idx,4,text_obj)

            
    ############################################################

    #################### 매수 종목 정보 조회 동작 ###############
    def show_bought_status(self):
        try:
            ## 프로그램 화면에 종목 정보 반영
            table_column = ['종목번호', '종목명', '보유수량', '매입가', '현재가', '현재 수익률(%)', '최소 수익률(%)', '손절가']
            
            self.bought_info_table.setColumnCount(len(table_column))
            self.bought_info_table.setRowCount(self.kiwoom.num_of_bought_stocks)
            self.bought_info_table.setHorizontalHeaderLabels(table_column)

            stock_codes = self.kiwoom.bought_stock_df.s_code.tolist()
            stock_names = self.kiwoom.bought_stock_df.s_name.tolist()
            stock_nums = self.kiwoom.bought_stock_df.s_bought_num.tolist()
            stock_buy_prices = self.kiwoom.bought_stock_df.s_bought_price.tolist()
            stock_current_prices = []
            stock_profit_rates = []
            target_profit_list = []
            
            for code in stock_codes:
                target_profit = self.kiwoom.loss_cut_df.loc[self.kiwoom.loss_cut_df.s_code == code, "target_profit_rate"].iloc[0]
                target_profit_list.append(target_profit)

            for idx, code in enumerate(stock_codes):
                try:
                    stock_idx = [idx for idx, obj in enumerate(self.kiwoom.stock_master_df_list) if obj['s_code'] == code][0]
                    stock_current_price = self.kiwoom.stock_master_df_list[stock_idx]['stock_master_df'].iloc[self.kiwoom.stock_master_df_list[stock_idx]['stock_master_df'].shape[0]-1].s_current_price
                    stock_profit_rate = round((stock_current_price - stock_buy_prices[idx]) / stock_buy_prices[idx] * 100, 3)

                    stock_current_prices.append(stock_current_price)
                    stock_profit_rates.append(stock_profit_rate)
                except IndexError:
                    ## 실시간 시세 등록 해제 (동일 종목 여러번 매수한 상황인 여부를 판단해서, 종목이 모두 팔렸으면 실시간 등록 해제)
                    stock_search_list = self.kiwoom.bought_stock_df.loc[self.kiwoom.bought_stock_df.s_code == code]

                    if stock_search_list.shape[0] == 0:
                        self.kiwoom.remove_real_reg(code)

            stock_info_list = [stock_codes, stock_names, stock_nums,
                            stock_buy_prices, stock_current_prices, stock_profit_rates,
                            target_profit_list]

            ## 매도 대기 종목 테이블 구성
            for row_idx in range(self.kiwoom.num_of_bought_stocks):
                for col_idx, info in enumerate(stock_info_list):
                    text = ""
                    if col_idx in (0, 1):
                        text = info[row_idx]
                    elif col_idx == 2:
                        text = str(info[row_idx])+"주"
                    elif col_idx in (3,4):
                        text = locale.currency(info[row_idx], grouping=True)
                    elif col_idx in (5,6):
                        text = str(info[row_idx])+"%"
                    
                    ## 종목 기본 정보 표시
                    text_obj = QTableWidgetItem(text)
                    text_obj.setTextAlignment(Qt.AlignCenter)
                    self.bought_info_table.setItem(row_idx,col_idx,text_obj)

                    ## 종목 손절가 정보 추가 표시
                    try:
                        loss_cut_price = self.kiwoom.loss_cut_df.loc[self.kiwoom.loss_cut_df.s_code == self.kiwoom.bought_stock_df.iloc[row_idx]['s_code']].loss_cut_price.iloc[0]
                    except:
                        loss_cut_price = 0

                    text = locale.currency(loss_cut_price, grouping=True)
                    text_obj = QTableWidgetItem(text)
                    text_obj.setTextAlignment(Qt.AlignCenter)
                    self.bought_info_table.setItem(row_idx,7,text_obj)
        except IndexError:
            pass
    ############################################################
    
    ############# 프로그램 종료 이전 가격 감시 삭제 ###############
    def closeEvent(self, event):
        ## 실시간 가격 등록 해제
        self.kiwoom.dynamicCall("SetRealRemove(QString, QString)", 'ALL', 'ALL')
        ## 금일 분할 매수 정보 포함한 종목 데이터 파일로 저장
        with open("./assets/bought_stock_df", "wb") as f:
            pickle.dump(self.kiwoom.bought_stock_df, f)
    ############################################################

    ############## 매수 종목 목표 수익률 지정 ####################
    def set_target_profit(self):
        try:
            target_profit = self.stock_profit_rate.text().strip()
            target_profit = float(target_profit)
        
            ## (선택) 메뉴 선택 시 목표 수익률 스핀 박스 값 다시 0으로
            self.kiwoom.target_profit_click_cnt += 1
            stock_code = self.profit_rate_stock_code.text().strip()
            
            change_num = 0
            for idx in range(self.kiwoom.loss_cut_df.shape[0]):
                stock_code_current = self.kiwoom.loss_cut_df.iloc[idx, 0]

                if stock_code_current == stock_code:
                    self.kiwoom.loss_cut_df.iloc[idx, 2] = target_profit
                    change_num += 1
            
            self.show_bought_status()
            self.show_waitlist()

            if change_num > 0:
                QMessageBox.information(self, "트레일링 시작 조건 변경 확인창", "트레일링 시작 조건 변경 완료")
            elif change_num == 0:
                QMessageBox.information(self, "트레일링 시작 조건 변경 오류창", "정확한 종목코드를 입력하세요")
        except ValueError:
            QMessageBox.information(self, "트레일링 시작 조건 변경 오류창", "정확한 최소 수익률을 입력하세요")
    ############################################################

    ################## 희망 매수 종목 수 지정 ####################
    def set_stock_nums_to_buy(self):
        stock_nums_to_buy = self.desired_num_stocks.value()
        self.kiwoom.stock_nums_to_buy = int(stock_nums_to_buy)
        self.waitlist_group.setTitle('매수 대기 종목 (구매 희망 종목 수 {0}개)'.format(self.kiwoom.stock_nums_to_buy))
        QMessageBox.information(self, "매수 희망 종목 수 변경 확인창", "매수 희망 종목 수 변경 완료")
    ############################################################

    ############## 매수 신청 처리 (시장가) ######################
    def send_buy_order(self, code, quantity):
        time.sleep(2)
        rqname="주식매수"
        order_type=1

        self.buy_request_stock_code = code
        self.buy_request_quantity = quantity

        self.kiwoom.rq_order(rqname, order_type, code, quantity)
    ############################################################

    ############## 매도 신청 처리 (시장가) ######################
    def send_sell_order(self, code, quantity, stock_idx):
        time.sleep(2)
        rqname="주식매도"
        order_type=2
        self.sell_request_stock_code = code
        self.kiwoom.rq_order(rqname, order_type, code, quantity)

        try:
            self.kiwoom.bought_stock_df.drop(stock_idx, inplace=True)

            ## rearrange index
            self.kiwoom.bought_stock_df.index = range(self.kiwoom.bought_stock_df.shape[0])

            ## 보유 종목 수 최신화
            self.kiwoom.num_of_bought_stocks = self.kiwoom.bought_stock_df.shape[0]

            ## stock_bought_list (real slot의 for 문에서 쓰이는 리스트) 데이터 최신화
            self.kiwoom.stock_bought_list = self.kiwoom.bought_stock_df.s_code.tolist()

        except IndexError:
            print('index error at send_sell_order')
        except KeyError:
            print('key error at send_sell_order') ## 초과 물량 매도 주문시
        
        ## 프로그램에 의한 매도 판단 변수 변경
        self.is_program_sell = True
        self.show_bought_status()
    #############################################################

    ################### 종목 실시간 가격 감시 #####################
    def real_slot(self, sJongmokCode, sRealType, sRealData):
        for jongmok_code_original in [sJongmokCode]:
            ### 현재가 실시간 조회
            current_price = ""

            if sRealType == "주식체결":
                try:
                    ## 종목 정보 실시간 업데이트
                    current_price = self.kiwoom.dynamicCall("GetCommRealData(QString, QString)", sJongmokCode, "10")
                    current_price = re.findall(r'\d+', current_price)
                    current_price = ''.join(current_price)
                    current_price = int(current_price)

                    self.kiwoom.update_current_price(s_code=jongmok_code_original, s_current_price=current_price)
                    self.kiwoom.calculate_smoothing_line(s_code=jongmok_code_original)
                                        
                    stock_master_df = [obj['stock_master_df'] for obj in self.kiwoom.stock_master_df_list if obj['s_code'] == jongmok_code_original][0] 
                except Exception as e:
                    continue

                ###### 매수 타점 잡기 ######
                try:
                    status_list = [obj[2] for obj in self.kiwoom.stock_waitlist if obj[0] == jongmok_code_original]

                    if status_list[0] == "not_bought":
                        
                        ## 현재 시점 이평선 정보
                        current_min2 = stock_master_df.iloc[stock_master_df.shape[0]-1].min2_smoothing
                        current_min15 = stock_master_df.iloc[stock_master_df.shape[0]-1].min15_smoothing
                        current_min60 = stock_master_df.iloc[stock_master_df.shape[0]-1].min60_smoothing
                        current_min90 = stock_master_df.iloc[stock_master_df.shape[0]-1].min90_smoothing
                        current_min120 = stock_master_df.iloc[stock_master_df.shape[0]-1].min120_smoothing
                        
                        ############## 매수 조건 1 : 저점에서 골든 크로스 확인 후 상승 추세에서 매수한다 ################                        
                        ## 15분 이평선 변곡점 조건
                        #buy_condition1 = (stock_master_df.iloc[stock_master_df.shape[0]-2].min15_smoothing_diff <= 0) & (stock_master_df.iloc[stock_master_df.shape[0]-1].min15_smoothing_diff > 0)

                        ## 15이평선 & 60이평선 2틱 전 골든크로스 조건
                        before_3ticks_smoothing = stock_master_df.iloc[stock_master_df.shape[0]-4].min15_smoothing < stock_master_df.iloc[stock_master_df.shape[0]-4].min60_smoothing
                        before_2ticks_smoothing = stock_master_df.iloc[stock_master_df.shape[0]-3].min15_smoothing >= stock_master_df.iloc[stock_master_df.shape[0]-3].min60_smoothing
                        before_1tick_smoothing = stock_master_df.iloc[stock_master_df.shape[0]-2].min15_smoothing > stock_master_df.iloc[stock_master_df.shape[0]-2].min60_smoothing
                        current_smoothing = stock_master_df.iloc[stock_master_df.shape[0]-1].min15_smoothing > stock_master_df.iloc[stock_master_df.shape[0]-1].min60_smoothing
                        ## 골든 크로스 이후 현재가 안정적 상승 유지 조건 (현재가 캔들 > 60이평선) 
                        before_1tick_candle = stock_master_df.iloc[stock_master_df.shape[0]-2].s_current_price > stock_master_df.iloc[stock_master_df.shape[0]-2].min60_smoothing
                        current_candle = stock_master_df.iloc[stock_master_df.shape[0]-1].s_current_price > stock_master_df.iloc[stock_master_df.shape[0]-1].min60_smoothing
                        ## 골든 크로스 이후 2이평선 기울기 상승 유지 조건
                        current_min2_slope = stock_master_df.iloc[stock_master_df.shape[0]-1].min2_smoothing_diff > 0

                        buy_condition1 = before_3ticks_smoothing & before_2ticks_smoothing & before_1tick_smoothing & current_smoothing & before_1tick_candle & current_candle & current_min2_slope
                        ## 장기 이평선 역배열 조건
                        #buy_condition2 = current_min15 < current_min60 < current_min90 < current_min120
                        buy_condition2 = current_min60 < current_min90 < current_min120                     
                        #############################################################################################
                        
                        ## 횡보장 회피 조건
                        max_val = max([current_min60, current_min90, current_min120]) #max([current_min15, current_min60, current_min90, current_min120])
                        min_val = min([current_min60, current_min90, current_min120]) #min([current_min15, current_min60, current_min90, current_min120])

                        disparity = abs(max_val - min_val)
                        disparity_ratio = disparity / max_val * 100
                        avoid_stable_spot = disparity_ratio >= 0.4 #1.0

                        # 장 초반 절반은 관망 (09:00 ~ 11:30 까지)
                        current_timestamp = datetime.today()
                        std_time = datetime(current_timestamp.year, current_timestamp.month, current_timestamp.day, 11, 30)
                        buy_condition3 = current_timestamp >= std_time

                        # 종목 가격 하락세 과도하게 오래 지속시 매수 금지
                        today_std_date = datetime.today().strftime('%Y-%m-%d')
                        idx_list = [idx for idx, t in enumerate(stock_master_df.s_time) if t[:10] == today_std_date]
                        stock_master_df_today = stock_master_df.iloc[idx_list]
                        stock_master_df_today.index = range(stock_master_df_today.shape[0])

                        reverse_token_list = []

                        for idx in range(stock_master_df_today.shape[0]):
                            current_min60_diff = stock_master_df_today.iloc[idx].min60_smoothing_diff
                            #current_min90 = stock_master_df_today.iloc[idx].min90_smoothing
                            #current_min120 = stock_master_df_today.iloc[idx].min120_smoothing

                            if current_min60_diff < 0:
                                reverse_token_list.append(1)
                            else:
                                reverse_token_list.append(0)

                        decrease_ratio = sum(reverse_token_list) / len(reverse_token_list) * 100
                        not_too_much_decrease = decrease_ratio < 80.0
                        #############################################################################################

                        ## 매수 여부 최종 판단
                        already_bought_nums = len(set(self.already_bought_stocks)) #len(set(self.kiwoom.bought_stock_df.s_code.tolist()))
                        try:
                            loss_cut_price = self.kiwoom.loss_cut_df.loc[self.kiwoom.loss_cut_df.s_code == jongmok_code_original].loss_cut_price.iloc[0]
                        except:
                            loss_cut_price = 0

                        over_loss_cut_price = current_price > loss_cut_price
                        
                        if already_bought_nums < self.kiwoom.stock_nums_to_buy:
                            if avoid_stable_spot & buy_condition1 & buy_condition2 & buy_condition3 & not_too_much_decrease & over_loss_cut_price:
                                available_money = self.kiwoom.trading_money / self.kiwoom.stock_nums_to_buy
                                unit_price = stock_master_df.iloc[stock_master_df.shape[0]-1].s_current_price
                                quantity = int(available_money / unit_price)

                                self.send_buy_order(code=jongmok_code_original, quantity=quantity)
                                self.already_bought_stocks.append(jongmok_code_original)

                                for obj in self.kiwoom.stock_waitlist:
                                    if obj[0] == jongmok_code_original:
                                        obj[2] = 'bought'
                        
                    ## 희망 종목 수 모두 채웠을 시 종목 정보 프로그램에서 삭제
                    if len(set(self.already_bought_stocks)) >= self.kiwoom.stock_nums_to_buy:
                        waitlist_remove_idx = []
                        master_df_remove_idx = []
                        for idx, obj in enumerate(self.kiwoom.stock_waitlist):
                            if obj[2] == 'not_bought':
                                ## 대기 목록 삭제 인덱스
                                waitlist_remove_idx.append(idx)

                                ## stock_master_df 정보 삭제 인덱스
                                remove_idx = [idx for idx, obj in enumerate(self.kiwoom.stock_master_df_list) if obj['s_code'] == jongmok_code_original][0]
                                master_df_remove_idx.append(remove_idx)

                                ## 실시간 가격 등록 해제
                                self.kiwoom.remove_real_reg(jongmok_code_original)
                                self.kiwoom.cancel_list.append(jongmok_code_original)
                        
                        ## waitlist 정보 삭제
                        waitlist_remove_idx = list(set(waitlist_remove_idx))
                        for idx in sorted(waitlist_remove_idx, reverse=True):
                            del self.kiwoom.stock_waitlist[idx]
                        
                        ## stock_master_df 내 종목 정보 삭제
                        master_df_remove_idx = list(set(master_df_remove_idx))
                        for idx in sorted(master_df_remove_idx, reverse=True):
                            del self.kiwoom.stock_master_df_list[idx]

                        ## waitlist 중복제거
                        self.kiwoom.stock_waitlist = list(set(tuple(item) for item in self.kiwoom.stock_waitlist))
                        self.kiwoom.stock_waitlist = [list(item) for item in self.kiwoom.stock_waitlist] 

                        ## UI 데이터 최신화
                        self.show_waitlist()

                except Exception as e:
                    pass

                ###### 매도 타점 잡기 ######
                if jongmok_code_original in self.kiwoom.stock_bought_list:             
                    try:
                        bought_stocks_list = self.kiwoom.bought_stock_df.loc[self.kiwoom.bought_stock_df.s_code == jongmok_code_original].index

                        for stock_idx_obj in bought_stocks_list:
                            stock_idx = int(stock_idx_obj)
                            bought_stock = self.kiwoom.bought_stock_df.iloc[stock_idx]
                            profit_rate = (current_price - bought_stock.s_bought_price) / bought_stock.s_bought_price * 100
                            quantity = bought_stock.s_bought_num

                            ## 익절 조건
                            try:
                                current_target_profit = self.kiwoom.loss_cut_df.loc[self.kiwoom.loss_cut_df.s_code == jongmok_code_original, "target_profit_rate"].iloc[0]
                            except:
                                current_target_profit = 2

                            sell_condition1 = profit_rate >= current_target_profit
                            sell_condition2 = (stock_master_df.iloc[stock_master_df.shape[0]-2].min2_smoothing_diff >= 0) & (stock_master_df.iloc[stock_master_df.shape[0]-1].min2_smoothing_diff < 0)
                            ## 손절 조건
                            try:
                                loss_cut_price = self.kiwoom.loss_cut_df.loc[self.kiwoom.loss_cut_df.s_code == jongmok_code_original].loss_cut_price.iloc[0]
                            except:
                                loss_cut_price = 0

                            # 손절조건 1 (응급 손절): 골든 크로스 이후 급격한 가격 하락으로 인한 지지 붕괴시
                            ## 1) 역배열 조건 (5이평 < 15이평 < 60이평)
                            current_min5 = stock_master_df.iloc[stock_master_df.shape[0]-1].min5_smoothing
                            current_min15 = stock_master_df.iloc[stock_master_df.shape[0]-1].min15_smoothing
                            current_min60 = stock_master_df.iloc[stock_master_df.shape[0]-1].min60_smoothing
                            reverse_arrangement = current_min5 < current_min15 < current_min60

                            ## 2) 5이평 기울기 감소 조건
                            current_min5_decrease = stock_master_df.iloc[stock_master_df.shape[0]-1].min5_smoothing_diff < 0

                            ## 3) 5이평, 15이평 이격도 조건 (이격도 >= 0.5%)
                            envelope_upper = current_min5 * 1.005
                            disparity_large = current_min15 > envelope_upper

                            sell_condition3 = reverse_arrangement & current_min5_decrease & disparity_large

                            # 손절조건 2: 정해 놓은 손절가 밑으로 가격 하락시
                            sell_condition4 =  current_price <= loss_cut_price

                            if (sell_condition1 & sell_condition2) | (sell_condition3 & (self.emergency_sell_mode == True)) | sell_condition4:
                                self.send_sell_order(code=jongmok_code_original, quantity=int(quantity), stock_idx=stock_idx)

                                for idx, obj in enumerate(self.kiwoom.stock_waitlist):
                                    if obj[0] == jongmok_code_original:
                                        bought_stock_remaining = [s for s in self.kiwoom.bought_stock_df.s_code if s == jongmok_code_original]
                                        if len(bought_stock_remaining) == 0:
                                            del self.kiwoom.stock_waitlist[idx]
                                
                                break

                    except Exception as e:
                        pass

                ## 지지선 붕괴 시 후보 종목 리스트에서 삭제
                try:
                    ## 분할 매수 시, 모든 분할 매수 종목 손절 후 종목 정보 삭제 처리
                    num_remaining_stocks = self.kiwoom.bought_stock_df.loc[self.kiwoom.bought_stock_df.s_code == jongmok_code_original].shape[0]
                    if num_remaining_stocks == 0:
                        try:
                            loss_cut_price = self.kiwoom.loss_cut_df.loc[self.kiwoom.loss_cut_df.s_code == jongmok_code_original].loss_cut_price.iloc[0]
                        except:
                            loss_cut_price = 0

                        if current_price <= loss_cut_price:
                            ## waitlist에서 종목 삭제
                            idx = [idx for idx, obj in enumerate(self.kiwoom.stock_waitlist) if obj[0] == jongmok_code_original][0]
                            del self.kiwoom.stock_waitlist[idx]
                            ## cancel list에 추가
                            self.kiwoom.cancel_list.append(jongmok_code_original)
                            ## loss_cut_df에서 종목 정보 삭제
                            self.kiwoom.loss_cut_df = self.kiwoom.loss_cut_df.loc[self.kiwoom.loss_cut_df.s_code != jongmok_code_original]
                            self.kiwoom.loss_cut_df.index = range(self.kiwoom.loss_cut_df.shape[0])
                            ## stock_master_df에서 종목정보 삭제
                            remove_idx = [idx for idx, wl in enumerate(self.kiwoom.stock_master_df_list)if wl['s_code'] == jongmok_code_original][0]
                            del self.kiwoom.stock_master_df_list[remove_idx]
                            ## UI 최신화
                            self.show_waitlist()
                except Exception:
                    pass
                
                ## 화면 UI 데이터 최신화
                try:
                    self.show_waitlist()
                    self.show_bought_status()
                except IndexError:
                    pass

    ############################################################
    def chejan_slot(self, gubun, item_cnt, fid):

        if gubun == "0":
            ## 체결 종목 데이터 수집
            buy_sell_gubun = self.kiwoom.dynamicCall('GetChejanData(907)')
            stock_code = self.kiwoom.dynamicCall('GetChejanData(9001)')
            stock_name = self.kiwoom.dynamicCall('GetChejanData(302)')
            stock_num = self.kiwoom.dynamicCall('GetChejanData(911)')
            stock_buy_price = self.kiwoom.dynamicCall('GetChejanData(910)')
            stock_current_price = self.kiwoom.dynamicCall('GetChejanData(10)')

            if stock_num != '':
                ## 수집 데이터 전처리
                stock_code_original = stock_code.strip()[1:]
                
                stock_num = re.findall(r'\d+', stock_num.strip())
                stock_num = ''.join(stock_num)
                stock_num = int(stock_num)

                stock_buy_price = re.findall(r'\d+', stock_buy_price.strip())
                stock_buy_price = ''.join(stock_buy_price)
                stock_buy_price = int(stock_buy_price)

                stock_current_price = re.findall(r'\d+', stock_current_price.strip())
                stock_current_price = ''.join(stock_current_price)
                stock_current_price = int(stock_current_price)

                ##### 매도 체결잔고 로직
                if buy_sell_gubun == "1":
                    try:
                        ## 프로그램이 아닌 사용자가 영웅문 앱을 통해 종목 매도 시 bought_stock_df에서 종목 관련 정보 삭제 (분할 매수 된 상황에서도 모든 종목 정보 삭제됨)
                        if self.is_program_sell == False:
                            self.kiwoom.bought_stock_df = self.kiwoom.bought_stock_df.loc[self.kiwoom.bought_stock_df.s_code != stock_code_original]
                            self.kiwoom.bought_stock_df.index = range(self.kiwoom.bought_stock_df.shape[0])
                            self.kiwoom.stock_bought_list = self.kiwoom.bought_stock_df.s_code.tolist()
                            self.kiwoom.num_of_bought_stocks = self.kiwoom.bought_stock_df.shape[0]
                            self.kiwoom.loss_cut_df = self.kiwoom.loss_cut_df.loc[self.kiwoom.loss_cut_df.s_code != stock_code_original]
                            
                            for idx, obj in enumerate(self.kiwoom.stock_waitlist):
                                if obj[0] == stock_code_original:
                                    del self.kiwoom.stock_waitlist[idx]

                        ## 분할 매수된 종목이 모두 매도 되었을 시에 종목 매도 후
                        if len([code for code in self.kiwoom.bought_stock_df.s_code if code == stock_code_original]) == 0:
                            # waiting list내 종목정보 제거
                            for idx, obj in enumerate(self.kiwoom.stock_waitlist):
                                if obj[0] == stock_code_original:
                                    del self.kiwoom.stock_waitlist[idx]
                            
                            # stock_master_df_list내 종목 정보 제거
                            for idx, obj in enumerate(self.kiwoom.stock_master_df_list):
                                if obj['s_code'] == stock_code_original:
                                    del self.kiwoom.stock_master_df_list[idx]
                            
                            # loss_cut_df내 종목정보 제거
                            self.kiwoom.loss_cut_df = self.kiwoom.loss_cut_df.loc[self.kiwoom.loss_cut_df.s_code != stock_code_original]
                            self.kiwoom.loss_cut_df.index = range(self.kiwoom.loss_cut_df.shape[0])
                    except Exception as e:
                        pass

                ##### 매수 체결잔고 로직
                elif buy_sell_gubun == "2":
                    ## 프로그램에 매수 종목 정보 추가                    
                    ## 사용자가 영웅문S를 통해 종목 매수시 프로그램에 반영 
                    if len([idx for idx, obj in enumerate(self.kiwoom.stock_master_df_list) if obj['s_code'] == stock_code_original]) == 0:
                        self.kiwoom.register_master_df(s_code=stock_code_original, s_name=stock_name.strip(), s_current_price=stock_current_price)
                        
                        ## 동일 종목에 대한 추가 매수가 아닌 경우 waitinglist 내 종목 정보 추가
                        self.kiwoom.stock_waitlist.append([stock_code_original, stock_name.strip(), 'bought'])

                        ## 동일 종목에 대한 추가 매수가 아닌 경우 loss_cut_df 내 종목 정보 추가 (default 손절가인 0원 손절가 부여)
                        self.kiwoom.loss_cut_df.loc[len(self.kiwoom.loss_cut_df)] = {'s_code': stock_code_original, 'loss_cut_price': 0, 'target_profit_rate': 2}

                    ## stock_waitlist 매수 여부 정보 최신화
                    for obj in self.kiwoom.stock_waitlist:
                        if obj[0] == stock_code_original:
                            obj[2] = 'bought'

                    ## stock_master_df_list 매수 여부 정보 최신화
                    for obj in self.kiwoom.stock_master_df_list:
                        if obj['s_code'] == stock_code_original:
                            obj['buy_status'] = 'bought'
                        
                    ## 매수 종목 정보 추가
                    self.kiwoom.update_bought_stock_df(s_code=stock_code_original, s_name=stock_name, s_bought_price=stock_buy_price, s_bought_num=stock_num)

                    ## 프로그램에 매수 종목 정보 추가
                    self.kiwoom.stock_bought_list = self.kiwoom.bought_stock_df.s_code.tolist()

                    ## 보유 종목 수 최신화
                    self.kiwoom.num_of_bought_stocks = self.kiwoom.bought_stock_df.shape[0]

        ## 프로그램 매도 여부 변수 초기화 & 화면 데이터 최신화
        self.is_program_sell = False
        self.show_bought_status()
        self.show_waitlist()

    #################### 이벤트 핸들러 슬롯 #####################
    def order_slot(self, scr_no, rqname, trcode, msg):
        print(msg)
        ## 매수 물량 부족시
        if re.search("855056", msg) is not None:
            print(msg)
            possible_quantity_group = re.search(r'(\d+)주 매수가능',msg)
            buy_stock_quantity = possible_quantity_group.group(1)
            self.send_buy_order(code=self.buy_request_stock_code, quantity=int(buy_stock_quantity))
        
        ## 매도 물량 부족시
        elif re.search("800033", msg) is not None:
            print(msg)
            possible_quantity_group = re.search(r'(\d+)주 매도가능',msg)
            sell_stock_quantity = possible_quantity_group.group(1)
            stock_idx = self.kiwoom.bought_stock_df.shape[0]+1 ## 매도 요청은 보내되, bought_stock_df에는 변경을 가하지 않기 위해 bought_stock_df의 최대 인덱스 보다 큰 인덱스 부여
            
            if int(sell_stock_quantity) > 0:
                self.send_sell_order(code=self.sell_request_stock_code, quantity=int(sell_stock_quantity), stock_idx = stock_idx)
        elif re.search("571445", msg) is not None:
            pass
    ############################################################

if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = BotWindow()
    myWindow.show()
    app.exec_()