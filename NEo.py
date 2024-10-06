from __future__ import print_function

import json
import math
import os.path
import pickle
from datetime import datetime
from io import StringIO
from time import sleep

import dataframe_image as dfi
import numpy as np
import pandas as pd
import requests
import telegram
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
# from ks_api_client import ks_api
from neo_api_client import NeoAPI

# Kotak API Access Details
kotak_live_user_id = "+918072965534"
kotak_live_password = "12Thunder@"
kotak_consumer_key = "QcDBVd2cpqjgorbyl0S19bgf2T4a"
kotak_access_token = "eyJ4NXQiOiJNbUprWWpVMlpETmpNelpqTURBM05UZ3pObUUxTm1NNU1qTXpNR1kyWm1OaFpHUTFNakE1TmciLCJraWQiOiJaalJqTUdRek9URmhPV1EwTm1WallXWTNZemRtWkdOa1pUUmpaVEUxTlRnMFkyWTBZVEUyTlRCaVlURTRNak5tWkRVeE5qZ3pPVGM0TWpGbFkyWXpOUV9SUzI1NiIsImFsZyI6IlJTMjU2In0.eyJzdWIiOiJjbGllbnQ2MTE2OSIsImF1dCI6IkFQUExJQ0FUSU9OIiwiYXVkIjoiUWNEQlZkMmNwcWpnb3JieWwwUzE5YmdmMlQ0YSIsIm5iZiI6MTcyODEzMjg2MiwiYXpwIjoiUWNEQlZkMmNwcWpnb3JieWwwUzE5YmdmMlQ0YSIsInNjb3BlIjoiZGVmYXVsdCIsImlzcyI6Imh0dHBzOlwvXC9uYXBpLmtvdGFrc2VjdXJpdGllcy5jb206NDQzXC9vYXV0aDJcL3Rva2VuIiwiZXhwIjoxNzI4MjE5MjYyLCJpYXQiOjE3MjgxMzI4NjIsImp0aSI6IjJkYzRhODA5LWFjOGItNGVkMi1iMDRhLThiMmVhNThhMjliYiJ9.CNphdlkGk3dNmgTPJ67odcZEYORkE0x0vOH0MNxPcdKSskAZmVpxzkGFvnQWa9mQVC0skye9n1l1TbAsakfZJQJywoR3Wequ3EDpTsdIkZuT3eXdedocKDjBOzMEBvd6T0kIOBhPQTm1UzbcKjkd1hAYp7tKaQmaJF1K9Wzof4m2Aqnh0NiN7BV4WdHN_pvbGIVN-vwwHNSb4PFhpDWhNXoeLug8dZwNcr0FdVi7Adghu7sT27-9p2GH2i0rSBUu3LCQQBanNSU8vyr2XhcrgLb0eZl4jBqU_4uHSl9N2a5qhTv7CpNwJh9YvXpqn_1w5E-tMS_sIIGjK9RXNqcd5w"
kotak_mpin = "170391"
kotak_consumer_secret = "pMASPTUiYbWe9skUdtoYKxyZG5Ea"

# Google sheet ID and scope, if modifying these scopes, delete the file token.pickle.
scopes = ['https://www.googleapis.com/auth/spreadsheets']
spreadsheet_id = '1F8gtC-Q97PBfczKxhDGZ7yA4jz2m3xVm8QfNur7YqSM'

# Telegram Connection
telegram_token = "7521663180:AAE6YPC0Wv4QUmYBSVvHV1NINu7AiJhAwjo"
telegram_channel = "ShoppersQuestTrendz"
chat_id_main_channel = "6201219849"
chat_id_error_channel = "6201219849"

# PD variables to display all the data and columns
pd.set_option('display.width', 1500)
pd.set_option('display.max_columns', 75)
pd.set_option('display.max_rows', 150)




def eod_auto_square_off_and_status_reporting(service):
    try:
        print("eod auto square off started.....")
        client = NeoAPI(consumer_key=kotak_consumer_key, consumer_secret=kotak_consumer_secret,
                        environment='prod')
        print("Session initiated")
        login_client_response = client.login(mobilenumber=kotak_live_user_id, password=kotak_live_password)
        print(login_client_response)
        login_client_response = client.session_2fa(OTP=kotak_mpin)
        print(login_client_response)
        # Generated session token successfully

        '''
        styles = [
            dict(selector="tr:hover",
                 props=[("background", "#f4f4f4")]),
            dict(selector="th", props=[("color", "#fff"),
                                       ("border", "1px solid #eee"),
                                       ("padding", "12px 35px"),
                                       ("border-collapse", "collapse"),
                                       ("background", "#583880"),
                                       # ("text-transform", "uppercase"),
                                       ("font-size", "20px")
                                       ]),
            dict(selector="td", props=[("color", "#101212"),
                                       ("border", "1px solid #eee"),
                                       ("padding", "12px 35px"),
                                       ("border-collapse", "collapse"),
                                       ("font-size", "18px")
                                       ]),
            dict(selector="table", props=[
                ("font-family", 'Arial'),
                ("margin", "25px auto"),
                ("border-collapse", "collapse"),
                ("border", "1px solid #eee"),
                ("border-bottom", "2px solid #00cccc"),
            ]),
            dict(selector="caption", props=[("caption-side", "bottom")])
        ]
        '''

        read_trading_inputs = read_gsheet_data(service, spreadsheet_id, 'ATP!A2:H2')
        trading_type = read_trading_inputs[0][2]
        read_trading_conditions = read_gsheet_data(service, spreadsheet_id, 'ATP!A7:H7')
        trading_symbol = read_trading_conditions[0][3]
        lot_size = int(read_trading_conditions[0][6])
        number_of_lots = int(read_trading_conditions[0][7])
        trade_quantity = lot_size * number_of_lots

        print("End of day Square Off at Cost")
        read_trade_details = read_gsheet_data(service, spreadsheet_id, 'ATP!A19:P21')
        print(read_trade_details)
        banknifty_strike_price = int(float(read_trade_details[1][3]))
        print(banknifty_strike_price)
        trade_leg_1_status = read_trade_details[1][14]
        trade_leg_2_status = read_trade_details[2][14]
        if trading_type == "R":
            ce_buy_order_id = int(float(read_trade_details[1][8]))
            print(ce_buy_order_id)
            pe_buy_order_id = int(float(read_trade_details[2][8]))
            print(pe_buy_order_id)
            if trade_leg_1_status != "YES":

                kotak_orders_data = client.order_report()
                for selected_order in kotak_orders_data['data']:
                    print(selected_order['nOrdNo'])
                    if selected_order['nOrdNo'] == str(ce_buy_order_id):
                        ce_trading_symbol_for_orders = str(selected_order['trdSym'])
                        print(ce_trading_symbol_for_orders)
                        ce_instrument_token = str(selected_order['tok'])
                        print(ce_instrument_token)

                modify_order_response = client.modify_order(order_id=str(ce_buy_order_id), exchange_segment="nse_fo",
                                                            product="MIS",
                                                            trading_symbol=ce_trading_symbol_for_orders,
                                                            instrument_token=ce_instrument_token,
                                                            price="0",
                                                            order_type="Market", amo="NO", trigger_price="0",
                                                            market_protection="0",
                                                            quantity=str(trade_quantity), validity="DAY",
                                                            disclosed_quantity="0",
                                                            transaction_type="B")

                print(modify_order_response)
                sleep(10)
                kotak_orders_data = client.order_report()
                print(kotak_orders_data)
                # lastOrderStatus = OrderDataFull['success'][-1]['status']
                # print(lastOrderStatus)
                ce_buy_order_closing_price = 0

                for current_order in kotak_orders_data['data']:
                    # print(currentOrder['orderId'])
                    if current_order['nOrdNo'] == str(ce_buy_order_id):
                        print(current_order['nOrdNo'])
                        ce_buy_order_status = current_order['ordSt']
                        if ce_buy_order_status == "complete":
                            ce_buy_order_closing_price = current_order['avgPrc']
                            print(ce_buy_order_closing_price)

                send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "YES", 'ATP!O20', "False", False)
                send_data_to_gsheet_for_one_cell(service, spreadsheet_id, ce_buy_order_closing_price, 'ATP!N20',
                                                 "False", False)
                print("square off at cost")
                message_text = 'New Passive Intraday Trade - ' + trading_symbol + ' - Leg 1 - EOD Forced Square Off'
                post_text_to_telegram(chat_id_main_channel, message_text)

            if trade_leg_2_status != "YES":

                kotak_orders_data = client.order_report()
                for selected_order in kotak_orders_data['data']:
                    print(selected_order['nOrdNo'])
                    if selected_order['nOrdNo'] == str(pe_buy_order_id):
                        pe_trading_symbol_for_orders = str(selected_order['trdSym'])
                        print(pe_trading_symbol_for_orders)
                        pe_instrument_token = str(selected_order['tok'])
                        print(pe_instrument_token)

                modify_order_response = client.modify_order(order_id=str(pe_buy_order_id), exchange_segment="nse_fo",
                                                            product="MIS",
                                                            trading_symbol=pe_trading_symbol_for_orders,
                                                            instrument_token=pe_instrument_token,
                                                            price="0",
                                                            order_type="Market", amo="NO", trigger_price="0",
                                                            market_protection="0",
                                                            quantity=str(trade_quantity), validity="DAY",
                                                            disclosed_quantity="0",
                                                            transaction_type="B")

                print(modify_order_response)
                sleep(10)
                kotak_orders_data = client.order_report()
                print(kotak_orders_data)
                # lastOrderStatus = OrderDataFull['success'][-1]['status']
                # print(lastOrderStatus)
                pe_buy_order_closing_price = 0
                for current_order in kotak_orders_data['data']:
                    # print(currentOrder['orderId'])
                    if current_order['nOrdNo'] == str(pe_buy_order_id):
                        print(current_order['nOrdNo'])
                        pe_buy_order_status = current_order['ordSt']
                        if pe_buy_order_status == "complete":
                            pe_buy_order_closing_price = current_order['avgPrc']
                            print(pe_buy_order_closing_price)
                send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "YES", 'ATP!O21', "False", False)
                send_data_to_gsheet_for_one_cell(service, spreadsheet_id, pe_buy_order_closing_price, 'ATP!N21',
                                                 "False",
                                                 False)
                print("square off at cost")
                message_text = 'New Passive Intraday Trade - ' + trading_symbol + ' - Leg 2 - EOD Forced Square Off'
                post_text_to_telegram(chat_id_main_channel, message_text)
        else:
            print(read_trade_details)
            ce_trade_symbol = int(float(read_trade_details[1][9]))
            print(ce_trade_symbol)
            instrument_tokens = '[{"instrument_token": "' + str(
                ce_trade_symbol) + '", "exchange_segment": "nse_fo"}]'
            print(instrument_tokens)
            ce_instrument_tokens_json = json.loads(instrument_tokens)
            print(ce_instrument_tokens_json)
            ce_ltp_price_quote = client.quotes(instrument_tokens=ce_instrument_tokens_json, quote_type="ltp")
            print(ce_ltp_price_quote)
            ce_ltp_data = int(float(ce_ltp_price_quote['message'][0]['ltp']))
            print(ce_ltp_data)
            ce_ltp = int(float(ce_ltp_data))

            pe_trade_symbol = int(float(read_trade_details[2][9]))
            print(pe_trade_symbol)
            instrument_tokens = '[{"instrument_token": "' + str(
                pe_trade_symbol) + '", "exchange_segment": "nse_fo"}]'
            print(instrument_tokens)
            pe_instrument_tokens_json = json.loads(instrument_tokens)
            print(pe_instrument_tokens_json)
            pe_ltp_price_quote = client.quotes(instrument_tokens=pe_instrument_tokens_json, quote_type="ltp")
            print(pe_ltp_price_quote)
            pe_ltp_data = int(float(pe_ltp_price_quote['message'][0]['ltp']))
            print(pe_ltp_data)
            pe_ltp = int(float(pe_ltp_data))

            if trade_leg_1_status != "YES":
                send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "YES", 'ATP!O20', "False", False)
                send_data_to_gsheet_for_one_cell(service, spreadsheet_id, ce_ltp, 'ATP!N20', "False", False)
                print("square off at cost")
                message_text = 'New Passive Intraday Trade - ' + trading_symbol + ' - Leg 1 - EOD Forced Square Off'
                post_text_to_telegram(chat_id_main_channel, message_text)

            if trade_leg_2_status != "YES":
                send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "YES", 'ATP!O21', "False", False)
                send_data_to_gsheet_for_one_cell(service, spreadsheet_id, pe_ltp, 'ATP!N21', "False", False)
                print("square off at cost")
                message_text = 'New Passive Intraday Trade - ' + trading_symbol + ' - Leg 2 - EOD Forced Square Off'
                post_text_to_telegram(chat_id_main_channel, message_text)

        read_updated_trade_details = read_gsheet_data(service, spreadsheet_id, 'ATP!A19:P21')
        ce_pl = float(read_updated_trade_details[1][15])
        pe_pl = float(read_updated_trade_details[2][15])
        net_pl = round(ce_pl + pe_pl, 2)
        print(net_pl)
        trade_details_dataframe = pd.DataFrame(read_updated_trade_details)
        print(trade_details_dataframe)
        header_row = 0
        trade_details_dataframe.columns = trade_details_dataframe.iloc[header_row]
        trade_details_dataframe = trade_details_dataframe.drop(header_row)
        trade_details_dataframe = trade_details_dataframe.reset_index(drop=True)
        trade_details_dataframe = trade_details_dataframe.drop(['LTP'], axis=1)
        trade_details_dataframe = trade_details_dataframe.drop(['SL'], axis=1)
        trade_details_dataframe = trade_details_dataframe.drop(['TSL'], axis=1)
        trade_details_dataframe = trade_details_dataframe.drop(['TSL Factor'], axis=1)
        trade_details_dataframe = trade_details_dataframe.drop(['SL OrderID'], axis=1)
        trade_details_dataframe = trade_details_dataframe.drop(['SymbolID'], axis=1)
        trade_details_dataframe = trade_details_dataframe.drop(['Low when entry'], axis=1)
        trade_details_dataframe = trade_details_dataframe.drop(['High when entry'], axis=1)
        trade_details_dataframe = trade_details_dataframe.drop(['Closed?'], axis=1)

        '''
        trade_details_dataframe_styled = trade_details_dataframe.style.set_precision(2).set_properties(
            **{'border': '2px solid green',
               'color': 'black',
               'text-align': 'center'}).hide_index().background_gradient().set_table_styles(
            styles)
        '''
        trade_details_dataframe_styled = trade_details_dataframe

        print(trade_details_dataframe_styled)
        file_name = "PassiveIntradayTradeBankNiftyStatus.png"
        dfi.export(trade_details_dataframe_styled, 'images/' + file_name)
        telegram_message_text = 'Passive Intraday Trade - ' + trading_symbol + ' - Final Status:\nNet Profile/Loss for the day = ' + str(
            net_pl)
        # Send to telegram
        post_image_to_telegram(chat_id_main_channel, telegram_message_text, 'images/' + file_name)

        send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "DONE", 'ATP!B16', "False", False)
        send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "DONE", 'ATP!C16', "False", False)
        log_last_updated_time_in_gsheets(service)

    except Exception as error:
        print("Error in fetch_and_submit_data {0}".format(error))
        message_text = 'Passive Trading - Forced square off and exit failed - ' + "Error, : {0}".format(error)
        post_text_to_telegram(chat_id_error_channel, message_text)


def monitor_trades_throughout_the_day(service):
    try:

        print("trade monitoring started.....")
        client = NeoAPI(consumer_key=kotak_consumer_key, consumer_secret=kotak_consumer_secret,
                        environment='prod')
        print("Session initiated")
        login_client_response = client.login(mobilenumber=kotak_live_user_id, password=kotak_live_password)
        print(login_client_response)
        login_client_response = client.session_2fa(OTP=kotak_mpin)
        print(login_client_response)
        # Generated session token successfully

        read_trading_inputs = read_gsheet_data(service, spreadsheet_id, 'ATP!A2:H2')
        trading_type = read_trading_inputs[0][2]
        read_trading_conditions = read_gsheet_data(service, spreadsheet_id, 'ATP!A7:H7')
        trading_symbol = read_trading_conditions[0][3]
        lot_size = int(read_trading_conditions[0][6])
        number_of_lots = int(read_trading_conditions[0][7])
        trade_quantity = lot_size * number_of_lots
        print(trade_quantity)
        print("Monitor Trades throughtout the day")
        read_trade_details = read_gsheet_data(service, spreadsheet_id, 'ATP!A19:P21')
        print(read_trade_details)
        additional_tsl_1_limit = 10
        additional_tsl_2_limit = 3.1

        if trading_type == "R":
            ce_buy_order_id = int(float(read_trade_details[1][8]))
            print(ce_buy_order_id)
            pe_buy_order_id = int(float(read_trade_details[2][8]))
            print(pe_buy_order_id)
            ce_instrument_id = int(float(read_trade_details[1][9]))
            print(ce_instrument_id)
            pe_instrument_id = int(float(read_trade_details[2][9]))
            print(pe_instrument_id)

            kotak_orders_data = client.order_report()
            print(kotak_orders_data)
            ce_buy_order_status = "N"
            pe_buy_order_status = "N"
            ce_buy_order_closing_price = 0
            pe_buy_order_closing_price = 0

            for current_order in kotak_orders_data['data']:
                #print(current_order)
                #print(current_order['nOrdNo'])
                #print(ce_buy_order_id)
                if current_order['nOrdNo'] == str(ce_buy_order_id):
                    print(current_order['nOrdNo'])
                    ce_buy_order_status = current_order['ordSt']
                    ce_trading_symbol_for_orders = str(current_order['trdSym'])
                    ce_instrument_token = str(current_order['tok'])
                    if ce_buy_order_status == "complete":
                        ce_buy_order_closing_price = current_order['avgPrc']
                        print(ce_buy_order_closing_price)
                if current_order['nOrdNo'] == str(pe_buy_order_id):
                    print(current_order['nOrdNo'])
                    pe_buy_order_status = current_order['ordSt']
                    pe_trading_symbol_for_orders = str(current_order['trdSym'])
                    pe_instrument_token = str(current_order['tok'])
                    if pe_buy_order_status == "complete":
                        pe_buy_order_closing_price = current_order['avgPrc']
                        print(pe_buy_order_closing_price)

            print(ce_buy_order_status)
            print(pe_buy_order_status)
            # bnfStrike = int(float(read_trade_details[1][3]))

            ce_leg_status = read_trade_details[1][14]
            ce_stoploss = read_trade_details[1][5]
            ce_tsl = read_trade_details[1][6]
            ce_tsl_factor = read_trade_details[1][7]
            ce_trading_symbol = int(float(read_trade_details[1][9]))
            ce_initial_low = int(float(read_trade_details[1][10]))

            # ### 2 lines below can be removed if not necessary - validate befpre removing
            # ### if this , along with pe line is removed, change line below print("Monitor Trades throughtout the day") , by changing range to P instead of R
            # ce_trading_symbol_for_orders= read_trade_details[1][17]
            # print(ce_trading_symbol_for_orders)

            instrument_tokens = '[{"instrument_token": "' + str(
                ce_trading_symbol) + '", "exchange_segment": "nse_fo"}]'
            print(instrument_tokens)
            ce_instrument_tokens_json = json.loads(instrument_tokens)
            print(ce_instrument_tokens_json)
            ce_ltp_price_quote = client.quotes(instrument_tokens=ce_instrument_tokens_json, quote_type="ltp")
            print(ce_ltp_price_quote)
            ce_ltp_data = int(float(ce_ltp_price_quote['message'][0]['ltp']))
            print(ce_ltp_data)
            ce_ltp = int(float(ce_ltp_data))
            send_data_to_gsheet_for_one_cell(service, spreadsheet_id, ce_ltp, 'ATP!M20', "False", False)
            ce_ohlc_price_quote = client.quotes(instrument_tokens=ce_instrument_tokens_json, quote_type='ohlc')
            ce_current_low = int(float(ce_ohlc_price_quote['message'][0]['ohlc']['low']))

            if ce_leg_status != "YES":
                if ce_buy_order_status == "rejected":
                    print("Major Emergency, running checks")
                    ce_emergency_flag = "Y"
                    ce_open_order_counter = 0

                    for current_order in kotak_orders_data['data']:
                        if str(current_order['tok']) == ce_instrument_id and (
                                current_order['ordSt'] == "complete") and \
                                current_order['trnsTp'] == "B":
                            new_ce_buy_order_id = current_order["nOrdNo"]
                            print(new_ce_buy_order_id)
                            send_data_to_gsheet_for_one_cell(service, spreadsheet_id, new_ce_buy_order_id, 'ATP!I20',
                                                             "False",
                                                             False)
                            print("New order details is updated")
                            ce_emergency_flag = "N"
                            ce_buy_order_status = current_order['ordSt']
                            ce_buy_order_closing_price = current_order['avgPrc']
                            print(ce_buy_order_closing_price)

                        if str(current_order['tok']) == ce_instrument_id and (
                                current_order['ordSt'] == "open") and \
                                current_order['trnsTp'] == "B":
                            ce_open_order_counter = ce_open_order_counter + 1

                    if ce_emergency_flag == "Y" and ce_open_order_counter == 0:
                        print("if there are no orders, place a new market order and sq off open position")

                        ce_buy_market_order_details = client.place_order(exchange_segment="nse_fo", product="MIS",
                                                                         price="0",
                                                                         order_type="Market",
                                                                         quantity=str(trade_quantity),
                                                                         validity="DAY",
                                                                         trading_symbol=str(
                                                                             ce_trading_symbol_for_orders),
                                                                         transaction_type="B", amo="NO",
                                                                         disclosed_quantity="0", market_protection="0",
                                                                         pf="N",
                                                                         trigger_price="0", tag=None)
                        print(ce_buy_market_order_details)

                    if ce_emergency_flag == "Y" and ce_open_order_counter > 0:
                        print("open unrecognized order exists, alert user")
                        message_text = 'Passive Intraday Trade - ' + trading_symbol + ' Major Emergency - open unrecognized order exists'
                        post_text_to_telegram(chat_id_error_channel, message_text)

                if ce_buy_order_status == "complete":
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "YES", 'ATP!O20', "False", False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, ce_buy_order_closing_price, 'ATP!N20',
                                                     "False", False)
                    print(ce_buy_order_closing_price)
                    print("SL Hit, enter SL price in Closed at sheet")
                    print("Updated closed? to closed and enter final exit price")
                    message_text = 'New Passive Intraday Trade - ' + trading_symbol + ' - Leg 1 - Exit Triggered'
                    post_text_to_telegram(chat_id_main_channel, message_text)

                elif ce_buy_order_status == "trigger pending":
                    # if (ce_ltp < ce_tsl) or (ce_initial_low > ce_tsl > ce_current_low):
                    if (ce_ltp < ce_tsl) or (ce_initial_low > ce_tsl > ce_current_low):
                        # if (ce_ltp < ce_tsl) or (ce_tsl < ce_initial_low and ce_current_low < ce_tsl)
                        print("TSL reached, change tsl to new tsl, SL to entry price ")
                        ce_new_stoploss = int(float(ce_stoploss - ce_tsl_factor))
                        print(ce_new_stoploss)
                        ce_new_stoploss_price = int(float(ce_new_stoploss + 30))
                        ce_new_tsl = int(float(ce_tsl - ce_tsl_factor))
                        print(ce_new_tsl)
                        # Adjust SL slightly
                        ce_sl_mod_value = ce_new_stoploss % 10
                        print(ce_sl_mod_value)
                        if ce_sl_mod_value > 7:
                            ce_new_stoploss = ce_new_stoploss + 2.1
                        else:
                            if ce_sl_mod_value == 0:
                                ce_new_stoploss = ce_new_stoploss + 1.1
                        print(ce_new_stoploss_price)
                        print(ce_new_stoploss)
                        if ce_new_stoploss_price > float(ce_new_stoploss * 1.19):
                            ce_new_stoploss_price = int(float(ce_new_stoploss * 1.19))
                        print(ce_new_stoploss_price)

                        print(ce_buy_order_id)

                        modify_order_response = client.modify_order(order_id=str(ce_buy_order_id),
                                                                    exchange_segment="nse_fo",
                                                                    product="MIS",
                                                                    trading_symbol=ce_trading_symbol_for_orders,
                                                                    instrument_token=ce_instrument_token,
                                                                    price=str(ce_new_stoploss_price),
                                                                    order_type="SL", amo="NO",
                                                                    trigger_price=str(ce_new_stoploss),
                                                                    market_protection="0",
                                                                    quantity=str(trade_quantity), validity="DAY",
                                                                    disclosed_quantity="0",
                                                                    transaction_type="B")

                        print(modify_order_response)

                        # trading_symbol = str(ce_trading_symbol_for_orders),
                        send_data_to_gsheet_for_one_cell(service, spreadsheet_id, ce_new_stoploss, 'ATP!F20', "False",
                                                         False)
                        send_data_to_gsheet_for_one_cell(service, spreadsheet_id, ce_new_tsl, 'ATP!G20', "False", False)
                        message_text = 'New Passive Intraday Trade - ' + trading_symbol + ' - Leg 1 - TSL Shifting'
                        post_text_to_telegram(chat_id_main_channel, message_text)

                    if additional_tsl_1_limit > ce_tsl > -additional_tsl_1_limit and (
                            (ce_ltp < additional_tsl_1_limit) or (
                            ce_initial_low > additional_tsl_1_limit > ce_current_low)):
                        # if(ce_tsl<10 and ce_tsl>-10 and ce_current_low<10):
                        print("TSL reached, change tsl to new tsl, SL to entry price ")
                        ce_new_stoploss = int(float(ce_stoploss - ce_tsl_factor))
                        print(ce_new_stoploss)
                        ce_new_stoploss_price = int(float(ce_new_stoploss + 30))
                        ce_new_tsl = int(float(ce_tsl - ce_tsl_factor))
                        print(ce_new_tsl)
                        # Adjust SL slightly
                        ce_sl_mod_value = ce_new_stoploss % 10
                        print(ce_sl_mod_value)
                        if ce_sl_mod_value > 7:
                            ce_new_stoploss = ce_new_stoploss + 2.1
                        else:
                            if ce_sl_mod_value == 0:
                                ce_new_stoploss = ce_new_stoploss + 1.1
                        print(ce_new_stoploss_price)
                        print(ce_new_stoploss)
                        if ce_new_stoploss_price > float(ce_new_stoploss * 1.19):
                            ce_new_stoploss_price = int(float(ce_new_stoploss * 1.19))
                        print(ce_new_stoploss_price)

                        modify_order_response = client.modify_order(order_id=str(ce_buy_order_id),
                                                                    exchange_segment="nse_fo",
                                                                    product="MIS",
                                                                    trading_symbol=ce_trading_symbol_for_orders,
                                                                    instrument_token=ce_instrument_token,
                                                                    price=str(ce_new_stoploss_price),
                                                                    order_type="SL", amo="NO",
                                                                    trigger_price=str(ce_new_stoploss),
                                                                    market_protection="0",
                                                                    quantity=str(trade_quantity), validity="DAY",
                                                                    disclosed_quantity="0",
                                                                    transaction_type="B")

                        print(modify_order_response)

                        send_data_to_gsheet_for_one_cell(service, spreadsheet_id, ce_new_stoploss, 'ATP!F20', "False",
                                                         False)
                        send_data_to_gsheet_for_one_cell(service, spreadsheet_id, ce_new_tsl, 'ATP!G20', "False", False)
                        message_text = 'New Passive Intraday Trade - ' + trading_symbol + ' - Leg 1 - TSL Shifting'
                        post_text_to_telegram(chat_id_main_channel, message_text)

                    if (ce_tsl < -additional_tsl_1_limit and ((ce_ltp < additional_tsl_2_limit) or (
                            ce_initial_low > additional_tsl_2_limit > ce_current_low))):
                        # if (ce_tsl < -10 and ce_current_low < 3.1):

                        modify_order_response = client.modify_order(order_id=str(ce_buy_order_id),
                                                                    exchange_segment="nse_fo",
                                                                    product="MIS",
                                                                    trading_symbol=ce_trading_symbol_for_orders,
                                                                    instrument_token=ce_instrument_token,
                                                                    price="0",
                                                                    order_type="Market", amo="NO",
                                                                    trigger_price="0",
                                                                    market_protection="0",
                                                                    quantity=str(trade_quantity), validity="DAY",
                                                                    disclosed_quantity="0",
                                                                    transaction_type="B")

                        print(modify_order_response)

                        message_text = 'New Passive Intraday Trade - ' + trading_symbol + ' - Leg 1 - Target Reached'
                        post_text_to_telegram(chat_id_main_channel, message_text)

            print("Leg1 Status Check complete")

            pe_leg_status = read_trade_details[2][14]
            pe_stoploss = read_trade_details[2][5]
            pe_tsl = read_trade_details[2][6]
            pe_tsl_factor = read_trade_details[2][7]
            pe_trading_symbol = int(float(read_trade_details[2][9]))
            pe_initial_low = int(float(read_trade_details[2][10]))

            # ### 1 lines below can be removed if not necessary - validate befpre removing
            # pe_trading_symbol_for_orders = read_trade_details[2][17]

            instrument_tokens = '[{"instrument_token": "' + str(
                pe_trading_symbol) + '", "exchange_segment": "nse_fo"}]'
            print(instrument_tokens)
            pe_instrument_tokens_json = json.loads(instrument_tokens)
            print(pe_instrument_tokens_json)
            pe_ltp_price_quote = client.quotes(instrument_tokens=pe_instrument_tokens_json, quote_type="ltp")
            print(pe_ltp_price_quote)
            pe_ltp_data = int(float(pe_ltp_price_quote['message'][0]['ltp']))
            print(pe_ltp_data)
            pe_ltp = int(float(pe_ltp_data))
            send_data_to_gsheet_for_one_cell(service, spreadsheet_id, pe_ltp, 'ATP!M21', "False", False)
            pe_ohlc_price_quote = client.quotes(instrument_tokens=pe_instrument_tokens_json, quote_type='ohlc')
            pe_current_low = int(float(pe_ohlc_price_quote['message'][0]['ohlc']['low']))

            if pe_leg_status != "YES":
                if pe_buy_order_status == "rejected":
                    print("Major Emergency, running checks")
                    pe_emergency_flag = "Y"
                    pe_open_order_counter = 0

                    for current_order in kotak_orders_data['data']:
                        if str(current_order['tok']) == pe_instrument_id and (
                                current_order['ordSt'] == "complete") and current_order['trnsTp'] == "B":
                            new_pe_buy_order_id = current_order["nOrdNo"]
                            print(new_pe_buy_order_id)
                            send_data_to_gsheet_for_one_cell(service, spreadsheet_id, new_pe_buy_order_id, 'ATP!I21',
                                                             "False",
                                                             False)
                            print("New order details is updated")
                            pe_emergency_flag = "N"
                            pe_buy_order_status = current_order['ordSt']
                            pe_buy_order_closing_price = current_order['avgPrc']
                            print(pe_buy_order_closing_price)

                        if str(current_order['tok']) == pe_instrument_id and (
                                current_order['ordSt'] == "open") and \
                                current_order['trnsTp'] == "B":
                            pe_open_order_counter = pe_open_order_counter + 1

                    if pe_emergency_flag == "Y" and pe_open_order_counter == 0:
                        print("if there are no orders, place a new market order and sq off open position")
                        pe_buy_market_order_details = client.place_order(exchange_segment="nse_fo", product="MIS",
                                                                         price="0",
                                                                         order_type="Market",
                                                                         quantity=str(trade_quantity),
                                                                         validity="DAY",
                                                                         trading_symbol=str(
                                                                             pe_trading_symbol_for_orders),
                                                                         transaction_type="B", amo="NO",
                                                                         disclosed_quantity="0", market_protection="0",
                                                                         pf="N",
                                                                         trigger_price="0", tag=None)
                        print(pe_buy_market_order_details)

                    if pe_emergency_flag == "Y" and pe_open_order_counter > 0:
                        print("open unrecognized order exists, alert user")
                        message_text = 'Passive Intraday Trade - ' + trading_symbol + ' Major Emergency - open unrecognized order exists'
                        post_text_to_telegram(chat_id_error_channel, message_text)

                if pe_buy_order_status == "complete":
                    # if (bnfLeg2LTP > bnfLeg2SL):
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "YES", 'ATP!O21', "False", False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, pe_buy_order_closing_price, 'ATP!N21',
                                                     "False",
                                                     False)
                    print(pe_buy_order_closing_price)
                    print("SL Hit, enter SL price in Closed at sheet")
                    print("Updated closed? to closed and enter final exit price")
                    message_text = 'New Passive Intraday Trade - ' + trading_symbol + ' - Leg 2 - Exit Triggered'
                    post_text_to_telegram(chat_id_main_channel, message_text)

                elif pe_buy_order_status == "trigger pending":
                    print("Check condition to shift SL?")
                    print(pe_tsl)
                    if (pe_ltp < pe_tsl) or (pe_initial_low > pe_tsl > pe_current_low):
                        print("TSL reached, change tsl to new tsl, SL to entry price ")
                        pe_new_stoploss = int(float(pe_stoploss - pe_tsl_factor))
                        pe_new_stoploss_price = int(float(pe_new_stoploss + 30))
                        pe_new_tsl = int(float(pe_tsl - pe_tsl_factor))
                        print(pe_new_tsl)
                        print("modified----------------")
                        # Adjust SL slightly
                        pe_sl_mod_value = pe_new_stoploss % 10
                        print(pe_sl_mod_value)
                        if pe_sl_mod_value > 7:
                            pe_new_stoploss = pe_new_stoploss + 2.1
                        else:
                            if pe_sl_mod_value == 0:
                                pe_new_stoploss = pe_new_stoploss + 1.1
                        print(pe_new_stoploss_price)
                        print(pe_new_stoploss)
                        if pe_new_stoploss_price > float(pe_new_stoploss * 1.19):
                            pe_new_stoploss_price = int(float(pe_new_stoploss * 1.19))
                        print(pe_new_stoploss_price)

                        modify_order_response = client.modify_order(order_id=str(pe_buy_order_id),
                                                                    exchange_segment="nse_fo",
                                                                    product="MIS",
                                                                    trading_symbol=pe_trading_symbol_for_orders,
                                                                    instrument_token=pe_instrument_token,
                                                                    price=str(pe_new_stoploss_price),
                                                                    order_type="SL", amo="NO",
                                                                    trigger_price=str(pe_new_stoploss),
                                                                    market_protection="0",
                                                                    quantity=str(trade_quantity), validity="DAY",
                                                                    disclosed_quantity="0",
                                                                    transaction_type="B")

                        print(modify_order_response)

                        # trading_symbol = str(pe_trading_symbol_for_orders),
                        send_data_to_gsheet_for_one_cell(service, spreadsheet_id, pe_new_stoploss, 'ATP!F21', "False",
                                                         False)
                        send_data_to_gsheet_for_one_cell(service, spreadsheet_id, pe_new_tsl, 'ATP!G21', "False",
                                                         False)
                        message_text = 'New Passive Intraday Trade - ' + trading_symbol + ' - Leg 2 - TSL Shifting'
                        post_text_to_telegram(chat_id_main_channel, message_text)

                    if (additional_tsl_1_limit > pe_tsl > -additional_tsl_1_limit and (
                            (pe_ltp < additional_tsl_1_limit) or (
                            pe_initial_low > additional_tsl_1_limit > pe_current_low))):
                        # if (pe_tsl<10 and pe_tsl>-10 and pe_current_low<10):
                        print("TSL reached, change tsl to new tsl, SL to entry price ")
                        pe_new_stoploss = int(float(pe_stoploss - pe_tsl_factor))
                        pe_new_stoploss_price = int(float(pe_new_stoploss + 30))
                        pe_new_tsl = int(float(pe_tsl - pe_tsl_factor))
                        print(pe_new_tsl)
                        print("modified----------------")
                        # Adjust SL slightly
                        pe_sl_mod_value = pe_new_stoploss % 10
                        print(pe_sl_mod_value)
                        if pe_sl_mod_value > 7:
                            pe_new_stoploss = pe_new_stoploss + 2.1
                        else:
                            if pe_sl_mod_value == 0:
                                pe_new_stoploss = pe_new_stoploss + 1.1
                        print(pe_new_stoploss_price)
                        print(pe_new_stoploss)
                        if pe_new_stoploss_price > float(pe_new_stoploss * 1.19):
                            pe_new_stoploss_price = int(float(pe_new_stoploss * 1.19))
                        print(pe_new_stoploss_price)

                        modify_order_response = client.modify_order(order_id=str(pe_buy_order_id),
                                                                    exchange_segment="nse_fo",
                                                                    product="MIS",
                                                                    trading_symbol=pe_trading_symbol_for_orders,
                                                                    instrument_token=pe_instrument_token,
                                                                    price=str(pe_new_stoploss_price),
                                                                    order_type="SL", amo="NO",
                                                                    trigger_price=str(pe_new_stoploss),
                                                                    market_protection="0",
                                                                    quantity=str(trade_quantity), validity="DAY",
                                                                    disclosed_quantity="0",
                                                                    transaction_type="B")

                        print(modify_order_response)

                        send_data_to_gsheet_for_one_cell(service, spreadsheet_id, pe_new_stoploss, 'ATP!F21', "False",
                                                         False)
                        send_data_to_gsheet_for_one_cell(service, spreadsheet_id, pe_new_tsl, 'ATP!G21', "False",
                                                         False)
                        message_text = 'New Passive Intraday Trade - ' + trading_symbol + ' - Leg 2 - TSL Shifting'
                        post_text_to_telegram(chat_id_main_channel, message_text)

                    if (pe_tsl < -additional_tsl_1_limit and ((pe_ltp < additional_tsl_2_limit) or (
                            pe_initial_low > additional_tsl_2_limit > pe_current_low))):
                        # if (pe_tsl < -10 and pe_current_low < 3.1):

                        modify_order_response = client.modify_order(order_id=str(pe_buy_order_id),
                                                                    exchange_segment="nse_fo",
                                                                    product="MIS",
                                                                    trading_symbol=pe_trading_symbol_for_orders,
                                                                    instrument_token=pe_instrument_token,
                                                                    price="0",
                                                                    order_type="Market", amo="NO",
                                                                    trigger_price="0",
                                                                    market_protection="0",
                                                                    quantity=str(trade_quantity), validity="DAY",
                                                                    disclosed_quantity="0",
                                                                    transaction_type="B")

                        print(modify_order_response)

                        message_text = 'New Passive Intraday Trade - ' + trading_symbol + ' - Leg 2 - Target Reached'
                        post_text_to_telegram(chat_id_main_channel, message_text)

            print("Leg2 Status Check complete")

        else:
            print("Paper trading")
            # banknifty_strike_price = int(float(read_trade_details[1][3]))
            ce_leg_status = read_trade_details[1][14]
            ce_stoploss = read_trade_details[1][5]
            ce_tsl = read_trade_details[1][6]
            ce_tsl_factor = read_trade_details[1][7]
            ce_trading_symbol = int(float(read_trade_details[1][9]))
            ce_initial_low = int(float(read_trade_details[1][10]))
            instrument_tokens = '[{"instrument_token": "' + str(
                ce_trading_symbol) + '", "exchange_segment": "nse_fo"}]'
            print(instrument_tokens)
            ce_instrument_tokens_json = json.loads(instrument_tokens)
            print(ce_instrument_tokens_json)
            ce_ltp_price_quote = client.quotes(instrument_tokens=ce_instrument_tokens_json, quote_type="ltp")
            print(ce_ltp_price_quote)
            ce_ltp_data = int(float(ce_ltp_price_quote['message'][0]['ltp']))
            print(ce_ltp_data)
            ce_ltp = int(float(ce_ltp_data))
            send_data_to_gsheet_for_one_cell(service, spreadsheet_id, ce_ltp, 'ATP!M20', "False", False)
            ce_ohlc_price_quote = client.quotes(instrument_tokens=ce_instrument_tokens_json, quote_type='ohlc')
            ce_current_low = int(float(ce_ohlc_price_quote['message'][0]['ohlc']['low']))
            ce_initial_high = int(float(read_trade_details[1][11]))
            ce_current_high = int(float(ce_ohlc_price_quote['message'][0]['ohlc']['high']))
            if ce_leg_status != "YES":
                if ce_ltp > ce_stoploss or (ce_current_high > ce_stoploss and ce_current_high > ce_initial_high):
                    # if (bnfLeg1LTP > bnfLeg1SL):
                    print(ce_ltp)
                    print(ce_stoploss)
                    print(ce_initial_high)
                    print(ce_current_high)

                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "YES", 'ATP!O20', "False", False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, ce_stoploss, 'ATP!N20', "False",
                                                     False)
                    print("SL Hit, enter SL price in Closed at sheet")
                    print("Updated closed? to closed and enter final exit price")
                    message_text = 'New Passive Intraday Trade - ' + trading_symbol + ' - Leg 1 - Exit Triggered'
                    post_text_to_telegram(chat_id_main_channel, message_text)

                else:
                    if (ce_ltp < ce_tsl) or (ce_initial_low > ce_tsl > ce_current_low):
                        print("TSL reached, change tsl to new tsl, SL to entry price ")
                        ce_new_stoploss = int(float(ce_stoploss - ce_tsl_factor))
                        print(ce_new_stoploss)
                        ce_new_tsl = int(float(ce_tsl - ce_tsl_factor))
                        print(ce_new_tsl)

                        # Adjust SL slightly
                        ce_sl_mod_value = ce_new_stoploss % 10
                        print(ce_sl_mod_value)
                        if ce_sl_mod_value > 7:
                            ce_new_stoploss = ce_new_stoploss + 2.1
                        else:
                            if ce_sl_mod_value == 0:
                                ce_new_stoploss = ce_new_stoploss + 1.1

                        send_data_to_gsheet_for_one_cell(service, spreadsheet_id, ce_new_stoploss, 'ATP!F20', "False",
                                                         False)
                        send_data_to_gsheet_for_one_cell(service, spreadsheet_id, ce_new_tsl, 'ATP!G20', "False", False)
                        message_text = 'New Passive Intraday Trade - ' + trading_symbol + ' - Leg 1 - TSL Shifting'
                        post_text_to_telegram(chat_id_main_channel, message_text)

                    if (additional_tsl_1_limit > ce_tsl > -additional_tsl_1_limit and (
                            (ce_ltp < additional_tsl_1_limit) or (
                            ce_initial_low > additional_tsl_1_limit > ce_current_low))):
                        # if(ce_tsl<10 and ce_tsl>-10 and ce_current_low<10):
                        print("TSL reached, change tsl to new tsl, SL to entry price ")
                        ce_new_stoploss = int(float(ce_stoploss - ce_tsl_factor))
                        print(ce_new_stoploss)
                        ce_new_tsl = int(float(ce_tsl - ce_tsl_factor))
                        print(ce_new_tsl)

                        # Adjust SL slightly
                        ce_sl_mod_value = ce_new_stoploss % 10
                        print(ce_sl_mod_value)
                        if ce_sl_mod_value > 7:
                            ce_new_stoploss = ce_new_stoploss + 2.1
                        else:
                            if ce_sl_mod_value == 0:
                                ce_new_stoploss = ce_new_stoploss + 1.1

                        send_data_to_gsheet_for_one_cell(service, spreadsheet_id, ce_new_stoploss, 'ATP!F20', "False",
                                                         False)
                        send_data_to_gsheet_for_one_cell(service, spreadsheet_id, ce_new_tsl, 'ATP!G20', "False", False)
                        message_text = 'New Passive Intraday Trade - ' + trading_symbol + ' - Leg 1 - TSL Shifting'
                        post_text_to_telegram(chat_id_main_channel, message_text)

                    if (ce_tsl < -additional_tsl_1_limit and ((ce_ltp < additional_tsl_2_limit) or (
                            ce_initial_low > additional_tsl_2_limit > ce_current_low))):
                        # if (ce_tsl < -10 and ce_current_low < 3.1):

                        # print(read_trade_details)
                        # ce_trade_symbol = int(float(read_trade_details[1][9]))
                        # print(ce_trade_symbol)
                        ce_ltp_price_quote = client.quotes(instrument_tokens=ce_instrument_tokens_json,
                                                           quote_type="ltp")
                        print(ce_ltp_price_quote)
                        ce_ltp_data = int(float(ce_ltp_price_quote['message'][0]['ltp']))
                        print(ce_ltp_data)
                        ce_ltp = int(float(ce_ltp_data))

                        send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "YES", 'ATP!O20', "False", False)
                        send_data_to_gsheet_for_one_cell(service, spreadsheet_id, ce_ltp, 'ATP!N20', "False", False)
                        message_text = 'New Passive Intraday Trade - ' + trading_symbol + ' - Leg 1 - Target hit'
                        post_text_to_telegram(chat_id_main_channel, message_text)

            print("Leg1 Status Check complete")

            pe_leg_status = read_trade_details[2][14]
            pe_stoploss = read_trade_details[2][5]
            pe_tsl = read_trade_details[2][6]
            pe_tsl_factor = read_trade_details[2][7]
            pe_trading_symbol = int(float(read_trade_details[2][9]))
            pe_initial_low = int(float(read_trade_details[2][10]))
            instrument_tokens = '[{"instrument_token": "' + str(
                pe_trading_symbol) + '", "exchange_segment": "nse_fo"}]'
            print(instrument_tokens)
            pe_instrument_tokens_json = json.loads(instrument_tokens)
            print(pe_instrument_tokens_json)
            pe_ltp_price_quote = client.quotes(instrument_tokens=pe_instrument_tokens_json, quote_type="ltp")
            print(pe_ltp_price_quote)
            pe_ltp_data = int(float(pe_ltp_price_quote['message'][0]['ltp']))
            print(pe_ltp_data)
            pe_ltp = int(float(pe_ltp_data))
            send_data_to_gsheet_for_one_cell(service, spreadsheet_id, pe_ltp, 'ATP!M21', "False", False)
            pe_ohlc_price_quote = client.quotes(instrument_tokens=pe_instrument_tokens_json, quote_type='ohlc')
            pe_current_low = int(float(pe_ohlc_price_quote['message'][0]['ohlc']['low']))
            pe_initial_high = int(float(read_trade_details[2][11]))
            pe_current_high = int(float(pe_ohlc_price_quote['message'][0]['ohlc']['low']))
            if pe_leg_status != "YES":
                if pe_ltp > pe_stoploss or (pe_current_high > pe_stoploss and pe_current_high > pe_initial_high):
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "YES", 'ATP!O21', "False", False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, pe_ltp, 'ATP!N21', "False", False)
                    print("SL Hit, enter SL price in Closed at sheet")
                    print("Updated closed? to closed and enter final exit price")
                    message_text = 'New Passive Intraday Trade - ' + trading_symbol + ' - Leg 2 - Exit Triggered'
                    post_text_to_telegram(chat_id_main_channel, message_text)
                else:
                    print("Check condition to shift SL?")
                    print(pe_tsl)
                    if (pe_ltp < pe_tsl) or (pe_initial_low > pe_tsl > pe_current_low):
                        print("TSL reached, change tsl to new tsl, SL to entry price ")
                        pe_new_stoploss = int(float(pe_stoploss - pe_tsl_factor))
                        print(pe_new_stoploss)
                        pe_new_tsl = int(float(pe_tsl - pe_tsl_factor))
                        print(pe_new_tsl)
                        print("modified----------------")

                        # Adjust SL slightly
                        pe_sl_mod_value = pe_new_stoploss % 10
                        print(pe_sl_mod_value)
                        if pe_sl_mod_value > 7:
                            pe_new_stoploss = pe_new_stoploss + 2.1
                        else:
                            if pe_sl_mod_value == 0:
                                pe_new_stoploss = pe_new_stoploss + 1.1

                        send_data_to_gsheet_for_one_cell(service, spreadsheet_id, pe_new_stoploss, 'ATP!F21', "False",
                                                         False)
                        send_data_to_gsheet_for_one_cell(service, spreadsheet_id, pe_new_tsl, 'ATP!G21', "False", False)
                        message_text = 'New Passive Intraday Trade - ' + trading_symbol + ' - Leg 2 - TSL Shifting'
                        post_text_to_telegram(chat_id_main_channel, message_text)

                    if (additional_tsl_1_limit > pe_tsl > -additional_tsl_1_limit and (
                            (pe_ltp < additional_tsl_1_limit) or (
                            pe_initial_low > additional_tsl_1_limit > pe_current_low))):
                        # if (pe_tsl<10 and pe_tsl>-10 and pe_current_low<10):
                        print("TSL reached, change tsl to new tsl, SL to entry price ")
                        pe_new_stoploss = int(float(pe_stoploss - pe_tsl_factor))
                        print(pe_new_stoploss)
                        pe_new_tsl = int(float(pe_tsl - pe_tsl_factor))
                        print(pe_new_tsl)
                        print("modified----------------")

                        # Adjust SL slightly
                        pe_sl_mod_value = pe_new_stoploss % 10
                        print(pe_sl_mod_value)
                        if pe_sl_mod_value > 7:
                            pe_new_stoploss = pe_new_stoploss + 2.1
                        else:
                            if pe_sl_mod_value == 0:
                                pe_new_stoploss = pe_new_stoploss + 1.1

                        send_data_to_gsheet_for_one_cell(service, spreadsheet_id, pe_new_stoploss, 'ATP!F21', "False",
                                                         False)
                        send_data_to_gsheet_for_one_cell(service, spreadsheet_id, pe_new_tsl, 'ATP!G21', "False", False)
                        message_text = 'New Passive Intraday Trade - ' + trading_symbol + ' - Leg 2 - TSL Shifting'
                        post_text_to_telegram(chat_id_main_channel, message_text)

                    if (pe_tsl < -additional_tsl_1_limit and ((pe_ltp < additional_tsl_2_limit) or (
                            pe_initial_low > additional_tsl_2_limit > pe_current_low))):
                        # if (pe_tsl < -10 and pe_current_low < 3.1):
                        # pe_trade_symbol = int(float(read_trade_details[2][9]))
                        # print(pe_trade_symbol)
                        pe_ltp_price_quote = client.quotes(instrument_tokens=pe_instrument_tokens_json,
                                                           quote_type="ltp")
                        print(pe_ltp_price_quote)
                        pe_ltp_data = int(float(pe_ltp_price_quote['message'][0]['ltp']))
                        print(pe_ltp_data)
                        pe_ltp = int(float(pe_ltp_data))

                        send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "YES", 'ATP!O21', "False", False)
                        send_data_to_gsheet_for_one_cell(service, spreadsheet_id, pe_ltp, 'ATP!N21', "False", False)
                        message_text = 'New Passive Intraday Trade - ' + trading_symbol + ' - Leg 2 - Target Hit'
                        post_text_to_telegram(chat_id_main_channel, message_text)

        log_last_updated_time_in_gsheets(service)

    except Exception as error:
        print("Error in fetch_and_submit_data {0}".format(error))
        message_text = 'Passive Trading - Error while monitoring trade - ' + "Error, : {0}".format(error)
        post_text_to_telegram(chat_id_error_channel, message_text)


def find_my_trades(service):
    try:
        print("find_trades started.....")
        client = NeoAPI(consumer_key=kotak_consumer_key, consumer_secret=kotak_consumer_secret,
                        environment='prod')
        print("Session initiated")
        login_client_response = client.login(mobilenumber=kotak_live_user_id, password=kotak_live_password)
        print(login_client_response)
        login_client_response = client.session_2fa(OTP=kotak_mpin)
        print(login_client_response)
        # Generated session token successfully

        '''
        styles = [
            dict(selector="tr:hover",
                 props=[("background", "#f4f4f4")]),
            dict(selector="th", props=[("color", "#fff"),
                                       ("border", "1px solid #eee"),
                                       ("padding", "12px 35px"),
                                       ("border-collapse", "collapse"),
                                       ("background", "#583880"),
                                       # ("text-transform", "uppercase"),
                                       ("font-size", "20px")
                                       ]),
            dict(selector="td", props=[("color", "#101212"),
                                       ("border", "1px solid #eee"),
                                       ("padding", "12px 35px"),
                                       ("border-collapse", "collapse"),
                                       ("font-size", "18px")
                                       ]),
            dict(selector="table", props=[
                ("font-family", 'Arial'),
                ("margin", "25px auto"),
                ("border-collapse", "collapse"),
                ("border", "1px solid #eee"),
                ("border-bottom", "2px solid #00cccc"),
            ]),
            dict(selector="caption", props=[("caption-side", "bottom")])
        ]
        '''

        read_last_update = read_gsheet_data(service, spreadsheet_id, 'ATP!A16:G16')
        read_flag1 = read_last_update[0][0]
        read_trading_inputs = read_gsheet_data(service, spreadsheet_id, 'ATP!A2:H2')
        trading_type = read_trading_inputs[0][2]
        read_trading_conditions = read_gsheet_data(service, spreadsheet_id, 'ATP!A7:H7')
        vix_lower_limit = float(read_trading_conditions[0][0])
        vix_upper_limit = float(read_trading_conditions[0][1])
        previous_day_eod_vix = float(read_trading_conditions[0][2])
        spot_symbol_name = read_trading_conditions[0][3]
        print(spot_symbol_name)
        spot_symbol_id = int(read_trading_conditions[0][4])
        print(spot_symbol_id)
        strike_price_difference = int(read_trading_conditions[0][5])
        print(strike_price_difference)
        lot_size = int(read_trading_conditions[0][6])
        number_of_lots = int(read_trading_conditions[0][7])
        trade_quantity = lot_size * number_of_lots
        spot_ltp = 0
        # final_dataframe = pd.DataFrame()
        print("------------------------------------------")

        # if ((read_flag1 == 0 or read_flag2 == 0 or read_flag3 == 0) and hour_now > 8 and hour_now < 16):
        if read_flag1 == read_flag1:
            # Get Spot Price
            print("Pull data..")
            instrument_tokens = '[{"instrument_token": "' + str(spot_symbol_id) + '", "exchange_segment": "nse_cm"}]'
            print(instrument_tokens)
            instrument_tokens_json = json.loads(instrument_tokens)
            print(instrument_tokens_json)
            spot_price_quote = client.quotes(instrument_tokens=instrument_tokens_json, quote_type="ltp")
            print(spot_price_quote)
            spot_ltp = spot_price_quote['message'][0]['ltp']
            print(spot_ltp)

            # Find all option strike price symbols
            trading_symbol_list = get_list_of_trading_symbols(spot_symbol_name)

            # Get trading symbol details and pull the symbols alone and append it and get the quote
            ce_trading_symbols = trading_symbol_list[0]
            pe_trading_symbols = trading_symbol_list[1]
            ce_trading_symbol_list = ce_trading_symbols['pSymbol'].tolist()
            pe_trading_symbol_list = pe_trading_symbols['pSymbol'].tolist()

            ce_trading_symbol_list_items = ce_trading_symbol_list[0:1]
            ce_jsonfile = []
            for index, value in enumerate(ce_trading_symbol_list_items):
                ce_jsonfile.append({
                    "instrument_token": ce_trading_symbol_list_items[index],
                    "exchange_segment": "nse_fo"
                })
            ce_instrument_tokens_list = json.dumps(ce_jsonfile)
            print(ce_instrument_tokens_list)
            print(json.dumps(ce_jsonfile, indent=4))
            ce_instrument_tokens_json = json.loads(ce_instrument_tokens_list)
            print(ce_instrument_tokens_json)
            ce_option_sample_quote = client.quotes(instrument_tokens=ce_instrument_tokens_json)
            print(ce_option_sample_quote)

            pe_trading_symbol_list_items = pe_trading_symbol_list[0:1]
            pe_jsonfile = []
            for index, value in enumerate(pe_trading_symbol_list_items):
                pe_jsonfile.append({
                    "instrument_token": pe_trading_symbol_list_items[index],
                    "exchange_segment": "nse_fo"
                })
            pe_instrument_tokens_list = json.dumps(pe_jsonfile)
            print(pe_instrument_tokens_list)
            print(json.dumps(pe_jsonfile, indent=4))
            pe_instrument_tokens_json = json.loads(pe_instrument_tokens_list)
            print(pe_instrument_tokens_json)
            pe_option_sample_quote = client.quotes(instrument_tokens=pe_instrument_tokens_json)
            print(pe_option_sample_quote)
            print("Data pulling completed")

        # if (hour_now == entry_hour and (mins_for_df > entry_minute_start and mins_for_df < entry_minute_end) and read_flag1 == "DO" and read_flag5 == "DONE" and read_flag6 == "DONE"):
        if read_flag1 == read_flag1:
            print("Find the trade at 9.30 AM everyday")
            fetch_requested_data = read_gsheet_data(service, "14BT6yWg0tJahYr2BhDxdhMU44nRaspYkNaWJ5K9s58Q",
                                                    'Vix!A3:B3')
            print(fetch_requested_data)
            vix_ltp = fetch_requested_data[0][0]
            print(vix_ltp)
            vix_updated_datetime_from_sheet = fetch_requested_data[0][1]
            vix_updated_datetime_from_sheet_datetime_object = datetime.strptime(vix_updated_datetime_from_sheet,
                                                                                "%m/%d/%Y, %H:%M:%S")
            print(vix_updated_datetime_from_sheet_datetime_object)
            print(datetime.now())
            difference_in_time = datetime.now() - vix_updated_datetime_from_sheet_datetime_object
            print(difference_in_time)
            duration_in_seconds = difference_in_time.total_seconds()
            print(duration_in_seconds)
            duration_in_minutes = divmod(duration_in_seconds, 60)[0]
            print("Time difference in minutes is below")
            print(duration_in_minutes)

            print(previous_day_eod_vix)
            vix_change = ((float(vix_ltp) - previous_day_eod_vix) / previous_day_eod_vix) * 100
            print(vix_change)

            # vix work around above

            print("vix_change: " + str(vix_change))
            print("vix_lower_limit: " + str(vix_lower_limit))
            print("vix_upper_limit: " + str(vix_upper_limit))
            if vix_upper_limit > vix_change > vix_lower_limit and duration_in_minutes < 5:
                # select Strike price to trade
                print(spot_ltp)
                spot_ltp_rounded = int(float(spot_ltp))
                print(spot_ltp_rounded)
                spot_ltp_floored = math.floor(spot_ltp_rounded / strike_price_difference)
                print(spot_ltp_floored)
                spot_ltp_floored = spot_ltp_floored * strike_price_difference
                print(spot_ltp_floored)
                price_difference = abs(spot_ltp_rounded - spot_ltp_floored)
                print(price_difference)



                selected_strike1=spot_ltp_floored
                selected_strike2 = spot_ltp_floored + strike_price_difference

                selected_row = ce_trading_symbols.loc[
                    ce_trading_symbols['dStrikePrice;'] == selected_strike1*100]
                print(selected_row)
                trading_symbol = int(float(selected_row['pSymbol'].iloc[0]))
                instrument_tokens = '[{"instrument_token": "' + str(
                    trading_symbol) + '", "exchange_segment": "nse_fo"}]'
                instrument_tokens_json = json.loads(instrument_tokens)
                print(instrument_tokens_json)
                spot_price_quote = client.quotes(instrument_tokens=instrument_tokens_json, quote_type="ltp")
                selected_strike1_ce_spot_ltp = spot_price_quote['message'][0]['ltp']
                print(selected_strike1_ce_spot_ltp)

                selected_row = pe_trading_symbols.loc[
                    pe_trading_symbols['dStrikePrice;'] == selected_strike1 * 100]
                print(selected_row)
                trading_symbol = int(float(selected_row['pSymbol'].iloc[0]))
                instrument_tokens = '[{"instrument_token": "' + str(
                    trading_symbol) + '", "exchange_segment": "nse_fo"}]'
                instrument_tokens_json = json.loads(instrument_tokens)
                print(instrument_tokens_json)
                spot_price_quote = client.quotes(instrument_tokens=instrument_tokens_json, quote_type="ltp")
                selected_strike1_pe_spot_ltp = spot_price_quote['message'][0]['ltp']
                print(selected_strike1_pe_spot_ltp)
                strike1_price_difference= abs(int(float(selected_strike1_pe_spot_ltp))-int(float(selected_strike1_ce_spot_ltp)))
                print(strike1_price_difference)

                selected_row = ce_trading_symbols.loc[
                    ce_trading_symbols['dStrikePrice;'] == selected_strike2 * 100]
                print(selected_row)
                trading_symbol = int(float(selected_row['pSymbol'].iloc[0]))
                instrument_tokens = '[{"instrument_token": "' + str(
                    trading_symbol) + '", "exchange_segment": "nse_fo"}]'
                instrument_tokens_json = json.loads(instrument_tokens)
                print(instrument_tokens_json)
                spot_price_quote = client.quotes(instrument_tokens=instrument_tokens_json, quote_type="ltp")
                selected_strike2_ce_spot_ltp = spot_price_quote['message'][0]['ltp']
                print(selected_strike2_ce_spot_ltp)

                selected_row = pe_trading_symbols.loc[
                    pe_trading_symbols['dStrikePrice;'] == selected_strike2 * 100]
                print(selected_row)
                trading_symbol = int(float(selected_row['pSymbol'].iloc[0]))
                instrument_tokens = '[{"instrument_token": "' + str(
                    trading_symbol) + '", "exchange_segment": "nse_fo"}]'
                instrument_tokens_json = json.loads(instrument_tokens)
                print(instrument_tokens_json)
                spot_price_quote = client.quotes(instrument_tokens=instrument_tokens_json, quote_type="ltp")
                selected_strike2_pe_spot_ltp = spot_price_quote['message'][0]['ltp']
                print(selected_strike2_pe_spot_ltp)
                strike2_price_difference = abs(int(float(selected_strike2_pe_spot_ltp)) - int(float(selected_strike2_ce_spot_ltp)))
                print(strike2_price_difference)
                if strike1_price_difference < strike2_price_difference:
                    selected_strike = selected_strike1
                else:
                    selected_strike = selected_strike2
                print(selected_strike)
                extended_selected_strike = selected_strike * 100
                print(extended_selected_strike)

                if trading_type == "R":
                    selected_strike_row = ce_trading_symbols.loc[
                        ce_trading_symbols['dStrikePrice;'] == extended_selected_strike]
                    print(selected_strike_row)
                    ce_trading_symbol = int(float(selected_strike_row['pSymbol'].iloc[0]))
                    print(ce_trading_symbol)
                    ce_trading_symbol_for_orders = selected_strike_row['pTrdSymbol'].iloc[0]
                    print(ce_trading_symbol_for_orders)
                    instrument_tokens = '[{"instrument_token": "' + str(
                        ce_trading_symbol) + '", "exchange_segment": "nse_fo"}]'
                    print(instrument_tokens)
                    ce_instrument_tokens_json = json.loads(instrument_tokens)
                    print(ce_instrument_tokens_json)

                    # In the below comment, doing indicates, in progress , this will avoid duplciate trades if network is slow
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "DOING", 'ATP!A16', "False", False)

                    # order_type="MIS" not working at times for options
                    try:
                        ce_sell_order_details = client.place_order(exchange_segment="nse_fo", product="MIS", price="0",
                                                                   order_type="Market", quantity=str(trade_quantity),
                                                                   validity="DAY",
                                                                   trading_symbol=str(ce_trading_symbol_for_orders),
                                                                   transaction_type="S", amo="NO",
                                                                   disclosed_quantity="0", market_protection="0",
                                                                   pf="N",
                                                                   trigger_price="0", tag=None)
                        print(ce_sell_order_details)

                        ce_sell_order_id = ce_sell_order_details['nOrdNo']
                        print(ce_sell_order_id)
                    except Exception as error:
                        print("Error, : {0}".format(error))
                        return
                    ce_sell_order_status = "Check"
                    ce_selling_price = 0
                    # kotak_orders_data = client.order_report()
                    # print(kotak_orders_data)

                    # wait here for the result to be available before continuing
                    while ce_sell_order_status != "complete":  # TRAD
                        try:
                            kotak_orders_data = client.order_report()
                            print(kotak_orders_data)
                            for selected_order in kotak_orders_data['data']:
                                print(selected_order['nOrdNo'])
                                if selected_order['nOrdNo'] == ce_sell_order_id:
                                    print(selected_order['ordSt'])
                                    ce_sell_order_status = selected_order['ordSt']
                                    print(ce_sell_order_status)
                                    # ce_sell_order_id = kotak_orders_data['success'][-1]['orderId']
                                    ce_selling_price = float(selected_order['avgPrc'])
                                    print(ce_selling_price)
                            # ce_sell_order_status="TRAD"

                        except Exception as error:
                            print("Error, : {0}".format(error))
                            pass
                        pass

                    print(ce_selling_price)
                    ce_tsl_component = int(float(ce_selling_price / 4))  # change to 4
                    ce_stoploss = ce_selling_price + ce_tsl_component + 1.1
                    print(ce_stoploss)
                    ce_stoploss_price = ce_stoploss + 30
                    print(ce_stoploss_price)
                    ce_tsl = ce_selling_price - ce_tsl_component
                    print(ce_tsl)
                    # Adjust SL slightly
                    ce_sl_mod_value = int(float(ce_stoploss)) % 10
                    print(ce_sl_mod_value)
                    if ce_sl_mod_value == 8:
                        ce_stoploss = ce_stoploss + 2.1
                    if ce_sl_mod_value == 9:
                        ce_stoploss = ce_stoploss + 1.1
                    if ce_sl_mod_value == 0:
                        ce_stoploss = ce_stoploss + 1.1
                    print(ce_stoploss)
                    print(ce_stoploss_price)
                    if ce_stoploss_price > float(ce_stoploss * 1.19):
                        ce_stoploss_price = int(float(ce_stoploss * 1.19))
                    print(ce_stoploss_price)

                    ce_buy_order_details = client.place_order(exchange_segment="nse_fo", product="MIS",
                                                              price=str(ce_stoploss_price),
                                                              order_type="SL", quantity=str(trade_quantity),
                                                              validity="DAY",
                                                              trading_symbol=str(ce_trading_symbol_for_orders),
                                                              transaction_type="B", amo="NO",
                                                              disclosed_quantity="0", market_protection="0", pf="N",
                                                              trigger_price=str(ce_stoploss), tag=None)
                    print(ce_buy_order_details)

                    ce_buy_order_id = ce_buy_order_details['nOrdNo']
                    print(ce_buy_order_id)

                    print("CE Orders set-up completed, move to PE-------------------------------------------")

                    selected_strike_row = pe_trading_symbols.loc[
                        pe_trading_symbols['dStrikePrice;'] == extended_selected_strike]
                    print(selected_strike_row)
                    pe_trading_symbol = int(float(selected_strike_row['pSymbol'].iloc[0]))
                    print(pe_trading_symbol)
                    pe_trading_symbol_for_orders = selected_strike_row['pTrdSymbol'].iloc[0]
                    print(pe_trading_symbol_for_orders)
                    instrument_tokens = '[{"instrument_token": "' + str(
                        pe_trading_symbol) + '", "exchange_segment": "nse_fo"}]'
                    print(instrument_tokens)
                    pe_instrument_tokens_json = json.loads(instrument_tokens)
                    print(pe_instrument_tokens_json)

                    try:
                        pe_sell_order_details = client.place_order(exchange_segment="nse_fo", product="MIS", price="0",
                                                                   order_type="Market", quantity=str(trade_quantity),
                                                                   validity="DAY",
                                                                   trading_symbol=str(pe_trading_symbol_for_orders),
                                                                   transaction_type="S", amo="NO",
                                                                   disclosed_quantity="0", market_protection="0",
                                                                   pf="N",
                                                                   trigger_price="0", tag=None)
                        print(pe_sell_order_details)
                        pe_sell_order_id = pe_sell_order_details['nOrdNo']
                        print(pe_sell_order_id)
                    except Exception as error:
                        print("Error, : {0}".format(error))
                        return
                    pe_sell_order_status = "Check"
                    pe_selling_price = 0

                    # wait here for the result to be available before continuing
                    while pe_sell_order_status != "complete":  # TRAD
                        try:
                            kotak_orders_data = client.order_report()
                            for selected_order in kotak_orders_data['data']:
                                print(selected_order['nOrdNo'])
                                if selected_order['nOrdNo'] == pe_sell_order_id:
                                    print(selected_order['ordSt'])
                                    pe_sell_order_status = selected_order['ordSt']
                                    print(pe_sell_order_status)
                                    # ce_sell_order_id = kotak_orders_data['success'][-1]['orderId']
                                    pe_selling_price = float(selected_order['avgPrc'])
                                    print(pe_selling_price)
                                    # ce_sell_order_status="TRAD"
                        except Exception as error:
                            print("Error, : {0}".format(error))
                            pass
                        pass
                    print(pe_selling_price)
                    pe_tsl_component = int(float(pe_selling_price / 4))  # change to 4
                    pe_stoploss = pe_selling_price + pe_tsl_component + 1.1
                    print(pe_stoploss)
                    pe_stoploss_price = pe_stoploss + 30
                    print(pe_stoploss_price)
                    pe_tsl = pe_selling_price - pe_tsl_component
                    print(pe_tsl)

                    # Adjust SL slightly
                    pe_sl_mod_value = int(float(pe_stoploss)) % 10
                    print(pe_sl_mod_value)
                    if pe_sl_mod_value == 8:
                        pe_stoploss = pe_stoploss + 2.1
                    if pe_sl_mod_value == 9:
                        pe_stoploss = pe_stoploss + 1.1
                    if pe_sl_mod_value == 0:
                        pe_stoploss = pe_stoploss + 1.1
                    print(pe_stoploss)
                    print(pe_stoploss_price)
                    if pe_stoploss_price > float(pe_stoploss * 1.19):
                        pe_stoploss_price = int(float(pe_stoploss * 1.19))
                    print(pe_stoploss_price)

                    pe_buy_order_details = client.place_order(exchange_segment="nse_fo", product="MIS",
                                                              price=str(pe_stoploss_price),
                                                              order_type="SL", quantity=str(trade_quantity),
                                                              validity="DAY",
                                                              trading_symbol=str(pe_trading_symbol_for_orders),
                                                              transaction_type="B", amo="NO",
                                                              disclosed_quantity="0", market_protection="0", pf="N",
                                                              trigger_price=str(pe_stoploss), tag=None)
                    print(pe_buy_order_details)

                    pe_buy_order_id = pe_buy_order_details['nOrdNo']
                    print(pe_buy_order_id)

                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, selected_strike, 'ATP!D20', "False",
                                                     False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, selected_strike, 'ATP!D21', "False",
                                                     False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, ce_selling_price, 'ATP!E20', "False",
                                                     False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, pe_selling_price, 'ATP!E21', "False",
                                                     False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, ce_stoploss, 'ATP!F20', "False", False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, pe_stoploss, 'ATP!F21', "False", False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, ce_tsl, 'ATP!G20', "False", False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, pe_tsl, 'ATP!G21', "False", False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, ce_buy_order_id, 'ATP!I20', "False",
                                                     False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, pe_buy_order_id, 'ATP!I21', "False",
                                                     False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, ce_trading_symbol, 'ATP!J20', "False",
                                                     False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, pe_trading_symbol, 'ATP!J21', "False",
                                                     False)
                    '''
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, ce_trading_symbol_for_orders, 'ATP!R20', "False",
                                                     False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, pe_trading_symbol_for_orders, 'ATP!R21', "False",
                                                     False)
                    '''
                    # ### ignore , not needed

                    ce_ohlc_price_quote = client.quotes(instrument_tokens=ce_instrument_tokens_json, quote_type="ohlc")
                    print(ce_ohlc_price_quote)
                    ce_current_low = int(float(ce_ohlc_price_quote['message'][0]['ohlc']['low']))
                    print(ce_current_low)
                    pe_ohlc_price_quote = client.quotes(instrument_tokens=pe_instrument_tokens_json, quote_type='ohlc')
                    print(pe_ohlc_price_quote)
                    pe_current_low = int(float(pe_ohlc_price_quote['message'][0]['ohlc']['low']))
                    print(pe_current_low)

                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, ce_current_low, 'ATP!K20', "False", False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, pe_current_low, 'ATP!K21', "False", False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "NA", 'ATP!L20', "False", False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "NA", 'ATP!L21', "False", False)

                    read_trade_details = read_gsheet_data(service, spreadsheet_id, 'ATP!A19:H21')
                    trade_details_dataframe = pd.DataFrame(read_trade_details)
                    print(trade_details_dataframe)
                    header_row = 0
                    trade_details_dataframe.columns = trade_details_dataframe.iloc[header_row]
                    trade_details_dataframe = trade_details_dataframe.drop(header_row)
                    trade_details_dataframe = trade_details_dataframe.reset_index(drop=True)
                    print(trade_details_dataframe)

                    '''
                    trade_details_dataframe_styled = trade_details_dataframe.style.set_precision(2).set_properties(
                        **{'border': '2px solid green', 'color': 'black',
                           'text-align': 'center'}).hide_index().background_gradient().set_table_styles(
                        styles)
                    '''
                    # ignore styling
                    trade_details_dataframe_styled = trade_details_dataframe
                    print(trade_details_dataframe_styled)
                    file_name = "NewPassiveIntradayTradeBankNifty.png"
                    dfi.export(trade_details_dataframe_styled, 'images/' + file_name)
                    telegram_message_text = "New Passive Intraday Trade - " + spot_symbol_name
                    # Send to telegram
                    post_image_to_telegram(chat_id_main_channel, telegram_message_text, 'images/' + file_name)

                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "DONE", 'ATP!A16', "False", False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "DO", 'ATP!B16', "False", False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "DO", 'ATP!C16', "False", False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "DO", 'ATP!D16', "False", False)
                    # print("Above script will post the trade details to google sheets")
                else:
                    selected_strike_row = ce_trading_symbols.loc[
                        ce_trading_symbols['dStrikePrice;'] == extended_selected_strike]
                    print(selected_strike_row)
                    ce_trading_symbol = int(float(selected_strike_row['pSymbol'].iloc[0]))
                    print(ce_trading_symbol)

                    instrument_tokens = '[{"instrument_token": "' + str(
                        ce_trading_symbol) + '", "exchange_segment": "nse_fo"}]'
                    print(instrument_tokens)
                    ce_instrument_tokens_json = json.loads(instrument_tokens)
                    print(ce_instrument_tokens_json)
                    spot_price_quote = client.quotes(instrument_tokens=ce_instrument_tokens_json, quote_type="ltp")
                    print(spot_price_quote)
                    spot_ltp = spot_price_quote['message'][0]['ltp']
                    print(spot_ltp)

                    ce_entry_price = float(spot_ltp)
                    ce_tsl_component = int(float(ce_entry_price / 4))  # change to 4
                    ce_stoploss = ce_entry_price + ce_tsl_component + 1.1
                    # ceSLPrice = ceSL + 20  # change to 20
                    ce_tsl = ce_entry_price - ce_tsl_component
                    # Adjust SL slightly
                    ce_sl_mod_value = int(float(ce_stoploss)) % 10
                    if ce_sl_mod_value == 8:
                        ce_stoploss = ce_stoploss + 2.1
                    if ce_sl_mod_value == 9:
                        ce_stoploss = ce_stoploss + 1.1
                    if ce_sl_mod_value == 0:
                        ce_stoploss = ce_stoploss + 1.1
                    # if ceSLPrice > float(ceSL * 1.19):
                    # ceSLPrice = int(float(ceSL * 1.19))

                    selected_strike_row = pe_trading_symbols.loc[
                        pe_trading_symbols['dStrikePrice;'] == extended_selected_strike]
                    print(selected_strike_row)
                    pe_trading_symbol = int(float(selected_strike_row['pSymbol'].iloc[0]))
                    print(pe_trading_symbol)

                    instrument_tokens = '[{"instrument_token": "' + str(
                        pe_trading_symbol) + '", "exchange_segment": "nse_fo"}]'
                    print(instrument_tokens)
                    pe_instrument_tokens_json = json.loads(instrument_tokens)
                    print(pe_instrument_tokens_json)
                    spot_price_quote = client.quotes(instrument_tokens=pe_instrument_tokens_json, quote_type="ltp")
                    print(spot_price_quote)
                    spot_ltp = spot_price_quote['message'][0]['ltp']
                    print(spot_ltp)
                    pe_entry_price = float(spot_ltp)

                    pe_tsl_component = int(float(pe_entry_price / 4))  # change to 4
                    pe_stoploss = pe_entry_price + pe_tsl_component + 1.1
                    # peSLPrice = peSL + 20  # change to 20
                    pe_tsl = pe_entry_price - pe_tsl_component
                    # Adjust SL slightly
                    pe_sl_mod_value = int(float(pe_stoploss)) % 10
                    if pe_sl_mod_value == 8:
                        pe_stoploss = pe_stoploss + 2.1
                    if pe_sl_mod_value == 9:
                        pe_stoploss = pe_stoploss + 1.1
                    if pe_sl_mod_value == 0:
                        pe_stoploss = pe_stoploss + 1.1
                    # if peSLPrice > float(peSL * 1.19):
                    # peSLPrice = int(float(peSL * 1.19))

                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, selected_strike, 'ATP!D20', "False",
                                                     False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, selected_strike, 'ATP!D21', "False",
                                                     False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, ce_entry_price, 'ATP!E20', "False", False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, pe_entry_price, 'ATP!E21', "False", False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, ce_stoploss, 'ATP!F20', "False", False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, pe_stoploss, 'ATP!F21', "False", False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, ce_tsl, 'ATP!G20', "False", False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, pe_tsl, 'ATP!G21', "False", False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "PaperTrade", 'ATP!I20', "False", False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "PaperTrade", 'ATP!I21', "False", False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, ce_trading_symbol, 'ATP!J20', "False",
                                                     False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, pe_trading_symbol, 'ATP!J21', "False",
                                                     False)

                    # ce_ohlc_price_quote = client.quote(instrument_token=ce_trading_symbol, quote_type='OHLC')
                    ce_ohlc_price_quote = client.quotes(instrument_tokens=ce_instrument_tokens_json, quote_type="ohlc")
                    print(ce_ohlc_price_quote)
                    ce_current_low = int(float(ce_ohlc_price_quote['message'][0]['ohlc']['low']))
                    ce_current_high = int(float(ce_ohlc_price_quote['message'][0]['ohlc']['high']))
                    pe_ohlc_price_quote = client.quotes(instrument_tokens=pe_instrument_tokens_json, quote_type='ohlc')
                    print(pe_ohlc_price_quote)
                    pe_current_low = int(float(pe_ohlc_price_quote['message'][0]['ohlc']['low']))
                    pe_current_high = int(float(pe_ohlc_price_quote['message'][0]['ohlc']['high']))

                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, ce_current_low, 'ATP!K20', "False", False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, pe_current_low, 'ATP!K21', "False", False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, ce_current_high, 'ATP!L20', "False",
                                                     False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, pe_current_high, 'ATP!L21', "False",
                                                     False)

                    read_trade_details = read_gsheet_data(service, spreadsheet_id, 'ATP!A19:H21')
                    trade_details_dataframe = pd.DataFrame(read_trade_details)
                    print(trade_details_dataframe)
                    header_row = 0
                    trade_details_dataframe.columns = trade_details_dataframe.iloc[header_row]
                    trade_details_dataframe = trade_details_dataframe.drop(header_row)
                    trade_details_dataframe = trade_details_dataframe.reset_index(drop=True)
                    print(trade_details_dataframe)
                    '''
                    trade_details_dataframe_styled = trade_details_dataframe.style.set_properties(
                        **{'border': '2px solid green', 'color': 'black',
                           'text-align': 'center'}).hide_index().background_gradient().set_table_styles(
                        styles)
                    '''
                    # ### ignore styling
                    trade_details_dataframe_styled = trade_details_dataframe
                    print(trade_details_dataframe_styled)
                    file_name = "NewPassiveIntradayTradeBankNifty.png"
                    dfi.export(trade_details_dataframe_styled, 'images/' + file_name)
                    telegram_message_text = "Paper Trade: New Passive Intraday Trade - " + spot_symbol_name
                    # Send to telegram
                    post_image_to_telegram(chat_id_main_channel, telegram_message_text, 'images/' + file_name)

                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "DONE", 'ATP!A16', "False", False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "DO", 'ATP!B16', "False", False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "DO", 'ATP!C16', "False", False)
                    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "DO", 'ATP!D16', "False", False)
                    # print("Above script will post the trade details to google sheets")

        log_last_updated_time_in_gsheets(service)

    except Exception as error:
        print("Error in fetch_and_submit_data {0}".format(error))
        message_text = 'Passive Trading - Issue during entry - ' + "Error, : {0}".format(error)
        post_text_to_telegram(chat_id_error_channel, message_text)


def collect_previous_day_eod_data(service):
    try:

        fetch_requested_data = read_gsheet_data(service, "14BT6yWg0tJahYr2BhDxdhMU44nRaspYkNaWJ5K9s58Q", 'Vix!A2:B2')
        vix_from_sheet = fetch_requested_data[0][0]
        print(vix_from_sheet)
        vix_updated_datetime_from_sheet = fetch_requested_data[0][1]
        vix_updated_datetime_from_sheet_datetime_object = datetime.strptime(vix_updated_datetime_from_sheet,
                                                                            "%m/%d/%Y, %H:%M:%S")
        print(vix_updated_datetime_from_sheet_datetime_object)
        print(datetime.now())
        difference_in_time = datetime.now() - vix_updated_datetime_from_sheet_datetime_object
        print(difference_in_time)
        duration_in_seconds = difference_in_time.total_seconds()
        print(duration_in_seconds)
        duration_in_minutes = divmod(duration_in_seconds, 60)[0]
        print(duration_in_minutes)

        if duration_in_minutes < 90:
            print("Vix is being updated using the value from the sheet")
            vix_ltp = vix_from_sheet
            send_data_to_gsheet_for_one_cell(service, spreadsheet_id, vix_ltp, 'ATP!C7', "False", False)
            send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "DONE", 'ATP!E16', "False", False)
            log_last_updated_time_in_gsheets(service)

    except Exception as error:
        print("Error in fetch_and_submit_data {0}".format(error))
        message_text = 'Passive Trading - Error during collecting previous day EOD data - ' + "Error, : {0}".format(
            error)
        post_text_to_telegram(chat_id_error_channel, message_text)


def find_platform_health(service):
    try:
        print("find_platform_health started.....")
        client = NeoAPI(consumer_key=kotak_consumer_key, consumer_secret=kotak_consumer_secret,
                        environment='prod')
        print("Session initiated")
        login_client_response = client.login(mobilenumber=kotak_live_user_id, password=kotak_live_password)
        print(login_client_response)
        login_client_response = client.session_2fa(OTP=kotak_mpin)
        print(login_client_response)
        # Generated session token successfully

        read_trading_conditions = read_gsheet_data(service, spreadsheet_id, 'ATP!A7:H7')
        spot_symbol_name = read_trading_conditions[0][3]
        spot_symbol_id = int(read_trading_conditions[0][4])
        print(spot_symbol_name)
        print(spot_symbol_id)
        print("------------------------------------------")
        print("Pull data")

        # Get Spot Price
        instrument_tokens = '[{"instrument_token": "' + str(spot_symbol_id) + '", "exchange_segment": "nse_cm"}]'
        print(instrument_tokens)
        instrument_tokens_json = json.loads(instrument_tokens)
        print(instrument_tokens_json)
        spot_price_quote = client.quotes(instrument_tokens=instrument_tokens_json, quote_type="ltp")
        print(spot_price_quote)
        spot_ltp = spot_price_quote['message'][0]['ltp']
        print(spot_ltp)

        # Find all option strike price symbols
        trading_symbol_list = get_list_of_trading_symbols(spot_symbol_name)

        # Get trading symbol details and pull the symbols alone and append it and get the quote
        ce_trading_symbols = trading_symbol_list[0]
        pe_trading_symbols = trading_symbol_list[1]
        ce_trading_symbol_list = ce_trading_symbols['pSymbol'].tolist()
        pe_trading_symbol_list = pe_trading_symbols['pSymbol'].tolist()

        ce_trading_symbol_list_items = ce_trading_symbol_list[0:1]
        ce_jsonfile = []
        for index, value in enumerate(ce_trading_symbol_list_items):
            ce_jsonfile.append({
                "instrument_token": ce_trading_symbol_list_items[index],
                "exchange_segment": "nse_fo"
            })
        ce_instrument_tokens_list = json.dumps(ce_jsonfile)
        print(ce_instrument_tokens_list)
        print(json.dumps(ce_jsonfile, indent=4))
        ce_instrument_tokens_json = json.loads(ce_instrument_tokens_list)
        print(ce_instrument_tokens_json)
        ce_option_sample_quote = client.quotes(instrument_tokens=ce_instrument_tokens_json)
        print(ce_option_sample_quote)

        pe_trading_symbol_list_items = pe_trading_symbol_list[0:1]
        pe_jsonfile = []
        for index, value in enumerate(pe_trading_symbol_list_items):
            pe_jsonfile.append({
                "instrument_token": pe_trading_symbol_list_items[index],
                "exchange_segment": "nse_fo"
            })
        pe_instrument_tokens_list = json.dumps(pe_jsonfile)
        print(pe_instrument_tokens_list)
        print(json.dumps(pe_jsonfile, indent=4))
        pe_instrument_tokens_json = json.loads(pe_instrument_tokens_list)
        print(pe_instrument_tokens_json)
        pe_option_sample_quote = client.quotes(instrument_tokens=pe_instrument_tokens_json)
        print(pe_option_sample_quote)

        print("Data pulling completed")
        log_last_updated_time_in_gsheets(service)

    except Exception as error:
        print("Error in fetch_and_submit_data {0}".format(error))
        message_text = 'Passive Trading - Issue during health check - ' + "Error, : {0}".format(error)
        post_text_to_telegram(chat_id_error_channel, message_text)


def get_list_of_trading_symbols(spot_trading_symbol):
    # Set up the URL and call the URL to get symbols
    year_as_of_today = datetime.now().strftime("%Y")
    month_as_of_today = datetime.now().strftime("%m")
    date_as_of_today = datetime.now().strftime("%d")
    datetime_value = year_as_of_today + "-" + month_as_of_today + "-" + date_as_of_today
    kotak_fno_symbols_url = "https://lapi.kotaksecurities.com/wso2-scripmaster/v1/prod/" + datetime_value + "/transformed/nse_fo.csv"
    print(kotak_fno_symbols_url)
    kotak_fno_symbols_response = requests.get(kotak_fno_symbols_url)
    print(kotak_fno_symbols_response)

    # Parse the response into a data frame
    kotak_fno_symbols_dataframe = pd.read_csv(StringIO(str(kotak_fno_symbols_response.text)))
    header_row = 0
    print(kotak_fno_symbols_dataframe)
    kotak_fno_symbols_dataframe.rename(columns=lambda x: x.strip(), inplace=True)
    print(kotak_fno_symbols_dataframe)
    print(spot_trading_symbol)
    print(kotak_fno_symbols_dataframe.iloc[header_row])

    # filter only the symbol needed and get the symbol list for CE and PE separately
    kotak_fno_symbols_dataframe_filtered = kotak_fno_symbols_dataframe[
        kotak_fno_symbols_dataframe['pSymbolName'] == spot_trading_symbol]
    print(kotak_fno_symbols_dataframe_filtered)
    kotak_fno_symbols_dataframe_filtered_ce_only = kotak_fno_symbols_dataframe_filtered[
        kotak_fno_symbols_dataframe_filtered['pOptionType'] == "CE"]
    print(kotak_fno_symbols_dataframe_filtered_ce_only)
    kotak_fno_symbols_dataframe_filtered_pe_only = kotak_fno_symbols_dataframe_filtered[
        kotak_fno_symbols_dataframe_filtered['pOptionType'] == "PE"]
    print(kotak_fno_symbols_dataframe_filtered_pe_only)

    # Define the Expiry Date
    expiry_date_list = kotak_fno_symbols_dataframe_filtered_ce_only.lExpiryDate.unique()
    print(expiry_date_list)
    expiry_date_list.sort()
    nearest_expiry_date = expiry_date_list[0]
    print("nearest_expiry_date =" + str(nearest_expiry_date))

    # Filter symbols using expiry date
    kotak_fno_symbols_dataframe_final_ce_only = kotak_fno_symbols_dataframe_filtered_ce_only[
        kotak_fno_symbols_dataframe_filtered_ce_only['lExpiryDate'] == nearest_expiry_date]
    kotak_fno_symbols_dataframe_final_pe_only = kotak_fno_symbols_dataframe_filtered_pe_only[
        kotak_fno_symbols_dataframe_filtered_pe_only['lExpiryDate'] == nearest_expiry_date]
    kotak_fno_symbols_dataframe_final_ce_only.sort_values('dStrikePrice;')
    kotak_fno_symbols_dataframe_final_pe_only.sort_values('dStrikePrice;')
    print(kotak_fno_symbols_dataframe_final_ce_only)
    print(kotak_fno_symbols_dataframe_final_pe_only)

    # Put it as 2 different items in an array and return
    trading_symbol_list = [kotak_fno_symbols_dataframe_final_ce_only, kotak_fno_symbols_dataframe_final_pe_only]
    return trading_symbol_list


def kotak_data_preparation(service, gsheet_spreadsheet_id, spreadsheet_range):
    fetch_requested_data = read_gsheet_data(service, gsheet_spreadsheet_id, spreadsheet_range)
    global kotak_live_user_id
    global kotak_live_password
    global kotak_consumer_key
    global kotak_access_token
    global kotak_mpin
    global kotak_consumer_secret
    kotak_live_user_id = fetch_requested_data[0][0]
    kotak_live_password = fetch_requested_data[1][0]
    kotak_consumer_key = fetch_requested_data[2][0]
    kotak_access_token = fetch_requested_data[3][0]
    kotak_mpin = fetch_requested_data[4][0]
    kotak_consumer_secret = fetch_requested_data[5][0]
    print(kotak_live_user_id)


def post_image_to_telegram(telegram_channel, caption, image_file):
    try:
        # use token generated in first step
        bot = telegram.Bot(token=telegram_token)
        status = bot.send_photo(chat_id=telegram_channel, photo=open(image_file, 'rb'), caption=caption)
        print(status)
        print('--------Just posted to telegram--------')
    except Exception as error:
        print("Error while posting to telegram : {0}".format(error))


def post_text_to_telegram(telegram_channel, message_text):
    try:
        # use token generated in first step
        bot = telegram.Bot(token=telegram_token)
        status = bot.send_message(chat_id=telegram_channel, text=message_text, parse_mode=telegram.ParseMode.HTML)
        print(status)
        print('--------Just posted to telegram--------')
    except Exception as error:
        print("Error while posting to telegram : {0}".format(error))


def telegram_data_preparation(service, gsheet_spreadsheet_id, spreadsheet_range):
    fetch_requested_data = read_gsheet_data(service, gsheet_spreadsheet_id, spreadsheet_range)
    global telegram_token
    global telegram_channel
    global chat_id_main_channel
    global chat_id_error_channel
    telegram_token = fetch_requested_data[0][0]
    telegram_channel = str(fetch_requested_data[1][0])
    chat_id_main_channel = str(fetch_requested_data[1][0])
    chat_id_error_channel = str(fetch_requested_data[1][0])
    print(telegram_token)
    print(telegram_channel)


def login():
    """
    Authentication into google sheets.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', scopes)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    # Create service for google sheets operations
    service = build('sheets', 'v4', credentials=creds)
    return service


def read_gsheet_data(service, gsheet_spreadsheet_id, spreadsheet_range):
    result = service.spreadsheets().values().get(
        spreadsheetId=gsheet_spreadsheet_id, range=spreadsheet_range, valueRenderOption='UNFORMATTED_VALUE').execute()
    data_pulled_from_gsheet = result.get('values')
    return data_pulled_from_gsheet


def send_data_to_gsheet_for_one_cell(service, gsheet_spreadsheet_id, data, spreadsheet_range, clear, drop_headers):
    # Call the Sheets API
    # sheet = service.spreadsheets()

    # Clear existing data if clear flag is true
    if clear == "True":
        service.spreadsheets().values().clear(
            spreadsheetId=gsheet_spreadsheet_id,
            range=spreadsheet_range
        ).execute()

    data_body = {
        "valueInputOption": "RAW",
        "data": [
            {
                'range': spreadsheet_range,
                'values': [[data]]
            },

        ]
    }

    # Send data to google sheets
    service.spreadsheets().values().batchUpdate(
        spreadsheetId=gsheet_spreadsheet_id,
        body=data_body
    ).execute()


def send_data_to_gsheet_multiple_cells(service, gsheet_spreadsheet_id, data, spreadsheet_range, clear, drop_headers):
    # Call the Sheets API
    # sheet = service.spreadsheets()

    # Clear existing data if clear flag is true
    if clear == "True":
        service.spreadsheets().values().clear(
            spreadsheetId=gsheet_spreadsheet_id,
            range=spreadsheet_range
        ).execute()

    data_body = {
        "valueInputOption": "RAW",
        "data": [
            {
                'range': spreadsheet_range,
                'values': data
            },

        ]
    }

    # Send data to google sheets
    service.spreadsheets().values().batchUpdate(
        spreadsheetId=gsheet_spreadsheet_id,
        body=data_body
    ).execute()


def convert_excel_time(excel_time):
    """
    converts excel float format to pandas datetime object
    round to '1min' with
    .dt.round('1min') to correct floating point conversion innaccuracy
    """

    return pd.to_datetime('1899-12-30') + pd.to_timedelta(excel_time, 'D')


def is_today_holiday(service, gsheet_spreadsheet_id, spreadsheet_range):
    """
    :param service:
    :param gsheet_spreadsheet_id:
    :param spreadsheet_range:
    :return: 1 if it is a holiday, 0 if it is working day
    """
    holiday_list_raw = read_gsheet_data(service, gsheet_spreadsheet_id, spreadsheet_range)
    holiday_list = np.hstack(holiday_list_raw)
    today_datetime = datetime.now()
    today_date = today_datetime.date()
    return_value = 0
    for each_holiday in holiday_list:
        current_holiday_as_integer = int(each_holiday)
        current_holiday_as_datetime = convert_excel_time(current_holiday_as_integer)
        current_holiday_as_date = current_holiday_as_datetime.date()
        # If today's date matches with any date in holiday list
        if today_date == current_holiday_as_date:
            return_value = 1
    return return_value


def log_last_updated_time_in_gsheets(service):
    hour_now = int(datetime.now().strftime("%H"))
    minute_now = int(datetime.now().strftime("%M"))
    if minute_now < 10:
        minute_now_updated = '0' + str(minute_now)
        log_time = str(hour_now) + ":" + str(minute_now_updated)
    else:
        log_time = str(hour_now) + ":" + str(minute_now)
    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, log_time, 'ATP!I7', "False", False)


def automated_trade_logging(service):
    print("Backup the trade results in Google Sheets")
    read_trade_details = read_gsheet_data(service, spreadsheet_id, 'ATP!A20:P21')
    print(read_trade_details)
    trading_symbol = read_trade_details[0][1]
    print(trading_symbol)
    banknifty_trade_leg1_details = "SOLD " + trading_symbol + " " + str(read_trade_details[0][3]) + " CE at " + str(
        read_trade_details[0][4]) + " and Exited at " + str(read_trade_details[0][13])
    banknifty_trade_leg2_details = "SOLD " + trading_symbol + " " + str(read_trade_details[1][3]) + " PE at " + str(
        read_trade_details[1][4]) + " and Exited at " + str(read_trade_details[1][13])
    print(banknifty_trade_leg1_details)
    print(banknifty_trade_leg2_details)
    banknifty_trade_leg1_pl = read_trade_details[0][15]
    banknifty_trade_leg2_pl = read_trade_details[1][15]
    banknifty_trade_net_pl = banknifty_trade_leg1_pl + banknifty_trade_leg2_pl
    print(banknifty_trade_net_pl)
    date_today = str(datetime.now().strftime('%d-%b-%Y'))
    print(date_today)
    read_trade_log_details = read_gsheet_data(service, spreadsheet_id, 'ATP!I7:J7')
    row_counter = read_trade_log_details[0][1]
    print(row_counter)

    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, date_today, 'TradeLog!A' + str(row_counter), "False",
                                     False)
    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, banknifty_trade_leg1_details,
                                     'TradeLog!B' + str(row_counter), "False", False)
    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, banknifty_trade_leg2_details,
                                     'TradeLog!C' + str(row_counter), "False", False)
    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, banknifty_trade_net_pl, 'TradeLog!D' + str(row_counter),
                                     "False", False)
    log_last_updated_time_in_gsheets(service)
    send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "DONE", 'ATP!D16', "False", False)


def main():
    service = login()



    # Prepare data to log into Kotak Securities
    kotak_data_preparation(service, spreadsheet_id, 'Kotak!B2:B7')
    telegram_data_preparation(service, spreadsheet_id, 'Telegram!B2:B5')
    weekday_now = datetime.now().weekday()
    print(weekday_now)



    # 0-4 indicates Monday to Friday, 5 & 6 is Saturday & Sunday
    if weekday_now < 5:
        holiday_check = is_today_holiday(service, spreadsheet_id, 'HolidayList!C2:C20')
        # Don't run the script on holidays (i.e) if returned value is 1
        if holiday_check == 0:
            hour_now = int(datetime.now().strftime("%H"))
            print(hour_now)
            minute_now = int(datetime.now().strftime("%M"))
            print(minute_now)
            read_trading_conditions = read_gsheet_data(service, spreadsheet_id, 'ATP!A2:H2')
            trading_switch = read_trading_conditions[0][1]
            read_trading_flags = read_gsheet_data(service, spreadsheet_id, 'ATP!A16:J16')

            if trading_switch == "Y":
                if hour_now > 7 or hour_now < 20:
                    print("Execution starts")
                    read_flag6 = read_trading_flags[0][5]
                    read_flag5 = read_trading_flags[0][4]
                    read_flag4 = read_trading_flags[0][3]
                    read_flag3 = read_trading_flags[0][2]
                    read_flag2 = read_trading_flags[0][1]
                    read_flag1 = read_trading_flags[0][0]
                    entry_hour = read_trading_conditions[0][3]
                    entry_minute_start = read_trading_conditions[0][4]
                    entry_minute_end = read_trading_conditions[0][5]
                    exit_hour = read_trading_conditions[0][6]
                    exit_hour_plus_one = exit_hour + 1
                    print(exit_hour_plus_one)
                    exit_minute = read_trading_conditions[0][7]

                    print(read_flag1)
                    print(read_flag5)
                    print(read_flag6)
                    print(entry_minute_start)
                    print(entry_minute_end)
                    print(entry_hour)
                    print(hour_now)
                    print(minute_now)
                    if hour_now == 8 and read_flag6 == "DO":
                        print("Reset Flags and data")
                        # Reset Data fields
                        send_data_to_gsheet_multiple_cells(service, spreadsheet_id, [["", "", "", ""]], 'ATP!D20',
                                                           "False", False)
                        send_data_to_gsheet_multiple_cells(service, spreadsheet_id, [["", "", "", ""]], 'ATP!D21',
                                                           "False", False)
                        send_data_to_gsheet_multiple_cells(service, spreadsheet_id, [["", "", "", ""]], 'ATP!I20',
                                                           "False", False)
                        send_data_to_gsheet_multiple_cells(service, spreadsheet_id, [["", "", "", ""]], 'ATP!I21',
                                                           "False", False)
                        send_data_to_gsheet_multiple_cells(service, spreadsheet_id, [["", ""]], 'ATP!N20',
                                                           "False", False)
                        send_data_to_gsheet_multiple_cells(service, spreadsheet_id, [["", ""]], 'ATP!N21',
                                                           "False", False)
                        # Reset Flags
                        send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "DO", 'ATP!A16', "False", False)
                        send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "NA", 'ATP!B16', "False", False)
                        send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "NA", 'ATP!C16', "False", False)
                        send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "NA", 'ATP!D16', "False", False)
                        send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "DO", 'ATP!E16', "False", False)
                        send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "DO", 'ATP!G16', "False", False)
                        send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "DO", 'ATP!J16', "False", False)
                        send_data_to_gsheet_for_one_cell(service, spreadsheet_id, "DONE", 'ATP!F16', "False", False)

                    if hour_now == 8 and read_flag5 == "DO":
                        collect_previous_day_eod_data(service)
                    if hour_now > exit_hour_plus_one and read_flag4 == "DO" and read_flag1 == "DONE":
                        automated_trade_logging(service)
                    if 8 < hour_now < 16 and read_flag2 == "DO":
                        monitor_trades_throughout_the_day(service)
                    if ((hour_now > exit_hour or (
                            hour_now == exit_hour and minute_now > exit_minute)) and read_flag3 == "DO"):
                        eod_auto_square_off_and_status_reporting(service)
                    if (hour_now == entry_hour and (
                            entry_minute_start < minute_now < entry_minute_end) and read_flag1 == "DO" and read_flag5 == "DONE" and read_flag6 == "DONE"):
                        print("minute_now")
                        find_my_trades(service)
                    if hour_now == 9 and 15 < minute_now < entry_minute_start:
                        find_platform_health(service)

    print("Execution Completed")



if __name__ == '__main__':
    main()
