import feedparser
import time
import pytz
from datetime import datetime
import telebot
from telebot import types
import os 
import time
import re

bot = telebot.TeleBot("6893650720:AAGadPW3xYHoRnH-7AxuL2DCKVuXmzbk1LM", parse_mode=None) # You can set parse_mode by default. HTML or MARKDOWN


chat_id=-1002059505388 #BB개발 테스트방

#1대1방
achat_id=1276134988

#bb그룹방
#chat_id= -1002057698933


#bb리서치방
#chat_id= -1001656364050

#bot.send_message()

first=True
lastEntryUpdated=None
lastEntryLink=None

def check_string1(input_string):
    # '단일판매'가 문자열에 포함되어 있는지 확인
    contains_required = '단일판매' in input_string

    # '기재정정'이 문자열에 포함되어 있지 않은지 확인
    does_not_contain_forbidden = '기재정정' not in input_string

    # 두 조건을 모두 만족하는지 반환
    return contains_required and does_not_contain_forbidden

def check_string2(input_string):
    contains_required = '신규시설투자' in input_string or '손익구조30%' in input_string or '공정공시' in input_string

    # '기재정정'이 문자열에 포함되어 있지 않은지 확인
    does_not_contain_forbidden = '기재정정' not in input_string

    # 두 조건을 모두 만족하는지 반환
    return contains_required and does_not_contain_forbidden

def check_string3(input_string):
    contains_required = '유상증자결정' in input_string or '전환사채권발행결정' in input_string or '교환사채권발행결정' in input_string

    # '기재정정'이 문자열에 포함되어 있지 않은지 확인
    does_not_contain_forbidden = '기재정정' not in input_string

    # 두 조건을 모두 만족하는지 반환
    return contains_required and does_not_contain_forbidden

def check_string4(input_string):
    contains_required = '자기주식취득결정' in input_string or '자기주식취득신탁계약체결결정' in input_string

    # '기재정정'이 문자열에 포함되어 있지 않은지 확인
    does_not_contain_forbidden = '기재정정' not in input_string

    # 두 조건을 모두 만족하는지 반환
    return contains_required and does_not_contain_forbidden

def struct_time_to_datetime(struct_time):
    # struct_time을 datetime 객체로 변환
    return datetime(*struct_time[:6])

def convert_to_kst(utc_dt):
    # UTC를 KST로 변환
    kst_tz = pytz.timezone('Asia/Seoul')
    return utc_dt.astimezone(kst_tz)

def check_feed(url):
    global first
    global lastEntryUpdated
    global lastEntryLink
    #print('Restart')
    while True:
        feed = feedparser.parse(url)
        if not feed.entries:  # 만약 항목이 없으면 예외 처리
            continue
        new_entries = []
        start=False
        for entry in feed.entries[::-1]:
            if start == False:
                if entry.updated == lastEntryUpdated and entry.link == lastEntryLink:
                    start=True
                    continue
                else:
                    continue
            if first == True:
                continue
            if check_string1(entry.title):
                message = ''
                message = message + '제목:' + entry.title + '\n'
                message = message + "링크:" + entry.link + '\n'
                #print(message)
                bot.send_message(chat_id, message)
            elif check_string2(entry.title):
                message = ''
                message = message + '제목:' + entry.title + '\n'
                message = message + "링크:" + entry.link + '\n'
                #print(message)
                bot.send_message(achat_id, message)
            elif check_string3(entry.title) or check_string4(entry.title):
                message = ''
                message = message + '제목:' + entry.title + '\n'
                message = message + "링크:" + entry.link + '\n'
                #print(message)
                bot.send_message(achat_id, message)
            #print(message)
            #bot.send_message(chat_id, message)
        if first == True:
            first=False
        if len(feed.entries) > 0:
            lastEntryUpdated=feed.entries[0].updated
            lastEntryLink=feed.entries[0].link
            print('lastEntry:' + str(feed.entries[0]))
            
        time.sleep(5)  # 5초 대기

bot.send_message(achat_id, 'restart')
url = "https://dart.fss.or.kr/api/todayRSS.xml"
#url = "http://kind.krx.co.kr:80/disclosure/rsstodaydistribute.do?method=searchRssTodayDistribute&repIsuSrtCd=&mktTpCd=0&searchCorpName=&currentPageSize=15"

check_feed(url)

# 실행
bot.polling()