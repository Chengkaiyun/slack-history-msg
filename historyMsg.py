#-*- coding: utf-8 -*-
from bs4 import BeautifulSoup
import requests
import json
import pandas as pd
from openpyxl.workbook import Workbook
import re
from datetime import datetime

#loads()：將json數據轉化成dict數據
#dumps()：將dict數據轉化成json數據

TOKEN = 'xoxp-18480229362-46671649537-1308403574613-f1a48edad831bf429bc8bf56c0649bb8'
CHANNEL = 'C0JGJUVLP'
INPUT_OLDEST = input("input time")
LIMIT = '1000'


m = re.search('(\d{4})[^0-9]*(\d{1,2})[^0-9]*(\d{1,2})',INPUT_OLDEST)
year = int(m.group(1))
month = int(m.group(2))
day = int(m.group(3))
dt = datetime(year, month ,day).timestamp()
OLDEST = str(dt).split(".")[0]


FILE = 'https://slack.com/api/conversations.history?token=' +TOKEN+ \
       '&channel=' + CHANNEL+ \
        '&limit=' + LIMIT + \
       '&oldest=' + OLDEST + \
       '&pretty=1'

#將此頁面的HTML GET下來
r = requests.get(FILE)
#印出HTML
str_data = r.text

data = json.loads(str_data)
alldata = []


for msg in data['messages']:
    text = msg['text']
    text = text.replace("&amp;","&").replace('`','').replace("&gt;",">")

    ts = int(msg['ts'].split(".")[0])
    time = datetime.fromtimestamp(ts)


    m1 = re.search(r'【分類 Type[^【]*', text, re.MULTILINE)
    m2 = re.search(r'【問題 Issue[^【]*', text, re.MULTILINE)
    m3 = re.search(r'【型號 Device[^【]*', text, re.MULTILINE)
    m4 = re.search(r'【位置 Where[^【]*', text, re.MULTILINE)

    n1 = re.search(r'【問題 Issue[^【]*', text, re.MULTILINE)
    n2 = re.search(r'【型號 Device[^【]*', text, re.MULTILINE)
    n3 = re.search(r'【位置 Where[^【]*', text, re.MULTILINE)
    n4 = re.search(r'【相關人 Who[^【]*', text, re.MULTILINE)

    if m1 and m2 and m3 and m4:
        text = [time, m1.group(), m2.group(), m3.group(), m4.group()]
    elif n1 and n2 and n3 and n4:
        text = [time, n1.group(), n2.group(), n3.group(), n4.group()]
    else:
        text = [time,text]
    alldata.append(text)


df = pd.DataFrame(alldata)
df.to_excel('result.xlsx',sheet_name='sheet1', \
            header=['【時間 Time】','【分類 Type】','【問題 Issue】','【型號 Device】','【位置 Where】'], \
            index = None)
