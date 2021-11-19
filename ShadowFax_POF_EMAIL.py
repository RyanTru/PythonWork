# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import pyodbc
import pandas as pd
from sqlalchemy import create_engine
import urllib
import xlwings as xw
import time 
import datetime
import win32com.client
import time
import matplotlib.pyplot as plt
import numpy as np
null = "NULL"
#'MELI','SHOP','FB','TWLO'

#COMP_SYMBOL = ('SPY')
#Initialize Connectho
params = urllib.parse.quote_plus(r'DRIVER={SQL Server Native Client 11.0};SERVER=localhost\SQLEXPRESS;DATABASE=DBN;Trusted_Connection=yes; timeout=300')
conn_str = 'mssql+pyodbc:///?odbc_connect={}'.format(params)
engine = create_engine(conn_str)
conn = engine.connect()
#Dates

SPY_Historic_Max = "select max(date_time) from SPY_ANALYTIC_UPLOAD"
Historic_Max_date = pd.read_sql_query(SPY_Historic_Max, engine)
Historic_Max_date = Historic_Max_date.iloc[0]

Signal_Detail = """ 
---RANK and rank history
 select format(SR.Date_Time,'MM/dd/yyyy') SIGNAL_DATE, SR.Symbol, SR.RANK, SR.Strat, Su.Last, B.Beta, (SU.PCT_Chg-SPY.PCT_Chg) Daily_SPY_ALPHA, ST.Sig_Count, ST.MAX_RANKING, format(ST.MAX_SIGDATE,'MM/dd/yyyy') MAX_SIGDATE, ST.MIN_RANKING, format(ST.MIN_SIGDATE,'MM/dd/yyyy') MIN_SIGDATE from 
(select * from SignalRank where Rank >= .8) SR
join
(select Symbol, last, dailynetChng, Open_PDClose, PD_Close, ((DailyNetChng+Open_PDClose)/PD_Close) PCT_Chg, high, low, ATR_7D, ATR_20D, HiSlope_7D, HiSlope_20D, LoSlope_7D, LoSlope_20D, date_time	 from SPY_ANALYTIC_UPLOAD where Date_Time in ((select max(date_time) from SPY_ANALYTIC_UPLOAD))) SU
on SR.Symbol = SU.Symbol and SR.Date_Time = SU.Date_Time and SR.Date_Time in (select max(date_time) from SPY_ANALYTIC_UPLOAD)
left join
(select distinct SMAX.sym, MAX(C1.Count_SYM) SIG_COUNT, max(SMAX.MAX_RANK) MAX_RANKING, max(SMAX.Date_Time) MAX_SIGDATE, min(MIN_RANK) MIN_RANKING, min(SMIN.Date_Time) MIN_SIGDATE  from 
(SELECT distinct S1.Sym, s1.MAX_RANK, d1.Date_Time from 
(select distinct symbol Sym, max(RANK) MAX_RANK  from signalrank where rank >= .0 GROUP BY SYMBOL) s1 
 join 
(select * from SignalRank) d1 
on s1.max_rank = d1.rank and s1.Sym = d1.Symbol) SMAX
 join
(SELECT distinct S1.Sym, s1.MIN_RANK, d1.Date_Time from 
(select distinct symbol Sym, min(RANK) MIN_RANK  from signalrank where rank >= .0 GROUP BY SYMBOL) s1 
join 
(select * from SignalRank) d1 on s1.min_rank = d1.rank and s1.Sym = d1.Symbol) SMIN
on SMAX.Sym = smin.Sym
join 
(select distinct symbol Sym, count(symbol) Count_SYM  from signalrank where rank >= .0 GROUP BY SYMBOL) c1 
on c1.Sym = Smax.Sym  group by SMAX.Sym) ST
on SR.Symbol = ST.Sym
join
(select symbol, date_time, dailynetChng, Open_PDClose, PD_Close, ((DailyNetChng+Open_PDClose)/PD_Close) PCT_Chg from SPY_ANALYTIC_UPLOAD where symbol in ('SPY') and Date_Time in ((select max(date_time) from SPY_ANALYTIC_UPLOAD))) SPY
on SPY.date_time = sr.Date_Time
join 
(select Symbol, Beta from Beta) B
on B.Symbol = sr.Symbol
order by rank desc
"""

try:
    Signal_Detail = pd.read_sql_query(Signal_Detail, engine)
    print('Signal Detail Success')
except Exception as e:
    print(e)
    print('Signal Detail Failed')




import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.text import MIMEText

import datetime as dt
contacts = ["r.trudeau.capesun@gmail.com"]

for x in range(1):
    
    
    me = "gfcapitalrt@gmail.com"
    my_password = r"password here"
    you = contacts[x]
    subj = "UPDATE - RYs POF Signals" + dt.datetime.today().strftime("%m.%d.%Y")

    msg = MIMEMultipart('alternative')
    msg['Subject'] = subj
    msg['From'] = me
    msg['To'] = you

    body1 = Signal_Detail.to_html()
    body1 = MIMEText(body1, 'html')
    msg.attach(body1)



    
    # Send the message via gmail's regular server, over SSL - passwords are being sent, afterall
    s = smtplib.SMTP_SSL('smtp.gmail.com')
    # uncomment if interested in the actual smtp conversation
    # s.set_debuglevel(1)
    # do the smtp auth; sends ehlo if it hasn't been sent already
    s.login(me, my_password)
    
    s.sendmail(me, you, msg.as_string())
    s.quit() 