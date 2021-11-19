# -*- coding: utf-8 -*-
"""
Created on Thu Jun 28 05:39:44 2018

@author: Ryan
"""
import xlwings as xw
import pandas as pd
import numpy as np
import time
import datetime as dt
from datetime import datetime
import pandas as pd
import pyodbc
import pandas as pd
from sqlalchemy import create_engine
import urllib

params = urllib.parse.quote_plus(r'DRIVER={SQL Server Native Client 11.0};SERVER=localhost\SQLEXPRESS;DATABASE=DBN;Trusted_Connection=yes; timeout=300')
conn_str = 'mssql+pyodbc:///?odbc_connect={}'.format(params)
engine = create_engine(conn_str)
conn = engine.connect()

app = xw.App(visible=True)
wb = xw.Book(R'C:\Users\Ryan\OneDrive\Desktop\REX_RUNS\Strategies\Scrape.xlsx')
time.sleep(10)
wb.close()
app.kill()
    #how many movers 
app = xw.App(visible=True)
wb = xw.Book(R'C:\Users\Ryan\OneDrive\Desktop\REX_RUNS\Strategies\Scrape.xlsx')
time.sleep(10)
wb.app.calculate()
time.sleep(40)
df = xw.Range('A1:O8450').options(pd.DataFrame).value
wb.close()
app.kill()


#df = pd.read_csv ("Book1.csv")
pd.to_numeric(df.NetChng, errors="coerce")
df = df.dropna()
dfc = df[['Symbol','NetChng','Volume','Last', 'Open', 'High', 'Low', '52High', '52Low']].copy() 
dfc['Rank'] = 0
dfc['Relslope'] = 0
dfc['ATRBreak'] = 0
dfc['CTR'] = 0
dfc['ATR'] = 0
dfc['Group'] = 0
dfc['Shares'] = 0
dfc['StopLoss'] = 0
dfc['Time'] = datetime.now()
dfc['Volume'] = pd.to_numeric(dfc.Volume)
dfc['NetChng'] = pd.to_numeric(dfc.NetChng)
dfc['Open'] = pd.to_numeric(dfc.Open)
dfc = dfc.where(dfc['Volume'] >= 50000)
dfc = dfc.where(dfc['Open'] > 0)
dfc = dfc.dropna()
dfdfc = pd.DataFrame({
        'Symbol':dfc['Symbol'], 
        'NetChng':dfc['NetChng'],
        'Last':dfc['Last'],
        'Volume':dfc['Volume'],
        'Open':dfc['Open'],
        'High':dfc['High'],
        'Low':dfc['Low'],
        '52High':dfc['52High'],                                                                                    
        '52Low':dfc['52Low'],
        'Rank':dfc['Rank'],
        'ATRBreak':dfc['ATRBreak'],
        'CTR':dfc['CTR'],
        'Relslope':dfc['Relslope'],
        'Group':dfc['Group'],
        'StopLoss':dfc['StopLoss'],
        'Shares':dfc['Shares'],
        'Time':dfc['Time'],
        'NetChng_pct_times_Volume':(dfc['Volume']*(dfc['NetChng']/dfc['Open']))
         })
dfdfc = dfdfc.dropna()
dfc = (dfdfc.sort_values(by=['NetChng_pct_times_Volume'], ascending=False))
dfc1 = dfc.where(dfc['Last'] > 1)
dfc1 = dfc1.mask(dfc1['Last'] > 10)
dfc1.Group = 1
dfc1 = dfc1.dropna()
dfc1 = dfc1.head(10)
dfc1.index = range(1,11)
dfc2 = dfc.where(dfc['Last'] >= 10)
dfc2 = dfc2.mask(dfc2['Last'] > 50)
dfc2.Group = 2
dfc2 = dfc2.dropna()
dfc2 = dfc2.head(10)
dfc2.index = range(11,21)
dfc3 = dfc.where(dfc['Last'] >= 50)
dfc3 = dfc3.mask(dfc3['Last'] > 100)
dfc3.Group = 3
dfc3 = dfc3.dropna()
dfc3 = dfc3.head(10)
dfc3.index = range(21,31)
dfc4 = dfc.where(dfc['Last'] >= 100)
dfc4 = dfc4.mask(dfc4['Last'] < 100)
dfc4.Group = 4
dfc4 = dfc4.dropna()
dfc4 = dfc4.head(10)
dfc4.index = range(31,41)
siglist = dfc1.append(dfc2).append(dfc3).append(dfc4)
siglist.to_sql('Siglist', con=engine, if_exists='replace')
print(siglist)

    
    




#print(dfcshrt.head(10))
 