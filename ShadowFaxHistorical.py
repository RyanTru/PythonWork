import re
from io import StringIO
from datetime import datetime, timedelta

import requests
import pandas as pd
import pyodbc
import pandas as pd
from sqlalchemy import create_engine
import urllib
import xlwings as xw
import time 
import win32com.client
import time


#Initialize
#Pull Historical Analytics From SQL
#Calc Prior Day Row
#Scrape all Day - minute intervals
#Calc Real Time Row     

#Initialize Connect
params = urllib.parse.quote_plus(r'DRIVER={SQL Server Native Client 11.0};SERVER=localhost\SQLEXPRESS;DATABASE=DBN;Trusted_Connection=yes; timeout=300')
conn_str = 'mssql+pyodbc:///?odbc_connect={}'.format(params)
engine = create_engine(conn_str)
conn = engine.connect()
error = 0
df_symbols = pd.DataFrame([['MARA']],columns = ['symbol'])
#Siglist_symbol = "select distinct symbol from siglist"
#df_symbols =  pd.read_sql_query(Siglist_symbol, engine)

for symbol in df_symbols['symbol']:
        error = 0
        try:
            #MAX DATE AND LIVE DATE CODE :) -- Experience counts
            SYMBOL = symbol
            SF_Live_Date = "select max(date_time) from ShadowFax_Live"
            SPY_Historic_Max = "select isnull(max(date_time),('12/01/2019')) As MaxHDate from SPY_ANALYTIC_UPLOAD Where Symbol in ('" + str(SYMBOL) +"') and Source in ('Historic')"
            app = xw.App
            #and Date_Time <= ('10/6/2019') 
            
            Live_date = pd.read_sql_query(SF_Live_Date, engine)
            Live_date = Live_date.copy()
            Historic_Max_date = pd.read_sql_query(SPY_Historic_Max, engine) 
            tday = datetime.now()
            print('Today: ', tday)
            print(Historic_Max_date)
            print('Live_date:')
            print(Live_date)
            Historic_Max_date = (Historic_Max_date.iloc[0])
            
            class YahooFinanceHistory:
                timeout = 2
                crumb_link = 'https://finance.yahoo.com/quote/{0}/history?p={0}'
                crumble_regex = r'CrumbStore":{"crumb":"(.*?)"}'
                quote_link = 'https://query1.finance.yahoo.com/v7/finance/download/{quote}?period1={dfrom}&period2={dto}&interval=1d&events=history&crumb={crumb}'
            
                def __init__(self, symbol, days_back=7):
                    self.symbol = symbol
                    self.session = requests.Session()
                    self.dt = timedelta(days=days_back)
            
                def get_crumb(self):   
                    response = self.session.get(self.crumb_link.format(self.symbol), timeout=self.timeout)
                    response.raise_for_status()
                    match = re.search(self.crumble_regex, response.text)
                    if not match:
                        raise ValueError('Could not get crumb from Yahoo Finance')
                    else:
                        self.crumb = match.group(1)
            
                def get_quote(self):
                    if not hasattr(self, 'crumb') or len(self.session.cookies) == 0:
                        self.get_crumb()
                    now = datetime.utcnow()
                    dateto = int(now.timestamp())
                    datefrom = int((now - self.dt).timestamp())
                    url = self.quote_link.format(quote=self.symbol, dfrom=datefrom, dto=dateto, crumb=self.crumb)
                    response = self.session.get(url)
                    response.raise_for_status()
                    return pd.read_csv(StringIO(response.text), parse_dates=['Date'])
                
            df = pd.DataFrame()
            while len(df) < 1:
                    df = YahooFinanceHistory(SYMBOL, days_back=300).get_quote()
                    df1 = df.copy()
                    df1 = df1.where(df1['Date']>=Historic_Max_date['MaxHDate']-timedelta(200)).dropna()
                    df1 = df1[['Date','Open','High','Low','Close','Adj Close','Volume']] 
                    df1['Symbol'] = SYMBOL
                    df1 = df1.sort_values(by=['Date'], ascending=True)
                    dfcalc_input = df1
                    dftarget = df1.where(df1['Date']>Historic_Max_date['MaxHDate'])
                    dftarget = dftarget.dropna()
                    dftarget_len = len(dftarget)
                    time.sleep(1)
                    print('Scrape Success')
                    time.sleep(1)
                    print('Pulled Time series Data from:')
                    print(min(dfcalc_input['Date']))
                    time.sleep(1)
                    print('Loading for Dates:')
                    print(dftarget)
                    time.sleep(1)
                    print('Target Length')
                    print(dftarget_len)
                    time.sleep(3)
                    error = error+1
            if error == 5:
                continue
            try:
                wb = xw.Book('SPY_Daily_Hist_Template.xlsm')
                spy_temp = wb.sheets['SPY_Daily_Hist']
                spy_temp.range('A1').value = dfcalc_input
                time.sleep(1)
                print('Time Series Data Pasted')
                time.sleep(3)
                #Historic_Max_date = Historic_Max_date.astype(datetime)
                New_Analytics = xw.Range('J1:AZ595').options(pd.DataFrame).value
                NW = New_Analytics.copy()
                NW2 = NW.where(NW['Symbol']==symbol)
                NW3 = NW2.dropna(thresh=10)
                NW3['Volume'].astype(int)
                NW4 = NW3.tail(40)
                time.sleep(1)
                print('Workbook Closed')
            except Exception as e:        
                print(e)
            try:
                print('Target Analytics Acquired')
                NW4.to_sql('SPY_ANALYTIC_UPLOAD', con=engine, if_exists='append')
                print('Data Loaded to SQL ANALYITC TABLE') 
                wb.app.kill()
            except Exception as e:      
                print(e)
        except Exception as e:
                print(e)
                print(SYMBOL)
                
dupe_cleanup = """WITH cte AS (
    SELECT 
        date_time, 
        Symbol, 
        source, 
        ROW_NUMBER() OVER (
            PARTITION BY 
        date_time, 
        Symbol, 
         source
            ORDER BY 
        date_time, 
        Symbol, 
        source 
        ) row_num
     FROM 
        SPY_ANALYTIC_UPLOAD
)
DELETE FROM cte
WHERE row_num > 1"""

try:
    conn.execute(dupe_cleanup)
    print('Dupe Cleanup Success')
except Exception as e:
    print(e)
    print('dupe cleanup failed')

