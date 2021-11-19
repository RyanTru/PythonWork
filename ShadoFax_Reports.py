# -*- coding: utf-8 -*-
"""
Created on Fri Apr 10 14:06:15 2020

@author: Ryan
"""

# -*- coding: utf-8 -*-
"""
Created on Sat Sep  7 13:34:23 2019

@author: Ryan
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
#'MELI','SHOP','FB','TWLO'
SYMBOL = ('ETHeUSD')
#COMP_SYMBOL = ('SPY')
#Initialize Connectho
params = urllib.parse.quote_plus(r'DRIVER={SQL Server Native Client 11.0};SERVER=localhost\SQLEXPRESS;DATABASE=DBN;Trusted_Connection=yes; timeout=300')
conn_str = 'mssql+pyodbc:///?odbc_connect={}'.format(params)
engine = create_engine(conn_str)
conn = engine.connect()
#Dates
SF_Live_Date = "select max(date_time) from ShadowFax_Live"
SPY_Historic_Max = "select max(date_time) from SPY_ANALYTIC_UPLOAD"

Live_date = pd.read_sql_query(SF_Live_Date, engine)
Historic_Max_date = pd.read_sql_query(SPY_Historic_Max, engine)


SF_HiLo = """ select Date_Time, Last, High, Low, Hi_7D, Hi_20D, Hi_90D Lo_7D, Lo_20D, Lo_90D from (
select Date_Time, Source, Time, Symbol, Last, DailyNetChng, Open_PDClose, Dollar_Chg, Bid, Ask, BidSize, AskSize, Volume, OpenCD, High, Low, High52, Low52, PD_Date, PD_NetChng, PD_Hi, PD_Lo, PD_Open, PD_Close, True_Range, ATR_7D, Hi_7D, Lo_7D, Vol_Avg_7D, ATR_20D, Hi_20D, Lo_20D, Vol_Avg_20D, Hi_90D, Lo_90D, Vol_Avg_90D
from SPY_ANALYTIC_UPLOAD where Symbol in ('""" + str(SYMBOL) + """') and Date_Time >= DATEADD(DAY, -300000, (select Date_Time from ShadowFax))) t1 order by date_time desc"""

SF_HiLo1 = """ select Date_Time, Last, High, Low, Hi_7D, Lo_7D from (
select Date_Time, Source, Time, Symbol, Last, DailyNetChng, Open_PDClose, Dollar_Chg, Bid, Ask, BidSize, AskSize, Volume, OpenCD, High, Low, High52, Low52, PD_Date, PD_NetChng, PD_Hi, PD_Lo, PD_Open, PD_Close, True_Range, ATR_7D, Hi_7D, Lo_7D, Vol_Avg_7D, ATR_20D, Hi_20D, Lo_20D, Vol_Avg_20D, Hi_90D, Lo_90D, Vol_Avg_90D
from SPY_ANALYTIC_UPLOAD where Symbol in ('""" + str(SYMBOL) + """') and Date_Time >= DATEADD(DAY, -300000, (select Date_Time from ShadowFax))) t1 order by date_time desc"""
 
SF_Volatility = """select Date_Time, ATR_7D, ATR_20D from (select Date_Time, Source, Time, Symbol, Last, DailyNetChng, Open_PDClose, Dollar_Chg, Bid, Ask, BidSize, AskSize, Volume, OpenCD, High, Low, High52, Low52, PD_Date, PD_NetChng, PD_Hi, PD_Lo, PD_Open, PD_Close, True_Range, ATR_7D, Hi_7D, Lo_7D, Vol_Avg_7D, ATR_20D, Hi_20D, Lo_20D, Vol_Avg_20D, Hi_90D, Lo_90D, Vol_Avg_90D
from ShadowFax where Symbol in  ('""" + str(SYMBOL) + """')
union
select Date_Time, Source, Time, Symbol, Last, DailyNetChng, Open_PDClose, Dollar_Chg, Bid, Ask, BidSize, AskSize, Volume, OpenCD, High, Low, High52, Low52, PD_Date, PD_NetChng, PD_Hi, PD_Lo, PD_Open, PD_Close, True_Range, ATR_7D, Hi_7D, Lo_7D, Vol_Avg_7D, ATR_20D, Hi_20D, Lo_20D, Vol_Avg_20D, Hi_90D, Lo_90D, Vol_Avg_90D
from SPY_ANALYTIC_UPLOAD where Symbol in ('""" + str(SYMBOL) + """') and Date_Time >= DATEADD(DAY, -300000, (select Date_Time from ShadowFax))) t1 order by date_time desc """

SF_COMP = """ """
 

SF_SlopeHi = """select Date_Time, HiSlope_20D, HiSlope_7D from SPY_ANALYTIC_UPLOAD where Symbol in ('""" + str(SYMBOL) + """') and Date_Time >= DATEADD(DAY, -300000, (select Date_Time from ShadowFax)) order by date_time desc """
SF_SlopeLo = """select Date_Time, LoSlope_20D, LoSlope_7D from SPY_ANALYTIC_UPLOAD where Symbol in ('""" + str(SYMBOL) + """') and Date_Time >= DATEADD(DAY, -300000, (select Date_Time from ShadowFax)) order by date_time desc """


#HiLo_Table = pd.read_sql_query(SF_COMP, engine)
#HiLo_Table0 = HiLo_Table.head(10)
#HiLo_Table0.set_index('Date_Time').plot(kind ='line', figsize=(15,10)).set_title("10 D Graph of Historic His/Los Great Indicator")
#plt.savefig('SF_COMP.png')


HiLo_Table = pd.read_sql_query(SF_HiLo1, engine)
HiLo_Table0 = HiLo_Table.head(10)
HiLo_Table0.set_index('Date_Time').plot(kind ='line', figsize=(15,10)).set_title("10 D Graph of Historic His/Los Great Indicator")
plt.savefig('HiLo_Table10.png')

HiLo_Table = pd.read_sql_query(SF_HiLo1, engine)
HiLo_Table1 = HiLo_Table.head(30)
HiLo_Table1.set_index('Date_Time').plot(kind ='line', figsize=(15,10)).set_title("30 D Graph of Historic His/Los Great Indicator")
plt.savefig('HiLo_Table30.png')

HiLo_Table = pd.read_sql_query(SF_HiLo1, engine)
HiLo_Table2 = HiLo_Table.head(60)
HiLo_Table2.set_index('Date_Time').plot(kind ='line', figsize=(15,10)).set_title("60 D Graph of Historic His/Los Great Indicator")
plt.savefig('HiLo_Table60.png')

HiLo_Table = pd.read_sql_query(SF_HiLo, engine)
HiLo_Table3 = HiLo_Table.head(90)
HiLo_Table3.set_index('Date_Time').plot(kind ='line', figsize=(15,10)).set_title("90 D Graph of Historic His/Los Great Indicator")
plt.savefig('HiLo_Table90.png')

HiLo_Table = pd.read_sql_query(SF_HiLo, engine)
HiLo_Table4 = HiLo_Table.head(300)
HiLo_Table4.set_index('Date_Time').plot(kind ='line', figsize=(15,10)).set_title("300 D Graph of Historic His/Los Great Indicator")
plt.savefig('HiLo_Table300.png')
            
Vol_Table = pd.read_sql_query(SF_Volatility, engine)
Vol_Table = Vol_Table.head(190)
Vol_Table.set_index('Date_Time').plot(figsize=(15,10)).set_title("Longer Manual Vol Measure :)")
plt.savefig('Vol_Table.png')


Slope_TableHi = pd.read_sql_query(SF_SlopeHi, engine)
Slope_TableHi = Slope_TableHi.head(30)
Slope_TableHi.set_index('Date_Time').plot(kind ='line', sharex = 'True', grid ='True', figsize=(15,10)).set_title("Slope of His")
plt.savefig('HiSlope_Table.png')


Slope_TableLo = pd.read_sql_query(SF_SlopeLo, engine)
Slope_TableLo = Slope_TableLo.head(120)
Slope_TableLo.set_index('Date_Time').plot(kind ='line', sharex = 'True', grid ='True', figsize=(15,10)).set_title("Longer Slope of los")
plt.savefig('LoSlope_Table.png')


# -*- coding: utf-8 -*-
"""
Created on Fri Dec 20 10:12:55 2019

@author: Ryan
"""

import matplotlib
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import matplotlib.dates as mdates
import numpy as np
from numpy import loadtxt
import time
from functools import reduce
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
#This version uses close prices to calculated percent change patterns - then compares the current 30 day pattern to those in the past - and shows average outcome


#Pull Historical Analytics From SQL

#Initialize Connect
params = urllib.parse.quote_plus(r'DRIVER={SQL Server Native Client 11.0};SERVER=localhost\SQLEXPRESS;DATABASE=DBN;Trusted_Connection=yes; timeout=300')
conn_str = 'mssql+pyodbc:///?odbc_connect={}'.format(params)
engine = create_engine(conn_str)
conn = engine.connect()

#Dates
SPY_Historic_Max = "select isnull(max(date_time),('10/1/2019')) As MaxHDate from SPY_ANALYTIC_UPLOAD Where Symbol in ('" + str(SYMBOL) +"') and Source in ('Historic')"
SPY_Historic_Max = pd.read_sql_query(SPY_Historic_Max, engine) 
#SPY_Historic_Max = SPY_Historic_Max.astype(datetime.datetime)
SPY_Historic_Max = SPY_Historic_Max-timedelta(100000)
SPY_Historic_Max = SPY_Historic_Max.astype(str)
SPY_Historic_Max = SPY_Historic_Max['MaxHDate'].iloc[0]

SF_Data = "select date_time, hislope_7d, loslope_7d, Last from SPY_ANALYTIC_UPLOAD where date_time > ('"+SPY_Historic_Max+"') and Symbol in ('" + str(SYMBOL) +"') order by date_time asc"


SF_Hist = pd.read_sql_query(SF_Data, engine)
SF_Hist = SF_Hist.copy()
SF_Last = np.array(SF_Hist['Last'])
SF_HiSlope = np.array(SF_Hist['hislope_7d'])

#print(SF_Hist)


avgLine = (SF_Last)

patternAr = []
performanceAr = []
patForRec = []


def percentChange(startPoint,currentPoint):
    try:
        x = ((float(currentPoint)-startPoint)/abs(startPoint))*100.00
        if x == 0.0:
            return 0.000000001
        else:
            return x
    except:
        return 0.0001

def patternStorage():
    '''
    The goal of patternFinder is to begin collection of %change patterns
    in the tick data. From there, we also collect the short-term outcome
    of this pattern. Later on, the length of the pattern, how far out we
    look to compare to, and the length of the compared range be changed,
    and even THAT can be machine learned to find the best of all 3 by
    comparing success rates.'''

    startTime = time.time()
    
    
    x = len(avgLine)-30
    y = 31
    currentStance = 'none'
    
    while y < x:
        pattern = []
        
        p1 = percentChange(avgLine[y-30], avgLine[y-29])
        p2 = percentChange(avgLine[y-30], avgLine[y-28])
        p3 = percentChange(avgLine[y-30], avgLine[y-27])
        p4 = percentChange(avgLine[y-30], avgLine[y-26])
        p5 = percentChange(avgLine[y-30], avgLine[y-25])
        p6 = percentChange(avgLine[y-30], avgLine[y-24])
        p7 = percentChange(avgLine[y-30], avgLine[y-23])
        p8 = percentChange(avgLine[y-30], avgLine[y-22])
        p9 = percentChange(avgLine[y-30], avgLine[y-21])
        p10= percentChange(avgLine[y-30], avgLine[y-20])
		
        p11 = percentChange(avgLine[y-30], avgLine[y-19])
        p12 = percentChange(avgLine[y-30], avgLine[y-18])
        p13 = percentChange(avgLine[y-30], avgLine[y-17])
        p14 = percentChange(avgLine[y-30], avgLine[y-16])
        p15 = percentChange(avgLine[y-30], avgLine[y-15])
        p16 = percentChange(avgLine[y-30], avgLine[y-14])
        p17 = percentChange(avgLine[y-30], avgLine[y-13])
        p18 = percentChange(avgLine[y-30], avgLine[y-12])
        p19 = percentChange(avgLine[y-30], avgLine[y-11])
        p20= percentChange(avgLine[y-30], avgLine[y-10])
		
        p21 = percentChange(avgLine[y-30], avgLine[y-9])
        p22 = percentChange(avgLine[y-30], avgLine[y-8])
        p23 = percentChange(avgLine[y-30], avgLine[y-7])
        p24 = percentChange(avgLine[y-30], avgLine[y-6])
        p25 = percentChange(avgLine[y-30], avgLine[y-5])
        p26 = percentChange(avgLine[y-30], avgLine[y-4])
        p27 = percentChange(avgLine[y-30], avgLine[y-3])
        p28 = percentChange(avgLine[y-30], avgLine[y-2])
        p29 = percentChange(avgLine[y-30], avgLine[y-1])
        p30= percentChange(avgLine[y-30], avgLine[y])


        outcomeRange = avgLine[y+20:y+30]
        currentPoint = avgLine[y]

        try:
            avgOutcome = reduce(lambda x, y: x + y, outcomeRange) / len(outcomeRange)
        except Exception as e:
            print (e)
            avgOutcome = 0
        futureOutcome = percentChange(currentPoint, avgOutcome)

        '''
        print 'where we are historically:',currentPoint
        print 'soft outcome of the horizon:',avgOutcome
        print 'This pattern brings a future change of:',futureOutcome
        print '_______'
        print p1, p2, p3, p4, p5, p6, p7, p8, p9, p10
        '''

        pattern.append(p1)
        pattern.append(p2)
        pattern.append(p3)
        pattern.append(p4)
        pattern.append(p5)
        pattern.append(p6)
        pattern.append(p7)
        pattern.append(p8)
        pattern.append(p9)
        pattern.append(p10)

        pattern.append(p11)
        pattern.append(p12)
        pattern.append(p13)
        pattern.append(p14)
        pattern.append(p15)
        pattern.append(p16)
        pattern.append(p17)
        pattern.append(p18)
        pattern.append(p19)
        pattern.append(p20)

        pattern.append(p21)
        pattern.append(p22)
        pattern.append(p23)
        pattern.append(p24)
        pattern.append(p25)
        pattern.append(p26)
        pattern.append(p27)
        pattern.append(p28)
        pattern.append(p29)
        pattern.append(p30)



        patternAr.append(pattern)
        performanceAr.append(futureOutcome)
        
        y+=1

    endTime = time.time()
    print (len(patternAr))
    print (len(performanceAr))
    print ('Pattern storing took:', endTime-startTime)

def currentPattern():
    mostRecentPoint = avgLine[-1]

    

    cp1 = percentChange(avgLine[-31],avgLine[-30])
    cp2 = percentChange(avgLine[-31],avgLine[-29])
    cp3 = percentChange(avgLine[-31],avgLine[-28])
    cp4 = percentChange(avgLine[-31],avgLine[-27])
    cp5 = percentChange(avgLine[-31],avgLine[-26])
    cp6 = percentChange(avgLine[-31],avgLine[-25])
    cp7 = percentChange(avgLine[-31],avgLine[-24])
    cp8 = percentChange(avgLine[-31],avgLine[-23])
    cp9 = percentChange(avgLine[-31],avgLine[-22])
    cp10= percentChange(avgLine[-31],avgLine[-21])
	
	
    cp11 = percentChange(avgLine[-31],avgLine[-20])
    cp12 = percentChange(avgLine[-31],avgLine[-19])
    cp13 = percentChange(avgLine[-31],avgLine[-18])
    cp14 = percentChange(avgLine[-31],avgLine[-17])
    cp15 = percentChange(avgLine[-31],avgLine[-16])
    cp16 = percentChange(avgLine[-31],avgLine[-15])
    cp17 = percentChange(avgLine[-31],avgLine[-14])
    cp18 = percentChange(avgLine[-31],avgLine[-13])
    cp19 = percentChange(avgLine[-31],avgLine[-12])
    cp20= percentChange(avgLine[-31],avgLine[-11])
	
    cp21 = percentChange(avgLine[-31],avgLine[-10])
    cp22 = percentChange(avgLine[-31],avgLine[-9])
    cp23 = percentChange(avgLine[-31],avgLine[-8])
    cp24 = percentChange(avgLine[-31],avgLine[-7])
    cp25 = percentChange(avgLine[-31],avgLine[-6])
    cp26 = percentChange(avgLine[-31],avgLine[-5])
    cp27 = percentChange(avgLine[-31],avgLine[-4])
    cp28 = percentChange(avgLine[-31],avgLine[-3])
    cp29 = percentChange(avgLine[-31],avgLine[-2])
    cp30= percentChange(avgLine[-31],avgLine[-1])

    patForRec.append(cp1)
    patForRec.append(cp2)
    patForRec.append(cp3)
    patForRec.append(cp4)
    patForRec.append(cp5)
    patForRec.append(cp6)
    patForRec.append(cp7)
    patForRec.append(cp8)
    patForRec.append(cp9)
    patForRec.append(cp10)
    patForRec.append(cp11)
    patForRec.append(cp12)
    patForRec.append(cp13)
    patForRec.append(cp14)
    patForRec.append(cp15)
    patForRec.append(cp16)
    patForRec.append(cp17)
    patForRec.append(cp18)
    patForRec.append(cp19)
    patForRec.append(cp20)
    patForRec.append(cp21)
    patForRec.append(cp22)
    patForRec.append(cp23)
    patForRec.append(cp24)
    patForRec.append(cp25)
    patForRec.append(cp26)
    patForRec.append(cp27)
    patForRec.append(cp28)
    patForRec.append(cp29)
    patForRec.append(cp30)



def graphRawFX():
    
    fig=plt.figure(figsize=(10,7))
    ax1 = plt.subplot2grid((40,40), (0,0), rowspan=40, colspan=40)
    ax1.plot(date,bid)
    ax1.plot(date,ask)
    ax1.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d %H:%M:%S'))
    plt.grid(True)
    for label in ax1.xaxis.get_ticklabels():
            label.set_rotation(45)
    plt.gca().get_yaxis().get_major_formatter().set_useOffset(False)
    ax1_2 = ax1.twinx()
    ax1_2.fill_between(date, 0, (ask-bid), facecolor='g',alpha=.3)

    plt.subplots_adjust(bottom=.23)
    plt.show()



    
def patternRecognition():
    ######
    predictedOutcomesAr = []
    ######
    plotPatAr = []
    patFound = 0

    for eachPattern in patternAr:
        sim1 = 100.00 - abs(percentChange(eachPattern[0], patForRec[0]))
        sim2 = 100.00 - abs(percentChange(eachPattern[1], patForRec[1]))
        sim3 = 100.00 - abs(percentChange(eachPattern[2], patForRec[2]))
        sim4 = 100.00 - abs(percentChange(eachPattern[3], patForRec[3]))
        sim5 = 100.00 - abs(percentChange(eachPattern[4], patForRec[4]))
        sim6 = 100.00 - abs(percentChange(eachPattern[5], patForRec[5]))
        sim7 = 100.00 - abs(percentChange(eachPattern[6], patForRec[6]))
        sim8 = 100.00 - abs(percentChange(eachPattern[7], patForRec[7]))
        sim9 = 100.00 - abs(percentChange(eachPattern[8], patForRec[8]))
        sim10 = 100.00 - abs(percentChange(eachPattern[9], patForRec[9]))
            
        sim11 = 100.00 - abs(percentChange(eachPattern[10], patForRec[10]))
        sim12 = 100.00 - abs(percentChange(eachPattern[11], patForRec[11]))
        sim13 = 100.00 - abs(percentChange(eachPattern[12], patForRec[12]))
        sim14 = 100.00 - abs(percentChange(eachPattern[13], patForRec[13]))
        sim15 = 100.00 - abs(percentChange(eachPattern[14], patForRec[14]))
        sim16 = 100.00 - abs(percentChange(eachPattern[15], patForRec[15]))
        sim17 = 100.00 - abs(percentChange(eachPattern[16], patForRec[16]))
        sim18 = 100.00 - abs(percentChange(eachPattern[17], patForRec[17]))
        sim19 = 100.00 - abs(percentChange(eachPattern[18], patForRec[18]))
        sim20 = 100.00 - abs(percentChange(eachPattern[19], patForRec[19]))
            
        sim21 = 100.00 - abs(percentChange(eachPattern[20], patForRec[20]))
        sim22 = 100.00 - abs(percentChange(eachPattern[21], patForRec[21]))
        sim23 = 100.00 - abs(percentChange(eachPattern[22], patForRec[22]))
        sim24 = 100.00 - abs(percentChange(eachPattern[23], patForRec[23]))
        sim25 = 100.00 - abs(percentChange(eachPattern[24], patForRec[24]))
        sim26 = 100.00 - abs(percentChange(eachPattern[25], patForRec[25]))
        sim27 = 100.00 - abs(percentChange(eachPattern[26], patForRec[26]))
        sim28 = 100.00 - abs(percentChange(eachPattern[27], patForRec[27]))
        sim29 = 100.00 - abs(percentChange(eachPattern[28], patForRec[28]))
        sim30 = 100.00 - abs(percentChange(eachPattern[29], patForRec[29]))



        
        howSim = (sim1+sim2+sim3+sim4+sim5+sim6+sim7+sim8+sim9+sim10
                  +sim11+sim12+sim13+sim14+sim15+sim16+sim17+sim18+sim19+sim20
                  +sim21+sim22+sim23+sim24+sim25+sim26+sim27+sim28+sim29+sim30)/30.00

        if howSim > 40:
            
            patdex = patternAr.index(eachPattern)
            patFound = 1
            xp = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30]
            #############

            plotPatAr.append(eachPattern)
        
    predArray = []
    if patFound == 1:
        fig = plt.figure(figsize=(15,10)).suptitle('Pattern REC >30% 1993 Pattern Lookback')
        plt.savefig('PatternRec_ShortTerm.png')

        for eachPatt in plotPatAr:
            futurePoints = patternAr.index(eachPatt)

            if performanceAr[futurePoints] > patForRec[29]:
                pcolor = '#24bc00'
                #####
                predArray.append(1.000)
            else:
                pcolor = '#d40000'
                #######
                predArray.append(-1.000)
            
            #plt.figure(figsize=(10,6))
            plt.plot(xp, eachPatt)
            predictedOutcomesAr.append(performanceAr[futurePoints])
            plt.scatter(35, performanceAr[futurePoints],c=pcolor,alpha=.4)


            
        
        predictedAvgOutcome = reduce(lambda x, y: x + y, predictedOutcomesAr) / len(predictedOutcomesAr)

        #######
      #  print (predArray)
        predictionAverage = reduce(lambda x, y: x + y, predArray) / len(predArray)
        print (predictionAverage)
        if predictionAverage < 0:
            print ('drop predicted')
            print (patForRec[29])
                
        if predictionAverage > 0:
            print ('rise predicted')
            print (patForRec[29])
                
        #######
        plt.scatter(40, predictedAvgOutcome, c='#008db8',s=25)
        plt.plot(xp, patForRec, '#54fff7', linewidth = 10)
        #plt.grid(True)
        #plt.title('Pattern Recognition.\nCyan line is the current pattern. Other lines are similar patterns from the past.\nPredicted outcomes are color-coded to reflect a positive or negative prediction.\nThe Cyan dot marks where the pattern went.\nOnly data in the past is used to generate patterns and predictions.')
        plt.savefig('PatternRec_ShortTerm.png')
        plt.show()

        
            
            


allData = (SF_Last)
toWhat = 6683
######
accuracyArray = []

####
samps = 0

patternStorage()
currentPattern()
patternRecognition()


# -*- coding: utf-8 -*-
"""
Created on Sat Aug 31 16:53:47 2019

@author: Ryan
"""

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.text import MIMEText

import datetime as dt
contacts = ["r.trudeau.capesun@gmail.com"] #,"cwalkley25@gmail.com","whmartin1823@gmail.com"]

for x in range(1):
    
    msgs = MIMEMultipart() 
    
    me = "gfcapitalrt@gmail.com"
    my_password = r""
    you = contacts[x]
    subj = "RYs Custom SPYs " + dt.datetime.today().strftime("%m.%d.%Y")

    msg = MIMEMultipart('alternative')
    msg['Subject'] = subj
    msg['From'] = me
    msg['To'] = you
    
    body =(r'C:\Users\Ryan\Desktop\PY\MASTERFILES-DONTRUNFROM-DONTFUCKWITH\'HiLo_Table.png')
    body1 = """\
<html>
  <head>CORRECTED - SLOPE GRAPHS WERE NOT EMAILING OUT CORRECTLY -  """ + str(SYMBOL) + """
 """+ dt.datetime.today().strftime("%m.%d.%Y") +"""  </head> 
  <body>
    <p>
    </p>
  </body>
</html>
"""
    fp = open('HiLo_Table10.png', 'rb')
    msgImage = MIMEImage(fp.read())
    fp.close()
    msgs.attach(msgImage)
    fp = open('HiLo_Table30.png', 'rb')
    msgImage = MIMEImage(fp.read())
    fp.close()
    msgs.attach(msgImage)
    fp = open('HiLo_Table60.png', 'rb')
    msgImage = MIMEImage(fp.read())
    fp.close()
    msgs.attach(msgImage)
    fp = open('HiLo_Table90.png', 'rb')
    msgImage = MIMEImage(fp.read())
    fp.close()
    msgs.attach(msgImage)
    fp = open('HiLo_Table300.png', 'rb')
    msgImage = MIMEImage(fp.read())
    fp.close()
    msgs.attach(msgImage)
    fp = open('LoSlope_Table.png', 'rb')
    msgImage = MIMEImage(fp.read())
    fp.close()
    msgs.attach(msgImage)
    fp = open('HiSlope_Table.png', 'rb')
    msgImage = MIMEImage(fp.read())
    fp.close()
    msgs.attach(msgImage)
    fp = open('Vol_Table.png', 'rb')
    msgImage = MIMEImage(fp.read())
    fp.close()
    msgs.attach(msgImage)
    fp = open('PatternRec_ShortTerm.png', 'rb')
    msgImage = MIMEImage(fp.read())
    fp.close()
    msgs.attach(msgImage)
    body1 = MIMEText(body1, 'text')
    msgs.attach(body1)

    #part2 = MIMEText(html, 'html')
    
    fp = 0
    #msg.attach(part2)
    

    
    # Send the message via gmail's regular server, over SSL - passwords are being sent, afterall
    s = smtplib.SMTP_SSL('smtp.gmail.com')
    # uncomment if interested in the actual smtp conversation
    # s.set_debuglevel(1)
    # do the smtp auth; sends ehlo if it hasn't been sent already
    s.login(me, my_password)
    
    s.sendmail(me, you, msgs.as_string())
    s.quit() 
