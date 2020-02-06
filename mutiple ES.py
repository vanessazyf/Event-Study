# -*- coding: utf-8 -*-
"""
Created on Tue Apr 30 10:35:54 2019

"""

# Importing the libraries
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.ticker import FormatStrFormatter
import statsmodels.api as sm  

#Define Dictionary TICKER = 'event_date'

# =============================================================================
# D = dict(PI='5/08/2018')
# =============================================================================
#Temporary
# =============================================================================
# D = dict(ADTN='4/18/2018',       
#          ESV='4/26/2018',
#          MTSC='5/08/2018',         
#          MNTA='5/08/2018',            
#          APO='5/03/2018',
#          FMS='5/03/2018',         
#          TPX='5/03/2018',
#          DRQ='4/27/2018',
#          INCY='5/01/2018')
# =============================================================================

#Permanant
# =============================================================================
# D = dict(GILD='5/2/2018',
#          HHC='5/2/2018',         
#          DBD='5/2/2018',
#          TAP='5/2/2018',
#          ANIK='5/3/2018',
#          GIL='5/3/2018',
#          SNBR='4/19/2018',
#          FE='4/23/2018',         
#          TRVG='4/25/2018',
#          AWR='5/08/2018',
#          EQC='5/08/2018',
#          ORA='5/08/2018',
#          MLM='5/08/2018',
#          MGI='5/08/2018',
#          SMP='5/03/2018',
#          PRA='5/04/2018',
#          RPT='5/04/2018',
#          IMGN='5/04/2018',
#          SRE='5/07/2018',
#          TSEM='5/07/2018',
#          CTB='4/30/2018',
#          VSTO='5/01/2018',
#          BLDP='5/02/2018')
# =============================================================================

#All
D = dict(GILD='5/2/2018',
         HHC='5/2/2018',
         ZVO='5/2/2018',
         AVA='5/2/2018',
         BPMC='5/2/2018',
         DBD='5/2/2018',
         TAP='5/2/2018',
         ANIK='5/3/2018',
         CRUS='5/3/2018',
         GIL='5/3/2018',
         ADTN='4/18/2018',
         SNBR='4/19/2018',
         FE='4/23/2018',
         HAS='4/23/2018',
         PHG='4/23/2018',
         TRVG='4/25/2018',
         ESV='4/26/2018',
         OII='4/26/2018',
         ASPS='4/26/2018',
         GNC='4/26/2018',
         AWR='5/08/2018',
         BKD='5/08/2018',
         EQC='5/08/2018',
         FRGI='5/08/2018',
         PI='5/08/2018',
         MTSC='5/08/2018',
         ORA='5/08/2018',
         MLM='5/08/2018',
         MGI='5/08/2018',
         MNTA='5/08/2018',            
         APO='5/03/2018',
         FMS='5/03/2018',
         SMP='5/03/2018',
         TPX='5/03/2018',
         PRA='5/04/2018',
         RPT='5/04/2018',
         SRG='5/04/2018',
         IMGN='5/04/2018',
         SRE='5/07/2018',
         TSEM='5/07/2018',
         NOK='4/26/2018',
         CBL='4/27/2018',
         DRQ='4/27/2018',
         MAT='4/27/2018',
         CTB='4/30/2018',
         RCII='5/01/2018',
         INCY='5/01/2018',
         VSTO='5/01/2018',
         BLDP='5/02/2018',
         AAOI='5/09/2018')



# Global Market Data
MrktData = pd.read_excel('SampleStocks.xlsx', sheet_name='S&P',index_col=0,usecols=['Date','Close'])
MrktReturn = np.log(MrktData.Close/MrktData.Close.shift(1)).dropna()

def Multiple_Event_Study(file_name,stock_name,event_date): 
    """
    file_name = 'SampleStocks.xlsx'
    stock_name = 'AWR' # As an example
    event_date = '5/8/2018'
    """
    t0 = event_date # Event Date
    StockData = pd.read_excel(file_name, sheet_name=stock_name,index_col=0,usecols=['Date','Close'])
    StockReturn = np.log(StockData.Close/StockData.Close.shift(1)).dropna()
    Y = StockReturn.iloc[StockReturn.index.get_loc(t0)-250:StockReturn.index.get_loc(t0)-10]
    X = MrktReturn.iloc[MrktReturn.index.get_loc(t0)-250:MrktReturn.index.get_loc(t0)-10]
    X = sm.add_constant(X)
    res = sm.OLS(Y,X).fit()

    Y1 = res.params[0] + res.params[1] * MrktReturn.iloc[MrktReturn.index.get_loc(t0)-10:MrktReturn.index.get_loc(t0)+11].values
    AR1 = StockReturn.iloc[StockReturn.index.get_loc(t0)-10:StockReturn.index.get_loc(t0)+11].values - Y1
    
    Y_hat = res.params[0] + res.params[1] * X.Close.values
    var = np.sum((Y_hat - Y)**2)/(240-2)

    se_ar = []
    for i in range(MrktReturn.index.get_loc(t0)-10,MrktReturn.index.get_loc(t0)+11):
        r_bar = X.Close.mean()
        sigma = np.sqrt(var)
        se_ar.append(sigma  * np.sqrt(1+1/240+(MrktReturn.iloc[i]-r_bar)**2/np.sum((X.Close.values - r_bar)**2)))
        
    se_car = []
    for j in range(-10,11):
        sigma = np.sqrt(var)
        se_car.append(sigma * np.sqrt(j+11))
    
    return AR1,np.array(se_ar)**2

event_info_1 = []
event_info_2 = []

for k,v in D.items():
    a, b = Multiple_Event_Study('SampleStocks.xlsx',k,v)
    event_info_1.append(a) #+= a
    event_info_2.append(b) #+= b

event_info1 = np.sum(event_info_1,axis=0)
event_info2 = np.sum(event_info_2,axis=0)
    
AAR = event_info1/len(D)
AAR_var = event_info2/len(D)**2
AAR_t = AAR/np.sqrt(AAR_var)

x = np.arange(-10, 11, 1)
y1 = AAR
y2 = AAR_t
fig = plt.figure()

ax1 = fig.add_subplot(111)
ax1.plot(x, y1,'r',label= 'AR')
ax1.legend(loc = 'lower left')
ax1.plot(x, y1, 'r*')
ax1.grid(True)
ax1.xaxis.set_ticks(np.arange(-10, 11, 2))
ax1.set_ylabel('Average Abnormal Return (%)')
ax1.yaxis.set_major_formatter(FormatStrFormatter('%.2f'))
ax1.set_title(" Average Abnormal Returns and t-Stats" )

ax2 = ax1.twinx()
ax2.plot(x, y2,'b',label='AAR t-score')
ax2.set_ylabel('AAR t-score')
ax2.yaxis.set_major_formatter(FormatStrFormatter('%.1f'))
ax2.legend()


CAAR = np.cumsum(AAR)
CAAR_var = np.cumsum(AAR_var)
CAAR_t=CAAR/np.sqrt(CAAR_var)


x = np.arange(-10, 11, 1)
y1 = CAAR
y2 = CAAR_t
fig = plt.figure()

ax1 = fig.add_subplot(111)
ax1.plot(x, y1,'r',label= 'CAAR')
ax1.legend(loc = 'lower left')
ax1.plot(x, y1, 'r*')
ax1.grid(True)
ax1.xaxis.set_ticks(np.arange(-10, 11, 2))
ax1.set_ylabel('Cumulative Average Abnormal Return (%)')
ax1.yaxis.set_major_formatter(FormatStrFormatter('%.2f'))
ax1.set_title(" Cumulative Average Abnormal Returns and t-Stats")

ax2 = ax1.twinx()
ax2.plot(x, y2,'b',label='CAAR t-score')
ax2.set_ylabel('CAAR t-score')
ax2.yaxis.set_major_formatter(FormatStrFormatter('%.1f'))
ax2.legend()


