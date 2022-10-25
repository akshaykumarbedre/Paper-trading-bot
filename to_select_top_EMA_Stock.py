import time
import statistics
import math
from rich import print
import xlwings as xw
import pandas as pd
import yfinance as yf
ws = xw.Book(f"information.xlsx")
SheetList=ws.sheets[0]
SheetPair=ws.sheets[2]

class Base:
    def PreProccesBase(self,pre_rate):
        self.rate = [pre_rate]
        self.rate_per = [1.0]
        self.no_of_rate=(len(self.rate))
        self.DataIndex=1
        self.OldSignal=1.0

class EmaProcces(Base):
    def PreProccesEma(self, s, l, lv,pre_rate):
        self.dic={}
        self.S_time_per = s
        self.L_time_per = l
        self.VL_time_per = lv   
        self.dic["VL_ema"] = self.dic["S_ema"] = self.dic["L_ema"] = pre_rate  
        self.VL_time_list=[pre_rate]
        self.VL_rate_list=[1]
        self.VL_diff_list=[pre_rate]

        
    def UpdateEma(self):
        self.dic["S_ema"] = self.CalEma(self.no_of_rate, self.S_time_per, self.dic["S_ema"])
        self.dic["L_ema"] = self.CalEma(self.no_of_rate, self.L_time_per, self.dic["L_ema"])
        self.dic["VL_ema"]=self.CalEma(self.no_of_rate, self.VL_time_per, self.dic["VL_ema"])
        self.VL_time_list.append(self.dic["VL_ema"]) 
        self.VL_diff_list.append(self.VL_time_list[-1]-self.VL_time_list[-2])
        self.VL_rate_list.append(self.VL_time_list[-1]/self.VL_time_list[-2])

    def CalEma(self, n, time_per, pre_ema):  # it caluctae EMA of given time perdoid(time_per)
        # multi=[2 ÷ (number of observations + 1)]
        multiplier = (2/(time_per+1))

        if n < time_per:
            ema = statistics.mean(self.rate[-n:])
        else:
            # EMA = Closing price x multiplier + EMA (previous day) x (1-multiplier)
            ema = ((self.rate[-1])*multiplier) + (pre_ema * (1-multiplier))
        return ema
    
class InformationToStore():
    def PreProccesInformationToStore(self):
        self.sheet=ws.sheets[1]
        self.cell=1
        
    def UpdateInformation(self,cur_rate,dic,VL_Dif,wt_to_do,after_buy,after_sell):
        self.sheet.range(f"a{self.cell}").value=[cur_rate,dic["S_ema"],dic["L_ema"],dic["VL_ema"],VL_Dif,wt_to_do,after_buy,after_sell]
        (self.cell)+=1
        

    def FinalList(self,VL_rate_list,row,lv):
        SheetList.range(f"o{row}").value=VL_rate_list[-1],(statistics.mean(VL_rate_list[-(lv):]))

    
class ProccessList(EmaProcces,Base,InformationToStore):
    def __init__(self,comapny,interval,period,start,end,s,l,lv,row):
        self.row=row 

        try:
            self.historical=yf.download(tickers=comapny,interval=interval,start=start,period=period,end=end)
        except:
            return 
        self.li = self.historical.values.tolist()
        self.pre_rate = float(self.li[0][3]) 
        super().PreProccesBase(self.pre_rate)
        super().PreProccesEma(s,l,lv,self.pre_rate)
        
    def procces(self):
        for i in range(0,len(self.li)):
            try:
                self.cur_rate=float(self.li[i][3])
            except: 
                return			 
            self.rate.append(self.cur_rate)
            self.no_of_rate=(len(self.rate))

            self.rate_per.append(self.cur_rate/self.pre_rate)
            self.UpdateEma()#to update new ema value 

            # self.Signal=self.dic["S_ema"]/self.dic["L_ema"] ***# self.wt_to_do=self.pre_load(self.no_of_rate,self.Signal) ***# no_of_rate_after_buying = len(self.after_buying_rate) ***# self.after_buying(self.cur_rate,no_of_rate_after_buying)
                        
            self.pre_rate=self.cur_rate     

        self.FinalList(self.VL_rate_list,self.row,lv)
     
    
interval = "1h"     # TIME INTERVAL FOR TESTING              #interval : str: ->1m,2m,5m,15m,30m,60m,90m,1h,1d,5d,1wk,1mo,3mo 
time_peroid=None    #period : str:-> 1d,5d,1mo,3mo,6mo,1y,2y,5y,10y,ytd,max 
start="2022-09-22"  #start: str:->Download start date string (YYYY-MM-DD) Default is 1900-01-01
end=None            #end: str->Download end date string (YYYY-MM-DD) Default is now
s=7             #THIS IS FOR SHORT TIMR EMA
l=21            #THIS IS FOR LONG EMA
lv=12           #NUMBER TO CALCUCATE TOP STOCK WHICH HAVING HIGH EMA(SLOAP)


# ***************for stock selction*******************
for i in range(2,199):#start from 2
    comapny=SheetList.range(f"B{i}").value
    print(comapny)
    s1 = ProccessList(comapny, interval,time_peroid,start,end,s,l,lv,i)
    s1.procces()
    print(i/200*100)