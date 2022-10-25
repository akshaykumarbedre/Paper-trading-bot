#this will cheak best pair of EMA for ag given time peroid & for given stock 
import time
import statistics
import math
from rich import print
import xlwings as xw
import pandas as pd
import yfinance as yf
ws = xw.Book(f"information.xlsx")
SheetPair=ws.sheets[2] #this is on 3rd page of excel

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
     
class BuySell:    
    def PreProccesBuySell(self):
        self.after_buy = "waiting"
        self.after_sell="waiting"
        self.wt_to_do = "None"
        self.f=0
        self.profit=[1]
        self.after_buying_rate=[]
        self.after_buy_rate_per=[1.0]
        self.pre_intersect=1
        self.after_selling_rate=[]
        self.after_sell_rate_per=[1.0]

    def pre_load(self,n,signal):
        if signal<=1  or self.f == 1 : #to prevent error while running code and become perfrect {start from "None" to get accurate}
            wt_to_do=self.cheak(signal)
            self.f=1
            return wt_to_do

    def cheak(self,intersect):
        wt_to_do=self.wt_to_do

        if(self.pre_intersect<1 and intersect>1):
            wt_to_do="Buy"
        if(self.pre_intersect>1 and intersect>1):
            wt_to_do="Hold"
        if(self.pre_intersect>1 and intersect<1):
            wt_to_do="Sell"
        if(self.pre_intersect<1 and intersect<1):
            wt_to_do="None"

        self.pre_intersect=intersect        
        # print(wt_to_do)        
        return wt_to_do

    def after_buying(self,b,n):
        if self.wt_to_do == "Buy":
            self.after_buying_rate.append(b)
            # print(f"BR = {math.prod(self.after_buy_rate_per)*100}")
            self.after_buy=f"	 BR =     {math.prod(self.after_buy_rate_per)*100}"	
            self.profit.append(math.prod(self.after_buy_rate_per))

        elif self.wt_to_do =="Hold":	
            self.profit.pop()
            self.after_buying_rate.append(b)
            per_buy=self.after_buying_rate[n]/self.after_buying_rate[n-1]
            self.after_buy_rate_per.append(per_buy)
            # print(f"BR = {math.prod(self.after_buy_rate_per)*100}")
            self.after_buy=f"	 BR =     {math.prod(self.after_buy_rate_per)*100}"
            self.profit.append(math.prod(self.after_buy_rate_per))

        elif self.wt_to_do =="Sell":
            self.profit.pop()
            self.after_buying_rate.append(b)
            per_buy=self.after_buying_rate[n]/self.after_buying_rate[n-1]
            self.after_buy_rate_per.append(per_buy)
            # print(f"BP= {math.prod(self.after_buy_rate_per)*100}")
            self.after_buy=(f"	BP= 	 {math.prod(self.after_buy_rate_per)*100}")
            self.after_buying_rate=[]
            self.profit.append(math.prod(self.after_buy_rate_per))
            self.after_buy_rate_per=[1.0]

        elif self.wt_to_do=="None":
            self.after_buy="waiting"

class InformationToStore():
    def PreProccesInformationToStore(self):
        self.sheet=ws.sheets[1]
        self.cell=1
        
    def UpdateInformation(self,cur_rate,dic,VL_Dif,wt_to_do,after_buy,after_sell):
        self.sheet.range(f"a{self.cell}").value=[cur_rate,dic["S_ema"],dic["L_ema"],dic["VL_ema"],VL_Dif,wt_to_do,after_buy,after_sell]
        (self.cell)+=1
        
    def FinalPair(self,profit,row,s,l):
        SheetPair.range(f"a{row}").value=f"{s,l}",math.prod(profit),len(profit),l/s

class ProccessPair(EmaProcces,Base,BuySell,InformationToStore):
    def __init__(self,historical,s,l,lv,row):
        self.row=row                  
        self.li = historical

        self.pre_rate = float(self.li[0][3]) 
       
        super().PreProccesBase(self.pre_rate)
        super().PreProccesEma(s,l,lv,self.pre_rate)
        super().PreProccesBuySell()
        super().PreProccesInformationToStore()

    def procces(self):
        for i in range(1,len(self.li)):
            try:
                self.cur_rate=float(self.li[i][3])
                pass
            except: 
                pass
        			 
            self.rate.append(self.cur_rate)
            self.no_of_rate=(len(self.rate))

            self.rate_per.append(self.cur_rate/self.pre_rate)
            self.UpdateEma()#to update new ema value 

            self.Signal=self.dic["S_ema"]/self.dic["L_ema"] 
            self.wt_to_do=self.pre_load(self.no_of_rate,self.Signal)
                    
            no_of_rate_after_buying = len(self.after_buying_rate)
            self.after_buying(self.cur_rate,no_of_rate_after_buying)

            self.pre_rate=self.cur_rate
            # self.DataIndex+=1
            
        #after proccese
        self.FinalPair(self.profit,self.row,self.S_time_per,self.L_time_per)

comapny = "BAJFINANCE.NS" #company selction this avavle on 1st page of excel
interval = "1d"     # TIME INTERVAL FOR TESTING              #interval : str: ->1m,2m,5m,15m,30m,60m,90m,1h,1d,5d,1wk,1mo,3mo 
time_peroid=None    #period : str:-> 1d,5d,1mo,3mo,6mo,1y,2y,5y,10y,ytd,max 
start="2022-09-22"  #start: str:->Download start date string (YYYY-MM-DD) Default is 1900-01-01
end=None            #end: str->Download end date string (YYYY-MM-DD) Default is now
s=7             #THIS IS FOR SHORT TIMR EMA
l=21            #THIS IS FOR LONG EMA
lv=12           #NUMBER TO CALCUCATE TOP STOCK WHICH HAVING HIGH EMA(SLOAP)           

# # ************for optimal pair *****************
no=2#start storeing  cell no 2 
historical=yf.download(tickers=comapny,interval=interval,start=start,period=time_peroid,end=end).values.tolist()
for i in range(1,50):
    for j in range(1,i):
        s1=ProccessPair(historical,j,i,lv,no)
        s1.procces()
        no+=1
        print(no)
#this produe combination of Ema pair which can be sort in excel file on 3rd page 