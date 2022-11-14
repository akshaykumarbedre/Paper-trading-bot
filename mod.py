import xlwings as xw
import pandas as pd
import math
import numpy as np
from talib.abstract import *
from talib import abstract
ws = xw.Book("INFY.csv")
SheetList=ws.sheets[0]

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

    def cheak(self,intersect): #to prevent error while running code and become perfrect {start from "None" to get accurate}
        if intersect<=1  or self.f == 1 : 
            if(self.pre_intersect<1 and intersect>1):
                self.wt_to_do="Buy"
            if(self.pre_intersect>1 and intersect>1):
                self.wt_to_do="Hold"
            if(self.pre_intersect>1 and intersect<1):
                self.wt_to_do="Sell"
            if(self.pre_intersect<1 and intersect<1):
                self.wt_to_do="None"
            self.pre_intersect=intersect 
        self.f=1
        print(self.wt_to_do)

    def after_buying(self,b):
        n=len(self.after_buying_rate)
        if self.wt_to_do == "Buy":
            self.after_buying_rate.append(b)
            print(f"BR = {math.prod(self.after_buy_rate_per)*100}")
            self.after_buy=f"	 BR =     {math.prod(self.after_buy_rate_per)*100}"	
            self.profit.append(math.prod(self.after_buy_rate_per))

        elif self.wt_to_do =="Hold":	
            self.profit.pop()
            self.after_buying_rate.append(b)
            per_buy=self.after_buying_rate[n]/self.after_buying_rate[n-1]
            self.after_buy_rate_per.append(per_buy)
            print(f"BR = {math.prod(self.after_buy_rate_per)*100}")
            self.after_buy=f"	 BR =     {math.prod(self.after_buy_rate_per)*100}"
            self.profit.append(math.prod(self.after_buy_rate_per))

        elif self.wt_to_do =="Sell":
            self.profit.pop()
            self.after_buying_rate.append(b)
            per_buy=self.after_buying_rate[n]/self.after_buying_rate[n-1]
            self.after_buy_rate_per.append(per_buy)
            print(f"BP= {math.prod(self.after_buy_rate_per)*100}")
            self.after_buy=(f"	BP= 	 {math.prod(self.after_buy_rate_per)*100}")
            self.after_buying_rate=[]
            self.profit.append(math.prod(self.after_buy_rate_per))
            self.after_buy_rate_per=[1.0]

        elif self.wt_to_do=="None":
            self.after_buy="waiting"

    def after_selling(self,s):
        n=len(self.after_selling_rate)            
        if self.wt_to_do == "Sell":
            self.after_selling_rate.append(s)
            print(f"SR = {math.prod(self.after_sell_rate_per)*100}")
            self.after_sell=f"	 SR =     {math.prod(self.after_sell_rate_per)*100}"	
            self.profit.append(math.prod(self.after_sell_rate_per))

        elif self.wt_to_do =="None":	
            self.profit.pop()
            self.after_selling_rate.append(s)
            per_sell=self.after_selling_rate[n-1]/self.after_selling_rate[n]
            self.after_sell_rate_per.append(per_sell)
            print(f"SR = {math.prod(self.after_sell_rate_per)*100}") #calcation is needed
            self.profit.append(math.prod(self.after_buy_rate_per))
            self.after_sell=f"	 SR =     {math.prod(self.after_sell_rate_per)*100}"

        elif self.wt_to_do =="Buy":
            self.profit.pop()
            self.after_selling_rate.append(s)
            per_sell=self.after_selling_rate[n-1]/self.after_selling_rate[n]
            self.after_sell_rate_per.append(per_sell)
            print(f"SP= {math.prod(self.after_sell_rate_per)*100}")
            self.after_sell=(f"	SP= 	 {math.prod(self.after_sell_rate_per)*100}")
            self.after_selling_rate=[]
            self.profit.append(math.prod(self.after_sell_rate_per))
            self.after_sell_rate_per=[1]

        elif self.wt_to_do=="Hold":
            self.after_sell="waiting"

class InformationToStore():
    def PreProccesInformationToStore(self):
        self.sheet=ws.sheets[0]
        self.cell=2
        
    def UpdateInformation(self,indicater_val,wt_to_do,after_buy,after_sell):
        self.sheet.range(f"g{self.cell}").value=[indicater_val,wt_to_do,after_buy,after_sell]
        (self.cell)+=1
        
    def FianlProof(self,profit):
        self.sheet.range(f"a{self.cell}").value=(profit)
        self.sheet.range(f"a{self.cell+1}").value=(math.prod(profit))

n=10000
class input_data(BuySell,InformationToStore):
    def __init__(self):
        self.tp=0
        self.PreProccesBuySell()
        self.PreProccesInformationToStore()
        input = {'open':[SheetList.range(f"b{2}").value] ,
            'high':[SheetList.range(f"c{2}").value],
            'low': [SheetList.range(f"d{2}").value],
            'close':[SheetList.range(f"e{2}").value],
            'volume':[SheetList.range(f"f{2}").value] }
        self.input=pd.DataFrame(input)

    def input_method(self,i):
        temp=pd.DataFrame({ 'open':[SheetList.range(f"b{i}").value] ,
            'high':[SheetList.range(f"c{i}").value],
            'low': [SheetList.range(f"d{i}").value],
            'close':[SheetList.range(f"e{i}").value],
            'volume':[SheetList.range(f"f{i}").value] })
        print(i/n*100)
        return temp

    def indicater_update(self,Sel_Indicate):
        if(Sel_Indicate=="MACD"):
                self.output = abstract.MACD(self.input)
        if(Sel_Indicate=="RSI"):
                self.output = abstract.RSI(self.input)
        if(Sel_Indicate=="EMA"):
                self.output = abstract.EMA(self.input)
        if(Sel_Indicate=="BBANDS"):
                self.output = abstract.BBANDS(self.input)

    def procces(self,Sel_Indicate):       
        for i in range(2,n):
            data=self.input_method(i)        
            self.input=pd.concat([self.input,data])

            self.indicater_update(Sel_Indicate)

            indicater_val=self.output.iloc[len(self.output)-1]
            print(indicater_val)
            
            price =self.input["close"].iloc[len(self.input["close"])-1]
            
            self.cheak(price/indicater_val)
            self.after_buying(price)
            self.after_selling(price)

            self.UpdateInformation(indicater_val,self.wt_to_do,self.after_buy,self.after_sell)

        print(math.prod(self.profit))
        self.FianlProof(self.profit)

i=input_data()
i.procces("EMA")# MACD RSI EMA BBANDS
