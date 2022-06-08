import pandas as pd 
import yfinance as yf
import numpy as np
import urllib
import json
from datetime import datetime
from sec_api import XbrlApi
from sec_api import QueryApi

api_key = "43086a924299cffd9b6ab60a22c3f3e378e13cb41ffba66b910cfb34c3b9e467"

class CompanyData:
    def __init__(self, ticker):
        self.ticker = yf.Ticker(ticker)
        self.name = ticker
    
    def get_stock_data(self):
        df = self.ticker.history(period = "6mo")
        df = df.loc[:, df.columns[0:5]]
        df = df.reset_index(level = 0)
        
        for i in ["Open", "High", "Low", "Close"]:
            df[i] = round(df[i], 2)
        
        for i in range(len(df)):
            df["Date"][i] = str(datetime.strptime
                           (str(df["Date"][i]),"%Y-%m-%d %H:%M:%S")
                           .date())
        
        df.loc[df["Volume"] == 0, "Volume"] = 1; 
        return df
    
    def get_generic_info(self):
        info = self.ticker.info
        desc = info["longBusinessSummary"]
        phone = info["phone"]
        addr = info["address1"] + ", " + info["city"] + ", " + info["state"]
        website = info["website"]
        name = info["shortName"]
        
        return {"desc":desc, "phone":phone, "address":addr, "website":website, "name":name}
    
    def get_shareholders(self):
        try:
            institution = self.ticker.institutional_holders.loc[:, ["Holder", "Date Reported", "% Out"]]
        except:
            institution = pd.DataFrame(columns = ["Holder", "Date Reported", "% Out"])
        try:
            mf = self.ticker.mutualfund_holders.loc[:, ["Holder", "Date Reported", "% Out"]]
        except:
            mf = pd.DataFrame(columns = ["Holder", "Date Reported", "% Out"])
        
        institution["type"] = "institution"
        mf["type"] = "mutual fund"
        
        df = mf.append(institution, ignore_index = True)
        
        data = pd.DataFrame(self.api_call("insiderHolders", "holders"))
        
        try:
            data = data[data["transactionDescription"] == "Purchase"]
            data["Holder"] = data["name"]
            data["% Out"] = np.nan
            data["Date Reported"] = np.nan
        
            if(not "positionIndirect" in data.columns):
                data["positionIndirect"] = np.nan
                data["positionIndirectDate"] = np.nan
            if(not "positionDirect" in data.columns):
                data["positionDirect"] = np.nan
                data["positionDirectDate"] = np.nan
        
            for i, _ in data.iterrows():
                if(not pd.isna(data.loc[i]["positionIndirect"]) and (not pd.isna(data.loc[i]["positionIndirectDate"]))):
                    data.loc[i,"% Out"] = data.loc[int(i)]["positionIndirect"]["raw"]
                    data.loc[i,"Date Reported"] = data.loc[i,"positionIndirectDate"]["fmt"]
                elif (not pd.isna(data.loc[i]["positionDirect"]) and (not pd.isna(data.loc[i]["positionDirectDate"]))):
                    data.loc[i,"% Out"] = int(data.loc[i,"positionDirect"]["raw"])
                    data.loc[i,"Date Reported"] = data.loc[i,"positionDirectDate"]["fmt"]
        
            data = data[["Holder","Date Reported", "% Out"]].dropna()
            data["type"] = "insider"
            data["% Out"] = data["% Out"]/self.get_share_data()
            df = df.append(data, ignore_index = True)
            
        except:
            pass
        
        df = df.sort_values("% Out", ascending = False, ignore_index = True)
        
        length = min(len(df), 10)
        df = df.iloc[0:length, :]
        
        for i,_ in df.iterrows():
            try: 
                df["Date Reported"][i] = str(datetime.strptime(str(df["Date Reported"][i]),"%Y-%m-%d %H:%M:%S").date())
            except:
                pass
        return df
    
    def api_call(self, module, category):
        response = urllib.request.urlopen(f"https://query2.finance.yahoo.com/v10/finance/quoteSummary/{self.name}?modules={module}")
        content = response.read()
        data = json.loads(content.decode('utf8'))['quoteSummary']['result'][0][module][category]
        return data 
    
    def get_management(self):
        data = self.api_call("assetProfile", "companyOfficers")
        df = pd.DataFrame(data)[["name","title"]][0:7]
        return df
    
    def get_news(self):
        lst_title = []
        lst_link = []
        data = self.ticker.news[0:40]
        for i in range(len(data)):
            lst_link.append(data[i]["link"])
            lst_title.append(data[i]["title"])
        return lst_link, lst_title
    
    def get_balance_sheet(self):
        df = self.ticker.quarterly_balance_sheet
             
        df = df.iloc[:,[0,3]].loc[["Cash","Total Current Assets", "Total Current Liabilities", "Retained Earnings"]]
        return df/1000
    
    def get_income_statement(self):
        annual = self.ticker.financials
        annual = annual.iloc[:,[0]].loc[["Total Revenue", "Cost Of Revenue", "Net Income"]]
        quarterly = self.ticker.quarterly_financials
        quarterly = quarterly.iloc[:,[0]].loc[["Total Revenue", "Cost Of Revenue", "Net Income"]]
        return pd.concat([annual,quarterly], axis = 1)/1000
    
    def get_cash_flow(self):
        data = self.ticker.quarterly_cashflow.iloc[:,[0,3]]
        data = data.loc[["Total Cash From Operating Activities", 
                         "Total Cashflows From Investing Activities", 
                         "Total Cash From Financing Activities"]]
        return data/1000
    
    def get_share_data(self):
        return self.ticker.info["sharesOutstanding"]
    
    def get_fiscal_end(self):
        data = self.api_call("defaultKeyStatistics","lastFiscalYearEnd")["fmt"][5::]
        return data
    
    def get_last_filing(self):
        data = self.api_call("secFilings", "filings")[0]
        return pd.DataFrame(data, index = [0])[["date", "type", "edgarUrl"]]
    
    def get_sec_file(self, form_type, size):
        queryApi = QueryApi(api_key)
        query = {
          "query": { "query_string": { 
              "query": "ticker:" + self.name +" AND formType:\""+form_type+"\"" 
            } },
          "from": "0",
          "size": str(size),
          "sort": [{ "filedAt": { "order": "desc" } }]
           }
        filings = queryApi.get_filings(query)
        return filings["filings"]
       
    
    def get_shares_auth(self):
        xbrlApi = XbrlApi(api_key)
        
        file_link = self.get_sef_file("10-K",1)
        
        xbrl_json = xbrlApi.xbrl_to_json(
            htm_url=file_link
        )
        
        return int(xbrl_json["BalanceSheetsParenthetical"]["CommonStockSharesAuthorized"][0]["value"])
        