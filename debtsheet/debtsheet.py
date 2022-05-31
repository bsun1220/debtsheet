from unittest import mock
p = mock.patch('openpyxl.styles.fonts.Font.family.max', new=100)
p.start()
from openpyxl import load_workbook
from company_data import *

class DebtSheet:
    def __init__(self, ticker):
        self.company = CompanyData(ticker)
        self.ticker = ticker
        self.workbook = load_workbook(filename="debtsheet.xlsm", read_only = False, keep_vba = True)
        self.fill()
        
    def fill_data(self):
        data = self.workbook["Data"]
        stock = self.company.get_stock_data()
        
        obj = {"A":0, "B":1, "C":2, "D":3, "E":4, "F":5}
        
        for i in range(len(stock)):
            index = i + 8
            
            for key in obj.keys():
                stock_val = stock.iloc[i, obj[key]]
                cell = key + str(index)
                data[cell] = stock_val
    
    def fill_management(self):
        manage = self.company.get_management()
        data = self.workbook["BNOW"]
        for i,_ in manage.iterrows():
            index = i + 50
            name = manage.loc[i]["name"]
            title = manage.loc[i]["title"]
            data["B"+str(index)] = name
            data["C"+str(index)] = title
    
    def fill_generic(self):
        excel = self.workbook["BNOW"]
        data = self.company.get_generic_info()
        
        excel["B59"] = data["desc"]
        excel["H3"] = data["address"]
        excel["H4"] = data["phone"]
        excel["H5"] = data["website"]
        excel["J4"] = self.ticker.upper()
        
        name = "" + data["name"].upper() + " ("+self.ticker.upper()+ ")"
        
        excel["B2"] = name
    
    def fill_shareholders(self):
        excel = self.workbook["BNOW"]
        data = self.company.get_shareholders()
        
        for i,_ in data.iterrows():
            index = i + 68
            holder = data.loc[i]["Holder"]
            type_ = data.loc[i]["type"]
            date = data.loc[i]["Date Reported"]
            percent = data.loc[i]["% Out"]
            
            excel["B"+str(index)] = holder
            excel["C"+str(index)] = type_
            excel["D"+str(index)] = date
            excel["E"+str(index)] = percent
        
    def fill_info(self):
        excel = self.workbook["BNOW"]
        
        filing = self.company.get_last_filing()
        
        excel["C3"] = filing.loc[0]["type"]
        excel["C4"] = filing.loc[0]["date"]
        excel["E4"] = filing.loc[0]["edgarUrl"]
        
        fiscal_end = self.company.get_fiscal_end()
        excel["C5"] = fiscal_end
    
    def fill_shares(self):
        excel = self.workbook["BNOW"]
        data = self.company.get_share_data()
        excel["J73"] = data
        excel["J74"] = data
    
    def fill_fundamentals(self):
        excel = self.workbook["BNOW"]
        bs = self.company.get_balance_sheet()
        
        excel["I55"] = bs.iloc[0,0]
        excel["I56"] = bs.iloc[1,0]
        excel["I57"] = bs.iloc[2,0]
        excel["I59"] = bs.iloc[3,0]
        excel["J55"] = bs.iloc[0,1]
        excel["J56"] = bs.iloc[1,1]
        excel["J57"] = bs.iloc[2,1]
        excel["J59"] = bs.iloc[3,1]
        
        is_ = self.company.get_income_statement()
        excel["I61"] = is_.iloc[0,1]
        excel["I62"] = is_.iloc[1,1]
        excel["I65"] = is_.iloc[2,1]
        excel["J61"] = is_.iloc[0,0]
        excel["J62"] = is_.iloc[1,0]
        excel["J65"] = is_.iloc[2,0]
        
        cf = self.company.get_cash_flow()
        excel["I67"] = cf.iloc[0,0]
        excel["I68"] = cf.iloc[1,0]
        excel["I69"] = cf.iloc[2,0]
        excel["J67"] = cf.iloc[0,1]
        excel["J68"] = cf.iloc[1,1]
        excel["J69"] = cf.iloc[2,1]
        
        for i in ["J49","J54","J66"]:
            excel[i] = "3 Quarters Ago"
        excel["I60"] = "Last Quarter"
    
    def fill_news(self):
        data_link, data_title = self.company.get_news()
        excel = self.workbook["BNOW"]
        for i in range(len(data_link)):
            index = i + 7
            excel["L"+str(index)] = data_title[i]
            excel["M"+str(index)] = data_link[i]
            
    def save_book(self):
        self.workbook.save("./completed/"+self.ticker + "_debtsheet.xlsm")
    
    def fill(self):
        self.fill_data()
        self.fill_management()
        self.fill_generic()
        self.fill_shareholders()
        self.fill_info()
        self.fill_shares()
        self.fill_fundamentals()
        try:
            self.fill_news()
        except:
            pass
        self.save_book()