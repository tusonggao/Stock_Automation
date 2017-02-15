# -*- coding: utf8 -*-

from __future__ import unicode_literals
from __future__ import print_function

import os.path
import sys
import time
import collections
import pandas as pd
import datetime

try:
    from pywinauto import application
except ImportError:
    pywinauto_path = os.path.abspath(__file__)
    pywinauto_path = os.path.split(os.path.split(pywinauto_path)[0])[0]
    sys.path.append(pywinauto_path)
    from pywinauto import application

import pywinauto
from pywinauto.timings import Timings
import win32gui
from win32.lib import win32con

import win32clipboard as w
import win32con

import tushare as ts


def getClipboardText():
    w.OpenClipboard()
    d = w.GetClipboardData(win32con.CF_TEXT)
    w.CloseClipboard()
    return d

def setClipboardText(aString):
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_TEXT, aString)
    w.CloseClipboard()
    

class THS_Automation:
    def __init__(self, account=None):
        if account is not None:
            self.startProcess(account)
        self.app = pywinauto.Application().Connect(path = "xiadan.exe")
        self.top_hwnd = win32gui.FindWindow(None, u'网上股票交易系统5.0')
        self.app[u'网上股票交易系统5.0'].maximize()
        self.updateTimeDelta = datetime.timedelta(5, 0, 0)
        self.lastUpdateTime = datetime.datetime(2000, 1, 1)
        self.slippage = 0.005  # 市价买入和卖出时取0.5%d的滑点，可以自行设置        
        

    def startProcess(self, account):
        app = pywinauto.Application().start(u'D:\\同花顺软件\\同花顺\\xiadan.exe')
        app.Dialog.ComboBox1.Select(account['stockCompany'])
        app.Dialog.ComboBox3.Edit.SetEditText(account['userId']) 
        app.Dialog.Edit2.SetEditText(account['password'])
        time.sleep(0.5)
        app.Dialog[u'登录'].click()
        app.Dialog[u'登录'].click()
        
        
    def updateMarketPrices(self):
        self.newestMarketPrices = ts.get_today_all()
        self.lastUpdateTime = datetime.datetime.now()           
        
    def buyStock_THS(self, stock_id, number, price=0.0):
        tree = self.app[u'网上股票交易系统5.0'].tree_view
        tree.get_item(u'\买入[F1]').click()
        
        if price==0.0:
            price = self.getMarketBuyPrice(stock_id)
            print("the sell price of stock_id %s is %.2f "%(stock_id, price))
            if price==0.0: # 停牌
                return
                                               
        setText_OK = True
        if setText_OK:
            for i in range(3):
                self.app[u'网上股票交易系统5.0'].Edit1.SetEditText(stock_id)
                if self.app[u'网上股票交易系统5.0'].Edit1.TextBlock()==stock_id:
                    break
                else:
                    print("i am useful")
                    self.app[u'网上股票交易系统5.0'].Edit1.SetEditText("")
            else:
                setText_OK = False
                
        if setText_OK:
            for i in range(3):
                self.app[u'网上股票交易系统5.0'].Edit2.SetEditText(str(price))
                if self.app[u'网上股票交易系统5.0'].Edit2.TextBlock()==str(price):
                    break
                else:
                    print("i am useful")
                    self.app[u'网上股票交易系统5.0'].Edit2.SetEditText("")
            else:
                setText_OK = False
                
        if setText_OK:
            for i in range(3):
                self.app[u'网上股票交易系统5.0'].Edit3.SetEditText(str(number))
                if self.app[u'网上股票交易系统5.0'].Edit3.TextBlock()==str(number):
                    break
                else:
                    print("i am useful")
                    self.app[u'网上股票交易系统5.0'].Edit3.SetEditText("")
            else:
                setText_OK = False
                
        if setText_OK:
            self.app[u'网上股票交易系统5.0'][u'买入[S]'].click()
            print("Buy order sent sucesessfully!!!") 
        else:
            print("Buy order sent unsucesessfully!!!")
        return setText_OK
        
    def getMarketBuyPrice(self, stock_id):
        if (datetime.datetime.now() - self.lastUpdateTime) > self.updateTimeDelta:
            self.updateMarketPrices()
        prices = self.newestMarketPrices.ix[self.newestMarketPrices.code==stock_id]
        print("prices is ", prices)
        price = prices.iloc[0]['trade']**(1 + self.slippage)      
        high_price = prices.iloc[0]['high']
        return min(price, high_price)
        
    def getMarketSellPrice(self, stock_id):
        if (datetime.datetime.now() - self.lastUpdateTime) > self.updateTimeDelta:
            self.updateMarketPrices()
        prices = self.newestMarketPrices.ix[self.newestMarketPrices.code==stock_id]
        price = prices.iloc[0]['trade']**(1 - self.slippage)      
        low_price = prices.iloc[0]['low']
        return round(max(price, low_price), 2)


    def sellAllPositionsByMarketPrice(self):
        self.getCurrentPositions()
        for index, row in self.currentPositons.iterrows():
            if int(row[u'可用余额'])>0:
                print("可用余额不为0")
                self.sellStock_THS(row[u'证券代码'], int(row[u'可用余额']))
            else:
                print("可用余额为0")
            print("sell stock_id: ", row[u'证券代码'], "number: ",  int(row[u'可用余额']))
    
    def sellStock_THS(self, stock_id, number, price=0.0):
        tree = self.app[u'网上股票交易系统5.0'].tree_view
        tree.get_item(u'\卖出[F1]').click()
        
        if price==0.0:
            price = self.getMarketSellPrice(stock_id)
            print("the sell price of stock_id %s is %.2f "%(stock_id, price))
            if price==0.0: # 停牌
                return
        
        setText_OK = True

        if setText_OK:
            for i in range(3):
                self.app[u'网上股票交易系统5.0'].Edit1.SetEditText(stock_id)
                if self.app[u'网上股票交易系统5.0'].Edit1.TextBlock()==stock_id:
                    break
                else:
                    print("i am useful sell 1")
                    self.app[u'网上股票交易系统5.0'].Edit1.SetEditText("")
            else:
                setText_OK = False
                
        if setText_OK:
            for i in range(3):
                self.app[u'网上股票交易系统5.0'].Edit2.SetEditText(str(price))
                if self.app[u'网上股票交易系统5.0'].Edit2.TextBlock()==str(price):
                    break
                else:
                    print("i am useful sell 2")
                    self.app[u'网上股票交易系统5.0'].Edit2.SetEditText("")
            else:
                setText_OK = False
                
        if setText_OK:
            for i in range(3):
                self.app[u'网上股票交易系统5.0'].Edit3.SetEditText(str(number))
                if self.app[u'网上股票交易系统5.0'].Edit3.TextBlock()==str(number):
                    break
                else:
                    print("i am useful sell 3")
                    self.app[u'网上股票交易系统5.0'].Edit3.SetEditText("")
            else:
                setText_OK = False
                
        if setText_OK:
            self.app[u'网上股票交易系统5.0'][u'卖出[S]'].click()
            print("Sell order sent sucesessfully!!!")            
        else:
            print("the set content is ", self.app[u'网上股票交易系统5.0'].Edit2.TextBlock())
            print("the price should be ", str(price))
            print("Sell order sent unsucesessfully!!!")       
        
        return setText_OK

    def getAccountValue_THS(self):
        tree = self.app[u'网上股票交易系统5.0'].tree_view    
        tree.get_item(u'\查询[F4]\资金股票').click()
        self.app[u'网上股票交易系统5.0'].TypeKeys("{F5}")
        time.sleep(0.5)
        
        account_value = collections.OrderedDict()
        account_value[u'资金余额'] = float(self.app[u'网上股票交易系统5.0'].Static4.WindowText())
        account_value[u'冻结金额'] = float(self.app[u'网上股票交易系统5.0'].Static5.WindowText())
        account_value[u'可用金额'] = float(self.app[u'网上股票交易系统5.0'].Static6.WindowText())
        account_value[u'可取金额'] = float(self.app[u'网上股票交易系统5.0'].Static10.WindowText())
        account_value[u'股票市值'] = float(self.app[u'网上股票交易系统5.0'].Static11.WindowText())
        account_value[u'总资产'] = float(self.app[u'网上股票交易系统5.0'].Static12.WindowText())
        account_value[u'持仓盈亏'] = float(self.app[u'网上股票交易系统5.0'].Static17.WindowText())
        account_value[u'当日盈亏'] = self.app[u'网上股票交易系统5.0'].Static15.WindowText()
        self.account_value = account_value
        
        print("the current account_value is:", self.account_value)

    def getCurrentPositions(self):
        tree = self.app[u'网上股票交易系统5.0'].tree_view
        tree.get_item(u'\查询[F4]\资金股票').click() # 可以通过ctrl+c来复制出来
        self.app[u'网上股票交易系统5.0'].CVirtualGridCtrl.TypeKeys("^C")
        self.app[u'网上股票交易系统5.0'].CVirtualGridCtrl.RightClickInput(coords=(800, 500))
        self.app[u'网上股票交易系统5.0'].CVirtualGridCtrl.ClickInput(coords=(830, 565))
        lines = getClipboardText().decode("gb2312").split('\r\n')
        titles = lines[0].split('\t')[:-1]
        positions = []
        for i in range(1, len(lines)):
            positions.append(lines[i].split('\t')[:-1])
        self.currentPositons = pd.DataFrame(positions, columns=titles)
        print(u"现在持有仓位")
        print(self.currentPositons)
        return self.currentPositons
        
        
    def getTodayTransaction(self):   
        tree = self.app[u'网上股票交易系统5.0'].tree_view
        tree.get_item(u'\查询[F4]\当日成交').click()
        self.app[u'网上股票交易系统5.0'].CVirtualGridCtrl.TypeKeys('{F5}')
        self.app[u'网上股票交易系统5.0'].CVirtualGridCtrl.set_focus()        
        self.app[u'网上股票交易系统5.0'].CVirtualGridCtrl.TypeKeys("^C")
        self.app[u'网上股票交易系统5.0'].CVirtualGridCtrl.RightClickInput(coords=(800, 500))
        self.app[u'网上股票交易系统5.0'].CVirtualGridCtrl.ClickInput(coords=(830, 565))        

        lines = getClipboardText().decode("gb2312").split('\r\n')
        titles = lines[0].split('\t')[:-1]
        transcations = []
        for i in range(1, len(lines)):
            transcations.append(lines[i].split('\t')[:-1])
        self.todayTranscations = pd.DataFrame(transcations, columns=titles)
        
        print(u"今日成交")
        print(self.todayTranscations)
        return self.todayTranscations
        





#=============================================================================
def demo():
    account1 = {'stockCompany': u'国泰君安-广州黄埔大道证券营业部', 
               'userId': '10335207',
               'password': '*******'}
    account2 = {'stockCompany': u'模拟炒股', 
               'userId': 'krazy47',
               'password': '*****'}
    auto = THS_Automation()

    auto.getAccountValue_THS()
#    auto.getCurrentPositions()
    auto.sellAllPositionsByMarketPrice()
    
    time.sleep(2.0)


    for i in range(3):
        print('get here', i)
        auto.getCurrentPositions()
        auto.getTodayTransaction()
        time.sleep(10.0)
        
        
if __name__ == '__main__':
    demo()    
        





#    win32gui.EnumWindows(getWindowTitles, 0)
#    lt = [t for t in getWindowTitles.titles if t]
#    lt.sort()
#    for t in lt:
#        print(t)
#        time.sleep(0.03)
