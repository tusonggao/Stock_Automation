# -*- coding: utf8 -*-

from __future__ import unicode_literals
from __future__ import print_function



import os.path
import sys
import time
import collections
import pandas as pd
import datetime


################################################################################
###I am doing this to remind myself not forgeting practicing coding EVERY DAY!!! 
###I am doing this to remind myself not forgeting practicing coding EVERY DAY!!! 
###############################################################################
###############################################################################




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

# import win32clipboard as w
# import win32con




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

################################################################

class THS_Automation:
    def __init__(self, account=None):
        if account is not None:
            self.startProcess(account)
        # self.app = pywinauto.Application().Connect(path = "xiadan.exe")
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
        
        
#################################################################################

        
    def updateMarketPrices(self):
        self.newestMarketPrices = ts.get_today_all()
        self.lastUpdateTime = datetime.datetime.now()           
        
        ##############################################
        
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
            
            
  
