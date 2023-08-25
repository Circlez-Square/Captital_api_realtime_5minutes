# 先把API com元件初始化
import os
import tkinter as tk

# 第二種讓群益API元件可導入Python code內用的物件宣告
import comtypes.client
import pandas as pd

comtypes.client.GetModule(os.path.split(os.path.realpath(__file__))[0] + r'\SKCOM.dll')  # 加此行需將API放與py同目錄
import comtypes.gen.SKCOMLib as sk

skC = comtypes.client.CreateObject(sk.SKCenterLib, interface=sk.ISKCenterLib)
skOOQ = comtypes.client.CreateObject(sk.SKOOQuoteLib, interface=sk.ISKOOQuoteLib)
skO = comtypes.client.CreateObject(sk.SKOrderLib, interface=sk.ISKOrderLib)
skOSQ = comtypes.client.CreateObject(sk.SKOSQuoteLib, interface=sk.ISKOSQuoteLib)
skQ = comtypes.client.CreateObject(sk.SKQuoteLib, interface=sk.ISKQuoteLib)
skR = comtypes.client.CreateObject(sk.SKReplyLib, interface=sk.ISKReplyLib)

# 畫視窗用物件
from tkinter import *
from tkinter.ttk import *
from tkinter import messagebox, colorchooser, font, Button, Frame, Label
from tkinter import ttk

# 數學計算用物件
import math

# 其它物件
import Config

#其他excel
import openpyxl

#時間類別
from datetime import datetime,timedelta

# 顯示各功能狀態用的function
def WriteMessage(strMsg, listInformation):
    listInformation.insert('end', strMsg)
    listInformation.see('end')

def SendReturnMessage(strType, nCode, strMessage, listInformation):
    GetMessage(strType, nCode, strMessage, listInformation)

def GetMessage(strType, nCode, strMessage, listInformation):
    strInfo = ""
    if (nCode != 0):
        strInfo = "【" + skC.SKCenterLib_GetLastLogInfo() + "】"
    WriteMessage("【" + strType + "】【" + strMessage + "】【" + skC.SKCenterLib_GetReturnCodeMessage(nCode) + "】" + strInfo,
                 listInformation)

# ----------------------------------------------------------------------------------------------------------------------------------------------------
# 上半部登入框
class FrameLogin(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.grid()
        # self.pack()
        self.place()
        self.FrameLogin = Frame(self)
        self.master["background"] = "#F5F5F5"
        self.FrameLogin.master["background"] = "#F5F5F5"
        self.createWidgets()

    def createWidgets(self):
        # 帳號
        self.labelID = Label(self)
        self.labelID["text"] = "帳號："
        self.labelID["background"] = "#F5F5F5"
        self.labelID["font"] = 20
        self.labelID.grid(column=0, row=0)
        # 輸入框
        self.textID = Entry(self)
        self.textID["width"] = 50
        self.textID.grid(column=1, row=0)

        # 密碼
        self.labelPassword = Label(self)
        self.labelPassword["text"] = "密碼："
        self.labelPassword["background"] = "#F5F5F5"
        self.labelPassword["font"] = 20
        self.labelPassword.grid(column=2, row=0)
        # 輸入框
        self.textPassword = Entry(self)
        self.textPassword["width"] = 50
        self.textPassword['show'] = '*'
        self.textPassword.grid(column=3, row=0)

        # 按鈕
        self.buttonLogin = Button(self)
        self.buttonLogin["text"] = "登入"
        self.buttonLogin["background"] = "#4169E1"
        self.buttonLogin["foreground"] = "#ffffff"
        self.buttonLogin["font"] = 20
        self.buttonLogin["command"] = self.buttonLogin_Click
        self.buttonLogin.grid(column=4, row=0)

        # ID
        self.labelID = Label(self)
        self.labelID["text"] = "<<ID>>"
        self.labelID["background"] = "#F5F5F5"
        self.labelID["font"] = 20
        self.labelID.grid(column=5, row=0)

        # 訊息欄
        self.listInformation = Listbox(root, height=5)
        self.listInformation.grid(column=0, row=1, sticky=E + W)

        global GlobalListInformation, Global_ID
        GlobalListInformation = self.listInformation
        Global_ID = self.labelID

    # 這裡是登入按鈕,使用群益API不管要幹嘛你都要先登入才行
    def buttonLogin_Click(self):
        try:
            m_nCode = skC.SKCenterLib_SetLogPath(os.path.split(os.path.realpath(__file__))[0] + "\\CapitalLog_Quote")
            m_nCode = skC.SKCenterLib_Login(self.textID.get().replace(' ', ''),
                                            self.textPassword.get().replace(' ', ''))
            if (m_nCode == 0):
                Global_ID["text"] = self.textID.get().replace(' ', '')
                WriteMessage("登入成功", self.listInformation)
            else:
                WriteMessage(m_nCode, self.listInformation)
        except Exception as e:
            messagebox.showerror("error！", e)

# 報價連線的按鈕
class FrameQuote(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.grid()
        self.FrameQuote = Frame(self)
        self.FrameQuote.master["background"] = "#F5F5F5"
        self.createWidgets()

    def createWidgets(self):
        # ID 若想顯示登入帳號
        # self.labelID = Label(self)
        # self.labelID["text"] = "ID："
        # self.labelID.grid(column = 0, row = 0)

        # Connect
        self.btnConnect = Button(self)
        self.btnConnect["text"] = "報價連線"
        self.btnConnect["background"] = "#d0d0d0"
        self.btnConnect["font"] = 20
        self.btnConnect["command"] = self.btnConnect_Click
        self.btnConnect.grid(column=0, row=1)

        # Disconnect
        self.btnDisconnect = Button(self)
        self.btnDisconnect["text"] = "報價斷線"
        self.btnDisconnect["background"] = "#d0d0d0"
        self.btnDisconnect["font"] = 20
        self.btnDisconnect["command"] = self.btnDisconnect_Click
        self.btnDisconnect.grid(column=1, row=1)

        # ServerTime
        self.btnTime = Button(self)
        self.btnTime["text"] = "主機時間"
        self.btnTime["background"] = "#d0d0d0"
        self.btnTime["font"] = 20
        self.btnTime["command"] = self.btnTime_Click
        self.btnTime.grid(column=2, row=1, sticky=E)

        self.timeshower = Label(self)
        self.timeshower["text"] = "00:00:00"
        self.timeshower["background"] = "#F5F5F5"
        self.timeshower["font"] = 20
        self.timeshower.grid(column=3, row=1, sticky=W)

        global Gobal_ServerTime_Information
        Gobal_ServerTime_Information = self.timeshower

        # TabControl
        self.TabControl = Notebook(self)
        # self.TabControl.add(Quote(master=self), text="行情_報價")
        self.TabControl.add(TickandBest5(master=self), text="Tick&Best5")
        self.TabControl.grid(column=0, row=2, sticky=E + W, columnspan=4)

    def btnConnect_Click(self):
        try:
            m_nCode = skQ.SKQuoteLib_EnterMonitorLONG()
            SendReturnMessage("Quote", m_nCode, "SKQuoteLib_EnterMonitorLONG", GlobalListInformation)
        except Exception as e:
            messagebox.showerror("error！", e)

    def btnDisconnect_Click(self):
        try:
            m_nCode = skQ.SKQuoteLib_LeaveMonitor()
            if (m_nCode != 0):
                strMsg = "SKQuoteLib_LeaveMonitor failed!", skC.SKCenterLib_GetReturnCodeMessage(m_nCode)
                WriteMessage(strMsg, GlobalListInformation)
            else:
                SendReturnMessage("Quote", m_nCode, "SKQuoteLib_LeaveMonitor", GlobalListInformation)
        except Exception as e:
            messagebox.showerror("error！", e)

    def btnTime_Click(self):
        try:
            m_nCode = skQ.SKQuoteLib_RequestServerTime()
            SendReturnMessage("Quote", m_nCode, "SKQuoteLib_RequestServerTime", GlobalListInformation)
        except Exception as e:
            messagebox.showerror("error！", e)

# 下半部-報價-TickandBest5項目
class TickandBest5(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.grid()
        self.Quote = Frame(self)
        self.Quote.master["background"] = "#F5F5F5"
        self.createWidgets()

    def createWidgets(self):
        # PageNo
        self.LabelPageNo = Label(self)
        self.LabelPageNo["text"] = "PageNo"
        self.LabelPageNo["background"] = "#F5F5F5"
        self.LabelPageNo["font"] = 20
        self.LabelPageNo.grid(column=0, row=0)
        # 輸入框
        self.strPageNo = StringVar()
        self.txtPageNo = Entry(self, textvariable=self.strPageNo)
        self.strPageNo.set("0")
        self.txtPageNo.grid(column=1, row=0)

        # 商品代碼
        self.LabelStocks = Label(self)
        self.LabelStocks["text"] = "商品代碼"
        self.LabelStocks["background"] = "#F5F5F5"
        self.LabelStocks["font"] = 20
        self.LabelStocks.grid(column=2, row=0)
        # 輸入框
        self.strStocks = StringVar()
        self.txtStocks = Entry(self, textvariable=self.strStocks)
        self.strStocks.set("TX00")
        self.txtStocks.grid(column=3, row=0)

        # 按鈕
        self.btnQueryStocks = Button(self)
        self.btnQueryStocks.config
        self.btnQueryStocks.config(text="查詢完整", fg="black", font="20", command=self.btnTick_Click)
        self.btnQueryStocks["background"] = "#d0d0d0"
        self.btnQueryStocks.grid(column=4, row=0)

        # LiveTick查詢按鈕
        self.btnQueryStocks = Button(self)
        self.btnQueryStocks.config(text="查詢即時", fg="black", font="20", command=self.btnLiveTick_Click)
        self.btnQueryStocks["background"] = "#d0d0d0"
        # self.btnQueryStocks.grid(column = 5, row = 0)

        self.btnQueryStocks = Button(self)
        self.btnQueryStocks.config(text="stopLive", fg="black", font="20", command=self.btnLiveStop_Click)
        self.btnQueryStocks["background"] = "#d0d0d0"
        # self.btnQueryStocks.grid(column = 7, row = 0)

        # 訊息欄
        self.listInformation = Listbox(self, height=25, width=100)
        self.listInformation.grid(column=0, row=1, sticky=E + W, columnspan=6, rowspan=7)

        global Gobal_Tick_ListInformation, Gobal_Best5TreeViewQ_Information, Gobal_Best5TreeViewP_Information, Gobal_Best5TreeViewQ_Information2, Gobal_Best5TreeViewP_Information2, Gobal_Best5TXTInfo
        # global Gobal_Best5TreeView, Gobal_Best5TreeView2
        Gobal_Tick_ListInformation = self.listInformation
        # Gobal_Best5TreeViewQ_Information = self.Quantity
        # Gobal_Best5TreeViewP_Information = self.Price
        # Gobal_Best5TreeViewQ_Information2 = self.Quantity2
        # Gobal_Best5TreeViewP_Information2 = self.Price2
        # Gobal_Best5TreeView = self.treeview
        # Gobal_Best5TreeView2 = self.treeview2

    def btnTick_Click(self):
        try:
            pn = 0
            if (self.txtPageNo.get().replace(' ', '') != ''):
                pn = int(self.txtPageNo.get())
            m_nCode = skQ.SKQuoteLib_RequestTicks(pn, self.txtStocks.get().replace(' ', ''))
            if (m_nCode == 0):
                SendReturnMessage("Tick&Best5", m_nCode[1], "SKQuoteLib_RequestTicks", GlobalListInformation)
        except Exception as e:
            messagebox.showerror("error！", e)

    def btnLiveTick_Click(self):
        try:
            if (self.txtPageNo.get().replace(' ', '') != ''):
                pn = int(self.txtPageNo.get())
            m_nCode = skQ.SKQuoteLib_RequestLiveTick(pn, self.txtStocks.get().replace(' ', ''))
        except Exception as e:
            messagebox.showerror("error！", e)

    # def btnStop_Click(self):
    #     try:
    #         pn = 0
    #         if (self.txtPageNo.get().replace(' ', '') != ''):
    #             pn = 50
    #         m_nCode = skQ.SKQuoteLib_RequestTicks(pn, self.txtStocks.get().replace(' ', ''))
    #     except Exception as e:
    #         messagebox.showerror("error！", e)

    def btnLiveStop_Click(self):
        try:
            pn = 2
            if (self.txtPageNo.get().replace(' ', '') != ''):
                pn = 50
            m_nCode = skQ.SKQuoteLib_RequestLiveTick(pn, self.txtStocks.get().replace(' ', ''))
        except Exception as e:
            messagebox.showerror("error！", e)

    def btnGetTick_Click(self):
        try:
            if (self.tickMarketNo.get() == "0 = 上市"):
                sMarketNo = 0
            elif (self.tickMarketNo.get() == "1 = 上櫃"):
                sMarketNo = 1
            elif (self.tickMarketNo.get() == "2 = 期貨"):
                sMarketNo = 2
            elif (self.tickMarketNo.get() == "3 = 選擇權"):
                sMarketNo = 3
            else:
                sMarketNo = 4

            pStock = sk.SKTICK()
            m_nCode = skQ.SKQuoteLib_GetTickLONG(sMarketNo, int(self.txtStocks3.get()), int(self.txtStocks4.get()),
                                                 pStock)
            SendReturnMessage("Quote", m_nCode[1], "SKQuoteLib_GetTickLONG", GlobalListInformation)

            # 處理時間顯示
            time = int(pStock.nTimehms)
            hour = time // 10000
            b = time - hour * 10000
            min = b // 100
            sec = b - min * 100
            millisec = pStock.nTimemillismicros // 1000
            microsec = pStock.nTimemillismicros - millisec * 1000

            # 判斷揭示類型
            if pStock.nSimulate == 0:
                tick_class = "一般揭示"
            elif pStock.nSimulate == 1:
                tick_class = "試算揭示"
            else:
                tick_class = "Error"

            # 輸出
            # Gobal_tickclass["text"] = "揭示類型: %s" % tick_class
            # Gobal_tickInfo["text"] = "日期:%d 買價:%d 賣價:%d 成交量:%d 成交價:%d" % (
            # pStock.nDate, pStock.nBid, pStock.nAsk, pStock.nQty, pStock.nClose)
            # Gobal_ticktime["text"] = "時間: %02d點 %02d分 %02d秒 %03d毫秒 %03d微秒" % (hour, min, sec, millisec, microsec)

        except Exception as e:
            messagebox.showerror("error！", e)

    def btnBest5_Click(self):
        try:
            if (self.boxMarketNo.get() == "0 = 上市"):
                sMarketNo = 0x00
            elif (self.boxMarketNo.get() == "1 = 上櫃"):
                sMarketNo = 0x01
            elif (self.boxMarketNo.get() == "2 = 期貨"):
                sMarketNo = 0x02
            elif (self.boxMarketNo.get() == "3 = 選擇權"):
                sMarketNo = 0x03
            else:
                sMarketNo = 0x04
            pStock = sk.SKBEST5()
            m_nCode = skQ.SKQuoteLib_GetBest5LONG(sMarketNo, int(self.txtBest5.get()), pStock)
            SendReturnMessage("Quote", m_nCode[1], "SKQuoteLib_GetBest5LONG", GlobalListInformation)

        except Exception as e:
            messagebox.showerror("error！", e)

    def btnGetALLInfo_Click(self):
        try:
            m_nCode = skQ.SKQuoteLib_RequestMACD(0, self.txtStocks.get())
            SendReturnMessage("Quote", m_nCode[1], "SKQuoteLib_RequestMACD", GlobalListInformation)

        except Exception as e:
            messagebox.showerror("error！", e)

        try:
            m_nCode = skQ.SKQuoteLib_RequestBoolTunel(0, self.txtStocks.get())
            SendReturnMessage("Quote", m_nCode[1], "SKQuoteLib_RequestBoolTunel", GlobalListInformation)

        except Exception as e:
            messagebox.showerror("error！", e)

        try:
            page = 0
            m_nCode = skQ.SKQuoteLib_RequestFutureTradeInfo(comtypes.automation.c_short(0), self.txtStocks.get())
            SendReturnMessage("Quote", m_nCode, "SKQuoteLib_RequestFutureTradeInfo", GlobalListInformation)

        except Exception as e:
            messagebox.showerror("error！", e)

    def btnCancel_Click(self):
        try:
            m_nCode = skQ.SKQuoteLib_RequestMACD(50, self.txtStocks.get())
            SendReturnMessage("Cancel", m_nCode[1], "CancelSKQuoteLib_RequestMACD", GlobalListInformation)

        except Exception as e:
            messagebox.showerror("error！", e)

        try:
            m_nCode = skQ.SKQuoteLib_RequestBoolTunel(50, self.txtStocks.get())
            SendReturnMessage("Cancel", m_nCode[1], "CancelSKQuoteLib_RequestBoolTunel", GlobalListInformation)

        except Exception as e:
            messagebox.showerror("error！", e)

        try:
            page = 50
            m_nCode = skQ.SKQuoteLib_RequestFutureTradeInfo(comtypes.automation.c_short(0), self.txtStocks.get())
            SendReturnMessage("Cancel", m_nCode, "CancelSKQuoteLib_RequestFutureTradeInfo", GlobalListInformation)

        except Exception as e:
            messagebox.showerror("error！", e)

# 事件
class SKQuoteLibEvents:

    def __init__(self):
        self.headers_created = False
        self.excel_file_name = 'data1.xlsx'
        self.create_excel_with_headers()
        self.data_current_min = {}
        self.data = pd.DataFrame(columns=['代碼','時間', '開盤價', '最大值', '最低值', '成交價'])
        self.current_interval_start = None
        self.open_updated = False
        # self.current_time = datetime.now()

    def OnConnection(self, nKind, nCode):
        if (nKind == 3001):
            strMsg = "Connected!"
        elif (nKind == 3002):
            strMsg = "DisConnected!"
        elif (nKind == 3003):
            strMsg = "Stocks ready!"
        elif (nKind == 3021):
            strMsg = "Connect Error!"
        WriteMessage(strMsg, GlobalListInformation)

    def create_excel_with_headers(self):
        if not self.headers_created:
            wb = openpyxl.Workbook()
            sheet = wb.active
            headers = ['代碼','時間', '開盤價', '最大值', '最低值', '成交價']
            sheet.append(headers)
            wb.save(self.excel_file_name)
            print("Excel file with headers created successfully.")
            self.headers_created = True

    def write_data_to_excel(self, data_list):
        wb = openpyxl.load_workbook(self.excel_file_name)
        sheet = wb.active
        sheet.append(data_list)
        wb.save(self.excel_file_name)
        print("Data successfully saved to '{}'".format(self.excel_file_name))

    #如果要回補歷史資料從這改
    # def OnNotifyHistoryTicksLONG(self, sMarketNo, nStockidx, nPtr, lDate, lTimehms, lTimemillismicros, nBid, nAsk, nClose, nQty, nSimulate):
    #     strMsg = "[OnNotifyHistoryTicksLONG]", sMarketNo,nStockidx
    #     WriteMessage(strMsg,Gobal_Tick_ListInformation)

    def OnNotifyTicksLONG(self, sMarketNo, nStockidx, nPtr, lDate, lTimehms, lTimemillismicros, nBid, nAsk, nClose,nQty, nSimulate):
        #夜盤台指代號4730
        # print("test", sMarketNo)
        # print("test1", nStockidx)
        new_name = nStockidx if nStockidx != 4730 else "TX00"
        # print("ok:",lTimehms)
        #使用python 本機時間.......有延遲
        # current_time = datetime.now()
        # minute = current_time.minute
        # second = current_time.second

        #看是不是在同一區間(這裡設定為5分k)
        minute = int((lTimehms / 100) % 100)
        second = int(lTimehms % 100)
        # print("test:",minute)
        # print ("test1:",second)
        current_interval_start = minute - minute % 5

        #當目前的時間區間 和 如果在這個名稱 new_name裡沒有值則回傳none/進一步檢查新名稱是否存在於當前數據中/最後檢查上一次的關閉時間是否存在
        if current_interval_start != self.data_current_min.get(new_name, {}).get("StartMinute") and new_name in self.data_current_min and self.data_current_min[new_name]["LastCloseTime"] is not None:
            # print("test3:", "ok3")
            last_close_time = self.data_current_min[new_name]["LastCloseTime"]
            hours = int(last_close_time / 10000)
            minutes = int((last_close_time / 100) % 100)
            seconds = int(last_close_time % 100)


            # 創建 datetime 對象
            last_close_datetime = datetime(year=lDate // 10000, month=(lDate // 100) % 100, day=lDate % 100,
                                           hour=hours, minute=minutes, second=seconds)

            # 計算下一個5分鐘區間的開始時間
            next_interval_start = last_close_datetime + timedelta(minutes=5 - (last_close_datetime.minute % 5),
                                                                  seconds=-last_close_datetime.second)

            # # 如果超過24小時，則將時間設置為下一天的00:00:00，可能會有換月問題要注意，換年等問題
            # if next_interval_start.hour == 24:
            #     next_interval_start = datetime(year=next_interval_start.year, month=next_interval_start.month,
            #                                    day=next_interval_start.day + 1,
            #                                    hour=0, minute=0, second=0)

            time_str = next_interval_start.strftime("%Y%m%d - %H:%M:%S")
            row_data = [new_name, time_str, (self.data_current_min[new_name]["Open"])/100,
                                (self.data_current_min[new_name]["High"])/100, (self.data_current_min[new_name]["Low"])/100,
                                (self.data_current_min[new_name]["Close"])/100]
            self.write_data_to_excel(row_data)
            del self.data_current_min[new_name]
            self.open_updated = True

        # 如果資料名稱不在或      當前周期和上一個已知的周期不同，則初始化新的周期，並建立欄位 (or current_interval_start != self.data_current_min[new_name].get("StartMinute"))
        if new_name not in self.data_current_min or current_interval_start != self.data_current_min[new_name].get("StartMinute"):
            self.data_current_min[new_name] = {"Open": nClose, "High": nClose, "Low": nClose,
                                               "StartMinute": current_interval_start, "LastCloseTime": None}
            self.open_updated = False
        # 更新最大值和最小值 時間和收盤價  ==0 可能會取 ==0 的最後一筆
        if new_name in self.data_current_min or (minute % 5 == 0 and second == 0):
            self.data_current_min[new_name]["LastCloseTime"] = lTimehms
            self.data_current_min[new_name]["High"] = max(self.data_current_min[new_name]["High"], nClose)
            self.data_current_min[new_name]["Low"] = min(self.data_current_min[new_name]["Low"], nClose)
            self.data_current_min[new_name]["Close"] = nClose

        # # 如果是5分鐘周期的最後一分鐘，則寫入收盤價和收盤時間的數據並清除當前數據.........這裡要注意商品性質 若最後一分鐘都沒有成交資料，視情況修改
        # if minute % 5 == 4 :


        strMsg = {
            "Newname" : nStockidx,
            "Date": lDate,
            "Hour":int(lTimehms/10000),
            "minute": int((lTimehms / 100) % 100),
            "Second": int(lTimehms % 100),
            "Time": str(lDate) + f" - {int(lTimehms/10000):02d}:{int((lTimehms / 100) % 100):02d}:{int(lTimehms % 100):02d}",
            "Close": nClose
        }
        WriteMessage(strMsg, Gobal_Tick_ListInformation)

    def OnNotifyServerTime(self, sHour, sMinute, sSecond, nTotal):
        strMsg = "%02d" % sHour, ":", "%02d" % sMinute, ":", "%02d" % sSecond
        Gobal_ServerTime_Information["text"] = strMsg

# SKQuoteLibEventHandler = win32com.client.WithEvents(SKQuoteLib, SKQuoteLibEvents)
SKQuoteEvent = SKQuoteLibEvents()
SKQuoteLibEventHandler = comtypes.client.GetEvents(skQ, SKQuoteEvent)


class SKReplyLibEvent():

    def OnReplyMessage(self, bstrUserID, bstrMessages):
        sConfirmCode = -1
        WriteMessage(bstrMessages, GlobalListInformation)
        return sConfirmCode


# comtypes使用此方式註冊callback
SKReplyEvent = SKReplyLibEvent()
SKReplyLibEventHandler = comtypes.client.GetEvents(skR, SKReplyEvent)

if __name__ == '__main__':
    # Globals.initialize()
    root = Tk()
    root.title("PythonExampleQuote")

    # Center
    FrameLogin(master=root)

    # TabControl
    root.TabControl = Notebook(root)
    root.TabControl.add(FrameQuote(master=root), text="報價功能")
    root.TabControl.grid(column=0, row=2, sticky='ew', padx=10, pady=10)

    root.mainloop()
