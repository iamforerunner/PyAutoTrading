# -*- encoding: utf8 -*-

import tkinter.messagebox
from tkinter import *
from tkinter.ttk import *
import datetime
import threading
import pickle
import time

import win32con
import tushare as ts

from winguiauto import (dumpWindow, dumpWindows, getWindowText,
                        getListViewInfo, setEditText, clickWindow,
                        click, closePopupWindows, findTopWindow,
                        maxFocusWindow, minWindow, getTableData, sendKeyEvent)

NUM_OF_STOCKS = 5  # 自定义股票数量
is_start = False
is_monitor = True
set_stocks_info = []
actual_stocks_info = []
consignation_info = []
is_ordered = [1] * NUM_OF_STOCKS  # 1：未下单  0：已下单
is_dealt = [0] * NUM_OF_STOCKS  # 0: 未成交   负整数：卖出数量， 正整数：买入数量
stock_codes = [''] * NUM_OF_STOCKS


class OperationThs:
    def __init__(self):

        self.__top_hwnd = findTopWindow(wantedText='网上股票交易系统5.0')
        temp_hwnds = dumpWindows(self.__top_hwnd)[0][0]
        temp_hwnds = dumpWindow(temp_hwnds)[4][0]
        self.__buy_sell_hwnds = dumpWindow(temp_hwnds)
        if len(self.__buy_sell_hwnds) != 73:
            tkinter.messagebox.showerror('错误', '无法获得同花顺双向委托界面的窗口句柄')

    def __buy(self, code, quantity):
        """买函数
        :param code: 代码， 字符串
        :param quantity: 数量， 字符串
        :return:
        """
        click(self.__buy_sell_hwnds[2][0])
        time.sleep(0.2)
        setEditText(self.__buy_sell_hwnds[2][0], code)
        time.sleep(0.2)
        if quantity != '0':
            setEditText(self.__buy_sell_hwnds[7][0], quantity)
            time.sleep(0.2)
        click(self.__buy_sell_hwnds[8][0])
        time.sleep(1)

    def __sell(self, code, quantity):
        """
        卖函数
        :param code: 股票代码， 字符串
        :param quantity: 数量， 字符串
        :return:
        """
        click(self.__buy_sell_hwnds[11][0])
        time.sleep(0.2)
        setEditText(self.__buy_sell_hwnds[11][0], code)
        time.sleep(0.2)
        if quantity != '0':
            setEditText(self.__buy_sell_hwnds[16][0], quantity)
            time.sleep(0.2)
        click(self.__buy_sell_hwnds[17][0])
        time.sleep(1)

    def order(self, code, direction, quantity):
        """
        下单函数
        :param code: 股票代码， 字符串
        :param direction: 买卖方向， 字符串
        :param quantity: 买卖数量， 字符串
        :return:
        """
        # restoreFocusWindow(self.__top_hwnd)
        if direction == 'B':
            self.__buy(code, quantity)
        if direction == 'S':
            self.__sell(code, quantity)
        closePopupWindows(self.__top_hwnd)

    def maximizeFocusWindow(self):
        """
        最大化窗口，获取焦点
        """
        maxFocusWindow(self.__top_hwnd)

    def minimizeWindow(self):
        """
        最小化窗体
        """
        minWindow(self.__top_hwnd)

    def clickRefreshButton(self):
        """
        点击刷新按钮
        """
        click(self.__buy_sell_hwnds[46][0])
        time.sleep(0.5)

    def getMoney(self):
        """
        获取可用资金
        """
        return float(self.__buy_sell_hwnds[51][1])

    def getPosition(self):
        """
        获取股票持仓
        """
        clickWindow(self.__buy_sell_hwnds[-2][0], 20)
        sendKeyEvent(win32con.VK_CONTROL, 0)
        sendKeyEvent(ord('C'), 0)
        sendKeyEvent(ord('C'), win32con.KEYEVENTF_KEYUP)
        sendKeyEvent(win32con.VK_CONTROL, win32con.KEYEVENTF_KEYUP)
        return getTableData(11)

    def getDeal(self, code, pre_position, cur_position):
        """
        获取成交数量
        :param code: 需检查的股票代码， 字符串
        :param pre_position: 下单前的持仓
        :param cur_position: 下单后的持仓
        :return: 0-未成交， 正整数是买入的数量， 负整数是卖出的数量
        """
        if pre_position == cur_position:
            return 0
        pre_len = len(pre_position)
        cur_len = len(cur_position)
        if pre_len == cur_len:
            for row in range(cur_len):
                if cur_position[row][0] == code:
                    return int(cur_position[row][1]) - int(pre_position[row][1])
        if cur_len > pre_len:
            return int(cur_position[-1][1])


class OperationTdx:
    def __init__(self):

        self.__top_hwnd = findTopWindow(wantedClass='TdxW_MainFrame_Class')
        self.__button = {'refresh': 180, 'position': 145, 'deal': 112, 'withdrawal': 83, 'sell': 50, 'buy': 20}
        windows = dumpWindows(self.__top_hwnd)
        temp_hwnd = 0
        for window in windows:
            child_hwnd, window_text, window_class = window
            if window_class == 'AfxMDIFrame42':
                temp_hwnd = child_hwnd
                break
        temp_hwnds = dumpWindow(temp_hwnd)
        temp_hwnds = dumpWindow(temp_hwnds[1][0])
        self.__menu_hwnds = dumpWindow(temp_hwnds[0][0])
        self.__buy_sell_hwnds = dumpWindow(temp_hwnds[4][0])
        if len(self.__buy_sell_hwnds) != 68:
            tkinter.messagebox.showerror('错误', '无法获得通达信对买对卖界面的窗口句柄')

    def __buy(self, code, quantity):
        """
        买入函数
        :param code: 股票代码，字符串
        :param quantity: 数量， 字符串
        """
        setEditText(self.__buy_sell_hwnds[0][0], code)
        time.sleep(0.2)
        if quantity != '0':
            setEditText(self.__buy_sell_hwnds[3][0], quantity)
            time.sleep(0.2)
        click(self.__buy_sell_hwnds[5][0])
        time.sleep(0.2)

    def __sell(self, code, quantity):
        """
        卖出函数
        :param code: 股票代码， 字符串
        :param quantity: 数量， 字符串
        """
        setEditText(self.__buy_sell_hwnds[24][0], code)
        time.sleep(0.2)
        if quantity != '0':
            setEditText(self.__buy_sell_hwnds[27][0], quantity)
            time.sleep(0.2)
        click(self.__buy_sell_hwnds[29][0])
        time.sleep(0.2)

    def order(self, code, direction, quantity):
        """
        下单函数
        :param code: 股票代码， 字符串
        :param direction: 买卖方向
        :param quantity: 数量， 字符串，数量为‘0’时，由交易软件指定数量
        """
        # restoreFocusWindow(self.__top_hwnd)
        if direction == 'B':
            self.__buy(code, quantity)
        if direction == 'S':
            self.__sell(code, quantity)
        closePopupWindows(self.__top_hwnd)

    def maximizeFocusWindow(self):
        '''
        最大化窗口，获取焦点
        '''
        maxFocusWindow(self.__top_hwnd)

    def minimizeWindow(self):
        """
        最小化窗体
        """
        minWindow(self.__top_hwnd)

    def clickRefreshButton(self):
        """点击刷新按钮
        """
        clickWindow(self.__menu_hwnds[0][0], self.__button['refresh'])
        time.sleep(0.5)

    def getMoney(self):
        """获取可用资金
        """
        setEditText(self.__buy_sell_hwnds[24][0], '999999')  # 测试时获得资金情况
        time.sleep(0.2)
        money = getWindowText(self.__buy_sell_hwnds[12][0]).strip()
        return float(money)

    def getPosition(self):
        """获取持仓股票信息
        """
        return getListViewInfo(self.__buy_sell_hwnds[-4][0], 3)

    def getDeal(self, code, pre_position, cur_position):
        """
        获取成交数量
        :param code: 股票代码
        :param pre_position: 下单前的持仓
        :param cur_position: 下单后的持仓
        :return: 0-未成交， 正整数是买入的数量， 负整数是卖出的数量
        """
        if pre_position == cur_position:
            return 0
        pre_len = len(pre_position)
        cur_len = len(cur_position)
        if pre_len == cur_len:
            for row in range(cur_len):
                if cur_position[row][0] == code:
                    return int(cur_position[row][1]) - int(pre_position[row][1])
        if cur_len > pre_len:
            return int(cur_position[-1][1])


def pickCodeFromItems(items_info):
    """
    提取股票代码
    :param items_info: UI下各项输入信息
    :return:股票代码列表
    """
    stock_codes = []
    for item in items_info:
        stock_codes.append(item[0])
    return stock_codes


# def getStockData(items_info):
#     """
#     获取股票实时数据
#     :param items_info: 股票信息，没写的股票用空字符代替
#     :return: 股票名称价格
#     """
#     global stock_codes
#     code_name_price = []
#     try:
#         df = ts.get_realtime_quotes(stock_codes)
#         df_len = len(df)
#         for stock_code in stock_codes:
#             is_found = False
#             for i in range(df_len):
#                 actual_code = df['code'][i]
#                 if stock_code == actual_code:
#                     actual_name = df['name'][i]
#                     pre_close = float(df['pre_close'][i])
#                     if 'ST' in actual_name:
#                         highest = str(round(pre_close * 1.05, 2))
#                         lowest = str(round(pre_close * 0.95, 2))
#                         code_name_price.append((actual_code, actual_name, float(df['price'][i]), (highest, lowest)))
#                     else:
#                         highest = str(round(pre_close * 1.1, 2))
#                         lowest = str(round(pre_close * 0.9, 2))
#                         code_name_price.append((actual_code, actual_name, float(df['price'][i]), (highest, lowest)))
#                     is_found = True
#                     break
#             if is_found is False:
#                 code_name_price.append(('', '', '', ('', '')))
#     except:
#         code_name_price = [('', '', '', ('', ''))] * NUM_OF_STOCKS
#     return code_name_price


def getStockData():
    """
    获取股票实时数据
    :return:股票实时数据
    """
    global stock_codes
    code_name_price = []
    try:
        df = ts.get_realtime_quotes(stock_codes)
        df_len = len(df)
        for stock_code in stock_codes:
            is_found = False
            for i in range(df_len):
                actual_code = df['code'][i]
                if stock_code == actual_code:
                    code_name_price.append((actual_code, df['name'][i], float(df['price'][i])))
                    is_found = True
                    break
            if is_found is False:
                code_name_price.append(('', '', 0))
    except:
        code_name_price = [('', '', 0)] * NUM_OF_STOCKS  # 网络不行，返回空
    return code_name_price


def monitor():
    """
    实时监控函数
    :return:
    """
    global actual_stocks_info, consignation_info, is_ordered, is_dealt, set_stocks_info
    count = 1
    try:
        try:
            operation = OperationTdx()
        except:
            operation = OperationThs()
    except:
        tkinter.messagebox.showerror('错误', '无法获得交易软件句柄')

    while is_monitor:

        if is_start:
            actual_stocks_info = getStockData()
            for row, (actual_code, actual_name, actual_price) in enumerate(actual_stocks_info):
                if actual_code and is_start and is_ordered[row] == 1 and actual_price > 0 \
                        and set_stocks_info[row][1] and set_stocks_info[row][2] > 0 \
                        and set_stocks_info[row][3] and set_stocks_info[row][4] \
                        and datetime.datetime.now().time() > set_stocks_info[row][5]:
                    if set_stocks_info[row][1] == '>' and actual_price > set_stocks_info[row][2]:
                        operation.maximizeFocusWindow()
                        pre_position = operation.getPosition()
                        operation.order(actual_code, set_stocks_info[row][3], set_stocks_info[row][4])
                        dt = datetime.datetime.now()
                        is_ordered[row] = 0
                        # time.sleep(1)
                        operation.clickRefreshButton()
                        cur_position = operation.getPosition()
                        is_dealt[row] = operation.getDeal(actual_code, pre_position, cur_position)
                        consignation_info.append(
                            (dt.strftime('%x'), dt.strftime('%X'), actual_code,
                             actual_name, set_stocks_info[row][3],
                             actual_price, set_stocks_info[row][4], '已委托', is_dealt[row]))

                    if set_stocks_info[row][1] == '<' and float(actual_price) < set_stocks_info[row][2]:
                        operation.maximizeFocusWindow()
                        pre_position = operation.getPosition()
                        operation.order(actual_code, set_stocks_info[row][3], set_stocks_info[row][4])
                        dt = datetime.datetime.now()
                        is_ordered[row] = 0
                        # time.sleep(1)
                        operation.clickRefreshButton()
                        cur_position = operation.getPosition()
                        is_dealt[row] = operation.getDeal(actual_code, pre_position, cur_position)
                        consignation_info.append(
                            (dt.strftime('%x'), dt.strftime('%X'), actual_code,
                             actual_name, set_stocks_info[row][3],
                             actual_price, set_stocks_info[row][4], '已委托', is_dealt[row]))

        if count % 200 == 0:
            operation.clickRefreshButton()
            time.sleep(2)
        time.sleep(3)
        count += 1


class StockGui:
    def __init__(self):
        self.window = Tk()
        self.window.title("自动化股票交易")
        self.window.resizable(0, 0)

        frame1 = Frame(self.window)
        frame1.pack(padx=10, pady=10)

        Label(frame1, text="股票代码", width=8, justify=CENTER).grid(
            row=1, column=1, padx=5, pady=5)
        Label(frame1, text="股票名称", width=8, justify=CENTER).grid(
            row=1, column=2, padx=5, pady=5)
        Label(frame1, text="实时价格", width=8, justify=CENTER).grid(
            row=1, column=3, padx=5, pady=5)
        Label(frame1, text="关系", width=4, justify=CENTER).grid(
            row=1, column=4, padx=5, pady=5)
        Label(frame1, text="设定价格", width=8, justify=CENTER).grid(
            row=1, column=5, padx=5, pady=5)
        Label(frame1, text="方向", width=4, justify=CENTER).grid(
            row=1, column=6, padx=5, pady=5)
        Label(frame1, text="数量", width=8, justify=CENTER).grid(
            row=1, column=7, padx=5, pady=5)
        Label(frame1, text="时间可选", width=8, justify=CENTER).grid(
            row=1, column=8, padx=5, pady=5)
        Label(frame1, text="委托", width=6, justify=CENTER).grid(
            row=1, column=9, padx=5, pady=5)
        Label(frame1, text="成交", width=6, justify=CENTER).grid(
            row=1, column=10, padx=5, pady=5)

        self.rows = NUM_OF_STOCKS
        self.cols = 10

        self.variable = []
        for row in range(self.rows):
            self.variable.append([])
            for col in range(self.cols):
                temp = StringVar()
                self.variable[row].append(temp)

        for row in range(self.rows):
            Entry(frame1, textvariable=self.variable[row][0],
                  width=8).grid(row=row + 2, column=1, padx=5, pady=5)
            Entry(frame1, textvariable=self.variable[row][1], state=DISABLED,
                  width=8).grid(row=row + 2, column=2, padx=5, pady=5)
            Entry(frame1, textvariable=self.variable[row][2], state=DISABLED, justify=RIGHT,
                  width=8).grid(row=row + 2, column=3, padx=5, pady=5)
            Combobox(frame1, values=('<', '>'), textvariable=self.variable[row][3],
                     width=2).grid(row=row + 2, column=4, padx=5, pady=5)
            Spinbox(frame1, from_=0, to=1000, textvariable=self.variable[row][4], justify=RIGHT,
                    increment=0.01, width=6).grid(row=row + 2, column=5, padx=5, pady=5)
            Combobox(frame1, values=('B', 'S'), textvariable=self.variable[row][5],
                     width=2).grid(row=row + 2, column=6, padx=5, pady=5)
            Spinbox(frame1, from_=0, to=100000, textvariable=self.variable[row][6], justify=RIGHT,
                    increment=100, width=6).grid(row=row + 2, column=7, padx=5, pady=5)
            Entry(frame1, textvariable=self.variable[row][7],
                  width=8).grid(row=row + 2, column=8, padx=5, pady=5)
            Entry(frame1, textvariable=self.variable[row][8], state=DISABLED, justify=CENTER,
                  width=6).grid(row=row + 2, column=9, padx=5, pady=5)
            Entry(frame1, textvariable=self.variable[row][9], state=DISABLED, justify=RIGHT,
                  width=6).grid(row=row + 2, column=10, padx=5, pady=5)

        frame3 = Frame(self.window)
        frame3.pack(padx=10, pady=10)
        self.start_bt = Button(frame3, text="开始", command=self.start)
        self.start_bt.pack(side=LEFT)
        self.set_bt = Button(frame3, text='重置买卖', command=self.setFlags)
        self.set_bt.pack(side=LEFT)
        Button(frame3, text="历史记录", command=self.displayHisRecords).pack(side=LEFT)
        Button(frame3, text='保存', command=self.save).pack(side=LEFT)
        self.load_bt = Button(frame3, text='载入', command=self.load)
        self.load_bt.pack(side=LEFT)

        self.window.protocol(name="WM_DELETE_WINDOW", func=self.close)
        self.window.after(100, self.updateControls)
        self.window.mainloop()

    def displayHisRecords(self):
        """
        显示历史信息
        :return:
        """
        global consignation_info
        tp = Toplevel()
        tp.title('历史记录')
        tp.resizable(0, 1)
        scrollbar = Scrollbar(tp)
        scrollbar.pack(side=RIGHT, fill=Y)
        col_name = ['日期', '时间', '证券代码', '证券名称', '方向', '价格', '数量', '委托', '成交']
        tree = Treeview(
            tp, show='headings', columns=col_name, height=30, yscrollcommand=scrollbar.set)
        tree.pack(expand=1, fill=Y)
        scrollbar.config(command=tree.yview)
        for name in col_name:
            tree.heading(name, text=name)
            tree.column(name, width=70, anchor=CENTER)

        for msg in consignation_info:
            tree.insert('', 0, values=msg)

    def save(self):
        """
        保存设置
        :return:
        """
        global set_stocks_info, consignation_info
        self.getItems()
        with open('stockInfo.dat', 'wb') as fp:
            pickle.dump(set_stocks_info, fp)
            pickle.dump(consignation_info, fp)

    def load(self):
        """
        载入设置
        :return:
        """
        global set_stocks_info, consignation_info
        try:
            with open('stockInfo.dat', 'rb') as fp:
                set_stocks_info = pickle.load(fp)
                consignation_info = pickle.load(fp)
        except FileNotFoundError as error:
            tkinter.messagebox.showerror('错误', error)

        for row in range(self.rows):
            for col in range(self.cols):
                if col == 0:
                    self.variable[row][col].set(set_stocks_info[row][0])
                elif col == 3:
                    self.variable[row][col].set(set_stocks_info[row][1])
                elif col == 4:
                    self.variable[row][col].set(set_stocks_info[row][2])
                elif col == 5:
                    self.variable[row][col].set(set_stocks_info[row][3])
                elif col == 6:
                    self.variable[row][col].set(set_stocks_info[row][4])
                elif col == 7:
                    temp = set_stocks_info[row][5].strftime('%X')
                    if temp == '01:00:00':
                        self.variable[row][col].set('')
                    else:
                        self.variable[row][col].set(temp)

    def setFlags(self):
        """
        重置买卖标志
        :return:
        """
        global is_start, is_ordered
        if is_start is False:
            is_ordered = [1] * NUM_OF_STOCKS

    def updateControls(self):
        """
        实时股票名称、价格、状态信息
        :return:
        """
        global actual_stocks_info, is_start
        if is_start:
            for row, (actual_code, actual_name, actual_price) in enumerate(actual_stocks_info):
                if actual_code:
                    self.variable[row][1].set(actual_name)
                    self.variable[row][2].set(str(actual_price))
                    if is_ordered[row] == 1:
                        self.variable[row][8].set('监控中')
                    elif is_ordered[row] == 0:
                        self.variable[row][8].set('已委托')
                    self.variable[row][9].set(str(is_dealt[row]))
                else:
                    self.variable[row][1].set('')
                    self.variable[row][2].set('')
                    self.variable[row][8].set('')
                    self.variable[row][9].set('')

        self.window.after(3000, self.updateControls)

    def start(self):
        """
        启动停止
        :return:
        """
        global is_start, stock_codes, set_stocks_info
        if is_start is False:
            is_start = True
        else:
            is_start = False

        if is_start:
            self.getItems()
            stock_codes = pickCodeFromItems(set_stocks_info)
            self.start_bt['text'] = '停止'
            self.set_bt['state'] = DISABLED
            self.load_bt['state'] = DISABLED
        else:
            self.start_bt['text'] = '开始'
            self.set_bt['state'] = NORMAL
            self.load_bt['state'] = NORMAL

    def close(self):
        """
        关闭程序时，停止monitor线程
        :return:
        """
        global is_monitor
        is_monitor = False
        self.window.quit()

    def getItems(self):
        """
        获取UI上用户输入的各项数据，
        """
        global set_stocks_info
        set_stocks_info = []

        # 获取买卖价格数量输入项等
        for row in range(self.rows):
            set_stocks_info.append([])
            for col in range(self.cols):
                temp = self.variable[row][col].get().strip()
                if col == 0:
                    if len(temp) == 6 and temp.isdigit():  # 判断股票代码是否为6位数
                        set_stocks_info[row].append(temp)
                    else:
                        set_stocks_info[row].append('')
                elif col == 3:
                    if temp in ('>', '<'):
                        set_stocks_info[row].append(temp)
                    else:
                        set_stocks_info[row].append('')
                elif col == 4:
                    try:
                        price = float(temp)
                        if price > 0:
                            set_stocks_info[row].append(price)  # 把价格转为数字
                        else:
                            set_stocks_info[row].append(0)
                    except ValueError:
                        set_stocks_info[row].append(0)
                elif col == 5:
                    if temp in ('B', 'S'):
                        set_stocks_info[row].append(temp)
                    else:
                        set_stocks_info[row].append('')
                elif col == 6:
                    if temp.isdigit() and int(temp) >= 0:
                        set_stocks_info[row].append(str(int(temp) // 100 * 100))
                    else:
                        set_stocks_info[row].append('')
                elif col == 7:
                    try:
                        set_stocks_info[row].append(datetime.datetime.strptime(temp, '%H:%M:%S').time())
                    except ValueError:
                        set_stocks_info[row].append(datetime.datetime.strptime('1:00:00', '%H:%M:%S').time())


if __name__ == '__main__':
    t1 = threading.Thread(target=StockGui)
    t1.start()
    t1.join(2)
    t2 = threading.Thread(target=monitor)
    t2.start()
