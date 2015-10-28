# -*- encoding: utf8 -*-
# QQ群： 486224275
__author__ = '人在江湖'
__email__ = 'ronghui.ding@outlook.com'

import time
import tkinter.messagebox
from tkinter import *
from tkinter.ttk import *
import datetime
import threading
import pickle

import win32con
import tushare as ts

from winguiauto import (findSpecifiedWindows, findPopupWindow,
                        findControl, clickButton, click, setEditText,
                        findSpecifiedTopWindow, sendKey)

is_start = False
is_monitor = True
set_stock_info = []
order_msg = []
actual_stock_info = []
is_ordered = [1] * 5  # 1：准备  0：交易成功 -1：交易失败


def pickHwndOfControls(top_hwnd, num_child_windows):
    cleaned_hwnd_controls = []
    hwnd_controls = findSpecifiedWindows(top_hwnd, num_child_windows)
    for Hwnd, text_name, class_name in hwnd_controls:
        if class_name in ('Button', 'Edit'):
            cleaned_hwnd_controls.append((Hwnd, text_name, class_name))
    return cleaned_hwnd_controls


def closePopupWindow(top_hwnd, wantedText=None, wantedClass=None):
    # 如果有弹出式窗口，点击它的确定按钮
    hwnd_popup = findPopupWindow(top_hwnd)
    if hwnd_popup:
        hwnd_control = findControl(hwnd_popup, wantedText, wantedClass)
        clickButton(hwnd_control)
        time.sleep(1)
        return True
    return False


def buy(hwnd_lst, code, stop_price, quantity):
    click(hwnd_lst[0][0])
    setEditText(hwnd_lst[0][0], code)
    setEditText(hwnd_lst[1][0], stop_price)
    time.sleep(0.2)
    setEditText(hwnd_lst[2][0], quantity)
    time.sleep(0.2)
    clickButton(hwnd_lst[3][0])
    time.sleep(1)


def sell(hwnd_lst, code, stop_price, quantity):
    click(hwnd_lst[4][0])
    setEditText(hwnd_lst[4][0], code)
    setEditText(hwnd_lst[5][0], stop_price)
    time.sleep(0.2)
    setEditText(hwnd_lst[6][0], quantity)
    time.sleep(0.2)
    clickButton(hwnd_lst[7][0])
    time.sleep(1)


def order(top_hwnd, hwnd_lst, code, stop_prices, quantity, direction):
    # 检测交易软件是否挂起或出错
    if closePopupWindow(top_hwnd, wantedClass='Button'):
        time.sleep(5)
    if direction == 'B':
        buy(hwnd_lst, code, stop_prices[0], quantity)
        return not closePopupWindow(top_hwnd, wantedClass='Button')
    if direction == 'S':
        sell(hwnd_lst, code, stop_prices[1], quantity)
        return not closePopupWindow(top_hwnd, wantedClass='Button')


def tradingInit():
    # 获取交易软件及子控件句柄
    hwnd = findSpecifiedTopWindow(wantedText='网上股票交易系统5.0')
    if hwnd == 0:
        tkinter.messagebox.showerror('错误', '请先打开华泰证券交易软件，再运行本软件')
        return hwnd, []
    else:
        sendKey(hwnd, win32con.VK_F6)
        hwnd_child_controls = pickHwndOfControls(hwnd, 70)
    return hwnd, hwnd_child_controls


def pickCodeFromItems(items_info):
    stock_codes = []
    for item in items_info:
        stock_codes.append(item[0])
    return stock_codes


def getStockData(items_info):
    '''
    获取股票数据
    :param items_info:
    :return:
    '''
    code_name_price = []
    stock_codes = pickCodeFromItems(items_info)
    try:
        df = ts.get_realtime_quotes(stock_codes)
        df_len = len(df)
        for stock_code in stock_codes:
            is_found = False
            for i in range(df_len):
                actual_code = df['code'][i]
                if stock_code == actual_code:
                    actual_name = df['name'][i]
                    pre_close = float(df['pre_close'][i])
                    if 'ST' in actual_name:
                        highest = str(round(pre_close * 1.05, 2))
                        lowest = str(round(pre_close * 0.95, 2))
                        code_name_price.append((actual_code, actual_name, df['price'][i], (highest, lowest)))
                    else:
                        highest = str(round(pre_close * 1.1, 2))
                        lowest = str(round(pre_close * 0.9, 2))
                        code_name_price.append((actual_code, actual_name, df['price'][i], (highest, lowest)))
                    is_found = True
                    break
            if is_found is False:
                code_name_price.append(('', '', '', ('', '')))
    except:
        code_name_price = [('', '', '', ('', ''))]
    return code_name_price


def monitor():
    # 股价监控函数
    global actual_stock_info, order_msg, is_ordered, set_stock_info
    count = 1
    hwnd, hwnd_child_controls = tradingInit()
    # 如果hwnd为零，直接终止循环
    while is_monitor and hwnd:
        if count % 100 == 0:
            clickButton(hwnd_child_controls[12][0])  # 点击刷新按钮
            time.sleep(1)
        time.sleep(3)
        count += 1
        if is_start:
            actual_stock_info = getStockData(set_stock_info)
            for row, (actual_code, actual_name, actual_price, stop_prices) in enumerate(actual_stock_info):
                if is_start and actual_code and is_ordered[row] == 1 \
                        and set_stock_info[row][1] and set_stock_info[row][2] > 0 \
                        and set_stock_info[row][3] and set_stock_info[row][4] \
                        and datetime.datetime.now().time() > set_stock_info[row][5]:
                    if is_start and set_stock_info[row][1] == '>' and float(actual_price) > set_stock_info[row][2]:
                        dt = datetime.datetime.now()
                        if order(hwnd, hwnd_child_controls, actual_code, stop_prices,
                                 set_stock_info[row][4], set_stock_info[row][3]):
                            order_msg.append(
                                (dt.strftime('%x'), dt.strftime('%X'), actual_code,
                                 actual_name, set_stock_info[row][3],
                                 actual_price, set_stock_info[row][4], '成功'))
                            is_ordered[row] = 0
                        else:
                            order_msg.append(
                                (dt.strftime('%x'), dt.strftime('%X'), actual_code,
                                 actual_name, set_stock_info[row][3],
                                 actual_price, set_stock_info[row][4], '失败'))
                            is_ordered[row] = -1
                    if is_start and set_stock_info[row][1] == '<' and float(actual_price) < set_stock_info[row][2]:
                        dt = datetime.datetime.now()
                        if order(hwnd, hwnd_child_controls, actual_code, stop_prices,
                                 set_stock_info[row][4], set_stock_info[row][3]):
                            order_msg.append(
                                (dt.strftime('%x'), dt.strftime('%X'), actual_code,
                                 actual_name, set_stock_info[row][3],
                                 actual_price, set_stock_info[row][4], '成功'))
                            is_ordered[row] = 0
                        else:
                            order_msg.append(
                                (dt.strftime('%x'), dt.strftime('%X'), actual_code,
                                 actual_name, set_stock_info[row][3],
                                 actual_price, set_stock_info[row][4], '失败'))
                            is_ordered[row] = -1


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
        Label(frame1, text="当前价格", width=8, justify=CENTER).grid(
            row=1, column=3, padx=5, pady=5)
        Label(frame1, text="关系", width=4, justify=CENTER).grid(
            row=1, column=4, padx=5, pady=5)
        Label(frame1, text="价格", width=8, justify=CENTER).grid(
            row=1, column=5, padx=5, pady=5)
        Label(frame1, text="方向", width=4, justify=CENTER).grid(
            row=1, column=6, padx=5, pady=5)
        Label(frame1, text="数量", width=8, justify=CENTER).grid(
            row=1, column=7, padx=5, pady=5)
        Label(frame1, text="时间可选", width=8, justify=CENTER).grid(
            row=1, column=8, padx=5, pady=5)
        Label(frame1, text="状态", width=4, justify=CENTER).grid(
            row=1, column=9, padx=5, pady=5)

        self.rows = 5
        self.cols = 9

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
            Entry(frame1, textvariable=self.variable[row][2], state=DISABLED,
                  width=8).grid(row=row + 2, column=3, padx=5, pady=5)
            Combobox(frame1, values=('<', '>'), textvariable=self.variable[row][3],
                     width=2).grid(row=row + 2, column=4, padx=5, pady=5)
            Spinbox(frame1, from_=0, to=1000, textvariable=self.variable[row][4],
                    increment=0.01, width=6).grid(row=row + 2, column=5, padx=5, pady=5)
            Combobox(frame1, values=('B', 'S'), textvariable=self.variable[row][5],
                     width=2).grid(row=row + 2, column=6, padx=5, pady=5)
            Spinbox(frame1, from_=0, to=100000, textvariable=self.variable[row][6],
                    increment=100, width=6).grid(row=row + 2, column=7, padx=5, pady=5)
            Entry(frame1, textvariable=self.variable[row][7],
                  width=8).grid(row=row + 2, column=8, padx=5, pady=5)
            Entry(frame1, textvariable=self.variable[row][8], state=DISABLED,
                  width=4).grid(row=row + 2, column=9, padx=5, pady=5)

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
        '''
        显示历史信息
        :return:
        '''
        global order_msg
        tp = Toplevel()
        tp.title('历史记录')
        tp.resizable(0, 1)
        scrollbar = Scrollbar(tp)
        scrollbar.pack(side=RIGHT, fill=Y)
        col_name = ['日期', '时间', '证券代码', '证券名称', '方向', '价格', '数量', '备注']
        tree = Treeview(
            tp, show='headings', columns=col_name, height=30, yscrollcommand=scrollbar.set)
        tree.pack(expand=1, fill=Y)
        scrollbar.config(command=tree.yview)
        for name in col_name:
            tree.heading(name, text=name)
            tree.column(name, width=70, anchor=CENTER)

        for msg in order_msg:
            tree.insert('', 0, values=msg)

    def save(self):
        '''
        保存设置
        :return:
        '''
        global set_stock_info, order_msg, actual_stock_info
        self.getItems()
        with open('stockInfo.dat', 'wb') as fp:
            pickle.dump(set_stock_info, fp)
            pickle.dump(actual_stock_info, fp)
            pickle.dump(order_msg, fp)

    def load(self):
        '''
        载入设置
        :return:
        '''
        global set_stock_info, order_msg, actual_stock_info
        with open('stockInfo.dat', 'rb') as fp:
            set_stock_info = pickle.load(fp)
            actual_stock_info = pickle.load(fp)
            order_msg = pickle.load(fp)
        for row in range(self.rows):
            for col in range(self.cols):
                if col == 0:
                    self.variable[row][col].set(set_stock_info[row][0])
                elif col == 3:
                    self.variable[row][col].set(set_stock_info[row][1])
                elif col == 4:
                    self.variable[row][col].set(set_stock_info[row][2])
                elif col == 5:
                    self.variable[row][col].set(set_stock_info[row][3])
                elif col == 6:
                    self.variable[row][col].set(set_stock_info[row][4])
                elif col == 7:
                    temp = set_stock_info[row][5].strftime('%X')
                    if temp == '01:00:00':
                        self.variable[row][col].set('')
                    else:
                        self.variable[row][col].set(temp)

    def setFlags(self):
        '''
        重置买卖标志
        :return:
        '''
        global is_start, is_ordered
        if is_start is False:
            is_ordered = [1] * 5

    def updateControls(self):
        '''
        实时股票名称、价格、状态信息
        :return:
        '''
        global set_stock_info, actual_stock_info, is_start
        if is_start:
            for row, (actual_code, actual_name, actual_price, _) in enumerate(actual_stock_info):
                if actual_code:
                    self.variable[row][1].set(actual_name)
                    self.variable[row][2].set(str(actual_price))
                    if is_ordered[row] == 1:
                        self.variable[row][8].set('监控')
                    elif is_ordered[row] == -1:
                        self.variable[row][8].set('失败')
                    elif is_ordered[row] == 0:
                        self.variable[row][8].set('成功')

        self.window.after(3000, self.updateControls)

    def start(self):
        global is_start

        if is_start is False:
            is_start = True
        else:
            is_start = False

        if is_start:
            self.getItems()
            # print(set_stock_info)
            self.start_bt['text'] = '停止'
            self.set_bt['state'] = DISABLED
            self.load_bt['state'] = DISABLED
        else:
            self.start_bt['text'] = '开始'
            self.set_bt['state'] = NORMAL
            self.load_bt['state'] = NORMAL

    def close(self):
        # 关闭软件时，停止monitor线程
        global is_monitor
        is_monitor = False
        self.window.quit()

    def getItems(self):
        global set_stock_info
        set_stock_info = []

        # 获取买卖价格数量输入项等
        for row in range(self.rows):
            set_stock_info.append([])
            for col in range(self.cols):
                temp = self.variable[row][col].get().strip()
                if col == 0:
                    if len(temp) == 6 and temp.isdigit():  # 判断股票代码是否为6位数
                        set_stock_info[row].append(temp)
                    else:
                        set_stock_info[row].append('')
                elif col == 3:
                    if temp in ('>', '<'):
                        set_stock_info[row].append(temp)
                    else:
                        set_stock_info[row].append('')
                elif col == 4:
                    try:
                        price = float(temp)
                        if price > 0:
                            set_stock_info[row].append(price)  # 把价格转为数字
                        else:
                            set_stock_info[row].append(0)
                    except ValueError:
                        set_stock_info[row].append(0)
                elif col == 5:
                    if temp in ('B', 'S'):
                        set_stock_info[row].append(temp)
                    else:
                        set_stock_info[row].append('')
                elif col == 6:
                    if temp.isdigit() and int(temp) >= 100:
                        set_stock_info[row].append(str(int(temp) // 100 * 100))
                    else:
                        set_stock_info[row].append('')
                elif col == 7:
                    try:
                        set_stock_info[row].append(datetime.datetime.strptime(temp, '%H:%M:%S').time())
                    except ValueError:
                        set_stock_info[row].append(datetime.datetime.strptime('1:00:00', '%H:%M:%S').time())


if __name__ == '__main__':
    t1 = threading.Thread(target=StockGui)
    t2 = threading.Thread(target=monitor)
    t1.start()
    t2.start()
    t1.join()
    t2.join()
