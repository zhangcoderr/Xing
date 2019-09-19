# coding=utf-8
from pykeyboard import PyKeyboard
from pymouse import PyMouse
import time
import pyHook
import pythoncom
import xlrd
import  pyperclip
from pynput import mouse,keyboard
import threading
import sys
def copy():

    k.press_key(k.control_l_key)
    k.tap_key("C")
    k.release_key(k.control_l_key)


class ExcelData:
    def __init__(self,keyArray,result,type):
        self.keyArray=keyArray
        self.result=result
        self.type=type

def getExcelData(excelUrl):
    datas=[]
    excel = xlrd.open_workbook(excelUrl)
    table = excel.sheets()[0]
    rowCount = table.nrows
    colCount = table.ncols


    for i in range(rowCount):
        if(i==0): continue
        keyArray=str(table.cell_value(i,0)).split('$')#不锈钢$地漏
        result=str(table.cell_value(i,1))#25
        type=str(table.cell_value(i,2))#暂定
        data=ExcelData(keyArray,result,type)
        datas.append(data)

    return datas

def getresult(str):
    print(str)
    datas= getExcelData()

    result=''
    contains_key=False
    for data in datas:

        for key in data.keyArray:
            if(key in str):
                contains_key=True
                continue
            else:
                contains_key=False
                break
        if(contains_key):
            result=data.result

    return result

def KeyBoardMove(key,count):
    for i in count:
        k.tap_key(key)

def Do():
    if start:
        print(1)
        #主代码---------------
        last = pyperclip.paste()
        maxTime = 3
        while (pyperclip.paste() == last and maxTime > 0):
            maxTime = maxTime - 1
            time.sleep(0.5)
            print('doing')
            copy()
        if (maxTime > 0):
            r = getresult(pyperclip.paste())
            if (r == ''):
                print('没有这个:' + pyperclip.paste() + ' ，需更新表格')
                #没有找到这个 跳过 to do ---------------------
                return
            else:
                #找到了 输入 to do ---------------------
                k.tap_key(k.escape_key)
                k.tap_key(k.right_key,5)
                k.tap_key(k.enter_key)
                k.type_string(r)
                k.tap_key(k.enter_key)
        else:
            print('add maxTime!!!!!!!!!!!')




#我的代码

def onKeyboardEvent(event):
    while True:
        if str(event.Key) == 'Capital':#开始
            global start
            start=True
        if str(event.Key) == 'Escape':#结束
            sys.exit()
        return True

def main():
    while True:

        #主程序在这
        Do()



if __name__ == '__main__':
    k = PyKeyboard()
    m = PyMouse()
    start=False
    excelUrl = r"C:\Users\Administrator\Desktop\guang.xlsx"#to do-------------
    threads = []
    t2 = threading.Thread(target=main, args=())
    threads.append(t2)
    for t in threads:
        t.setDaemon(True)
        t.start()
    print('press Capital to start')

    hm = pyHook.HookManager()
    hm.KeyDown = onKeyboardEvent
    hm.HookKeyboard()
    pythoncom.PumpMessages(10000)
    # with keyboard.Listener(on_press=onpressed) as listener:
    #     listener.join()
