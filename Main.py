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
        result=str(table.cell_value(i,2))#25     价格
        type=str(table.cell_value(i,1))#DN50  规格型号
        data=ExcelData(keyArray,result,type)
        datas.append(data)

    return datas

def getresult(str):
    print(str)
    datas= getExcelData(excelUrl)

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

#判断规格型号是否对应
def getresult_2(str):
    print(str)
    datas= getExcelData(excelUrl)

    result=''
    contains_key=False
    for data in datas:

        for key in data.type:
            if(str in key):
                contains_key=True
                continue
            else:
                contains_key=False
                break
        if(contains_key):
            result=data.result

    return result

def tapkey(key,count=1):
    for i in range(0,count):
        k.tap_key(key)
        time.sleep(0.1)

def Do():
    if start:
        #主代码---------------
        maxTime = 3#3秒复制 调用copy() 不管结果对错
        while (maxTime > 0):
            maxTime = maxTime - 0.5
            time.sleep(0.5)
            print('doing')
            copy()

        targetName=pyperclip.paste()
        result = getresult(pyperclip.paste())
        if (result == ''):
            print('没有这个:' + pyperclip.paste() + ' ，需更新表格')
            # 没有找到这个 跳过 to do ---------------------
            tapkey(k.escape_key)
            tapkey(k.down_key)
            tapkey(k.left_key)
            tapkey(k.enter_key)
            return
        else:
            print('find')
            tapkey(k.enter_key,5)
            k.type_string(result)
            tapkey(k.enter_key)
            tapkey(k.escape_key)
            tapkey(k.left_key,6)
            tapkey(k.enter_key)

            return
            #判断型号 TODO---------------------
            while (maxTime > 0):
                maxTime = maxTime - 0.5
                time.sleep(0.5)
                print('doing')
                copy()
            r2 = getresult_2(pyperclip.paste())
            if(r2==''):
                print('not same type')
            else:
                print('same')

            # 找到了 输入 to do ---------------------
            # k.tap_key(k.escape_key)
            # k.tap_key(k.right_key,5)
            # k.tap_key(k.enter_key)
            # k.type_string(r)
            # k.tap_key(k.enter_key)


#我的代码
def onpressed(Key):
    while True:
        #print(Key)
        if (Key==keyboard.Key.caps_lock):#开始
            global start
            start=True
            print('go')
        if (Key==keyboard.Key.f3):#结束
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
    excelUrl = r"C:\Users\Administrator\Desktop\Xing.xlsx"#to do-------------
    threads = []
    t2 = threading.Thread(target=main, args=())
    threads.append(t2)
    for t in threads:
        t.setDaemon(True)
        t.start()
    print('press Capital to start')


    with keyboard.Listener(on_press=onpressed) as listener:
        listener.join()

    # hm = pyHook.HookManager()
    # hm.KeyDown = onKeyboardEvent
    # hm.HookKeyboard()
    # pythoncom.PumpMessages(10000)


# def onKeyboardEvent(event):
#     while True:
#         print(event.Key)
#         if str(event.Key) == 'Capital':#开始
#             global start
#             start=True
#         if str(event.Key) == 'F3':#结束
#             sys.exit()
#         return True
