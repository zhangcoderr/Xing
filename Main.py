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
    def __init__(self,keyArray,result,type,compareType=''):
        self.keyArray=keyArray
        self.result=result
        self.type=type
        self.compareType=compareType
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
        compareType=str(table.cell_value(i,3))#0 1 匹配方式 1是完全匹配 比如 水==水
        data=ExcelData(keyArray,result,type,compareType)
        datas.append(data)

    return datas

def getresult(str):
    datas= getExcelData(excelUrl)

    dataResult=ExcelData('','','')
    contains_key=False
    for data in datas:

        for key in data.keyArray:
            if(data.compareType=='1.0'):
                if(key==str):
                    contains_key=True
                    break
                else:
                    contains_key=False
            elif(data.compareType==''):
                if(key in str):
                    contains_key=True
                    continue
                else:
                    contains_key=False
                    break
        if(contains_key):
            dataResult=data
            break

    return dataResult

#判断规格型号是否对应
def getresult_2(name,type):
    print(type)
    datas= getExcelData(excelUrl)

    result=''
    resultDatas=[]
    for data in datas:
        contains_key = False

        for key in data.keyArray:
            if(key in name):
                contains_key = True
                continue
            else:
                contains_key = False
                break
        if(contains_key):
            resultDatas.append(data)

    for d in resultDatas:
        if (d.type in type or type == d.type):
            result=d.result
        elif(d.type in name):
            result=d.result

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
            #print('doing')
            copy()

        targetName=pyperclip.paste()

        dataResult=getresult(pyperclip.paste())
        if (dataResult.result == ''):
            print('没有这个:' + pyperclip.paste() + ' ，需更新表格')
            # 没有找到这个 跳过 to do ---------------------
            tapkey(k.escape_key)
            tapkey(k.down_key)
            tapkey(k.left_key)
            tapkey(k.enter_key)
            return
        else:

            if(dataResult.type==''):
                print('无规格直接输入')
                tapkey(k.enter_key, 5)
                k.type_string(dataResult.result)
                tapkey(k.enter_key)
                time.sleep(2)

                tapkey(k.escape_key)
                tapkey(k.left_key, 6)
                tapkey(k.enter_key)
            else:
                print('表格匹配名字，判断且有规格')
                tapkey(k.enter_key)
                maxTime = 3
                while (maxTime > 0):
                    maxTime = maxTime - 0.5
                    time.sleep(0.5)
                    #print('doing')
                    copy()
                targetType = pyperclip.paste()
                if(targetType==targetName):#相等则targetType为空 规格型号没填 复制为上次结果
                    targetType=''
                result_2 = getresult_2(targetName, targetType)
                if(result_2==''):
                    print('无匹配   '+targetName+'   无匹配   '+targetType)
                    tapkey(k.escape_key)
                    tapkey(k.down_key)
                    tapkey(k.left_key,2)
                    tapkey(k.enter_key)
                else:
                    print('规格匹配输入'+targetName+'     '+targetType)
                    tapkey(k.enter_key,4)
                    k.type_string(result_2)
                    tapkey(k.enter_key)
                    time.sleep(2)
                    tapkey(k.escape_key)
                    tapkey(k.left_key, 6)
                    tapkey(k.enter_key)


            #判断型号 TODO---------------------
            return
            # tapkey(k.enter_key)
            # while (maxTime > 0):
            #     maxTime = maxTime - 0.5
            #     time.sleep(0.5)
            #     print('doing')
            #     copy()
            # targetType = pyperclip.paste()
            # result_2= getresult_2(targetName, targetType)
            # if(result_2==''):
            #
            # if(targetType==''):
            #     tapkey(k.escape_key)
            #     tapkey(k.down_key)
            #     tapkey(k.left_key)
            #     tapkey(k.enter_key)
            #
            # else:
            #
            #     r2 = getresult_2(targetName,targetType)
            #     if(r2==''):
            #         print('not same type')
            #     else:
            #         print('same')

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
    excelUrl = r"C:\Users\123\Desktop\广联达\安装\Xing.xlsx"#to do-------------
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
