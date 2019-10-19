# coding=utf-8
from pykeyboard import PyKeyboard
from pymouse import PyMouse
import time
import pyHook
import pythoncom
import xlrd
import xlwt
import  pyperclip
from pynput import mouse,keyboard
import threading
import sys
from openpyxl import Workbook,load_workbook

def copy():

    k.press_key(k.control_l_key)
    k.tap_key("C")
    k.release_key(k.control_l_key)

def getCopy(maxTime=3):
    #maxTime = 3  # 3秒复制 调用copy() 不管结果对错
    while (maxTime > 0):
        maxTime = maxTime - 0.5
        time.sleep(0.5)
        # print('doing')
        copy()

    result = pyperclip.paste()
    return result


class ExcelData:
    def __init__(self,keyArray,result,typeArray,compareType=''):
        self.keyArray=keyArray
        self.result=result
        self.typeArray=typeArray
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

        typeArray=(table.cell_value(i,4).split('$'))#有吊顶 DN25 规格型号
        compareType=str(table.cell_value(i,3))#0 1 匹配方式 1是完全匹配 比如 水==水
        data=ExcelData(keyArray,result,typeArray,compareType)
        datas.append(data)

    return datas

def getresult(str):
    #datas= getExcelData(excelUrl)

    dataResult=ExcelData('','','')
    contains_key=False
    for data in datas:

        for key in data.keyArray:
            if(key==''):
                contains_key=False
                continue
            if(data.compareType=='1.0'):
                if(key==str):
                    contains_key=True
                    break
                else:
                    contains_key=False
            #elif(data.compareType==''):
            else:
                if(key in str):
                    contains_key=True
                    continue
                else:
                    contains_key=False
                    break
        if(contains_key):
            dataResult=data

    #print(dataResult)
    return dataResult

def calc_result(data,type,keyword):
    result=0
    A=0
    B=0
    array=[]
    try:
        array=type.split(keyword)
        A=float(array[0].strip())/1000
        B=float(array[1].strip())/1000
        if (data.compareType == '2.0'):
            calc_array=str(data.result).split('$')
            X=int(calc_array[0])
            Y=int(calc_array[1])

            result=A*B*X+Y
        elif(data.compareType=='3.0'):
            clac_arrays=str(data.result).split('/')
            calc_array0=clac_arrays[0].split('$')
            calc_array1=clac_arrays[1].split('$')
            X=0
            Y=0
            if(A*B<=1):
                X = int(calc_array0[0])
                Y = int(calc_array0[1])
            else:
                X = int(calc_array1[0])
                Y = int(calc_array1[1])
            result=A*B*X+Y
        elif(data.compareType=='4.0'):
            try:
                L=float(array[2].strip())/1000
            except:
                print('无法找到L！！！！！！')
                L=1
            clac=float(data.result)
            result=(A*B+A*L+B*L)*2*clac
        elif(data.compareType=='5.0'):
            calc_array = str(data.result).split('$')
            X = int(calc_array[0])
            Y = int(calc_array[1])

            result = A * (B+0.2) * X + Y

        else:#TODO-----------------------------------
            result=''
    except:
        result=''
        print('计算规则无法识别')
        print('key:'+keyword)
    finally:
        result=str(result)
    print(result)
    return result

#判断规格型号是否对应
def getresult_2(name,type):
    print(type)

    result=''
    result_data=ExcelData('','','')
    resultDatas=[]
    for data in datas:
        contains_key = False

        for key in data.keyArray:
            if(key==''): continue
            if(key in name):
                contains_key = True

                continue
            else:
                contains_key = False
                break
        if(contains_key and data.compareType!='1.0'):
            resultDatas.append(data)
        if(len(data.keyArray)==1 and data.keyArray[0]==name and data.compareType=='1.0'):
            return data.result
    for d in resultDatas:
        hasResult = False
        for datatype in d.typeArray:
            if(datatype in name):
                hasResult=True
                continue
            else:
                hasResult=False
            if (datatype in type or type == datatype):
                if(data.compareType=='0.0' and type!=datatype):#给电力电缆用————————-
                    hasResult=False
                    break
                hasResult=True
            else:
                hasResult=False
                break
        if(hasResult):
            result=d.result
            result_data=d
    # print('-------')
    # print(result_data.keyArray)
    # print(result_data.typeArray)
    # print(result_data.result)
    # print('-------')
    if(result_data.compareType=='2.0' or result_data.compareType=='3.0' or result_data.compareType=='4.0' or result_data.compareType=='5.0'):

        result= calc_result(result_data,type,result_data.typeArray[0])
    return result

def tapkey(key,count=1):
    for i in range(0,count):
        k.tap_key(key)
        time.sleep(0.2)

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

            if(dataResult.typeArray==''):
                print('无规格直接输入')
                tapkey(k.enter_key, 5)

                k.type_string(dataResult.result)
                tapkey(k.enter_key)
                time.sleep(6)

                tapkey(k.escape_key)
                tapkey(k.left_key, 11)#适当修改
                tapkey(k.enter_key,3)#适当修改
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
                    tapkey(k.enter_key,4)#适当修改
                    k.type_string(result_2)
                    tapkey(k.enter_key)
                    time.sleep(5)
                    tapkey(k.escape_key)
                    tapkey(k.left_key, 11)#适当修改
                    tapkey(k.enter_key,3)#适当修改


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



def saveToExcel(name,type,value):
    #saveworkbook = xlrd.open_workbook(saveExcelUrl)
    #wb = excel_copy(saveworkbook)  # 利用xlutils.copy下的copy函数复制
    wb= load_workbook(filename=saveExcelUrl)
    worksheet=wb.active
    worksheet=wb['Sheet1']

    global rowMaxCount
    #print(rowMaxCount)
    worksheet.cell(row=rowMaxCount+1,column=1,value=name)
    worksheet.cell(row=rowMaxCount+1,column=2,value=type)
    worksheet.cell(row=rowMaxCount+1,column=3,value=value)

    wb.save(saveExcelUrl)
    rowMaxCount=rowMaxCount+1

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

        if (Key == keyboard.Key.f4):
            if(not start):
                targetName = getCopy(1)
                tapkey(k.enter_key)
                targetType = getCopy(1)
                tapkey(k.enter_key, 4)
                targetValue=getCopy(1)
                if(targetType==targetName):
                    targetType=''
                saveToExcel(targetName,targetType,targetValue)
                tapkey(k.enter_key)
                tapkey(k.escape_key)
                tapkey(k.left_key,6)
                tapkey(k.enter_key)
                print('save  '+targetName)

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
    #excelUrl = r"C:\Users\123\Desktop\广联达\安装\Xing.xlsx"  # to do-------------
    #saveExcelUrl = r"C:\Users\123\Desktop\广联达\安装\save.xlsx"  # to do-------------
    saveExcelUrl = r"C:\Users\Administrator\Desktop\save.xlsx"  # to do-------------
    saveworkbook = xlrd.open_workbook(saveExcelUrl)
    rowMaxCount=saveworkbook.sheets()[0].nrows

    datas= getExcelData(excelUrl)

    threads = []
    t2 = threading.Thread(target=main, args=())
    threads.append(t2)
    for t in threads:
        t.setDaemon(True)
        t.start()
    print('press Capital to start')


    with keyboard.Listener(on_press=onpressed) as listener:
        listener.join()


