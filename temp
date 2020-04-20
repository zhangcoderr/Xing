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
import re
from openpyxl import Workbook,load_workbook

def copy():

    k.press_key(k.control_l_key)
    k.tap_key("c")#改小写！！！！ 大写的话由于单进程会触发shift键 ctrl键就失效了
    k.release_key(k.control_l_key)

def getCopy(maxTime=2):
    #maxTime = 3  # 3秒复制 调用copy() 不管结果对错
    while (maxTime > 0):
        maxTime = maxTime - 0.5
        time.sleep(0.5)
        # print('doing')
        copy()

    result = pyperclip.paste()
    return result



replaceDic=\
    {
        '×':'*',
        'WDZN-':'NH-',
        '-BY-':'-BYJ-',
        '-BYR-':'-BYJR-',
        '-RYS-':'-RVS-',
        'TBTRZY-':'BTTVZ-',
        'WDZB-':'WDZBN-',
        'WDZN-':'WDZCN-',
        'WDZ-':'WDZC-',
        'FS-':'',
        '-KYY-':'-KYJY-',
        '-BYRJ-':'-BYJR-',
        'WDZCN-RVS-':'NH-RVS-',
        'WDZBN-RVS-':'NH-RVS-',
        '-RYJS-':'-RVS-',
        'BBTRZ-':'BTTVZ-',
        '-0.6/1KV':'',
        'WDZAN-RVS-':'NH-RVS-',
        'WDZAN-RYJS':'NH-RVS',

        'WDZAN-KVV':'NH-KVV',
        'ZB-YJV':'ZRB-YJV',
        'ZBN-YJV':'NH-YJV',
        'ZBN-KVV':'HN-KVV',
        'WDZB-KYJ':'HN-KVV',




    }
def typeStringReplace(typeString):

    for key in replaceDic:
        if(key in typeString):
            typeString= typeString.replace(key,replaceDic[key])
    return typeString


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

def calc_result(data,name,type,keyword):
    result=0
    A=0
    B=0
    array=[]
    isCircle=False

    try:
        regex_string = ''
        if (keyword in name):
            regex_string=name
        elif(keyword in type):
            regex_string=type
        else:
            result=''
            return result
        if ('φ' in type):  # 对圆单独处理 φ670格式固定----------todo
            isCircle = True
            regex_string=type.replace('φ','')

        if(isCircle):
            array.append(regex_string)
            array.append(regex_string)
        else:
            compile = r'\d+[*xX×]{1,1}\d+'
            split_string=''

            regex = re.compile(compile)
            #TEMP=regex.search(regex_string)
            split_string = regex.search(regex_string).group()

            array=split_string.split(keyword)

            #array=type.split(keyword)
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
        if (isCircle):
            result = str(result * 1.15)
        else:
            result=str(result)

    print(result)
    return result

#判断规格型号是否对应
def getresult_2(name,type):
    print(type)

    result=''
    result_data=ExcelData('','','')
    resultDatas=[]
    pre_type=type
    type = typeStringReplace(type)

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
            #FOR TEST----------type in 则匹配todo-----------
            total_same=False
            for datatype in data.typeArray:
                if (datatype in type):
                    total_same = True
                else:
                    total_same = False
                    break

            if(total_same):
                return data.result
    for d in resultDatas:#贪婪匹配 后面覆盖前面的
        hasResult = False
        for datatype in d.typeArray:
            if(datatype in name):
                hasResult=True
                continue
            else:
                hasResult=False
            if (datatype in type or type == datatype):

                hasResult=True
            else:
                hasResult=False
                break
        if(hasResult):#贪婪匹配 后面覆盖前面的
            result=d.result
            result_data=d

    # for 电力电缆
    for d in resultDatas:#非贪婪 有1个解就跳出
        if(d.compareType=='0.0'):

            if(type==''):
                break
            hasResult_type=False
            if (len(d.typeArray) == 1):
                if (type == d.typeArray[0] or pre_type==d.typeArray[0]):
                    hasResult_type = True
            else:
                split_type_string = type
                for datatype in d.typeArray:
                    split_type_string= split_type_string.replace(datatype,'')
                split_type_string= split_type_string.replace('-','')
                split_type_string= split_type_string.replace(' ','')
                split_type_string = split_type_string.replace('mm2', '')
                split_type_string = split_type_string.replace('1KV', '')
                split_type_string = split_type_string.replace('1kV', '')
                split_type_string = split_type_string.replace('1.0kV', '')


                if(len(split_type_string)>0):
                    #print(split_type_string)
                    hasResult_type=False
                elif(len(split_type_string)==0):
                    hasResult_type=True


            if (hasResult_type):
                result = d.result
                result_data = d
                break#非贪婪 有1个解就跳出
            else:
                result=''#?不能删
                result_data=ExcelData('','','')
    # print('-------')
    # print(result_data.keyArray)
    # print(result_data.typeArray)
    # print(result_data.result)
    # print('-------')
    if(result_data.compareType=='2.0' or result_data.compareType=='3.0' or result_data.compareType=='4.0' or result_data.compareType=='5.0'):

        result= calc_result(result_data,name,type,result_data.typeArray[0])
    return result

def tapkey(key,count=1):
    for i in range(0,count):
        k.tap_key(key)
        time.sleep(0.2)

def Do():
    if start:
        #主代码---------------
        # maxTime = 3#3秒复制 调用copy() 不管结果对错
        # while (maxTime > 0):
        #     maxTime = maxTime - 0.5
        #     time.sleep(0.5)
        #     #print('doing')
        #     copy()
        #
        # targetName=pyperclip.paste()
        targetName=getCopy()


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
                time.sleep(10)

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
                    time.sleep(10)


                    tapkey(k.escape_key)
                    tapkey(k.left_key, 11)#适当修改
                    tapkey(k.enter_key,3)#适当修改


            #判断型号 TODO---------------------
            return





#我的代码
def onpressed(Key):
    while True:
        #print(Key)
        if (Key==keyboard.Key.caps_lock):#开始
            global start
            if(start==True):
                start=False
                print('stop')
            else:
                start=True
                print('go')
        if (Key==keyboard.Key.f3):#结束
            sys.exit()

        #print(Key)
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

