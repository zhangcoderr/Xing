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
import xml.dom.minidom
import re
import codecs
import random
def copy():

    k.press_key(k.control_l_key)
    k.tap_key("c")
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

def replaceXmlEncoding(filepath):
    try:
        f = open(filepath, mode='r')
        content = f.read()#文本方式读入
        content = re.sub("GB2312", "UTF-8", content)#替换encoding头
        f.close()
        f = open(filepath, 'w')#写入
        f.write(content)
        f.close()
        f = codecs.open(filepath, 'rb', 'mbcs')#二进制方式读入
        text = f.read().encode("utf-8")#使用utf-8方式编码
        f.close()
        f = open(filepath, 'wb')#二进制方式写入
        f.write(text)
        f.close()
    except:
        return

def getXMLDic(url,hourse_name=''):
    dic={}
    dom=xml.dom.minidom.parse(url)
    root = dom.documentElement
    for node in root.childNodes[1].childNodes:
        if(node.nodeName=='单位工程'):
            for child in node.childNodes:

                if (child.nodeName == '专业工程' ):
                    tag1=hourse_name==''
                    tag2=hourse_name in child.attributes._attrs['名称']._value
                    if(tag1==False and tag2==False):
                        continue
                    if(tag1 or (tag1==False and tag2==True)):
                        temp = child.attributes._attrs['名称']._value
                        print('dic:'+temp)
                        for c in child.childNodes[1].childNodes:
                            for end in c.childNodes:

                                if (end.nodeName == '清单'):
                                    attrs = end._attrs

                                    value = attrs['预算价']
                                    key = attrs['项目特征']
                                    # if('胶片灯' in key._value):
                                    #     print(1)
                                    dic[key._value] = value._value

    return dic



def tapkey(key,count=1):
    for i in range(0,count):
        k.tap_key(key)
        time.sleep(0.05)

def Do():
    global start
    try:
        if start==True:
            # #主代码---------------
            # maxTime = 0.6#3秒复制 调用copy() 不管结果对错
            #
            #
            # name=getCopy(maxTime)
            # name=str(name).replace('\r\n',' ')
            # k.type_string('6')
            # value=dic[name]
            # tapkey(k.enter_key,4)
            # tapkey(k.down_key)
            #
            #
            # minus=getCopy(maxTime)
            # tapkey(k.down_key)
            # random_float=random.randrange(5,10)
            # random_float=2
            # calc=float(value)-float(minus)+random_float
            #
            # if(calc<0):
            #     calc=random.randint(1,10)
            # result=str(calc)
            # int_result=result.split('.')[0]
            # k.type_string(int_result)
            # tapkey(k.enter_key,2)
            # start=False

            # 主代码---------------
            maxTime = 0.6  # 3秒复制 调用copy() 不管结果对错

            name = getCopy(maxTime)
            name = str(name).replace('\r\n', ' ')
            k.type_string('6')
            value = float(name)
            tapkey(k.escape_key)
            tapkey(k.left_key)
            tapkey(k.down_key)

            minus = getCopy(maxTime)
            tapkey(k.down_key)
            random_float = random.randrange(5, 10)
            random_float = 2
            calc = float(value) - float(minus) + random_float

            if (calc < 0):
                calc = random.randint(1, 10)
            result = str(calc)
            int_result = result.split('.')[0]

            #int_result=result#float
            #判断是否已经填写数据  delete
            value_result=getCopy()
            if(value_result==minus):
                k.type_string(int_result)
            else:
                print('跳过，已经填写数据：')
                print(value_result)
            #delete
            #k.type_string(int_result)
            tapkey(k.enter_key, 2)
            start = False
    except:
        start=False

#我的代码
def onpressed(Key):
    global start
    while True:
        #print(Key)
        if (Key==keyboard.Key.caps_lock):#单输入控制价
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
    url = r"C:\Users\Administrator\Desktop\Temp\深国际华东智慧城二期地块项目施工总承包工程(版本号：1).xml"#to do-------------
    replaceXmlEncoding(url)

    #dic=getXMLDic(url,'')#---------------------------------------------

    threads = []
    t2 = threading.Thread(target=main, args=())
    threads.append(t2)
    for t in threads:
        t.setDaemon(True)
        t.start()
    #print('读取完成 ')
    #print(dic)
    print('go')
    with keyboard.Listener(on_press=onpressed) as listener:
        listener.join()
