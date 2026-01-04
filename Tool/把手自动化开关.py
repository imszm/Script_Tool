#coding=gb2312
#各种函数实现

import win32api
import win32con
import os
import time
import serial
import re
import datetime
import random



#获取当前时间  时间数组
def strGetNowTime():
    now = datetime.datetime.now()
    strTime = now.strftime("%Y-%m-%d %H:%M:%S")
    return strTime

#提示窗口
def vShowForm(strDetails="OK",strTitle="Result"):
    win32api.MessageBox(0,
                        str(strDetails),
                        str(strTitle)+" "+strGetNowTime(),
                        win32con.MB_ICONINFORMATION)

#写文件并选择方式和是否打印写入信息
def bWritePrint(strWriteInfo="",strFileName="log.txt",strWriteWay='a',bIsPrint=True):
    try:
        with open(strFileName,strWriteWay) as fl:
            fl.write(strWriteInfo.replace("\r",""))
        if bIsPrint:
            print (strWriteInfo.strip())
        return True
    except:
        return False



#向固定端口发送指令 返回值为指令反馈数组
def strSendPort(strCmd,strPort="com20",fTime=0.5):
    try:
        ser = serial.Serial(strPort,9600,timeout=float(fTime))
        if strCmd != "":
            ser.write(strCmd.encode("utf-8"))
        strInfo = ser.readlines()
        ser.close()
        return strInfo
    except:
        raise AssertionError(strPort+" open error!")



#控制继电器
def bControlRelay(intInput=1):
    if intInput == 1:
        strSendPort("O")
        #vShowForm("All On")
    elif intInput == 0:
        strSendPort("P")
        #vShowForm("All Off")
    elif intInput == 2:
        strSendPort("D")
        #vShowForm("开第二个继电器")
    elif intInput == 3:
        strSendPort("H")
        #vShowForm("开第一个继电器")  
    else:
        vShowForm("Input fail,pleas input 0 or 1!")
        return False
    return True

#随机断电操作：
def Random_On_Off():
    TimeNum=random.randint(1,2)
    bControlRelay(1)
    time.sleep(3)
    print(OffLineTime)
    bControlRelay(1)
    time.sleep(OffLineTime)
   

#控制继电器实现循环随机断电操作

def vLoop(Times=200000):
    bWritePrint("###### 断电测试 ######\n")
    for i in range(Times):        
        bWritePrint("Test No."+str(i+1)+"        "+strGetNowTime()+"\n")                 
        bControlRelay(1)
        time.sleep(3)
        bControlRelay(0)

            
            
        


        
def main():
    vLoop(200000)
    return

if __name__ == "__main__":
    main()

