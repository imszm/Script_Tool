import pyautogui
import win32api
import win32con
import os
import time
import serial
import re
import datetime
import random
 
try:
    while True:
        # 获取鼠标当前位置
        x, y = pyautogui.position()
        print(f"Mouse position: X={x}, Y={y}")
        time.sleep(2)  # 每0.5秒打印一次鼠标坐标
 
except KeyboardInterrupt:
    print("Program terminated.")
