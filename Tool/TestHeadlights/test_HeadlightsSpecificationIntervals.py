import serial
import time

# 打开串口
ser = serial.Serial("COM4", 9600, timeout=0.5)
time.sleep(0.5)  # 给继电器模块一点准备时间

# 循环右灯测试 10 次
for i in range(10):

    # 右灯亮
    ser.write(bytes([0b10100000, 0b00000011, 0b00000001, 0b10100100]))
    ser.flush()
    time.sleep(0.5)  # 右灯亮持续时间

    # 右灯熄灭
    ser.write(bytes([0b10100000, 0b00000011, 0b00000000, 0b10100011]))
    ser.flush()
    time.sleep(1)  # 下次循环前间隔

ser.close()
