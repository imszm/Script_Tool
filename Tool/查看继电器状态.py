import serial
# 修改为你的继电器端口
ser = serial.Serial("COM9", 9600) 

print("发送 0x4F (期望吸合/亮灯)...")
ser.write(bytes([0x4F]))
# 此时观察：继电器是否“咔哒”一声吸合？指示灯是否亮起？

input("按回车发送 0x50 (期望断开/灭灯)...")
ser.write(bytes([0x50]))
# 此时观察：继电器是否“咔哒”一声断开？指示灯是否熄灭？

ser.close()
