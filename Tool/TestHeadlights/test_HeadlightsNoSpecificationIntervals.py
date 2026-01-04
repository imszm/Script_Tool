"""
左右转向灯交替闪烁测试（循环10次）
第2路继电器（左灯）：A0 02 01 A3 打开，A0 02 00 A2 关闭
第3路继电器（右灯）：A0 03 01 A4 打开，A0 03 00 A3 关闭
"""
import time
import serial
import allure, pytest

ser = None
numbers = list(range(1, 11))  # 循环次数 10 次

def setup_module(module):
    global ser
    ser = serial.Serial('COM4', 9600, timeout=1)
    print("串口初始化成功")

    # 初始化阶段：关闭所有继电器
    ser.write(bytes([0xA0, 0x02, 0x00, 0xA2]))
    ser.write(bytes([0xA0, 0x03, 0x00, 0xA3]))
    time.sleep(1)


def teardown_module(module):
    global ser
    if ser is not None:
        # 测试结束关闭所有继电器
        ser.write(bytes([0xA0, 0x02, 0x00, 0xA2]))
        ser.write(bytes([0xA0, 0x03, 0x00, 0xA3]))
        ser.close()
        print("串口已关闭") 


@allure.epic('L1项目')
@allure.feature('左右转向灯压力测试')
@allure.story('左右灯交替闪烁')
@allure.title('压力测试') 
@allure.severity('blocker') 
@pytest.mark.parametrize('cnahsu', numbers)
def test_leftRightBlink(cnahsu):
    global ser
    
    print('执行次数：', cnahsu)

    # 左灯亮
    ser.write(bytes([0xA0, 0x02, 0x01, 0xA3]))  # 打开第2路继电器
    time.sleep(2)  # 左灯亮持续时间

    # 左灯熄灭
    ser.write(bytes([0xA0, 0x02, 0x00, 0xA2]))  # 关闭第2路继电器
    time.sleep(0.5)  # 左右灯切换间隔

    # 右灯亮
    ser.write(bytes([0xA0, 0x03, 0x01, 0xA4]))  # 打开第3路继电器
    time.sleep(2)  # 右灯亮持续时间

    # 右灯熄灭
    ser.write(bytes([0xA0, 0x03, 0x00, 0xA3]))  # 关闭第3路继电器
    time.sleep(0.5)  # 下次循环前间隔
