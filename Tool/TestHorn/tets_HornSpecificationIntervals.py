
"""

喇叭规间隔时间开关自动化

"""
import time
import serial
import allure, pytest

numbers = list(range(1, 101))


def setup_module(module):
    global ser
    ser = serial.Serial('COM15', 9600, timeout=1)
    ser.write(bytes([0x50]))
    time.sleep(0.5)
    ser.write(bytes([0x51]))

def teardown_module(module):
    global ser
    if ser is not None:
        ser.close()   
        

@allure.epic('L1项目')
@allure.feature('喇叭机压力测试')
@allure.story('喇规间隔时间开关')
@allure.title('压力测试') 
@allure.severity('blocker') 
@pytest.mark.parametrize('cnahsu', numbers)
def test_onAndOff(cnahsu):
    global ser

    print('执行次数：',cnahsu)
    ser.write(bytes([0x4F]))
    time.sleep(0.2)
    ser.write(bytes([0x50]))
    time.sleep(1)