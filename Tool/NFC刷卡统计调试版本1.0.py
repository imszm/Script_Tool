# -*- coding: utf-8 -*-
import serial
import time
import random
import datetime
import logging
import sys
import win32api
import win32con

# ================= 配置区域 =================
RELAY_PORT = "COM14"  # 继电器控制端口
DEVICE_PORT = "COM10"  # 被测设备端口
RELAY_BAUDRATE = 9600
DEVICE_BAUDRATE = 115200
LOG_FILE = "nfc_test.log"

# ================= 日志配置 =================
# 配置日志同时输出到文件和控制台，替代原来的 bWritePrint
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)


# ================= 辅助函数 =================
def show_message_box(details, title="Alert"):
    """
    非阻塞或仅在致命错误时使用的弹窗
    注意：在自动化循环中尽量少用，以免阻断测试
    """
    win32api.MessageBox(0, str(details), str(title), win32con.MB_ICONINFORMATION)


# ================= 核心类定义 =================

class RelayController:
    """继电器控制类，负责长连接管理"""

    def __init__(self, port, baudrate):
        self.port = port
        self.baudrate = baudrate
        self.ser = None

    def connect(self):
        try:
            if self.ser is None or not self.ser.is_open:
                self.ser = serial.Serial(self.port, self.baudrate, timeout=1)
        except Exception as e:
            logging.error(f"继电器串口 {self.port} 打开失败: {e}")
            raise

    def close(self):
        if self.ser and self.ser.is_open:
            self.ser.close()

    def send_cmd(self, cmd_char):
        try:
            if not self.ser or not self.ser.is_open:
                self.connect()
            self.ser.write(cmd_char.encode("utf-8"))
            time.sleep(0.1)  # 给继电器一点反应时间
        except Exception as e:
            logging.error(f"继电器指令发送失败: {e}")

    def power_on(self):
        self.send_cmd("o")  # 对应原代码 intInput=1

    def power_off(self):
        self.send_cmd("P")  # 对应原代码 intInput=0

    def reset_motor(self):
        logging.warning("执行电机复位操作...")
        self.power_on()
        time.sleep(2.5)
        self.power_off()


class DeviceMonitor:
    """被测设备监听类"""

    def __init__(self, port, baudrate):
        self.port = port
        self.baudrate = baudrate
        self.ser = None

    def connect(self):
        try:
            self.ser = serial.Serial(self.port, self.baudrate, timeout=0.5)
            logging.info(f"设备串口 {self.port} 已连接")
        except Exception as e:
            logging.error(f"设备串口 {self.port} 打开失败: {e}")
            raise

    def close(self):
        if self.ser and self.ser.is_open:
            self.ser.close()

    def read_clean_data(self):
        """读取并清洗数据，替换原有的复杂 replace 链"""
        if not self.ser or not self.ser.is_open:
            return ""

        try:
            # 读取所有缓冲区数据
            raw_data = self.ser.readlines()
            decoded_lines = []
            for line in raw_data:
                try:
                    # 尝试用 utf-8 解码，忽略错误字符
                    s = line.decode('utf-8', errors='ignore').strip()
                    # 去除颜色代码和其他杂质 (根据你原代码的逻辑简化)
                    s = s.replace('\x1b', '').replace('[33m', '')
                    if s:
                        decoded_lines.append(s)
                except:
                    continue
            return "".join(decoded_lines).replace(" ", "").lower()  # 转小写方便查找
        except Exception as e:
            logging.error(f"读取数据异常: {e}")
            return ""


# ================= 主逻辑 =================

def run_nfc_test(loop_times=10000):
    logging.info("###### NFC开关锁测试开始 ######")

    relay = RelayController(RELAY_PORT, RELAY_BAUDRATE)
    device = DeviceMonitor(DEVICE_PORT, DEVICE_BAUDRATE)

    success_count = 0
    reset_count = 0  # 因电机位置错误导致的复位次数

    try:
        relay.connect()
        device.connect()
    except Exception as e:
        show_message_box(f"串口初始化失败: {e}", "Error")
        return

    try:
        for i in range(loop_times):
            current_idx = i + 1
            logging.info(f"Test No.{current_idx} Starting...")

            # 1. 生成随机时间并操作继电器
            random_sleep = random.uniform(6.2, 6.5)  # 原代码逻辑
            # logging.info(f"随机断电延时: {random_sleep:.2f}s")

            relay.power_on()
            time.sleep(random_sleep)
            relay.power_off()

            # 2. 读取设备反馈
            # 这里的逻辑稍微调整：先断电，再读之前积累的log，还是断电后等待设备反应？
            # 原代码是断电后立刻 readlines。
            # 建议稍微等待设备吐出log，或者 serial 的 timeout 会自动处理

            rsp_str = device.read_clean_data()
            # print(rsp_str) # 调试时可开启

            # 3. 结果判断 (统一转为小写判断，提高健壮性)
            is_on = "ui_pm_acc0nfc1on0" in rsp_str
            is_off = "ui_pm_acc0nfc0off1" in rsp_str
            is_invalid = "paramisinvalid" in rsp_str  # 原代码 Mak

            if is_invalid:
                logging.error("当前组包错误，显示E0，测试暂停")
                # 遇到严重错误是否需要 input() 暂停由你决定，这里改为记录日志
                # show_message_box("当前组包错误")

            elif is_on and not is_off:
                logging.info("结果: 恭喜！~ NFC解锁成功，车辆开机")
                success_count += 1

            elif not is_on and is_off:
                logging.info("结果: 恭喜！~ NFC解锁成功，车辆关机")
                success_count += 1

            elif is_on and is_off:
                logging.warning("异常: 电机位置不正确，执行复位...")
                relay.reset_motor()
                reset_count += 1
            else:
                logging.warning(f"结果: 解锁失败或无数据. (Raw len: {len(rsp_str)})")

            # 4. 计算并显示成功率
            valid_tests = current_idx - reset_count
            if valid_tests > 0:
                percentage = (success_count / valid_tests) * 100
                logging.info(
                    f"当前统计: 成功率 {percentage:.2f}% (测试数:{current_idx}, 成功:{success_count}, 复位:{reset_count})")
            else:
                logging.info(f"当前统计: 有效测试数为0")

            # 5. 循环间隔
            time.sleep(2)

    except KeyboardInterrupt:
        logging.info("用户强制停止测试")
    except Exception as e:
        logging.error(f"测试过程中发生未捕获异常: {e}", exc_info=True)
    finally:
        # 确保程序退出时关闭串口
        relay.close()
        device.close()
        logging.info("###### 测试结束，资源已释放 ######")


if __name__ == "__main__":
    run_nfc_test(10000)