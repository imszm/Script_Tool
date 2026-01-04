# -*- coding: utf-8 -*-
import serial
import serial.tools.list_ports
import time
import datetime
import win32api
import win32con

# ================= 测试参数配置 =================
RELAY_BAUDRATE = 9600
DEVICE_BAUDRATE = 115200
SERIAL_TIMEOUT = 1.0
TEST_CYCLES = 500000
LOG_FILENAME = "relay_random_test_log.txt"
EXCEPTION_LOG_FILENAME = "relay_exception_log.txt"

# 按键时间配置
BUTTON_PRESS_TIME_ON = 2.5   # 开机：长按 2.5 秒
BUTTON_PRESS_TIME_OFF = 1.0  # 关机：短按 1.0 秒

POWER_HOLD_TIME = 10.0        # 正常开机后保持时间（秒）
LOG_FLUSH_INTERVAL = 60
SAVE_LOG_TO_FILE = True

# ================= 关键字配置 =================
FORCE_LOSS_KEYWORD = "force_main_polling, communication loss"
STOP_KEYWORD = "[I/voice] voice_msg num: 6"

# ================= 指令配置 (基于您的实际测试结果) =================
# 现象：0x50(灭灯)是按下/开机，0x4F(亮灯)是松开/关机
# 意味着需要一直给0x4F保持松开状态，给0x50才是按下动作
CMD_PRESS = bytes([0x50])    # 模拟按下 (继电器灭/NC导通)
CMD_RELEASE = bytes([0x4F])  # 模拟松开 (继电器亮/NC断开)

# =================================================

class RelayTester:
    def __init__(self):
        self.relay_ser = None
        self.device_ser = None
        self.stop_flag = False
        self.log_cache_normal = []
        self.log_cache_exception = []
        self.last_flush_time = time.time()
        self.device_port = None
        self.relay_port = None

    # ---------------- 通用函数 ----------------
    def get_time(self):
        return datetime.datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")

    def log(self, message, show=True, is_exception=False):
        msg = f"{self.get_time()} {message}"
        if show:
            print(msg)
        if is_exception:
            self.log_cache_exception.append(msg)
        else:
            self.log_cache_normal.append(msg)

        if SAVE_LOG_TO_FILE and (time.time() - self.last_flush_time >= LOG_FLUSH_INTERVAL):
            self.save_logs_to_file()
            self.last_flush_time = time.time()

    def save_logs_to_file(self):
        if not SAVE_LOG_TO_FILE:
            return
        if self.log_cache_normal:
            with open(LOG_FILENAME, 'a', encoding='utf-8') as f:
                f.write("\n".join(self.log_cache_normal) + "\n")
            self.log_cache_normal.clear()
        if self.log_cache_exception:
            with open(EXCEPTION_LOG_FILENAME, 'a', encoding='utf-8') as f:
                f.write("\n".join(self.log_cache_exception) + "\n")
            self.log_cache_exception.clear()

    def show_message(self, message, title="提示"):
        win32api.MessageBox(0, str(message), f"{title} {self.get_time()}", win32con.MB_ICONINFORMATION)

    # ---------------- 串口操作 ----------------
    def detect_ports(self):
        ports = list(serial.tools.list_ports.comports())
        relay_port = None
        device_port = None
        for p in ports:
            desc = p.description.lower()
            if "9" in desc:  # 根据实际情况调整(CH340/USB-SERIAL)
                relay_port = p.device
            elif "cp210x" in desc:
                device_port = p.device
        self.log(f"检测到继电器串口: {relay_port}, 通信串口: {device_port}")
        return device_port, relay_port

    def open_serial_ports(self):
        self.device_port, self.relay_port = self.detect_ports()
        if not self.device_port or not self.relay_port:
            self.log("未检测到完整的串口设备", is_exception=True)
            return False
        try:
            self.relay_ser = serial.Serial(self.relay_port, RELAY_BAUDRATE, timeout=SERIAL_TIMEOUT)
            self.device_ser = serial.Serial(self.device_port, DEVICE_BAUDRATE, timeout=SERIAL_TIMEOUT)
            self.log(f"串口打开成功: 继电器={self.relay_port}, 设备={self.device_port}")
            return True
        except Exception as e:
            self.log(f"串口打开失败: {e}", is_exception=True)
            return False

    def close_serial_ports(self):
        try:
            # 【安全退出】程序结束时，必须发送松开指令 (CMD_RELEASE = 0x4F)
            # 注意：这意味着继电器会保持亮灯状态
            if self.relay_ser and self.relay_ser.is_open:
                self.relay_ser.write(CMD_RELEASE) 
                time.sleep(0.1)
            
            if self.relay_ser and self.relay_ser.is_open:
                self.relay_ser.close()
            if self.device_ser and self.device_ser.is_open:
                self.device_ser.close()
            self.log("串口已安全关闭 (继电器维持松开状态)")
        except Exception as e:
            self.log(f"关闭串口异常: {e}", is_exception=True)

    # ---------------- 继电器动作 ----------------
    def init_relay_state(self):
        """
        初始化继电器状态
        根据您的测试：需要发送 0x4F (CMD_RELEASE) 才是松手
        """
        if self.relay_ser and self.relay_ser.is_open:
            self.log("正在初始化继电器状态 (强制松开/亮灯)...")
            self.relay_ser.write(CMD_RELEASE)  
            time.sleep(2.0) # 等待2秒确保状态稳定
            self.log("继电器初始化完成")

    def relay_press_button(self, duration, action_name="按钮动作"):
        if self.stop_flag:
            return

        if self.relay_ser and self.relay_ser.is_open:
            try:
                # 1. 按下 (发送 0x50 / 灭灯)
                self.relay_ser.write(CMD_PRESS)  
                self.log(f"【{action_name}】按钮按下 (保持 {duration}s)")
                
                # 2. 保持
                start_t = time.time()
                while time.time() - start_t < duration:
                    if self.stop_flag:
                        self.log(f"【警告】在{action_name}期间脚本终止，立即松开！", is_exception=True)
                        break
                    time.sleep(0.05)

                # 3. 松开 (发送 0x4F / 亮灯)
                self.relay_ser.write(CMD_RELEASE)  
                self.log(f"【{action_name}】按钮松开")
                
            except Exception as e:
                self.log(f"继电器控制异常: {e}", is_exception=True)
                self.stop_flag = True
                # 尝试补救松开
                try:
                    self.relay_ser.write(CMD_RELEASE)
                except:
                    pass

    # ---------------- 读取设备日志 ----------------
    def read_device_logs(self, duration):
        end_time = time.time() + duration
        while time.time() < end_time and not self.stop_flag:
            try:
                if self.device_ser and self.device_ser.in_waiting:
                    line = self.device_ser.readline().decode('gb2312', errors='replace').strip()
                    if line:
                        if FORCE_LOSS_KEYWORD in line:
                            self.log(f"【严重】检测到通信丢失: {line}", is_exception=True)
                            self.stop_flag = True
                            break

                        if STOP_KEYWORD in line:
                            self.log(f"【捕获目标】检测到关键词: {STOP_KEYWORD}", is_exception=True)
                            self.log("停止脚本执行，保留现场！", is_exception=True)
                            self.stop_flag = True
                            break
                else:
                    time.sleep(0.01)
            except Exception as e:
                self.log(f"读取设备日志异常: {e}", is_exception=True)
                self.stop_flag = True
                break

    # ---------------- 测试逻辑 ----------------
    def run_single_cycle(self, cycle_num):
        if self.stop_flag:
            return
        self.log(f"\n=== 第 {cycle_num} 次测试 ===")

        # --- 步骤 1: 开机 (长按 2.5s) ---
        self.relay_press_button(BUTTON_PRESS_TIME_ON, action_name="开机")

        if self.stop_flag: 
            return

        # --- 步骤 2: 保持开机状态 ---
        self.log(f"【系统运行】保持待机 {POWER_HOLD_TIME} 秒...")
        self.read_device_logs(POWER_HOLD_TIME)

        if self.stop_flag:
            return

        # --- 步骤 3: 关机 (短按 1.0s) ---
        self.relay_press_button(BUTTON_PRESS_TIME_OFF, action_name="关机")
        
        self.read_device_logs(2.0)

    def run_test(self):
        if not self.open_serial_ports():
            self.show_message("串口打开失败，测试无法开始", "错误")
            return

        # 初始化：发送 0x4F (亮灯) 以松开按钮
        self.init_relay_state()

        self.log(f"初始化完成，开始循环测试。目标次数: {TEST_CYCLES}")
        
        try:
            for i in range(1, TEST_CYCLES + 1):
                if self.stop_flag:
                    break
                self.run_single_cycle(i)
        except KeyboardInterrupt:
            self.log("用户手动中断", is_exception=True)
        finally:
            if self.stop_flag:
                self.log("检测到中止信号，循环已中断。", is_exception=True)
                self.show_message(f"测试中止！\n检测到: {STOP_KEYWORD}\n或通信丢失", "警告")
            else:
                self.log("测试正常结束")
                self.show_message("测试循环完成", "完成")

            self.save_logs_to_file()
            self.close_serial_ports() 
            # 退出后，Relay应该保持在亮灯(0x4F)状态

if __name__ == "__main__":
    tester = RelayTester()
    tester.run_test()
