# -*- coding: utf-8 -*-
import serial
import serial.tools.list_ports
import time
import datetime
import win32api
import win32con
import os

# ==========================================================
#                    串口与硬件配置 (调试请改这里)
# ==========================================================
# 1. 串口识别关键词 (填具体COM号如 "COM3" 或 驱动关键词如 "CH340")
RELAY_PORT_KEYWORD = "4"  # 继电器串口特征
DEVICE_PORT_KEYWORD = "cp210x"  # 设备通信串口特征

# 2. 波特率配置
RELAY_BAUDRATE = 9600
DEVICE_BAUDRATE = 115200
SERIAL_TIMEOUT = 1.0

# 3. 按键时长配置 (秒)
BUTTON_PRESS_TIME_ON = 2.5  # 开机：长按 2.5 秒
BUTTON_PRESS_TIME_OFF = 1.0  # 关机：短按 1.0 秒

# 4. 测试流程配置
TEST_CYCLES = 500000  # 目标测试循环次数
POWER_HOLD_TIME = 10.0  # 正常开机后保持监控的时间（秒）
LOG_FLUSH_INTERVAL = 10  # 每隔多少秒将缓存写入硬盘 (防止意外丢失)

# ==========================================================
#                    关键指令与停止条件
# ==========================================================
# 停止脚本的关键字 (优化为核心特征码，避免因模块名微调导致匹配失败)
FORCE_LOSS_KEYWORD = "force_main_polling, communication loss"
STOP_KEYWORD = "voice_msg num: 6"

# 继电器动作指令 (基于 NC 常闭接法)
# 逻辑：0x50(灭灯)=按下, 0x4F(亮灯)=松开
CMD_PRESS = bytes([0x50])  # 模拟按下 (继电器灭/NC导通)
CMD_RELEASE = bytes([0x4F])  # 模拟松开 (继电器亮/NC断开/安全状态)
CMD_ENABLE = bytes([0x51])  # 初始化使能指令


# ==========================================================

class RelayTester:
    def __init__(self):
        self.relay_ser = None
        self.device_ser = None
        self.stop_flag = False

        # 日志缓存
        self.log_cache_normal = []
        self.log_cache_exception = []
        self.last_flush_time = time.time()

        self.device_port = None
        self.relay_port = None

        # 生成带时间戳的独立日志文件名
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_filename = f"TestLog_{timestamp}.txt"
        self.exception_filename = f"ExceptionLog_{timestamp}.txt"

        print(f"=== 测试初始化 ===")
        print(f"正常日志文件: {self.log_filename}")
        print(f"异常日志文件: {self.exception_filename}")
        print(f"==================\n")

    # ---------------- 通用函数 ----------------
    def get_time(self):
        return datetime.datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")

    def log(self, message, show=True, is_exception=False):
        """
        日志记录函数
        """
        timestamp_msg = f"{self.get_time()} {message}"

        # 1. 打印到控制台
        if show:
            print(timestamp_msg)

        # 2. 存入缓存
        if is_exception:
            self.log_cache_exception.append(timestamp_msg)
            # 异常日志同时也记录到普通日志中
            self.log_cache_normal.append(timestamp_msg)
        else:
            self.log_cache_normal.append(timestamp_msg)

        # 3. 定时写入文件
        if time.time() - self.last_flush_time >= LOG_FLUSH_INTERVAL:
            self.save_logs_to_file()
            self.last_flush_time = time.time()

    def save_logs_to_file(self):
        """将内存中的日志写入硬盘"""
        try:
            if self.log_cache_normal:
                with open(self.log_filename, 'a', encoding='utf-8') as f:
                    f.write("\n".join(self.log_cache_normal) + "\n")
                self.log_cache_normal.clear()

            if self.log_cache_exception:
                with open(self.exception_filename, 'a', encoding='utf-8') as f:
                    f.write("\n".join(self.log_cache_exception) + "\n")
                self.log_cache_exception.clear()
        except Exception as e:
            print(f"系统错误 日志写入失败: {e}")

    def show_message(self, message, title="提示"):
        win32api.MessageBox(0, str(message), f"{title} {self.get_time()}", win32con.MB_ICONINFORMATION)

    # ---------------- 串口操作 ----------------
    def detect_ports(self):
        ports = list(serial.tools.list_ports.comports())
        relay_port = None
        device_port = None

        self.log(f"正在扫描串口... (关键字: 继电器='{RELAY_PORT_KEYWORD}', 设备='{DEVICE_PORT_KEYWORD}')")

        for p in ports:
            desc = p.description.lower()
            name = p.device.lower()

            # 匹配继电器
            if RELAY_PORT_KEYWORD.lower() in desc or RELAY_PORT_KEYWORD.lower() in name:
                relay_port = p.device
            # 匹配设备
            elif DEVICE_PORT_KEYWORD.lower() in desc or DEVICE_PORT_KEYWORD.lower() in name:
                device_port = p.device

        self.log(f"检测结果 -> 继电器: {relay_port}, 通信串口: {device_port}")
        return device_port, relay_port

    def open_serial_ports(self):
        self.device_port, self.relay_port = self.detect_ports()
        if not self.device_port or not self.relay_port:
            self.log("错误 未检测到完整的串口设备，请检查配置区的关键字", is_exception=True)
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
            # 安全退出：必须发送松开指令 (亮灯 0x4F)
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
        1. 确保松开 (0x4F)
        2. 发送使能/查询 (0x51)
        """
        if self.relay_ser and self.relay_ser.is_open:
            # Step 1: 强制复位
            self.log("初始化 Step 1: 强制复位继电器 (发送 0x4F/松开)...")
            self.relay_ser.write(CMD_RELEASE)
            time.sleep(1.0)

            # Step 2: 发送使能
            self.log("初始化 Step 2: 发送使能/查询指令 (0x51)...")
            self.relay_ser.write(CMD_ENABLE)
            time.sleep(1.0)

            self.log("继电器初始化完成，状态: 松开 (安全)")

    def relay_press_button(self, duration, action_name="按钮动作"):
        if self.stop_flag:
            return

        if self.relay_ser and self.relay_ser.is_open:
            try:
                # 1. 按下 (发送 0x50 / 灭灯 / NC导通)
                self.relay_ser.write(CMD_PRESS)
                self.log(f"【{action_name}】按钮按下 (保持 {duration}s)")

                # 2. 保持
                start_t = time.time()
                while time.time() - start_t < duration:
                    if self.stop_flag:
                        self.log(f"警告 在{action_name}期间脚本终止，立即执行松开！", is_exception=True)
                        break
                    time.sleep(0.05)

                    # 3. 松开 (发送 0x4F / 亮灯 / NC断开)
                self.relay_ser.write(CMD_RELEASE)
                self.log(f"【{action_name}】按钮松开")

            except Exception as e:
                self.log(f"继电器控制异常: {e}", is_exception=True)
                self.stop_flag = True
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
                    # 使用 errors='replace' 忽略无法解码的乱码
                    line = self.device_ser.readline().decode('gb2312', errors='replace').strip()
                    if line:
                        # 打印日志到屏幕
                        print(f"[设备日志] {line}")
                        self.log_cache_normal.append(f"{self.get_time()} [设备] {line}")

                        # 检测关键字
                        if FORCE_LOSS_KEYWORD in line:
                            self.log(f"严重 检测到通信丢失: {line}", is_exception=True)
                            self.stop_flag = True
                            break

                        if STOP_KEYWORD in line:
                            self.log(f"捕获目标 检测到关键词: {STOP_KEYWORD}", is_exception=True)
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
        self.log(f"系统运行 保持待机监控 {POWER_HOLD_TIME} 秒...")
        self.read_device_logs(POWER_HOLD_TIME)

        if self.stop_flag:
            return

        # --- 步骤 3: 关机 (短按 1.0s) ---
        self.relay_press_button(BUTTON_PRESS_TIME_OFF, action_name="关机")

        # 关机后短暂等待并读取日志
        self.read_device_logs(3.0)

    def run_test(self):
        if not self.open_serial_ports():
            self.show_message("串口打开失败，请检查配置区的串口关键字", "错误")
            return

        # 初始化：确保安全松开 + 发送使能0x51
        self.init_relay_state()

        self.log(f"初始化完成，开始循环测试。目标次数: {TEST_CYCLES}")
        self.log(f"开机长按: {BUTTON_PRESS_TIME_ON}s, 关机短按: {BUTTON_PRESS_TIME_OFF}s")

        try:
            for i in range(1, TEST_CYCLES + 1):
                if self.stop_flag:
                    break
                self.run_single_cycle(i)
        except KeyboardInterrupt:
            self.log("用户手动中断 (Ctrl+C)", is_exception=True)
        except Exception as e:
            self.log(f"主循环发生未捕获异常: {e}", is_exception=True)
        finally:
            if self.stop_flag:
                self.log("测试中止 检测到异常或停止信号。", is_exception=True)
                self.show_message(f"测试中止！\n检测到: {STOP_KEYWORD}\n或通信丢失", "警告")
            else:
                self.log("测试结束 正常完成所有循环或手动停止。")
                self.show_message("测试循环完成", "完成")

            self.save_logs_to_file()
            self.close_serial_ports()


if __name__ == "__main__":
    tester = RelayTester()
    tester.run_test()