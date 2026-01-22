# -*- coding: utf-8 -*-
import serial
import serial.tools.list_ports
import time
import datetime
import random
import sys
import win32api
import win32con
import re
import os  # 引入操作系统模块，用于创建文件夹和路径处理
from collections import deque

# ================= 测试参数配置 =================
RELAY_BAUDRATE = 9600     # 继电器串口波特率
DEVICE_BAUDRATE = 115200  # 设备串口波特率
SERIAL_TIMEOUT = 0.1      # 串口读取超时
TEST_CYCLES = 200000      # 测试循环次数

# 上电时间范围（秒）
POWER_ON_MIN = 2
POWER_ON_MAX = 2

# 断电时间范围（秒）
POWER_OFF_MIN = 2
POWER_OFF_MAX = 2

DEVICE_RETRY_DELAY = 3.0  # 设备串口重连等待时间（秒）

# ================= 日志配置 =================
LOG_DIR_NAME = "Test_Logs"  # 定义日志存放的文件夹名称
SAVE_LOG_TO_FILE = True     # 是否保存日志到文件
LOG_FLUSH_INTERVAL = 60     # 内存缓存落盘间隔（秒）

# ================= 关键字逻辑配置 =================
# 1. 普通异常关键字 (发现即记录异常，但不停止)
EXCEPTION_KEYWORDS = [
    "assertion faile datfunction",
]

# 2. 普通信息关键字 (仅记录，不报错)
INFO_KEYWORDS = [
    "voice_msgnum",
    "voice_msgcutoff",
    "ui_pm_acc"
]

# 3. 累计错误关键字 (逻辑：window秒内 >= count次 -> 停止测试)
ERROR_CONFIG = {
    "keyword": "param is invalid".replace(" ", ""),
    "window": 3.0,
    "count": 3
}

# 4. 致命错误关键字 (逻辑：window秒内 >= count次 -> 立即停止测试)
CRITICAL_CONFIG = {
    "keyword": "[e/motor]reg_addr(00)isunviald",
    "window": 1.0,
    "count": 3
}

# 5. 开机成功判定关键字 (满足任意一个即可认为开机成功)
SUCCESS_KEYWORDS = [
    "motorpoweron",
    "poweron",
    "voice_msg num: 0",
    "uipmacc:1:acc1:on0"
]


# =================================================

class StopTestException(Exception):
    """用于触发停止测试的自定义异常"""
    pass


class RelayTester:
    def __init__(self):
        self.relay_ser = None
        self.device_ser = None
        self.total_success = 0
        self.total_exceptions = 0
        self.device_disconnect_count = 0
        self.relay_port = None
        self.device_port = None

        # ANSI 颜色去除正则预编译 (用于去除串口日志中的颜色代码，方便逻辑匹配)
        self.ansi_escape = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')

        # === 初始化日志路径逻辑 ===
        # 获取当前脚本所在的绝对路径，确保在任何地方运行脚本都能找到正确位置
        base_path = os.path.dirname(os.path.abspath(__file__))
        # 拼接日志文件夹路径
        self.log_dir_path = os.path.join(base_path, LOG_DIR_NAME)

        # 如果文件夹不存在，则创建
        if not os.path.exists(self.log_dir_path):
            try:
                os.makedirs(self.log_dir_path)
                print(f"日志文件夹已创建: {self.log_dir_path}")
            except Exception as e:
                print(f"创建日志文件夹失败: {e}")
                # 如果创建失败，回退到当前目录，防止程序崩溃
                self.log_dir_path = base_path

        # 获取当前启动时间，用于文件名
        current_time_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

        # 使用 os.path.join 拼接完整的文件路径
        # 1. 简要测试过程日志 (给测试人员看)
        self.log_filename = os.path.join(self.log_dir_path, f"relay_summary_{current_time_str}.txt")
        # 2. 完整原始开发日志 (给开发人员看)
        self.raw_log_filename = os.path.join(self.log_dir_path, f"relay_dev_raw_{current_time_str}.txt")
        # 3. 异常日志
        self.exception_log_filename = os.path.join(self.log_dir_path, f"relay_exception_{current_time_str}.txt")

        print(f"日志将保存在: {self.log_dir_path}")

        # 日志缓存 (减少IO操作频率)
        self.log_cache_normal = []     # 摘要日志缓存
        self.log_cache_exception = []  # 异常日志缓存
        self.log_cache_raw = []        # 原始开发日志缓存
        self.last_flush_time = time.time()

        # 错误计数器 (使用 deque 存储时间戳，用于滑动窗口频率检测)
        self.error_timestamps = deque()
        self.critical_timestamps = deque()

    def get_time(self):
        return datetime.datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")

    def log(self, message, show=True, is_exception=False):
        """
        摘要日志记录：记录测试步骤和结果
        """
        log_entry = f"{self.get_time()} {message}"
        if show:
            print(log_entry)

        target_cache = self.log_cache_exception if is_exception else self.log_cache_normal
        target_cache.append(log_entry)

        # 摘要日志同时也写入一份到原始日志中，方便开发对照时间轴查看操作
        self.log_cache_raw.append(f"{self.get_time()} [TEST_ACTION] {message}\n")

        self.check_and_flush_logs()

    def log_raw_data(self, raw_text):
        """
        原始日志记录：专门记录串口原始输出，不打印到控制台
        此处包含 ANSI 颜色去除逻辑
        """
        # 1. 去除 ANSI 颜色代码
        clean_text = self.ansi_escape.sub('', raw_text)

        # 2. 给每一行原始数据加上时间戳
        timestamp = self.get_time()
        formatted_line = f"{timestamp} {clean_text}"

        # 如果原始行末尾没有换行符，补一个，保持格式整洁
        if not formatted_line.endswith('\n'):
            formatted_line += '\n'

        self.log_cache_raw.append(formatted_line)

    def check_and_flush_logs(self):
        """检查时间间隔并落盘所有日志"""
        if SAVE_LOG_TO_FILE and (time.time() - self.last_flush_time >= LOG_FLUSH_INTERVAL):
            self.save_logs_to_file()

    def save_logs_to_file(self):
        """强制刷写日志到磁盘"""
        if not SAVE_LOG_TO_FILE: return

        try:
            # 1. 保存摘要日志
            if self.log_cache_normal:
                with open(self.log_filename, 'a', encoding='utf-8') as f:
                    f.write("\n".join(self.log_cache_normal) + "\n")
                self.log_cache_normal.clear()

            # 2. 保存异常日志
            if self.log_cache_exception:
                with open(self.exception_log_filename, 'a', encoding='utf-8') as f:
                    f.write("\n".join(self.log_cache_exception) + "\n")
                self.log_cache_exception.clear()

            # 3. 保存原始开发日志
            if self.log_cache_raw:
                # 原始日志量大，使用 utf-8 存储，errors='ignore' 防止特殊字符导致写入失败
                with open(self.raw_log_filename, 'a', encoding='utf-8', errors='ignore') as f:
                    f.write("".join(self.log_cache_raw))
                self.log_cache_raw.clear()

            self.last_flush_time = time.time()
        except IOError as e:
            # 捕获IO错误（例如文件被用户打开占用时），仅打印不崩溃
            print(f"警告：日志写入被拒绝（文件可能被占用）: {e}")
        except Exception as e:
            print(f"日志写入发生未知错误: {e}")

    def show_message(self, message, title="提示"):
        """弹窗提示"""
        try:
            win32api.MessageBox(0, str(message), f"{title} {self.get_time()}", win32con.MB_ICONINFORMATION)
        except Exception:
            # 如果在无界面环境运行，退化为控制台打印
            print(f"[{title}] {message}")

    def detect_ports(self):
        """自动检测串口"""
        ports = list(serial.tools.list_ports.comports())
        relay_port = None
        device_port = None

        for p in ports:
            # 将描述转为小写，提高匹配容错率
            desc = p.description.lower()
            # 注意：这里的关键字 "4" 和 "com25" 可能需要根据实际电脑情况调整
            if "4" in desc:  # 继电器驱动标识
                relay_port = p.device
            elif "com25" in desc:  # 设备通信线标识
                device_port = p.device

        self.log(f"检测结果 -> 继电器: {relay_port} | 通信线: {device_port}")
        return device_port, relay_port

    def open_serial_ports(self):
        """打开串口连接"""
        self.device_port, self.relay_port = self.detect_ports()
        if not self.device_port or not self.relay_port:
            self.log("未检测到完整设备，无法启动", is_exception=True)
            return False

        try:
            self.relay_ser = serial.Serial(self.relay_port, RELAY_BAUDRATE, timeout=SERIAL_TIMEOUT)
            self.device_ser = serial.Serial(self.device_port, DEVICE_BAUDRATE, timeout=SERIAL_TIMEOUT)
            self.log("串口打开成功")
            return True
        except Exception as e:
            self.log(f"串口打开失败: {e}", is_exception=True)
            return False

    def check_frequency(self, timestamps_deque, window_seconds, threshold_count):
        """
        检查特定时间内关键字出现的频率
        :param timestamps_deque: 存储时间戳的双端队列
        :param window_seconds: 时间窗口大小
        :param threshold_count: 触发阈值
        """
        now = time.time()
        timestamps_deque.append(now)

        # 移除超出时间窗口的旧记录
        while timestamps_deque and timestamps_deque[0] < now - window_seconds:
            timestamps_deque.popleft()

        return len(timestamps_deque) >= threshold_count

    def process_log_line(self, line):
        """处理单行日志，实时检测停止条件"""
        # 去除ANSI颜色代码，便于逻辑匹配
        clean_line = self.ansi_escape.sub('', line)
        line_check = clean_line.lower().replace(" ", "")

        # 信息关键字检测
        for kw in INFO_KEYWORDS:
            if kw in line_check:
                self.log(f"信息关键字检测: {kw} -> {clean_line.strip()}", show=False)

        # 普通异常关键字检测
        for kw in EXCEPTION_KEYWORDS:
            if kw in line_check:
                self.total_exceptions += 1
                self.log(f"异常检测触发: 发现关键字: {kw}", is_exception=True)

        # 累计错误监控
        if ERROR_CONFIG["keyword"] in line_check:
            if self.check_frequency(self.error_timestamps, ERROR_CONFIG["window"], ERROR_CONFIG["count"]):
                return True, f"触发停止条件：{ERROR_CONFIG['window']}秒内出现{ERROR_CONFIG['count']}次 '{ERROR_CONFIG['keyword']}'"

        # 致命错误监控
        if CRITICAL_CONFIG["keyword"] in line_check:
            if self.check_frequency(self.critical_timestamps, CRITICAL_CONFIG["window"], CRITICAL_CONFIG["count"]):
                return True, f"触发致命停止：{CRITICAL_CONFIG['window']}秒内出现{CRITICAL_CONFIG['count']}次 '{CRITICAL_CONFIG['keyword']}'"

        return False, None

    def control_relay(self, action):
        """控制继电器动作"""
        if not self.relay_ser or not self.relay_ser.is_open:
            return
        try:
            cmd = bytes([0x50]) if action == 'on' else bytes([0x4F])
            self.relay_ser.write(cmd)
            # 稍作延时等待继电器动作
            time.sleep(0.1)
            # 清空输入缓存，防止读取到旧数据
            self.relay_ser.read_all()
            self.log(f"继电器执行动作 -> {action.upper()}", show=False)
        except Exception as e:
            self.log(f"继电器控制失败: {e}", is_exception=True)

    def monitor_serial_stream(self, duration, stop_on_success=True):
        """
        [优化版] 监控串口流
        :param duration: 最大监控时长
        :param stop_on_success: 是否在检测到开机成功后立即停止监控 (提高效率)
        :return: (collected_logs, is_error_stop, stop_reason, is_success)
        """
        end_time = time.time() + duration
        collected_logs_for_cycle = []
        is_success_detected = False  # 新增：记录是否实时检测到了成功

        while time.time() < end_time:
            try:
                if self.device_ser and self.device_ser.in_waiting:
                    raw_bytes = self.device_ser.readline()
                    if not raw_bytes: continue

                    decoded_line = raw_bytes.decode('gb2312', errors='replace')
                    self.log_raw_data(decoded_line)  # 写入原始日志

                    stripped_line = decoded_line.strip()
                    if stripped_line:
                        collected_logs_for_cycle.append(stripped_line)

                        # 1. 实时检测：错误/停止条件
                        should_stop, reason = self.process_log_line(stripped_line)
                        if should_stop:
                            return "\n".join(collected_logs_for_cycle), True, reason, False

                        # 2. 实时检测：成功条件 (新增逻辑)
                        # 只要当前行包含任意一个成功关键字，就认为成功
                        line_check = stripped_line.lower().replace(" ", "")
                        if not is_success_detected:  # 如果还没成功过，才去检测
                            for kw in SUCCESS_KEYWORDS:
                                if kw in line_check:
                                    is_success_detected = True
                                    if stop_on_success:
                                        # 发现成功，提前结束监听！
                                        return "\n".join(collected_logs_for_cycle), False, None, True

                    self.check_and_flush_logs()
                    continue

                time.sleep(0.005)

            except serial.SerialException:
                self.log("警告: 串口断开，尝试重连...", is_exception=True)
                self.try_reconnect_device()
                break
            except Exception as e:
                self.log(f"读取流异常: {e}", is_exception=True)

        # 时间到了也没发现成功(或者不要求立即停止)，返回当前状态
        return "\n".join(collected_logs_for_cycle), False, None, is_success_detected

    def try_reconnect_device(self):
        """尝试重新连接设备串口"""
        self.device_disconnect_count += 1
        if self.device_ser:
            try:
                self.device_ser.close()
            except:
                pass

        time.sleep(DEVICE_RETRY_DELAY)
        new_dev, _ = self.detect_ports()

        if new_dev:
            try:
                self.device_port = new_dev
                self.device_ser = serial.Serial(self.device_port, DEVICE_BAUDRATE, timeout=SERIAL_TIMEOUT)
                self.log(f"状态恢复: 设备串口重连成功: {new_dev}")
            except Exception as e:
                self.log(f"重连失败: {e}", is_exception=True)

    def analyze_cycle_result(self, full_logs):
        """
        分析单次循环的最终结果
        判定逻辑：必须出现 motor power on 或 ui_pm_acc: 1:acc 1:on 0 之一
        """
        # 预处理日志：转小写并去空格
        logs_processed = full_logs.lower().replace(" ", "")

        # 检查是否包含预设的成功关键字
        for target in SUCCESS_KEYWORDS:
            if target in logs_processed:
                return True, f"正常 (检测到关键字段: {target})"

        return False, "未检测到开机必要字段 (motor power on 或 ui_pm_acc...)"

    def run_single_cycle(self, cycle_num):
        """执行单次开关机循环"""
        on_time = round(random.uniform(POWER_ON_MIN, POWER_ON_MAX), 1)
        off_time = round(random.uniform(POWER_OFF_MIN, POWER_OFF_MAX), 1)

        # 监控时间设长一点作为超时保护(例如5秒)，因为我们会提前返回，所以不用担心浪费时间
        timeout_duration = 5.0

        self.log(f"\n--- 第 {cycle_num} 次循环 (上电保持: {on_time}s, 超时设定: {timeout_duration}s) ---")

        # 2. 继电器上电
        self.control_relay('on')

        # 记录开始时间用于计算实际启动耗时
        t0 = time.time()

        # 3. 监控串口流 (开启 stop_on_success=True)
        # 注意：这里返回值变成了 4 个
        logs, stop_triggered, stop_reason, is_success = self.monitor_serial_stream(timeout_duration,
                                                                                   stop_on_success=True)

        boot_time = time.time() - t0  # 计算实际启动花费时间

        # 4. 继电器断电
        self.control_relay('off')

        # 5. 错误处理
        if stop_triggered:
            self.log(f"严重错误触发停止: {stop_reason}", is_exception=True)
            raise StopTestException(stop_reason)

        # 6. 结果判定 (直接使用实时检测的结果)
        if is_success:
            self.total_success += 1
            # 可以在这里打印实际启动耗时，非常有用的数据
            self.log(f"单次测试结果: 成功 (启动耗时: {boot_time:.2f}s)")
        else:
            self.log(f"单次测试结果: 失败 - {timeout_duration}秒内未检测到开机关键字", is_exception=True)
            raise StopTestException(f"第 {cycle_num} 次循环开机超时")

        # 7. 断电等待
        time.sleep(off_time)

        rate = (self.total_success / cycle_num) * 100
        print(f"当前累计成功率: {rate:.2f}%")

    def run_test(self):
        """主测试循环逻辑"""
        if not self.open_serial_ports():
            self.show_message("串口打开失败，请检查连接及端口占用", "错误")
            return

        self.log("正在初始化测试环境...")

        # 1. 先开机
        self.log("初始化步骤 1: 执行开机动作")
        self.control_relay('on')
        time.sleep(3.0)

        # 2. 再关机，进入初始状态
        self.log("初始化步骤 2: 执行关机动作")
        self.control_relay('off')
        self.log("初始化完成: 已进入初始断电状态，等待 2 秒后开始压力测试")
        time.sleep(2.0)

        self.log(f"测试正式开始，目标总次数: {TEST_CYCLES}")
        start_time = time.time()

        cycle_count = 0
        try:
            for i in range(1, TEST_CYCLES + 1):
                cycle_count = i
                self.run_single_cycle(i)
        except StopTestException as e:
            # 捕获异常，测试停止
            self.show_message(f"测试已停止以保留现场\n原因: {e}", "异常中止")
        except KeyboardInterrupt:
            self.log("用户手动通过键盘(Ctrl+C)中断测试")
        except Exception as e:
            self.log(f"程序运行发生未捕获异常: {e}", is_exception=True)
        finally:
            self.save_logs_to_file()
            # 退出时关闭串口
            if self.relay_ser: self.relay_ser.close()
            if self.device_ser: self.device_ser.close()

        # 生成并输出测试报告
        elapsed = time.time() - start_time
        summary = (
            f"\n{'=' * 10} 测试统计报告 {'=' * 10}\n"
            f"执行总循环: {cycle_count}\n"
            f"符合开机条件次数: {self.total_success}\n"
            f"异常关键字触发数: {self.total_exceptions}\n"
            f"总耗时: {elapsed:.1f} 秒\n"
            f"{'=' * 30}"
        )
        self.log(summary)


if __name__ == "__main__":
    tester = RelayTester()
    tester.run_test()