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
from collections import deque

# ================= 测试参数配置 =================
RELAY_BAUDRATE = 9600  # 继电器串口波特率
DEVICE_BAUDRATE = 115200  # 设备串口波特率
SERIAL_TIMEOUT = 0.1  # 串口读取超时
TEST_CYCLES = 10000  # 测试循环次数
POWER_ON_MIN = 2.0  # 最小供电时间（秒）
POWER_ON_MAX = 5.0  # 最大供电时间（秒）
POWER_OFF_TIME = 5.0  # 断电时间（秒）

# RESET_TIME 已删除
LOG_FILENAME = "relay_random_test_log.txt"  # 正常日志文件名
EXCEPTION_LOG_FILENAME = "relay_exception_log.txt"  # 异常日志文件名
DEVICE_RETRY_DELAY = 3.0  # 设备串口重连等待时间（秒）

# ================= 开关配置 =================
SAVE_LOG_TO_FILE = True  # 是否保存日志到文件
LOG_FLUSH_INTERVAL = 60  # 内存缓存落盘间隔（秒）

# ================= 关键字逻辑配置 =================
# 1. 普通异常关键字 (发现即记录异常)
EXCEPTION_KEYWORDS = [
    "assertion faile datfunction",
]

# 2. 普通信息关键字 (仅记录，不报错)
INFO_KEYWORDS = [
    "voice_msgnum",
    "voice_msgcutoff",
    "ui_pm_acc"
]

# 3. 累计错误关键字 (逻辑：3秒内 ≥ 3次 → 停止测试)
ERROR_CONFIG = {
    "keyword": "param is invalid".replace(" ", ""),
    "window": 3.0,
    "count": 3
}

# 4. 致命错误关键字 (逻辑：1秒内 ≥ 3次 → 立即停止测试)
CRITICAL_CONFIG = {
    "keyword": "[e/motor]reg_addr(00)isunviald",
    "window": 1.0,
    "count": 3
}


# =================================================

class StopTestException(Exception):
    """用于触发停止测试的自定义异常"""
    pass


class RelayTester:
    def __init__(self):
        self.relay_ser = None
        self.device_ser = None
        self.total_success = 0
        # self.total_resets 已删除
        self.total_exceptions = 0
        self.device_disconnect_count = 0
        self.relay_port = None
        self.device_port = None

        # ANSI 颜色去除正则预编译
        self.ansi_escape = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')

        # 日志缓存
        self.log_cache_normal = []
        self.log_cache_exception = []
        self.last_flush_time = time.time()

        # 错误计数器 (使用 deque 存储时间戳)
        self.error_timestamps = deque()
        self.critical_timestamps = deque()

    def get_time(self):
        return datetime.datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")

    def log(self, message, show=True, is_exception=False):
        """日志记录：内存缓存 + 控制台输出 + 自动落盘"""
        log_entry = f"{self.get_time()} {message}"
        if show:
            print(log_entry)

        target_cache = self.log_cache_exception if is_exception else self.log_cache_normal
        target_cache.append(log_entry)

        if SAVE_LOG_TO_FILE and (time.time() - self.last_flush_time >= LOG_FLUSH_INTERVAL):
            self.save_logs_to_file()

    def save_logs_to_file(self):
        """强制刷写日志到磁盘"""
        if not SAVE_LOG_TO_FILE: return

        try:
            if self.log_cache_normal:
                with open(LOG_FILENAME, 'a', encoding='utf-8') as f:
                    f.write("\n".join(self.log_cache_normal) + "\n")
                self.log_cache_normal.clear()

            if self.log_cache_exception:
                with open(EXCEPTION_LOG_FILENAME, 'a', encoding='utf-8') as f:
                    f.write("\n".join(self.log_cache_exception) + "\n")
                self.log_cache_exception.clear()

            self.last_flush_time = time.time()
        except Exception as e:
            print(f"日志写入失败: {e}")

    def show_message(self, message, title="提示"):
        """弹窗提示"""
        try:
            win32api.MessageBox(0, str(message), f"{title} {self.get_time()}", win32con.MB_ICONINFORMATION)
        except Exception:
            print(f"[{title}] {message}")  # 防止非Windows环境报错

    def detect_ports(self):
        """自动检测串口"""
        ports = list(serial.tools.list_ports.comports())
        relay_port = None
        device_port = None

        for p in ports:
            desc = p.description.lower()
            if "4" in desc:  # 根据实际驱动名称调整
                relay_port = p.device
            elif "cp210x" in desc:
                device_port = p.device

        self.log(f"检测结果 -> 继电器: {relay_port} | 通信线: {device_port}")
        return device_port, relay_port

    def open_serial_ports(self):
        """打开所有串口"""
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
        """通用频率检查函数"""
        now = time.time()
        timestamps_deque.append(now)

        # 移除超出时间窗口的旧记录
        while timestamps_deque and timestamps_deque[0] < now - window_seconds:
            timestamps_deque.popleft()

        return len(timestamps_deque) >= threshold_count

    def process_log_line(self, line):
        """处理单行日志，返回 (是否触发停止, 停止原因)"""
        # 1. 预处理：去色、转小写、去空格
        clean_line = self.ansi_escape.sub('', line)
        line_check = clean_line.lower().replace(" ", "")

        # 2. 信息关键字检测
        for kw in INFO_KEYWORDS:
            if kw in line_check:
                self.log(f"【信息】{kw} -> {clean_line.strip()}", show=False)

        # 3. 普通异常关键字检测
        for kw in EXCEPTION_KEYWORDS:
            if kw in line_check:
                self.total_exceptions += 1
                self.log(f"【异常检测】发现关键字: {kw}", is_exception=True)

        # 4. 累计错误 (3秒 >=3次)
        if ERROR_CONFIG["keyword"] in line_check:
            if self.check_frequency(self.error_timestamps, ERROR_CONFIG["window"], ERROR_CONFIG["count"]):
                return True, f"触发停止条件：{ERROR_CONFIG['window']}秒内出现{ERROR_CONFIG['count']}次 '{ERROR_CONFIG['keyword']}'"

        # 5. 致命错误 (1秒 >=3次)
        if CRITICAL_CONFIG["keyword"] in line_check:
            if self.check_frequency(self.critical_timestamps, CRITICAL_CONFIG["window"], CRITICAL_CONFIG["count"]):
                return True, f"触发致命停止：{CRITICAL_CONFIG['window']}秒内出现{CRITICAL_CONFIG['count']}次 '{CRITICAL_CONFIG['keyword']}'"

        return False, None

    def control_relay(self, action):
        """控制继电器 action='on' or 'off'"""
        if not self.relay_ser or not self.relay_ser.is_open:
            return
        try:
            cmd = bytes([0x50]) if action == 'on' else bytes([0x4F])
            self.relay_ser.write(cmd)
            time.sleep(0.1)
            self.relay_ser.read_all()  # 清空缓冲区
            self.log(f"继电器 -> {action.upper()}", show=False)
        except Exception as e:
            self.log(f"继电器控制失败: {e}", is_exception=True)

    def monitor_serial_stream(self, duration):
        """
        实时监控串口流
        duration: 监控持续时长(秒)
        返回: (full_log_str, stop_triggered, stop_reason)
        """
        end_time = time.time() + duration
        collected_logs = []

        while time.time() < end_time:
            try:
                if self.device_ser and self.device_ser.in_waiting:
                    # 使用 errors='replace' 防止乱码崩溃
                    raw_line = self.device_ser.readline().decode('gb2312', errors='replace')
                    if not raw_line: continue

                    collected_logs.append(raw_line.strip())

                    # 实时分析
                    should_stop, reason = self.process_log_line(raw_line)
                    if should_stop:
                        return "\n".join(collected_logs), True, reason

                    # 优化：如果有数据，不sleep，直接进行下一次读取，加快处理速度
                    continue

                    # 只有无数据时才休眠，减少CPU占用
                time.sleep(0.005)

            except serial.SerialException:
                self.log("【警告】串口断开，尝试重连...", is_exception=True)
                self.try_reconnect_device()
                break
            except Exception as e:
                self.log(f"读取流异常: {e}", is_exception=True)

        return "\n".join(collected_logs), False, None

    def try_reconnect_device(self):
        """断线重连逻辑"""
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
                self.log(f"【恢复】设备串口重连成功: {new_dev}")
            except Exception as e:
                self.log(f"重连失败: {e}", is_exception=True)

    def analyze_cycle_result(self, full_logs):
        """
        分析单次循环的最终结果
        优化后：移除了复位判断，只要有 MotorOn 或 PM/Off 任意一个即视为成功
        """
        logs_lower = full_logs.lower().replace(" ", "")

        has_motor_on = "motorpoweron..." in logs_lower
        has_pm_acc = "pm_acc_tim," in logs_lower
        has_power_off = "power_off_system" in logs_lower

        # 简化后的判断逻辑
        if has_motor_on or has_pm_acc or has_power_off:
            details = []
            if has_motor_on: details.append("MotorOn")
            if has_pm_acc: details.append("PM_ACC")
            if has_power_off: details.append("PowerOff")
            return True, f"正常 ({', '.join(details)})"

        return False, "无有效响应"

    def run_single_cycle(self, cycle_num):
        self.log(f"\n--- 第 {cycle_num} 次循环 ---")

        # 1. 随机生成上电时间
        on_time = round(random.uniform(POWER_ON_MIN, POWER_ON_MAX), 1)

        # 2. 继电器上电
        self.control_relay('on')

        # 3. 实时监控 (上电时间 + 缓冲时间)
        logs, stop_triggered, stop_reason = self.monitor_serial_stream(on_time + 1.0)

        # 4. 继电器断电
        self.control_relay('off')

        # 5. 如果触发了停止条件，抛出异常以终止测试
        if stop_triggered:
            self.log(f"【严重】{stop_reason}", is_exception=True)
            self.log("触发中止条件，正在保存日志并退出...")
            raise StopTestException(stop_reason)

        # 6. 分析本次结果 (简化版)
        success, reason = self.analyze_cycle_result(logs)

        if success:
            self.total_success += 1
            self.log(f"【结果】成功: {reason}")
        else:
            self.log(f"【结果】失败: {reason}", is_exception=True)

        # 断电等待
        time.sleep(POWER_OFF_TIME)

        # 打印当前成功率
        rate = (self.total_success / cycle_num) * 100
        print(f"当前成功率: {rate:.2f}%")

    def run_test(self):
        if not self.open_serial_ports():
            self.show_message("串口打开失败", "错误")
            return

        self.log("正在初始化...预先断电")
        self.control_relay('off')
        time.sleep(2.0)

        self.log(f"测试开始，目标循环: {TEST_CYCLES} 次")
        start_time = time.time()

        try:
            for i in range(1, TEST_CYCLES + 1):
                self.run_single_cycle(i)
        except StopTestException as e:
            self.show_message(f"测试自动中止\n原因: {e}", "测试中止")
        except KeyboardInterrupt:
            self.log("用户手动中断测试")
        except Exception as e:
            self.log(f"发生未捕获异常: {e}", is_exception=True)
        finally:
            self.save_logs_to_file()
            self.control_relay('off')  # 确保结束时断电
            if self.relay_ser: self.relay_ser.close()
            if self.device_ser: self.device_ser.close()

        # 统计信息
        elapsed = time.time() - start_time
        summary = (
            f"\n{'=' * 10} 测试报告 {'=' * 10}\n"
            f"总循环: {TEST_CYCLES if i == TEST_CYCLES else i}\n"  # 使用实际运行的次数
            f"成功次数: {self.total_success}\n"
            f"异常关键字数: {self.total_exceptions}\n"
            f"耗时: {elapsed:.1f} 秒"
        )
        self.log(summary)


if __name__ == "__main__":
    tester = RelayTester()
    tester.run_test()