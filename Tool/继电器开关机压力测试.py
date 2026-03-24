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
import os
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
    "assertionfailedatfunction",
]

# 2. 普通信息关键字 (仅记录，不报错)
INFO_KEYWORDS = [
    "voice_msgnum",
    "voice_msgcutoff",
    "ui_pm_acc"
]

# 3. 累计错误关键字 (逻辑：window秒内 >= count次 -> 停止测试)
ERROR_CONFIG = {
    "keyword": "paramisinvalid".replace(" ", ""),
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
#    !! 这些关键字只会在「本次上电之后」的日志里匹配，不会误读历史缓存 !!
SUCCESS_KEYWORDS = [
    "motorpoweron",            # motor power on
    "poweron",                 # [I/power_on]
    "voice_msgnum:0",
    "threadoperatingsystem",
    "motor_svc_init",
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

        # ANSI 颜色去除正则预编译
        self.ansi_escape = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')

        # === 初始化日志路径 ===
        base_path = os.path.dirname(os.path.abspath(__file__))
        self.log_dir_path = os.path.join(base_path, LOG_DIR_NAME)

        if not os.path.exists(self.log_dir_path):
            try:
                os.makedirs(self.log_dir_path)
                print(f"日志文件夹已创建: {self.log_dir_path}")
            except Exception as e:
                print(f"创建日志文件夹失败: {e}")
                self.log_dir_path = base_path

        current_time_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_filename           = os.path.join(self.log_dir_path, f"relay_summary_{current_time_str}.txt")
        self.raw_log_filename       = os.path.join(self.log_dir_path, f"relay_dev_raw_{current_time_str}.txt")
        self.exception_log_filename = os.path.join(self.log_dir_path, f"relay_exception_{current_time_str}.txt")

        print(f"日志将保存在: {self.log_dir_path}")

        # 日志缓存
        self.log_cache_normal    = []
        self.log_cache_exception = []
        self.log_cache_raw       = []
        self.last_flush_time     = time.time()

        # 错误计数器（滑动时间窗口）
        self.error_timestamps    = deque()
        self.critical_timestamps = deque()

    # ------------------------------------------------------------------
    # 工具方法
    # ------------------------------------------------------------------
    def get_time(self):
        return datetime.datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")

    def log(self, message, show=True, is_exception=False):
        log_entry = f"{self.get_time()} {message}"
        if show:
            print(log_entry)

        target_cache = self.log_cache_exception if is_exception else self.log_cache_normal
        target_cache.append(log_entry)
        self.log_cache_raw.append(f"{self.get_time()} [TEST_ACTION] {message}\n")
        self.check_and_flush_logs()

    def log_raw_data(self, raw_text):
        clean_text = self.ansi_escape.sub('', raw_text)
        timestamp  = self.get_time()
        formatted_line = f"{timestamp} {clean_text}"
        if not formatted_line.endswith('\n'):
            formatted_line += '\n'
        self.log_cache_raw.append(formatted_line)

    def check_and_flush_logs(self):
        if SAVE_LOG_TO_FILE and (time.time() - self.last_flush_time >= LOG_FLUSH_INTERVAL):
            self.save_logs_to_file()

    def save_logs_to_file(self):
        if not SAVE_LOG_TO_FILE:
            return
        try:
            if self.log_cache_normal:
                with open(self.log_filename, 'a', encoding='utf-8') as f:
                    f.write("\n".join(self.log_cache_normal) + "\n")
                self.log_cache_normal.clear()

            if self.log_cache_exception:
                with open(self.exception_log_filename, 'a', encoding='utf-8') as f:
                    f.write("\n".join(self.log_cache_exception) + "\n")
                self.log_cache_exception.clear()

            if self.log_cache_raw:
                with open(self.raw_log_filename, 'a', encoding='utf-8', errors='ignore') as f:
                    f.write("".join(self.log_cache_raw))
                self.log_cache_raw.clear()

            self.last_flush_time = time.time()
        except IOError as e:
            print(f"警告：日志写入被拒绝（文件可能被占用）: {e}")
        except Exception as e:
            print(f"日志写入发生未知错误: {e}")

    def show_message(self, message, title="提示"):
        try:
            win32api.MessageBox(0, str(message), f"{title} {self.get_time()}", win32con.MB_ICONINFORMATION)
        except Exception:
            print(f"[{title}] {message}")

    # ------------------------------------------------------------------
    # 串口管理
    # ------------------------------------------------------------------
    def detect_ports(self):
        ports      = list(serial.tools.list_ports.comports())
        relay_port  = None
        device_port = None

        for p in ports:
            desc = p.description.lower()
            if "11" in desc:
                relay_port = p.device
            elif "14" in desc:
                device_port = p.device

        self.log(f"检测结果 -> 继电器: {relay_port} | 通信线: {device_port}")
        return device_port, relay_port

    def open_serial_ports(self):
        self.device_port, self.relay_port = self.detect_ports()
        if not self.device_port or not self.relay_port:
            self.log("未检测到完整设备，无法启动", is_exception=True)
            return False
        try:
            self.relay_ser  = serial.Serial(self.relay_port,  RELAY_BAUDRATE,  timeout=SERIAL_TIMEOUT)
            self.device_ser = serial.Serial(self.device_port, DEVICE_BAUDRATE, timeout=SERIAL_TIMEOUT)
            self.log("串口打开成功")
            return True
        except Exception as e:
            self.log(f"串口打开失败: {e}", is_exception=True)
            return False

    def try_reconnect_device(self):
        """尝试重新连接设备串口"""
        self.device_disconnect_count += 1
        if self.device_ser:
            try:
                self.device_ser.close()
            except Exception:
                pass

        time.sleep(DEVICE_RETRY_DELAY)
        new_dev, _ = self.detect_ports()

        if new_dev:
            try:
                self.device_port = new_dev
                self.device_ser  = serial.Serial(self.device_port, DEVICE_BAUDRATE, timeout=SERIAL_TIMEOUT)
                self.log(f"状态恢复: 设备串口重连成功: {new_dev}")
            except Exception as e:
                self.log(f"重连失败: {e}", is_exception=True)

    # ★★★ 新增：上电前清空串口缓冲区 ★★★
    def flush_device_input_buffer(self):
        """
        清空设备串口输入缓冲区。
        必须在每次继电器上电之前调用，确保 monitor_serial_stream
        只读取「本次上电之后」产生的日志，不会误判历史残留数据。
        """
        if self.device_ser and self.device_ser.is_open:
            try:
                self.device_ser.reset_input_buffer()
                self.log("已清空串口输入缓冲区，开始监听本次上电日志", show=False)
            except Exception as e:
                self.log(f"清空串口缓冲区失败: {e}", is_exception=True)

    # ------------------------------------------------------------------
    # 继电器控制
    # ------------------------------------------------------------------
    def control_relay(self, action):
        if not self.relay_ser or not self.relay_ser.is_open:
            return
        try:
            cmd = bytes([0x50]) if action == 'on' else bytes([0x4F])
            self.relay_ser.write(cmd)
            time.sleep(0.1)
            self.relay_ser.read_all()
            self.log(f"继电器执行动作 -> {action.upper()}", show=False)
        except Exception as e:
            self.log(f"继电器控制失败: {e}", is_exception=True)

    # ------------------------------------------------------------------
    # 日志分析
    # ------------------------------------------------------------------
    def check_frequency(self, timestamps_deque, window_seconds, threshold_count):
        now = time.time()
        timestamps_deque.append(now)
        while timestamps_deque and timestamps_deque[0] < now - window_seconds:
            timestamps_deque.popleft()
        return len(timestamps_deque) >= threshold_count

    def process_log_line(self, line):
        """处理单行日志，实时检测停止条件"""
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
                return True, (f"触发停止条件：{ERROR_CONFIG['window']}秒内出现"
                              f"{ERROR_CONFIG['count']}次 '{ERROR_CONFIG['keyword']}'")

        # 致命错误监控
        if CRITICAL_CONFIG["keyword"] in line_check:
            if self.check_frequency(self.critical_timestamps, CRITICAL_CONFIG["window"], CRITICAL_CONFIG["count"]):
                return True, (f"触发致命停止：{CRITICAL_CONFIG['window']}秒内出现"
                              f"{CRITICAL_CONFIG['count']}次 '{CRITICAL_CONFIG['keyword']}'")

        return False, None

    # ------------------------------------------------------------------
    # 串口流监控（只读取上电之后的新数据）
    # ------------------------------------------------------------------
    def monitor_serial_stream(self, duration, stop_on_success=True):
        """
        监控串口流，只处理本次调用之后到达的新数据。
        调用前已通过 flush_device_input_buffer() 清空了缓冲区，
        因此这里读到的每一行都是「本次上电之后」真实产生的日志。

        :param duration:        最大监控时长（秒）
        :param stop_on_success: 检测到开机成功关键字后是否立即返回
        :return: (collected_logs, is_error_stop, stop_reason, is_success)
        """
        end_time             = time.time() + duration
        collected_logs       = []
        is_success_detected  = False

        while time.time() < end_time:
            try:
                if self.device_ser and self.device_ser.in_waiting:
                    raw_bytes = self.device_ser.readline()
                    if not raw_bytes:
                        continue

                    decoded_line = raw_bytes.decode('gb2312', errors='replace')
                    self.log_raw_data(decoded_line)

                    stripped_line = decoded_line.strip()
                    if not stripped_line:
                        continue

                    collected_logs.append(stripped_line)

                    # 检测停止/错误条件
                    should_stop, reason = self.process_log_line(stripped_line)
                    if should_stop:
                        return "\n".join(collected_logs), True, reason, False

                    # 检测开机成功关键字（只在未成功前检测，避免重复触发）
                    if not is_success_detected:
                        line_check = stripped_line.lower().replace(" ", "")
                        for kw in SUCCESS_KEYWORDS:
                            if kw in line_check:
                                is_success_detected = True
                                self.log(f"成功关键字命中: [{kw}] -> {stripped_line}", show=False)
                                if stop_on_success:
                                    return "\n".join(collected_logs), False, None, True
                                break  # 同一行命中一个即可，跳出 for 继续读后续行

                    self.check_and_flush_logs()

                else:
                    time.sleep(0.005)

            except serial.SerialException:
                self.log("警告: 串口断开，尝试重连...", is_exception=True)
                self.try_reconnect_device()
                break
            except Exception as e:
                self.log(f"读取流异常: {e}", is_exception=True)

        return "\n".join(collected_logs), False, None, is_success_detected

    # ------------------------------------------------------------------
    # 单次开关机循环
    # ------------------------------------------------------------------
    def run_single_cycle(self, cycle_num):
        """执行单次开关机循环"""
        on_time      = round(random.uniform(POWER_ON_MIN, POWER_ON_MAX), 1)
        off_time     = round(random.uniform(POWER_OFF_MIN, POWER_OFF_MAX), 1)
        # 超时保护：上电最长等待时间，可以比 on_time 稍长
        timeout_duration = max(on_time + 3.0, 5.0)

        self.log(f"\n--- 第 {cycle_num} 次循环 "
                 f"(上电保持: {on_time}s, 超时设定: {timeout_duration}s) ---")

        # ★★★ 关键修复：上电前清空串口缓冲区 ★★★
        # 确保后续监听到的日志全部来自「本次上电」之后，
        # 彻底杜绝旧数据命中成功关键字导致误判的问题。
        self.flush_device_input_buffer()

        # 继电器上电
        self.control_relay('on')
        t0 = time.time()

        # 监控串口流（仅读取上电后新产生的数据）
        logs, stop_triggered, stop_reason, is_success = self.monitor_serial_stream(
            timeout_duration, stop_on_success=True
        )

        boot_time = time.time() - t0

        # 继电器断电
        self.control_relay('off')

        # 错误处理
        if stop_triggered:
            self.log(f"严重错误触发停止: {stop_reason}", is_exception=True)
            raise StopTestException(stop_reason)

        # 结果判定
        if is_success:
            self.total_success += 1
            self.log(f"单次测试结果: 成功 (启动耗时: {boot_time:.2f}s)")
        else:
            self.log(
                f"单次测试结果: 失败 - {timeout_duration}秒内未检测到开机关键字",
                is_exception=True
            )
            raise StopTestException(f"第 {cycle_num} 次循环开机超时")

        # 断电等待
        time.sleep(off_time)

        rate = (self.total_success / cycle_num) * 100
        print(f"当前累计成功率: {rate:.2f}%")

    # ------------------------------------------------------------------
    # 主测试流程
    # ------------------------------------------------------------------
    def run_test(self):
        if not self.open_serial_ports():
            self.show_message("串口打开失败，请检查连接及端口占用", "错误")
            return

        self.log("正在初始化测试环境...")

        # 初始化：先上电再断电，进入已知初始状态
        self.log("初始化步骤 1: 执行开机动作")
        self.control_relay('on')
        time.sleep(3.0)

        self.log("初始化步骤 2: 执行关机动作")
        self.control_relay('off')

        # ★★★ 初始化结束后也清空一次缓冲区，防止初始化日志污染第一次循环 ★★★
        self.flush_device_input_buffer()

        self.log("初始化完成: 已进入初始断电状态，等待 2 秒后开始压力测试")
        time.sleep(2.0)

        self.log(f"测试正式开始，目标总次数: {TEST_CYCLES}")
        start_time  = time.time()
        cycle_count = 0

        try:
            for i in range(1, TEST_CYCLES + 1):
                cycle_count = i
                self.run_single_cycle(i)
        except StopTestException as e:
            self.show_message(f"测试已停止以保留现场\n原因: {e}", "异常中止")
        except KeyboardInterrupt:
            self.log("用户手动通过键盘(Ctrl+C)中断测试")
        except Exception as e:
            self.log(f"程序运行发生未捕获异常: {e}", is_exception=True)
        finally:
            self.save_logs_to_file()
            if self.relay_ser:  self.relay_ser.close()
            if self.device_ser: self.device_ser.close()

        elapsed = time.time() - start_time
        summary = (
            f"\n{'=' * 10} 测试统计报告 {'=' * 10}\n"
            f"执行总循环:          {cycle_count}\n"
            f"符合开机条件次数:    {self.total_success}\n"
            f"异常关键字触发数:    {self.total_exceptions}\n"
            f"设备串口断连次数:    {self.device_disconnect_count}\n"
            f"总耗时:              {elapsed:.1f} 秒\n"
            f"{'=' * 30}"
        )
        self.log(summary)


if __name__ == "__main__":
    tester = RelayTester()
    tester.run_test()