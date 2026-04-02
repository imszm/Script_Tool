# -*- coding: utf-8 -*-
import serial
import serial.tools.list_ports
import time
import datetime
import random
import sys
import re
import os
import logging
from collections import deque
from typing import Optional, Tuple, List, Dict, Deque

# ================= 兼容性处理 =================
# 增加对 win32api 的兼容性捕获，避免在非 Windows 环境下直接崩溃
try:
    import win32api
    import win32con
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

# ================= 测试参数配置 =================
RELAY_BAUDRATE: int = 9600        # 继电器串口波特率
DEVICE_BAUDRATE: int = 115200     # 设备串口波特率
SERIAL_TIMEOUT: float = 0.1       # 串口读取超时
TEST_CYCLES: int = 200000         # 测试循环次数

# 上电时间范围（秒）
POWER_ON_MIN: float = 2.0
POWER_ON_MAX: float = 2.0

# 断电时间范围（秒）
POWER_OFF_MIN: float = 2.0
POWER_OFF_MAX: float = 2.0

DEVICE_RETRY_DELAY: float = 3.0   # 设备串口重连等待时间（秒）

# ================= 日志配置 =================
LOG_DIR_NAME: str = "Test_Logs"   # 定义日志存放的文件夹名称

# ================= 关键字逻辑配置 =================
# 1. 普通异常关键字 (发现即记录异常，但不停止)
EXCEPTION_KEYWORDS: List[str] = [
    "assertionfailedatfunction",
]

# 2. 普通信息关键字 (仅记录，不报错)
INFO_KEYWORDS: List[str] = [
    "voice_msgnum",
    "voice_msgcutoff",
    "ui_pm_acc"
]

# 3. 累计错误关键字 (逻辑：window秒内 >= count次 -> 停止测试)
ERROR_CONFIG: Dict[str, Union[str, float, int]] = {
    "keyword": "paramisinvalid".replace(" ", ""),
    "window": 3.0,
    "count": 3
}

# 4. 致命错误关键字 (逻辑：window秒内 >= count次 -> 立即停止测试)
CRITICAL_CONFIG: Dict[str, Union[str, float, int]] = {
    "keyword": "[e/motor]reg_addr(00)isunviald",
    "window": 1.0,
    "count": 3
}

# 5. 开机成功判定关键字 (满足任意一个即可认为开机成功)
SUCCESS_KEYWORDS: List[str] = [
    "motorpoweron",            
    "poweron",                 
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
    def __init__(self) -> None:
        self.relay_ser: Optional[serial.Serial] = None
        self.device_ser: Optional[serial.Serial] = None
        self.total_success: int = 0
        self.total_exceptions: int = 0
        self.device_disconnect_count: int = 0
        self.relay_port: Optional[str] = None
        self.device_port: Optional[str] = None

        # ANSI 颜色去除正则预编译
        self.ansi_escape: re.Pattern = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')

        # 错误计数器（滑动时间窗口）
        self.error_timestamps: Deque[float] = deque()
        self.critical_timestamps: Deque[float] = deque()

        # 初始化日志系统
        self._init_logging()

    def _init_logging(self) -> None:
        """初始化 Python 标准 logging 模块"""
        base_path = os.path.dirname(os.path.abspath(__file__))
        self.log_dir_path = os.path.join(base_path, LOG_DIR_NAME)

        if not os.path.exists(self.log_dir_path):
            try:
                os.makedirs(self.log_dir_path)
                print(f"日志文件夹已创建: {self.log_dir_path}")
            except OSError as e:
                print(f"创建日志文件夹失败: {e}")
                self.log_dir_path = base_path

        current_time_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # 定义日志格式
        formatter = logging.Formatter('[%(asctime)s] %(message)s', datefmt="%Y-%m-%d %H:%M:%S")
        raw_formatter = logging.Formatter('[%(asctime)s] %(message)s', datefmt="%Y-%m-%d %H:%M:%S")

        # 1. 主日志配置 (Summary)
        self.main_logger = logging.getLogger('MainLogger')
        self.main_logger.setLevel(logging.INFO)
        main_handler = logging.FileHandler(os.path.join(self.log_dir_path, f"relay_summary_{current_time_str}.txt"), encoding='utf-8')
        main_handler.setFormatter(formatter)
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        self.main_logger.addHandler(main_handler)
        self.main_logger.addHandler(console_handler)

        # 2. 异常日志配置 (Exception)
        self.error_logger = logging.getLogger('ErrorLogger')
        self.error_logger.setLevel(logging.ERROR)
        error_handler = logging.FileHandler(os.path.join(self.log_dir_path, f"relay_exception_{current_time_str}.txt"), encoding='utf-8')
        error_handler.setFormatter(formatter)
        self.error_logger.addHandler(error_handler)

        # 3. 原始串口数据日志配置 (Raw)
        self.raw_logger = logging.getLogger('RawLogger')
        self.raw_logger.setLevel(logging.INFO)
        raw_handler = logging.FileHandler(os.path.join(self.log_dir_path, f"relay_dev_raw_{current_time_str}.txt"), encoding='utf-8')
        raw_handler.setFormatter(raw_formatter)
        self.raw_logger.addHandler(raw_handler)

        self.main_logger.info(f"日志系统初始化完成，路径: {self.log_dir_path}")

    # ------------------------------------------------------------------
    # 工具方法
    # ------------------------------------------------------------------
    def log(self, message: str, show: bool = True, is_exception: bool = False) -> None:
        """
        统一日志记录接口，替代原有的手动缓存机制。
        """
        if is_exception:
            self.error_logger.error(message)
            if show:
                # 即使是异常，也同步到控制台和主日志，方便回溯
                self.main_logger.error(f"[EXCEPTION] {message}")
        else:
            if show:
                self.main_logger.info(message)
        
        # 动作日志同步记录到 Raw 文件中
        self.raw_logger.info(f"[TEST_ACTION] {message}")

    def log_raw_data(self, raw_text: str) -> None:
        """记录设备端纯净的原始数据"""
        clean_text = self.ansi_escape.sub('', raw_text).strip('\r\n')
        if clean_text:
            self.raw_logger.info(clean_text)

    def show_message(self, message: str, title: str = "提示") -> None:
        """弹窗提示，加入跨平台降级处理"""
        if HAS_WIN32:
            try:
                time_str = datetime.datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
                win32api.MessageBox(0, str(message), f"{title} {time_str}", win32con.MB_ICONINFORMATION)
            except Exception as e:
                self.log(f"弹窗调用失败: {e}", is_exception=True)
        else:
            self.main_logger.warning(f"[{title}] {message}")

    # ------------------------------------------------------------------
    # 串口管理
    # ------------------------------------------------------------------
    def detect_ports(self) -> Tuple[Optional[str], Optional[str]]:
        """检测并分配继电器与设备的串口号"""
        ports = list(serial.tools.list_ports.comports())
        relay_port: Optional[str] = None
        device_port: Optional[str] = None

        for p in ports:
            desc = p.description.lower()
            # 依赖于具体硬件的描述字段特征进行区分
            if "11" in desc:
                relay_port = p.device
            elif "14" in desc:
                device_port = p.device

        self.log(f"检测结果 -> 继电器: {relay_port} | 通信线: {device_port}")
        return device_port, relay_port

    def open_serial_ports(self) -> bool:
        """打开所需的串口连接"""
        self.device_port, self.relay_port = self.detect_ports()
        if not self.device_port or not self.relay_port:
            self.log("未检测到完整设备，无法启动", is_exception=True)
            return False
            
        try:
            self.relay_ser = serial.Serial(self.relay_port, RELAY_BAUDRATE, timeout=SERIAL_TIMEOUT)
            self.device_ser = serial.Serial(self.device_port, DEVICE_BAUDRATE, timeout=SERIAL_TIMEOUT)
            self.log("串口打开成功")
            return True
        except serial.SerialException as e:
            self.log(f"串口打开失败，请检查端口是否被占用: {e}", is_exception=True)
            return False
        except Exception as e:
            self.log(f"串口初始化发生未知错误: {e}", is_exception=True)
            return False

    def try_reconnect_device(self) -> None:
        """尝试重新连接设备串口"""
        self.device_disconnect_count += 1
        if self.device_ser:
            try:
                self.device_ser.close()
            except Exception as e:
                self.log(f"关闭失效串口时发生异常 (可忽略): {e}", show=False)

        time.sleep(DEVICE_RETRY_DELAY)
        new_dev, _ = self.detect_ports()

        if new_dev:
            try:
                self.device_port = new_dev
                self.device_ser = serial.Serial(self.device_port, DEVICE_BAUDRATE, timeout=SERIAL_TIMEOUT)
                self.log(f"状态恢复: 设备串口重连成功: {new_dev}")
            except serial.SerialException as e:
                self.log(f"重连失败: {e}", is_exception=True)

    def flush_device_input_buffer(self) -> None:
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
    def control_relay(self, action: str) -> None:
        """控制继电器开关"""
        if not self.relay_ser or not self.relay_ser.is_open:
            return
            
        try:
            cmd = bytes([0x50]) if action == 'on' else bytes([0x4F])
            self.relay_ser.write(cmd)
            time.sleep(0.1)
            # 读取所有返回以清空继电器响应缓冲区
            self.relay_ser.read_all()
            self.log(f"继电器执行动作 -> {action.upper()}", show=False)
        except Exception as e:
            self.log(f"继电器控制失败: {e}", is_exception=True)

    # ------------------------------------------------------------------
    # 日志分析
    # ------------------------------------------------------------------
    def check_frequency(self, timestamps_deque: Deque[float], window_seconds: float, threshold_count: int) -> bool:
        """检查滑动时间窗口内的发生频率"""
        now = time.time()
        timestamps_deque.append(now)
        while timestamps_deque and timestamps_deque[0] < now - window_seconds:
            timestamps_deque.popleft()
        return len(timestamps_deque) >= threshold_count

    def process_log_line(self, line: str) -> Tuple[bool, Optional[str]]:
        """
        处理单行日志，实时检测停止条件。
        返回: (是否需要停止测试, 停止原因)
        """
        clean_line = self.ansi_escape.sub('', line)
        line_check = clean_line.lower().replace(" ", "")

        # 1. 信息关键字检测
        for kw in INFO_KEYWORDS:
            if kw in line_check:
                self.log(f"信息关键字检测: {kw} -> {clean_line.strip()}", show=False)

        # 2. 普通异常关键字检测
        for kw in EXCEPTION_KEYWORDS:
            if kw in line_check:
                self.total_exceptions += 1
                self.log(f"异常检测触发: 发现关键字: {kw}", is_exception=True)

        # 3. 累计错误监控
        if str(ERROR_CONFIG["keyword"]) in line_check:
            window = float(ERROR_CONFIG["window"])
            count = int(ERROR_CONFIG["count"])
            if self.check_frequency(self.error_timestamps, window, count):
                return True, f"触发停止条件：{window}秒内出现{count}次 '{ERROR_CONFIG['keyword']}'"

        # 4. 致命错误监控
        if str(CRITICAL_CONFIG["keyword"]) in line_check:
            window = float(CRITICAL_CONFIG["window"])
            count = int(CRITICAL_CONFIG["count"])
            if self.check_frequency(self.critical_timestamps, window, count):
                return True, f"触发致命停止：{window}秒内出现{count}次 '{CRITICAL_CONFIG['keyword']}'"

        return False, None

    # ------------------------------------------------------------------
    # 串口流监控
    # ------------------------------------------------------------------
    def monitor_serial_stream(self, duration: float, stop_on_success: bool = True) -> Tuple[str, bool, Optional[str], bool]:
        """
        监控串口流，只处理本次调用之后到达的新数据。
        
        :param duration: 最大监控时长（秒）
        :param stop_on_success: 检测到开机成功关键字后是否立即返回
        :return: (收集到的日志字符串, 是否触发错误停止, 停止原因, 是否开机成功)
        """
        end_time = time.time() + duration
        collected_logs: List[str] = []
        is_success_detected: bool = False

        while time.time() < end_time:
            try:
                if self.device_ser and self.device_ser.in_waiting:
                    raw_bytes = self.device_ser.readline()
                    if not raw_bytes:
                        continue

                    # 处理串口编码问题，遇到无法解析的字节进行替换防崩溃
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

                    # 检测开机成功关键字
                    if not is_success_detected:
                        line_check = stripped_line.lower().replace(" ", "")
                        for kw in SUCCESS_KEYWORDS:
                            if kw in line_check:
                                is_success_detected = True
                                self.log(f"成功关键字命中: [{kw}] -> {stripped_line}", show=False)
                                if stop_on_success:
                                    return "\n".join(collected_logs), False, None, True
                                break
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
    def run_single_cycle(self, cycle_num: int) -> None:
        """执行单次开关机循环"""
        on_time = round(random.uniform(POWER_ON_MIN, POWER_ON_MAX), 1)
        off_time = round(random.uniform(POWER_OFF_MIN, POWER_OFF_MAX), 1)
        
        # 超时保护：上电最长等待时间，设定为上电时间+3秒或最低5秒
        timeout_duration = max(on_time + 3.0, 5.0)

        self.log(f"\n--- 第 {cycle_num} 次循环 "
                 f"(上电保持: {on_time}s, 超时设定: {timeout_duration}s) ---")

        self.flush_device_input_buffer()

        # 继电器上电
        self.control_relay('on')
        t0 = time.time()

        # 监控串口流
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
        self.log(f"当前累计成功率: {rate:.2f}%", show=False)
        print(f"当前累计成功率: {rate:.2f}%")

    # ------------------------------------------------------------------
    # 主测试流程
    # ------------------------------------------------------------------
    def run_test(self) -> None:
        """程序主入口"""
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

        self.flush_device_input_buffer()

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
            self.show_message(f"测试已停止以保留现场\n原因: {e}", "异常中止")
        except KeyboardInterrupt:
            self.log("用户手动通过键盘(Ctrl+C)中断测试")
        except Exception as e:
            self.main_logger.exception(f"程序运行发生未捕获异常: {e}")
        finally:
            if self.relay_ser:
                self.relay_ser.close()
            if self.device_ser:
                self.device_ser.close()

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