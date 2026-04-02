# -*- coding: utf-8 -*-
import serial
import serial.tools.list_ports
import time
import datetime
import random
import re
import logging
import sys
from collections import deque
from typing import Optional, Tuple, Deque, List, Dict, Any

# 尝试导入 win32api，用于在 Windows 环境下弹出提示框
try:
    import win32api
    import win32con
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

# ================= 测试参数配置 =================
RELAY_BAUDRATE: int = 9600        # 继电器串口波特率
DEVICE_BAUDRATE: int = 460800     # 设备串口波特率
SERIAL_TIMEOUT: float = 0.1       # 串口读取超时时间（秒）
TEST_CYCLES: int = 10000          # 测试总循环次数
POWER_ON_MIN: float = 3.0         # 最小供电时间（秒）
POWER_ON_MAX: float = 4.0         # 最大供电时间（秒）
POWER_OFF_TIME: float = 1.0       # 断电等待时间（秒）
DEVICE_RETRY_DELAY: float = 3.0   # 设备串口重连等待时间（秒）

# ================= 日志配置 =================
LOG_FILENAME: str = "relay_random_test_log.txt"       # 正常运行日志文件名
EXCEPTION_LOG_FILENAME: str = "relay_exception_log.txt" # 异常追踪日志文件名

# ================= 关键字逻辑配置 =================
# 1. 普通异常关键字 (发现即记录异常)
EXCEPTION_KEYWORDS: List[str] = [
    "assertion faile datfunction",
]

# 2. 普通信息关键字 (仅记录，不报错)
INFO_KEYWORDS: List[str] = [
    "voice_msgnum",
    "voice_msgcutoff",
    "ui_pm_acc"
]

# 3. 累计错误关键字配置 (逻辑：指定时间窗口内达到指定次数 -> 停止测试)
ERROR_CONFIG: Dict[str, Any] = {
    "keyword": "param is invalid".replace(" ", ""),
    "window": 3.0,
    "count": 3
}

# 4. 致命错误关键字配置 (逻辑：指定时间窗口内达到指定次数 -> 立即停止测试)
CRITICAL_CONFIG: Dict[str, Any] = {
    "keyword": "[e/motor]reg_addr(00)isunviald",
    "window": 1.0,
    "count": 3
}


class StopTestException(Exception):
    """用于触发停止测试的自定义异常类"""
    pass


class RelayTester:
    def __init__(self) -> None:
        """初始化测试器，配置串口对象、统计变量、正则匹配器及日志系统"""
        self.relay_ser: Optional[serial.Serial] = None
        self.device_ser: Optional[serial.Serial] = None
        self.relay_port: Optional[str] = None
        self.device_port: Optional[str] = None
        
        # 统计数据
        self.total_success: int = 0
        self.total_exceptions: int = 0
        self.device_disconnect_count: int = 0

        # ANSI 颜色去除正则预编译，用于清洗串口吐出的带色日志
        self.ansi_escape: re.Pattern = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')

        # 错误计数器 (使用 deque 存储时间戳以实现滑动窗口)
        self.error_timestamps: Deque[float] = deque()
        self.critical_timestamps: Deque[float] = deque()

        # 初始化日志系统
        self.logger = self._setup_logger()

    def _setup_logger(self) -> logging.Logger:
        """配置并返回标准 logging 实例"""
        logger = logging.getLogger("RelayTester")
        logger.setLevel(logging.DEBUG)
        
        # 避免重复添加 Handler
        if not logger.handlers:
            formatter = logging.Formatter('[%(asctime)s] %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

            # 1. 控制台输出处理器
            console_handler = logging.StreamHandler(sys.stdout)
            console_handler.setLevel(logging.INFO)
            console_handler.setFormatter(formatter)

            # 2. 全量日志文件处理器 (INFO 及以上级别)
            file_handler = logging.FileHandler(LOG_FILENAME, encoding='utf-8', mode='a')
            file_handler.setLevel(logging.INFO)
            file_handler.setFormatter(formatter)

            # 3. 异常日志文件处理器 (ERROR 级别)
            error_file_handler = logging.FileHandler(EXCEPTION_LOG_FILENAME, encoding='utf-8', mode='a')
            error_file_handler.setLevel(logging.ERROR)
            error_file_handler.setFormatter(formatter)

            logger.addHandler(console_handler)
            logger.addHandler(file_handler)
            logger.addHandler(error_file_handler)
            
        return logger

    def show_message(self, message: str, title: str = "提示") -> None:
        """
        调用系统弹窗提示信息
        
        :param message: 提示内容
        :param title: 弹窗标题
        """
        if HAS_WIN32:
            try:
                # 使用系统模态对话框，确保置顶可见
                win32api.MessageBox(0, str(message), f"{title}", 
                                    win32con.MB_ICONINFORMATION | win32con.MB_SYSTEMMODAL)
            except Exception as e:
                self.logger.warning(f"Win32API 弹窗调用失败: {e}。退回控制台输出: [{title}] {message}")
        else:
            self.logger.info(f"[{title}] {message}")

    def detect_ports(self) -> Tuple[Optional[str], Optional[str]]:
        """
        自动遍历并检测继电器和被测设备的系统串口号
        
        :return: (被测设备串口号, 继电器串口号)
        """
        ports = list(serial.tools.list_ports.comports())
        relay_port: Optional[str] = None
        device_port: Optional[str] = None

        for p in ports:
            desc = p.description.lower()
            # 注意：此处依赖具体驱动名称，需确保测试环境一致
            if "4" in desc:  
                relay_port = p.device
            elif "cp210x" in desc:
                device_port = p.device

        self.logger.info(f"串口检测结果 -> 继电器: {relay_port} | 通信线: {device_port}")
        return device_port, relay_port

    def open_serial_ports(self) -> bool:
        """
        打开继电器与被测设备的串口连接并清空缓冲区
        
        :return: 布尔值，指示串口是否全部成功打开
        """
        self.device_port, self.relay_port = self.detect_ports()
        if not self.device_port or not self.relay_port:
            self.logger.error("未检测到完整的设备（继电器或测试设备缺失），无法启动测试。")
            return False

        try:
            self.relay_ser = serial.Serial(self.relay_port, RELAY_BAUDRATE, timeout=SERIAL_TIMEOUT)
            self.device_ser = serial.Serial(self.device_port, DEVICE_BAUDRATE, timeout=SERIAL_TIMEOUT)
            self.relay_ser.reset_input_buffer()
            self.device_ser.reset_input_buffer()
            self.logger.info("继电器与设备串口均打开成功。")
            return True
        except serial.SerialException as e:
            self.logger.error(f"打开串口时发生错误: {e}")
            return False

    def init_relay_hardware(self) -> None:
        """
        继电器硬件初始化逻辑
        流程：复位(0x50) -> 握手(0x51) -> 设备识别 -> 强制断电(0x50)
        """
        if not self.relay_ser or not self.relay_ser.is_open:
            self.logger.error("初始化跳过：继电器串口未打开或未初始化。")
            return

        self.logger.info(">>> 开始执行继电器硬件初始化...")
        try:
            # 1. 发送 0x50 (复位信号)
            self.logger.info("STEP 1: 发送复位指令 (0x50)...")
            self.relay_ser.write(bytes([0x50]))
            time.sleep(1.0)
            
            # 清理硬件复位可能产生的冗余响应数据
            if self.relay_ser.in_waiting:
                self.relay_ser.read(self.relay_ser.in_waiting)

            # 2. 发送 0x51 (使能/查询指令)
            self.logger.info("STEP 2: 发送使能/查询指令 (0x51)...")
            self.relay_ser.write(bytes([0x51]))
            time.sleep(1.0)

            # 3. 读取响应并判断硬件通道类型
            if self.relay_ser.in_waiting:
                resp = self.relay_ser.read(self.relay_ser.in_waiting)
                resp_hex = resp.hex().lower()

                type_str = "未知通道继电器"
                if "ac" in resp_hex:
                    type_str = "8路继电器"
                elif "ab" in resp_hex:
                    type_str = "4路继电器"
                elif "ad" in resp_hex:
                    type_str = "2路继电器"

                self.logger.info(f"=== 成功检测到硬件：{type_str} (响应: {resp_hex}) ===")
            else:
                self.logger.error("=== 警告：继电器未返回任何握手数据，可能通信异常 ===")

            # 4. 初始化完成后，置于安全状态（强制关闭）
            self.logger.info("STEP 3: 初始化完成，下发强制关闭指令 (0x50)...")
            self.relay_ser.write(bytes([0x50]))
            time.sleep(2.0)
            self.logger.info(">>> 继电器已就绪 (处于 OFF 状态)")

        except serial.SerialException as e:
            self.logger.exception(f"继电器串口读写异常: {e}")

    def check_frequency(self, timestamps_deque: Deque[float], window_seconds: float, threshold_count: int) -> bool:
        """
        通用的基于时间窗口的频率检查算法
        
        :param timestamps_deque: 存储历史事件时间戳的双端队列
        :param window_seconds: 统计的时间窗口大小（秒）
        :param threshold_count: 触发阈值次数
        :return: 是否达到触发频率
        """
        now = time.time()
        timestamps_deque.append(now)

        # 剔除滑动窗口之外的陈旧时间戳记录
        while timestamps_deque and timestamps_deque[0] < now - window_seconds:
            timestamps_deque.popleft()

        return len(timestamps_deque) >= threshold_count

    def process_log_line(self, line: str) -> Tuple[bool, Optional[str]]:
        """
        分析单行串口日志，检索关键字并进行频率计算
        
        :param line: 原始串口日志字符串
        :return: (是否触发停止条件, 具体的停止原因描述)
        """
        # 1. 预处理：清洗ANSI控制符，转为小写并去除空格以增强匹配容错率
        clean_line = self.ansi_escape.sub('', line)
        line_check = clean_line.lower().replace(" ", "")

        # 2. 信息类关键字检测（仅记录日志调试用）
        for kw in INFO_KEYWORDS:
            if kw in line_check:
                # 使用 DEBUG 级别避免刷屏，实际测试中可根据需要调整
                self.logger.debug(f"【匹配信息】{kw} -> {clean_line.strip()}")

        # 3. 异常关键字检测（记录错误但不立即停止）
        for kw in EXCEPTION_KEYWORDS:
            if kw in line_check:
                self.total_exceptions += 1
                self.logger.error(f"【异常检测】发现目标关键字: {kw} | 原始日志: {clean_line.strip()}")

        # 4. 累计型错误检测 (如：3秒内 >= 3次)
        if ERROR_CONFIG["keyword"] in line_check:
            if self.check_frequency(self.error_timestamps, ERROR_CONFIG["window"], ERROR_CONFIG["count"]):
                return True, f"触发停止条件：{ERROR_CONFIG['window']}秒内出现{ERROR_CONFIG['count']}次 '{ERROR_CONFIG['keyword']}'"

        # 5. 致命型错误检测 (如：1秒内 >= 3次)
        if CRITICAL_CONFIG["keyword"] in line_check:
            if self.check_frequency(self.critical_timestamps, CRITICAL_CONFIG["window"], CRITICAL_CONFIG["count"]):
                return True, f"触发致命停止：{CRITICAL_CONFIG['window']}秒内出现{CRITICAL_CONFIG['count']}次 '{CRITICAL_CONFIG['keyword']}'"

        return False, None

    def control_relay(self, action: str) -> None:
        """
        下发指令控制继电器开关
        
        :param action: 'on' 打开 或 'off' 关闭
        """
        if not self.relay_ser or not self.relay_ser.is_open:
            self.logger.warning(f"试图执行继电器操作 '{action}' 失败: 继电器串口未就绪。")
            return
            
        try:
            cmd = bytes([0x4F]) if action == 'on' else bytes([0x50])
            self.relay_ser.write(cmd)
            time.sleep(0.1)
            self.logger.info(f"继电器物理动作执行 -> {action.upper()}")
        except serial.SerialException as e:
            self.logger.error(f"向继电器下发控制指令失败: {e}")

    def monitor_serial_stream(self, duration: float) -> Tuple[str, bool, Optional[str]]:
        """
        在指定时间内持续读取并分析被测设备的串口流数据
        
        :param duration: 持续监控时长（秒）
        :return: (全量日志拼接字符串, 是否触发停止逻辑, 停止原因文本)
        """
        end_time = time.time() + duration
        collected_logs: List[str] = []

        while time.time() < end_time:
            try:
                if self.device_ser and self.device_ser.in_waiting:
                    # 读取时指定 errors='replace' 可有效防止乱码导致程序崩溃
                    raw_line = self.device_ser.readline().decode('gb2312', errors='replace')
                    if not raw_line: 
                        continue

                    collected_logs.append(raw_line.strip())

                    # 实时分析此行日志内容
                    should_stop, reason = self.process_log_line(raw_line)
                    if should_stop:
                        return "\n".join(collected_logs), True, reason

                    # 如果读取到了数据，紧接着尝试下一次读取以清空缓冲区，不执行 sleep
                    continue

                # 避免 CPU 空转
                time.sleep(0.005)

            except serial.SerialException:
                self.logger.error("【警告】被测设备串口连接断开，尝试触发重连机制...")
                self.try_reconnect_device()
                break
            except Exception as e:
                self.logger.exception(f"读取串口数据流时发生未预期异常: {e}")

        return "\n".join(collected_logs), False, None

    def try_reconnect_device(self) -> None:
        """当设备串口发生物理断开或通信异常时的自动重连逻辑"""
        self.device_disconnect_count += 1
        
        # 尝试清理旧的句柄
        if self.device_ser:
            try:
                self.device_ser.close()
            except Exception:
                pass

        self.logger.info(f"等待 {DEVICE_RETRY_DELAY} 秒后重新扫描设备端口...")
        time.sleep(DEVICE_RETRY_DELAY)
        
        new_dev, _ = self.detect_ports()

        if new_dev:
            try:
                self.device_port = new_dev
                self.device_ser = serial.Serial(self.device_port, DEVICE_BAUDRATE, timeout=SERIAL_TIMEOUT)
                self.logger.info(f"【恢复】设备串口重连成功，挂载端口: {new_dev}")
            except serial.SerialException as e:
                self.logger.error(f"设备端口重连阶段引发异常: {e}")
        else:
            self.logger.error("设备重连失败：系统未扫描到对应的硬件节点。")

    def analyze_cycle_result(self, full_logs: str) -> Tuple[bool, str]:
        """
        对单次测试循环内收集到的所有日志进行最终成功率判定
        
        :param full_logs: 该周期内收集的串口文本
        :return: (判定是否成功, 判定详情摘要)
        """
        logs_lower = full_logs.lower().replace(" ", "")

        has_motor_on = "motorpoweron..." in logs_lower
        has_pm_acc = "pm_acc_tim," in logs_lower
        has_power_off = "power_off_system" in logs_lower

        # 当前判定逻辑：只要出现其中一种特征日志即视为有响应
        if has_motor_on or has_pm_acc or has_power_off:
            details: List[str] = []
            if has_motor_on: details.append("MotorOn")
            if has_pm_acc: details.append("PM_ACC")
            if has_power_off: details.append("PowerOff")
            return True, f"状态正常 (匹配项: {', '.join(details)})"

        return False, "无有效响应日志"

    def run_single_cycle(self, cycle_num: int) -> None:
        """
        执行单次压力测试循环（包含上电、监控、判定、断电流程）
        
        :param cycle_num: 当前所处的循环轮次
        """
        self.logger.info(f"\n{'-'*15} 启动第 {cycle_num} 次压力循环 {'-'*15}")

        # 1. 随机生成设备持续上电的时间长度
        on_time = round(random.uniform(POWER_ON_MIN, POWER_ON_MAX), 1)

        # 2. 控制继电器闭合（设备上电）
        self.control_relay('on')

        # 3. 阻塞式实时监控 (上电时间 + 1.0秒的冗余缓冲时间)
        logs, stop_triggered, stop_reason = self.monitor_serial_stream(on_time + 1.0)

        # 4. 控制继电器断开（设备断电）
        self.control_relay('off')

        # 5. 安全检查：若触发了致命/累计异常阈值，抛出异常阻断后续循环
        if stop_triggered:
            self.logger.error(f"【严重阻断】{stop_reason}")
            self.logger.info("测试任务触发了预设的中止条件，即将退出主循环...")
            raise StopTestException(stop_reason)

        # 6. 分析本次循环的响应正确性
        success, reason = self.analyze_cycle_result(logs)

        if success:
            self.total_success += 1
            self.logger.info(f"【单次判定】成功: {reason}")
        else:
            self.logger.error(f"【单次判定】失败: {reason}")

        # 断电静置缓冲，等待电容放电或系统彻底归零
        time.sleep(POWER_OFF_TIME)

        # 打印当前测试统计面板
        rate = (self.total_success / cycle_num) * 100
        self.logger.info(f"当前整体成功率进度: {rate:.2f}% ({self.total_success}/{cycle_num})")

    def run_test(self) -> None:
        """测试任务主控流程入口"""
        if not self.open_serial_ports():
            self.show_message("串口资源请求失败，请检查硬件连接并关闭占用串口的其他软件。", "错误")
            return

        # 执行硬件就绪校验与初始化
        self.init_relay_hardware()

        self.logger.info(f"========== 自动化压力测试任务开始，目标执行总次数: {TEST_CYCLES} 次 ==========")
        start_time = time.time()
        final_count = 0

        try:
            for i in range(1, TEST_CYCLES + 1):
                final_count = i
                self.run_single_cycle(i)
                
        except StopTestException as e:
            self.show_message(f"测试由于触发规则已自动中止\n具体原因: {e}", "测试中止")
        except KeyboardInterrupt:
            self.logger.info("测试主循环被用户通过键盘强制中断。")
        except Exception as e:
            self.logger.exception(f"主控流程发生未被捕获的严重异常: {e}")
        finally:
            # 无论如何退出，确保测试结束时继电器处于断电的安全状态
            self.control_relay('off')  
            
            # 稳妥地释放串口句柄资源
            if self.relay_ser and self.relay_ser.is_open: 
                self.relay_ser.close()
            if self.device_ser and self.device_ser.is_open: 
                self.device_ser.close()

            # 汇总测试报告数据
            elapsed = time.time() - start_time
            summary = (
                f"\n{'=' * 15} 最终测试报告 {'=' * 15}\n"
                f"实际执行总循环: {final_count} 次\n"
                f"成功响应次数: {self.total_success} 次\n"
                f"捕获异常关键字总数: {self.total_exceptions} 次\n"
                f"被测设备重连次数: {self.device_disconnect_count} 次\n"
                f"测试总耗时: {elapsed:.1f} 秒\n"
                f"{'=' * 44}"
            )
            self.logger.info(summary)


if __name__ == "__main__":
    tester = RelayTester()
    tester.run_test()