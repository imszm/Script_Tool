# -*- coding: utf-8 -*-
import serial
import serial.tools.list_ports
import time
import datetime
import random
import sys
import re
import logging
import collections

# 尝试导入 win32api，如果没有安装也不影响脚本运行
try:
    import win32api
    import win32con

    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False


# ================= 配置类 =================
class Config:
    # 串口配置
    RELAY_BAUDRATE = 9600
    DEVICE_BAUDRATE = 115200
    SERIAL_TIMEOUT = 0.1  # 读操作超时缩短，提高循环响应速度

    # 测试流程参数
    TEST_CYCLES = 100000  # 总测试次数
    POWER_ON_MIN = 7.0  # 最小供电时间
    POWER_ON_MAX = 7.0  # 最大供电时间
    POWER_OFF_TIME = 1.0  # 断电等待时间
    DEVICE_RETRY_DELAY = 3.0  # 掉线重连等待时间

    # 文件名
    LOG_FILE = "test_normal.log"  # 详细日志
    ERR_FILE = "test_error.log"  # 错误日志

    # ---------------- 异常检测阈值 ----------------
    # 严重错误：继电器开启状态下，1秒内出现5次 -> 立即停止
    CRITICAL_KEYWORD = "[e/motor]reg_addr(00)isunviald"
    CRITICAL_WINDOW = 1.0
    CRITICAL_MAX_COUNT = 5

    # 普通错误：3秒内出现3次 -> 停止
    ERROR_KEYWORD = "paramisinvalid"
    ERROR_WINDOW = 3.0
    ERROR_MAX_COUNT = 3

    # 信息关键字（仅打印）
    INFO_KEYWORDS = ["voice_msgnum", "voice_msgcutoff", "ui_pm_acc"]

    # 异常关键字（记录到统计）
    EXCEPTION_KEYWORDS = ["assertionfailedatfunction"]


# ================= 日志系统初始化 =================
def setup_logger():
    # 创建主logger
    logger = logging.getLogger("RelayTester")
    logger.setLevel(logging.DEBUG)
    logger.handlers = []  # 清除旧handler

    # 格式器
    formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

    # 1. 控制台 Handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # 2. 正常日志文件 Handler
    file_handler = logging.FileHandler(Config.LOG_FILE, encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    # 3. 错误日志文件 Handler (只记录 ERROR 及以上)
    err_handler = logging.FileHandler(Config.ERR_FILE, encoding='utf-8')
    err_handler.setLevel(logging.ERROR)
    err_handler.setFormatter(formatter)
    logger.addHandler(err_handler)

    return logger


log = setup_logger()


# ================= 辅助工具 =================
def show_alert(message, title="测试提示"):
    """跨平台的弹窗提示"""
    if HAS_WIN32:
        win32api.MessageBox(0, str(message), title, win32con.MB_ICONINFORMATION)
    else:
        print(f"\n{'!' * 20}\n[{title}] {message}\n{'!' * 20}\n")


class ErrorCounter:
    """滑动窗口计数器，用于检测 X 秒内 Y 次错误"""

    def __init__(self, window_seconds, max_count):
        self.window = window_seconds
        self.max_count = max_count
        self.timestamps = collections.deque()

    def add(self):
        now = time.time()
        self.timestamps.append(now)
        self.cleanup(now)
        return len(self.timestamps)

    def cleanup(self, now):
        """移除超出时间窗口的记录"""
        while self.timestamps and (now - self.timestamps[0] > self.window):
            self.timestamps.popleft()

    def clear(self):
        self.timestamps.clear()


# ================= 核心测试类 =================
class RelayTester:
    def __init__(self):
        self.relay_ser = None
        self.device_ser = None
        self.relay_port = None
        self.device_port = None

        # 统计数据
        self.stats = {
            "success": 0,
            "fail": 0,
            "exception": 0,
            "disconnect": 0
        }

        # 错误检测器
        self.critical_checker = ErrorCounter(Config.CRITICAL_WINDOW, Config.CRITICAL_MAX_COUNT)
        self.error_checker = ErrorCounter(Config.ERROR_WINDOW, Config.ERROR_MAX_COUNT)

        # 状态标记
        self.is_relay_on = False

    def detect_ports(self):
        """自动侦测端口"""
        ports = list(serial.tools.list_ports.comports())
        r_port, d_port = None, None

        for p in ports:
            desc = p.description.lower()
            if "4" in desc and "usb" in desc:  # 适配常见USB继电器描述
                r_port = p.device
            elif "cp210x" in desc or "ch340" in desc:  # 适配常见串口芯片
                d_port = p.device

        return d_port, r_port

    def open_ports(self):
        """打开串口"""
        self.device_port, self.relay_port = self.detect_ports()

        if not self.relay_port or not self.device_port:
            log.error(f"端口检测失败: 设备={self.device_port}, 继电器={self.relay_port}")
            return False

        try:
            # 继电器串口
            if not self.relay_ser or not self.relay_ser.is_open:
                self.relay_ser = serial.Serial(self.relay_port, Config.RELAY_BAUDRATE, timeout=1)

            # 设备串口
            if not self.device_ser or not self.device_ser.is_open:
                self.device_ser = serial.Serial(self.device_port, Config.DEVICE_BAUDRATE, timeout=Config.SERIAL_TIMEOUT)

            log.info(f"串口就绪: 继电器[{self.relay_port}] 设备[{self.device_port}]")
            return True
        except Exception as e:
            log.error(f"打开串口异常: {e}")
            return False

    def close_ports(self):
        """安全关闭串口"""
        for s in [self.relay_ser, self.device_ser]:
            if s and s.is_open:
                try:
                    s.close()
                except:
                    pass
        log.info("所有串口已关闭")

    def reconnect_device(self):
        """设备串口重连逻辑"""
        self.stats["disconnect"] += 1
        log.warning("正在尝试重连设备串口...")

        if self.device_ser:
            try:
                self.device_ser.close()
            except:
                pass

        time.sleep(Config.DEVICE_RETRY_DELAY)
        # 重新侦测
        new_d_port, _ = self.detect_ports()
        if new_d_port:
            try:
                self.device_port = new_d_port
                self.device_ser = serial.Serial(self.device_port, Config.DEVICE_BAUDRATE, timeout=Config.SERIAL_TIMEOUT)
                log.info(f"设备串口重连成功: {new_d_port}")
                return True
            except Exception as e:
                log.error(f"重连失败: {e}")
        return False

    def control_relay(self, state):
        """控制继电器 True=开, False=关"""
        cmd = b'\x50' if state else b'\x4F'  # 50=On, 4F=Off
        action = "开启" if state else "关闭"

        try:
            if self.relay_ser and self.relay_ser.is_open:
                self.relay_ser.write(cmd)
                time.sleep(0.1)  # 给一点硬件反应时间
                self.relay_ser.reset_input_buffer()  # 清空回显
                self.is_relay_on = state
                log.debug(f"继电器已{action}")
            else:
                log.error("继电器串口未打开，无法控制")
        except Exception as e:
            log.error(f"继电器控制失败: {e}")

    def process_log_line(self, line):
        """
        单行日志处理核心
        返回: (is_critical_error, should_stop_test, cleaned_line)
        """
        raw_line = line.strip()
        # 去除ANSI颜色码
        clean_line = re.sub(r'\x1B\[[0-?]*[ -/]*[@-~]', '', raw_line)
        if not clean_line:
            return False, False, None

        lower_line = clean_line.lower().replace(" ", "")

        # 1. 严重错误检测 (1秒5次)
        if Config.CRITICAL_KEYWORD in lower_line:
            cnt = self.critical_checker.add()
            log.warning(f"检测到严重关键字 ({cnt}/{Config.CRITICAL_MAX_COUNT})")

            if self.is_relay_on and cnt >= Config.CRITICAL_MAX_COUNT:
                log.critical(f"【致命错误】严重故障触发阈值！立即停机。")
                return True, True, clean_line

        # 2. 普通错误检测 (3秒3次)
        if Config.ERROR_KEYWORD in lower_line:
            cnt = self.error_checker.add()
            log.warning(f"检测到普通错误 ({cnt}/{Config.ERROR_MAX_COUNT})")
            if cnt >= Config.ERROR_MAX_COUNT:
                log.critical("【错误中止】普通错误触发阈值。")
                return True, True, clean_line

        # 3. 异常关键字记录
        for kw in Config.EXCEPTION_KEYWORDS:
            if kw in lower_line:
                log.error(f"检测到异常日志: {clean_line}")
                self.stats["exception"] += 1

        # 4. 信息关键字打印
        for kw in Config.INFO_KEYWORDS:
            if kw in lower_line:
                log.info(f"捕获信息: {clean_line}")

        return False, False, clean_line

    def monitor_loop(self, duration):
        """
        实时监控循环
        在 duration 时间内持续读取串口，同时进行实时错误检测
        """
        end_time = time.time() + duration
        collected_logs = []

        # 清空计数器，开始新的一轮检测
        self.critical_checker.clear()

        while time.time() < end_time:
            try:
                if not self.device_ser or not self.device_ser.is_open:
                    if not self.reconnect_device():
                        break  # 重连失败跳出循环

                if self.device_ser.in_waiting:
                    # 读取一行，设置errors='ignore'防止乱码炸掉脚本
                    line_bytes = self.device_ser.readline()
                    try:
                        line_str = line_bytes.decode('gb2312', errors='ignore')
                    except:
                        line_str = line_bytes.decode('utf-8', errors='ignore')

                    is_crit, stop_test, clean_line = self.process_log_line(line_str)

                    if clean_line:
                        collected_logs.append(clean_line)

                    if stop_test:
                        # 立即执行紧急动作
                        self.control_relay(False)
                        self.close_ports()
                        show_alert("检测到达到阈值的错误，测试已紧急停止！", "致命错误")
                        sys.exit(1)
                else:
                    # 避免CPU空转，短暂休眠
                    time.sleep(0.01)

            except Exception as e:
                log.error(f"读取循环异常: {e}")
                self.reconnect_device()

        return "\n".join(collected_logs)

    def analyze_cycle_result(self, logs):
        """
        分析结果：
        不再判断 '同时出现' 的情况。
        只要有 motorpoweron 或者 关机信号，就算成功。
        """
        logs_lower = logs.lower().replace(" ", "")

        has_power_on = "motorpoweron" in logs_lower
        has_shutdown = "pm_acc_tim" in logs_lower or "power_off_system" in logs_lower

        if has_power_on or has_shutdown:
            return True
        else:
            return False

    def run_single_test(self, cycle_idx):
        log.info(f"--- 循环 {cycle_idx} / {Config.TEST_CYCLES} ---")

        on_time = round(random.uniform(Config.POWER_ON_MIN, Config.POWER_ON_MAX), 2)

        # 1. 开启继电器
        self.control_relay(True)

        # 2. 实时监控（供电期间）
        logs = self.monitor_loop(on_time)

        # 3. 关闭继电器
        self.control_relay(False)

        # 4. 读取剩余日志（断电后缓冲）
        time.sleep(Config.POWER_OFF_TIME)
        extra_logs = self.monitor_loop(1.0)  # 额外读1秒
        full_logs = logs + "\n" + extra_logs

        # 5. 结果判定
        is_success = self.analyze_cycle_result(full_logs)

        if is_success:
            self.stats["success"] += 1
            log.info("【判定】本轮成功")
        else:
            self.stats["fail"] += 1
            log.error("【判定】本轮失败 (未检测到有效启动/关机信号)")

        # 计算成功率时排除已掉线的次数，或者直接除以总循环数，这里采用除以当前循环数
        success_rate = (self.stats['success'] / cycle_idx) * 100
        log.info(f"当前成功率: {success_rate:.2f}%")

    def run(self):
        print(f"{'=' * 30}\n开始自动化压力测试 (无复位版)\n{'=' * 30}")
        if not self.open_ports():
            return

        # 初始状态复位
        self.control_relay(False)
        time.sleep(1)

        try:
            for i in range(1, Config.TEST_CYCLES + 1):
                self.run_single_test(i)
        except KeyboardInterrupt:
            log.warning("用户手动停止测试")
        except Exception as e:
            log.exception(f"未捕获的全局异常: {e}")
        finally:
            self.control_relay(False)
            self.close_ports()
            self.report_summary()

    def report_summary(self):
        msg = (
            f"\n{'=' * 30}\n测试结束\n"
            f"成功次数: {self.stats['success']}\n"
            f"失败次数: {self.stats['fail']}\n"
            f"异常行数: {self.stats['exception']}\n"
            f"掉线次数: {self.stats['disconnect']}\n"
            f"{'=' * 30}"
        )
        log.info(msg)
        show_alert(msg, "测试报告")


if __name__ == "__main__":
    tester = RelayTester()
    tester.run()