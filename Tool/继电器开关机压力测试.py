# -*- coding: utf-8 -*-
"""
继电器开关机压力测试脚本 - v3 优化版

【本次修复的问题 —— 针对「第1次即失败」日志的分析】
────────────────────────────────────────────────────────────────────────
  [BUG-4] 失败时强制断电，导致现场丢失（直接致命）
    原代码逻辑（run_single_cycle）：
      1. control_relay('on')    ← 上电
      2. monitor_serial_stream() ← 监听
      3. control_relay('off')   ← 【无条件先关！】← BUG 所在
      4. if is_success: ...
         else: raise StopTestException()  ← 此时设备已经断电，现场全无

    修复：只在 is_success=True 时调用 control_relay('off')，
         失败时保持继电器闭合，设备持续上电，供现场排查。

  [BUG-5] 缺乏零数据预判（无法区分「硬件断路」与「设备启动失败」）
    上次失败日志中，5s 监听窗口内【零字节】来自 COM13，说明串口连接
    本身有问题（接线、端口识别错误等），而非设备启动失败。
    原脚本无法区分这两种情况，统一报"疑似启动挂死"，误导排查方向。

    修复：区分零数据故障和关键字未命中故障，给出针对性诊断提示。

  [关于「继电器逻辑是否反了」的判断]
    根据日志：19:07:01 发送 OFF 后设备断电（用户确认"继电器直接关了，
    导致设备也关了"）。这说明 0x4F→断电、0x50→上电 的逻辑是正确的，
    继电器并未反接。
    但为防止后续更换继电器模块出现极性问题，本版新增 RELAY_CMD_ON /
    RELAY_CMD_OFF 显式配置，两个字节互换即可适配极性相反的模块。

【新增功能】
────────────────────────────────────────────────────────────────────────
  1. 修复 BUG-4：失败时不断电，保留现场
  2. 修复 BUG-5：区分零数据 vs 关键字未命中，给出精准诊断建议
  3. 新增 verify_serial_connectivity()：压测开始前上电一次，确认能收到
     串口数据，零数据则中止，彻底避免「带着硬件故障跑压测」的盲跑
  4. 新增 RELAY_CMD_ON / RELAY_CMD_OFF 配置项，一键适配不同极性的继电器
  5. 监听窗口内超过 NO_DATA_WARNING_SECS 秒无数据，立即触发预警日志
"""

import serial
import serial.tools.list_ports
import time
import datetime
import re
import os
import logging
from collections import deque
from typing import Optional, Tuple, List, Dict, Deque, Union

# ================= 兼容性处理 =================
try:
    import win32api
    import win32con
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

# ================= 测试参数配置 =================
RELAY_BAUDRATE: int = 9600
DEVICE_BAUDRATE: int = 115200
SERIAL_TIMEOUT: float = 0.1
TEST_CYCLES: int = 200000

POWER_ON_TIME: float = 5.0        # 继电器上电并保持监听的固定时长（秒）
POWER_OFF_TIME: float = 5.0       # 继电器断电复位，给电机控制器电容放电的固定时长（秒）
DEVICE_RETRY_DELAY: float = 3.0

# ─────────────────────────────────────────────────────────────────────────────
# 继电器指令字节配置
# 默认: 0x50('P') = 闭合上电，0x4F('O') = 断开断电
# 若发现 ON 指令反而断电（模块极性相反），将两个值互换即可：
#   RELAY_CMD_ON  = 0x4F
#   RELAY_CMD_OFF = 0x50
# ─────────────────────────────────────────────────────────────────────────────
RELAY_CMD_ON:  int = 0x50
RELAY_CMD_OFF: int = 0x4F

# 上电后 N 秒内仍未收到任何串口字节 → 触发预警日志
NO_DATA_WARNING_SECS: float = 2.5

# 压测正式开始前的串口连通性验证：等待首字节数据的最长超时（秒）
CONNECTIVITY_VERIFY_TIMEOUT: float = 8.0

# ================= 日志配置 =================
LOG_DIR_NAME: str = "Test_Logs"

# ================= 关键字逻辑配置 =================
EXCEPTION_KEYWORDS: List[str] = [
    "assertionfailedatfunction",
]

INFO_KEYWORDS: List[str] = [
    "voice_msgnum",
    "voice_msgcutoff",
    "ui_pm_acc"
]

ERROR_CONFIG: Dict[str, Union[str, float, int]] = {
    "keyword": "paramisinvalid".replace(" ", ""),
    "window": 3.0,
    "count": 3
}

CRITICAL_CONFIG: Dict[str, Union[str, float, int]] = {
    "keyword": "[e/motor]reg_addr(00)isunviald",
    "window": 1.0,
    "count": 3
}

# ─────────────────────────────────────────────────────────────────────────────
# 成功判定关键字列表（v2 已修复下划线和前缀问题，此处延续）
# ─────────────────────────────────────────────────────────────────────────────
SUCCESS_KEYWORDS: List[str] = [
    "motorpoweron",
    "poweron",
    "voice_msgnum:",         # 前缀匹配，兼容任意数量
    "threadoperatingsystem",
    "motor_svc_init",
    "ui_pm_acc:1:acc1:on0",  # 已修复：补回下划线
    "exitsleep:acc_key1",    # 早期可靠信号
    "acc_cb,key_evt:1",      # 启动最早信号
]

CONCAT_BUFFER_MAX_LEN: int = 2048

# =============================================================================


class StopTestException(Exception):
    """自定义异常类，用于在达到致命错误条件时触发并停止测试"""
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

        self.ansi_escape: re.Pattern = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')
        self.error_timestamps: Deque[float] = deque()
        self.critical_timestamps: Deque[float] = deque()

        self._init_logging()

    def _init_logging(self) -> None:
        base_path: str = os.path.dirname(os.path.abspath(__file__))
        self.log_dir_path: str = os.path.join(base_path, LOG_DIR_NAME)

        current_time_str: str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        formatter: logging.Formatter = logging.Formatter(
            '[%(asctime)s] %(message)s', datefmt="%Y-%m-%d %H:%M:%S"
        )

        self.main_logger: logging.Logger = logging.getLogger('MainLogger')
        self.main_logger.setLevel(logging.INFO)
        console_handler: logging.StreamHandler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        if not self.main_logger.handlers:
            self.main_logger.addHandler(console_handler)

        if not os.path.exists(self.log_dir_path):
            try:
                os.makedirs(self.log_dir_path)
                self.main_logger.info(f"日志文件夹已创建: {self.log_dir_path}")
            except OSError as e:
                self.main_logger.error(f"创建日志文件夹失败，将回退到根目录: {e}")
                self.log_dir_path = base_path

        main_handler: logging.FileHandler = logging.FileHandler(
            os.path.join(self.log_dir_path, f"relay_summary_{current_time_str}.txt"),
            encoding='utf-8'
        )
        main_handler.setFormatter(formatter)
        self.main_logger.addHandler(main_handler)

        self.error_logger: logging.Logger = logging.getLogger('ErrorLogger')
        self.error_logger.setLevel(logging.ERROR)
        error_handler: logging.FileHandler = logging.FileHandler(
            os.path.join(self.log_dir_path, f"relay_exception_{current_time_str}.txt"),
            encoding='utf-8'
        )
        error_handler.setFormatter(formatter)
        if not self.error_logger.handlers:
            self.error_logger.addHandler(error_handler)

        self.raw_logger: logging.Logger = logging.getLogger('RawLogger')
        self.raw_logger.setLevel(logging.INFO)
        raw_handler: logging.FileHandler = logging.FileHandler(
            os.path.join(self.log_dir_path, f"relay_dev_raw_{current_time_str}.txt"),
            encoding='utf-8'
        )
        raw_handler.setFormatter(formatter)
        if not self.raw_logger.handlers:
            self.raw_logger.addHandler(raw_handler)

        self.main_logger.info(f"日志系统初始化完成，存储路径: {self.log_dir_path}")

    def log(self, message: str, show: bool = True, is_exception: bool = False) -> None:
        if is_exception:
            self.error_logger.error(message)
            if show:
                self.main_logger.error(f"[EXCEPTION] {message}")
        else:
            if show:
                self.main_logger.info(message)
        self.raw_logger.info(f"[TEST_ACTION] {message}")

    def log_raw_data(self, raw_text: str) -> None:
        clean_text: str = self.ansi_escape.sub('', raw_text).strip('\r\n')
        if clean_text:
            self.raw_logger.info(clean_text)

    def show_message(self, message: str, title: str = "提示") -> None:
        if HAS_WIN32:
            try:
                time_str: str = datetime.datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
                win32api.MessageBox(
                    0, str(message), f"{title} {time_str}", win32con.MB_ICONINFORMATION
                )
            except Exception as e:
                self.log(f"Win32 弹窗调用失败，转换为控制台输出: {e}", is_exception=True)
        else:
            self.main_logger.warning(f"[{title}] {message}")

    def detect_ports(self) -> Tuple[Optional[str], Optional[str]]:
        ports = list(serial.tools.list_ports.comports())
        relay_port: Optional[str] = None
        device_port: Optional[str] = None

        for p in ports:
            desc: str = p.description.lower()
            if "ch340" in desc:
                relay_port = p.device
            elif "cp210x" in desc or "cp210" in desc:
                device_port = p.device

        self.log(
            f"串口检测结果 -> 继电器管控端口(CH340): {relay_port} "
            f"| 设备通信端口(CP210x): {device_port}"
        )
        return device_port, relay_port

    def open_serial_ports(self) -> bool:
        self.device_port, self.relay_port = self.detect_ports()
        if not self.device_port or not self.relay_port:
            self.log("未检测到完整的硬件设备，无法启动测试流", is_exception=True)
            return False

        try:
            self.relay_ser = serial.Serial(self.relay_port, RELAY_BAUDRATE, timeout=SERIAL_TIMEOUT)
            self.device_ser = serial.Serial(self.device_port, DEVICE_BAUDRATE, timeout=SERIAL_TIMEOUT)
            self.log("全部串口初始化打开成功")
            return True
        except serial.SerialException as e:
            self.log(f"串口打开失败，请检查端口占用: {e}", is_exception=True)
            return False
        except Exception as e:
            self.log(f"串口初始化发生未预料错误: {e}", is_exception=True)
            return False

    def try_reconnect_device(self) -> None:
        self.device_disconnect_count += 1
        if self.device_ser:
            try:
                self.device_ser.close()
            except Exception as e:
                self.log(f"关闭失效串口发生异常 (可忽略): {e}", show=False)

        time.sleep(DEVICE_RETRY_DELAY)
        new_dev, _ = self.detect_ports()

        if new_dev:
            try:
                self.device_port = new_dev
                self.device_ser = serial.Serial(self.device_port, DEVICE_BAUDRATE, timeout=SERIAL_TIMEOUT)
                self.log(f"状态恢复: 设备串口重连成功: {new_dev}")
            except serial.SerialException as e:
                self.log(f"重连硬件失败: {e}", is_exception=True)

    def flush_device_input_buffer(self) -> None:
        if self.device_ser and self.device_ser.is_open:
            try:
                self.device_ser.reset_input_buffer()
                self.log("已清空串口输入缓冲区，开始专注监听本次上电日志", show=False)
            except Exception as e:
                self.log(f"清空串口缓冲区失败: {e}", is_exception=True)

    def control_relay(self, action: str) -> None:
        """
        向继电器发送控制指令。
        ON  → 发送 RELAY_CMD_ON  字节（默认 0x50）→ 继电器闭合，设备上电
        OFF → 发送 RELAY_CMD_OFF 字节（默认 0x4F）→ 继电器断开，设备断电

        如果发现 ON 反而断电，只需在配置区将 RELAY_CMD_ON/RELAY_CMD_OFF 互换。
        """
        if not self.relay_ser or not self.relay_ser.is_open:
            return

        try:
            cmd: bytes = bytes([RELAY_CMD_ON]) if action == 'on' else bytes([RELAY_CMD_OFF])
            self.relay_ser.write(cmd)
            time.sleep(0.1)
            self.relay_ser.read_all()
            self.log(f"继电器执行动作指令 -> {action.upper()}", show=False)
        except Exception as e:
            self.log(f"继电器控制指令下发失败: {e}", is_exception=True)

    def check_frequency(
        self,
        timestamps_deque: Deque[float],
        window_seconds: float,
        threshold_count: int
    ) -> bool:
        now: float = time.time()
        timestamps_deque.append(now)
        while timestamps_deque and timestamps_deque[0] < now - window_seconds:
            timestamps_deque.popleft()
        return len(timestamps_deque) >= threshold_count

    def process_log_line(self, line: str) -> Tuple[bool, Optional[str]]:
        clean_line: str = self.ansi_escape.sub('', line)
        line_check: str = clean_line.lower().replace(" ", "")

        for kw in INFO_KEYWORDS:
            if kw in line_check:
                self.log(f"信息关键字记录: {kw} -> {clean_line.strip()}", show=False)

        for kw in EXCEPTION_KEYWORDS:
            if kw in line_check:
                self.total_exceptions += 1
                self.log(f"异常报警触发: 解析到异常关键字: {kw}", is_exception=True)

        if str(ERROR_CONFIG["keyword"]) in line_check:
            window: float = float(ERROR_CONFIG["window"])
            count: int = int(ERROR_CONFIG["count"])
            if self.check_frequency(self.error_timestamps, window, count):
                return True, (
                    f"测试阻断：触发频繁错误规则 "
                    f"({window}秒内检测到 {count} 次 '{ERROR_CONFIG['keyword']}')"
                )

        if str(CRITICAL_CONFIG["keyword"]) in line_check:
            window = float(CRITICAL_CONFIG["window"])
            count = int(CRITICAL_CONFIG["count"])
            if self.check_frequency(self.critical_timestamps, window, count):
                return True, (
                    f"致命阻断：触发严重报错规则 "
                    f"({window}秒内检测到 {count} 次 '{CRITICAL_CONFIG['keyword']}')"
                )

        return False, None

    def _check_success_keywords(
        self,
        text_to_search: str,
        source_label: str,
        original_line: str
    ) -> Optional[str]:
        for kw in SUCCESS_KEYWORDS:
            if kw in text_to_search:
                self.log(
                    f"成功校验点命中 [{source_label}]: [{kw}] -> {original_line.strip()}",
                    show=False
                )
                return kw
        return None

    def monitor_serial_stream(
        self,
        duration: float,
        stop_on_success: bool = False
    ) -> Tuple[str, bool, Optional[str], bool]:
        """
        阻塞监听来自 CP210x 的串口流数据。

        新增：
          - 零数据预警：上电后 NO_DATA_WARNING_SECS 秒内无任何字节 → 打印预警日志
          - 双层成功检测（逐行 + 滑动拼接缓冲区容错）
        """
        end_time: float = time.time() + duration
        monitor_start: float = time.time()
        collected_logs: List[str] = []
        is_success_detected: bool = False
        any_data_received: bool = False
        no_data_warned: bool = False
        concat_buffer: str = ""

        while time.time() < end_time:
            try:
                if self.device_ser and self.device_ser.in_waiting:
                    raw_bytes: bytes = self.device_ser.readline()
                    if not raw_bytes:
                        continue

                    any_data_received = True
                    decoded_line: str = raw_bytes.decode('gb2312', errors='replace')
                    self.log_raw_data(decoded_line)

                    stripped_line: str = decoded_line.strip()
                    if not stripped_line:
                        continue

                    collected_logs.append(stripped_line)

                    should_stop, reason = self.process_log_line(stripped_line)
                    if should_stop:
                        return "\n".join(collected_logs), True, reason, False

                    if not is_success_detected:
                        clean_for_match: str = (
                            self.ansi_escape.sub('', stripped_line)
                            .lower()
                            .replace(" ", "")
                        )

                        matched_kw = self._check_success_keywords(
                            clean_for_match, "逐行", stripped_line
                        )

                        if not matched_kw:
                            concat_buffer += clean_for_match
                            if len(concat_buffer) > CONCAT_BUFFER_MAX_LEN:
                                concat_buffer = concat_buffer[-CONCAT_BUFFER_MAX_LEN:]
                            matched_kw = self._check_success_keywords(
                                concat_buffer, "缓冲区容错", stripped_line
                            )
                        else:
                            concat_buffer += clean_for_match
                            if len(concat_buffer) > CONCAT_BUFFER_MAX_LEN:
                                concat_buffer = concat_buffer[-CONCAT_BUFFER_MAX_LEN:]

                        if matched_kw:
                            is_success_detected = True
                            if stop_on_success:
                                return "\n".join(collected_logs), False, None, True
                else:
                    # ── 零数据预警 ──────────────────────────────────────────
                    if (not any_data_received
                            and not no_data_warned
                            and time.time() - monitor_start >= NO_DATA_WARNING_SECS):
                        self.log(
                            f"[串口静默预警] 上电 {NO_DATA_WARNING_SECS:.1f}s 后"
                            f" {self.device_port} 仍未收到任何字节。"
                            f"请检查: ①串口接线是否松动 ②设备UART-TX是否接通"
                            f" ③串口号 {self.device_port} 是否正确",
                            is_exception=True
                        )
                        no_data_warned = True
                    time.sleep(0.005)

            except serial.SerialException:
                self.log("硬件警告: 设备串口失去连接，启动重连机制...", is_exception=True)
                self.try_reconnect_device()
                break
            except Exception as e:
                self.log(f"数据流读取逻辑发生不可预料异常: {e}", is_exception=True)
                break

        return "\n".join(collected_logs), False, None, is_success_detected

    def verify_serial_connectivity(self) -> bool:
        """
        压测正式开始前的串口连通性预验证。

        上电一次，等待最多 CONNECTIVITY_VERIFY_TIMEOUT 秒，确认能从设备串口
        收到至少一个字节数据。零数据 → 中止测试，避免带着硬件故障空转压测。
        """
        self.log(
            f"串口连通性预验证: 上电，等待最多 {CONNECTIVITY_VERIFY_TIMEOUT:.0f}s "
            f"确认 {self.device_port} 能收到数据..."
        )
        self.control_relay('on')
        start = time.time()
        received = False

        try:
            deadline = time.time() + CONNECTIVITY_VERIFY_TIMEOUT
            while time.time() < deadline:
                if (self.device_ser
                        and self.device_ser.is_open
                        and self.device_ser.in_waiting > 0):
                    received = True
                    elapsed = time.time() - start
                    self.log(
                        f"连通性验证通过: 上电 {elapsed:.2f}s 后"
                        f"收到首字节数据，串口链路正常"
                    )
                    break
                time.sleep(0.05)
        except Exception as e:
            self.log(f"连通性验证过程异常: {e}", is_exception=True)

        # 无论结果如何，断电复位
        self.control_relay('off')

        if not received:
            self.log(
                f"[连通性验证失败] 上电 {CONNECTIVITY_VERIFY_TIMEOUT:.0f}s 后"
                f" {self.device_port} 仍未收到任何数据。\n"
                f"  可能原因:\n"
                f"    ① 串口接线问题（TX/RX 线松动或未接）\n"
                f"    ② 串口号识别错误（当前识别为 {self.device_port}，"
                f"请确认是否正确）\n"
                f"    ③ 继电器极性问题（ON 指令实际断电）——"
                f"可尝试对调 RELAY_CMD_ON / RELAY_CMD_OFF\n"
                f"  压力测试中止，请排查后重新启动。",
                is_exception=True
            )
            return False

        # 等待电容放电
        self.log(f"等待 {POWER_OFF_TIME:.0f}s 电容放电后进入压测循环...")
        time.sleep(POWER_OFF_TIME)
        self.flush_device_input_buffer()
        return True

    def run_single_cycle(self, cycle_num: int) -> None:
        """
        执行一次完整的断电/上电测试。

        【核心修复 BUG-4】
        relay OFF 指令只在测试通过（is_success=True）时发送。
        测试失败时，继电器保持闭合，设备持续上电，供现场排查。
        """
        self.log(
            f"\n--- [流程标记] 第 {cycle_num} 次压力循环 "
            f"(固定上电监听: {POWER_ON_TIME}s, 断电时长: {POWER_OFF_TIME}s) ---"
        )

        self.flush_device_input_buffer()
        self.control_relay('on')
        t0: float = time.time()

        logs, stop_triggered, stop_reason, is_success = self.monitor_serial_stream(
            POWER_ON_TIME, stop_on_success=False
        )

        boot_time: float = time.time() - t0

        # ── 异常关键字触发阻断：保留上电现场 ────────────────────────────────
        if stop_triggered:
            self.log(
                f"严重错误触发中断（设备保持上电以供现场排查）: {stop_reason}",
                is_exception=True
            )
            raise StopTestException(stop_reason)

        # ── 成功路径：正常断电，进入下一个循环 ──────────────────────────────
        if is_success:
            self.control_relay('off')          # ← 只有成功时才断电
            self.total_success += 1
            self.log(f"单次结论: 测试通过 (固定监听周期结束，耗时: {boot_time:.2f}s)")
            time.sleep(POWER_OFF_TIME)
            rate: float = (self.total_success / cycle_num) * 100
            self.log(
                f"状态更新: 当前累计通过率为 {rate:.2f}% "
                f"(总次数 {cycle_num}，成功 {self.total_success})",
                show=True
            )
            return

        # ── 失败路径：不断电，区分故障类型，给出针对性诊断 ──────────────────
        # 注意：此处故意【不】调用 control_relay('off')
        # 继电器保持闭合 → 设备持续上电 → 可立即排查现场

        no_data = (len(logs.strip()) == 0)

        if no_data:
            fail_type = "零数据故障"
            fail_detail = (
                f"{POWER_ON_TIME}s 监听窗口内未从 {self.device_port} 收到任何字节\n"
                f"  诊断建议:\n"
                f"    ① 检查 CP210x 模块到设备 UART-TX 的接线是否断路\n"
                f"    ② 确认 {self.device_port} 串口号是否被正确识别\n"
                f"    ③ 若怀疑继电器极性反了，对调配置区"
                f" RELAY_CMD_ON / RELAY_CMD_OFF 后重试"
            )
        else:
            line_count = len(logs.splitlines())
            fail_type = "关键字未命中"
            fail_detail = (
                f"收到 {line_count} 行串口数据但未匹配任何成功关键字\n"
                f"  诊断建议:\n"
                f"    ① 检查 SUCCESS_KEYWORDS 配置是否与实际日志格式匹配\n"
                f"    ② 确认设备启动序列是否发生变化"
            )

        self.log(
            f"单次结论: 测试失败 [{fail_type}] — {fail_detail}\n"
            f"  【现场保留】继电器保持闭合，设备持续上电，请即时排查！",
            is_exception=True
        )
        raise StopTestException(f"第 {cycle_num} 次循环 [{fail_type}]: {fail_detail.splitlines()[0]}")

    def run_test(self) -> None:
        if not self.open_serial_ports():
            self.show_message("通信串口建立连接失败，请检查线路及系统端口占用情况", "初始化失败")
            return

        self.log("开始部署并初始化压力测试流环境...")

        # ── 串口连通性预验证（替代原来的 ON→3s→OFF 初始化序列）────────────
        # 该步骤同时完成物理预热 + 验证串口链路是否通畅
        if not self.verify_serial_connectivity():
            self.show_message(
                f"串口连通性验证失败，压力测试中止。\n"
                f"请检查 {self.device_port} 的串口连接后重新启动。",
                "测试中止"
            )
            return

        self.log(f"====== 压力测试正式启动 (目标执行总次数: {TEST_CYCLES}) ======")
        start_time: float = time.time()
        cycle_count: int = 0

        try:
            for i in range(1, TEST_CYCLES + 1):
                cycle_count = i
                self.run_single_cycle(i)

        except StopTestException as e:
            self.log(
                f"压力测试异常熔断。继电器当前状态: 保持闭合（设备仍处于上电状态）。\n"
                f"请排查完毕后手动断电或重启脚本。"
            )
            self.show_message(
                f"测试异常停止，设备保持上电以供现场排查\n阻断原因: {e}",
                "压力测试异常熔断"
            )
        except KeyboardInterrupt:
            self.log("收到外部键盘强行阻断信号 (Ctrl+C)，当前测试停止")
            self.control_relay('off')
            self.log("已手动断电")
        except Exception as e:
            self.main_logger.exception(f"不可预料的系统级别崩溃异常: {e}")
        finally:
            # 注意：此处仅关闭串口句柄，不主动发送继电器指令。
            # 正常结束时，最后一次 relay OFF 已在 run_single_cycle 中发送。
            # 异常结束时，继电器保持上一状态（闭合），设备上电待排查。
            if self.relay_ser:
                self.relay_ser.close()
            if self.device_ser:
                self.device_ser.close()

        elapsed: float = time.time() - start_time
        summary: str = (
            f"\n{'=' * 15} 测试统计总结报告 {'=' * 15}\n"
            f"  执行跑机总循环:        {cycle_count} 次\n"
            f"  符合开机有效条件次数:  {self.total_success} 次\n"
            f"  截获异常关键字次数:    {self.total_exceptions} 次\n"
            f"  设备通信中途断连次数:  {self.device_disconnect_count} 次\n"
            f"  流程整体执行时间:      {elapsed:.1f} 秒\n"
            f"{'=' * 46}"
        )
        self.log(summary)


if __name__ == "__main__":
    tester = RelayTester()
    tester.run_test()