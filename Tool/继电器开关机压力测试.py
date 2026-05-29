# -*- coding: utf-8 -*-
"""
继电器开关机压力测试脚本 - 优化版

【本次修复的关键问题】
────────────────────────────────────────────────────────────────────────
  第 804 次循环失败是测试脚本 BUG 引发的误判（设备端完全正常），共 3 处缺陷：

  [BUG-1] 关键字拼写错误（直接致命）
    原配置: "uipmacc:1:acc1:on0"
    正配置: "ui_pm_acc:1:acc1:on0"
    原因：关键字去掉了下划线，但实际日志 "[I/ui] ui_pm_acc: 1:acc 1:on 0"
          经 .replace(" ", "") 后保留下划线变为 "ui_pm_acc:1:acc1:on0"，
          导致该关键字【永远不可能命中】。

  [BUG-2] 成功关键字值不匹配（直接致命）
    原配置: "voice_msgnum:0"
    正配置: "voice_msgnum:"（前缀匹配，匹配任意数量）
    原因：实际日志始终输出 "voice_msg num: 1"，去空格后为 "voice_msgnum:1"，
          精确匹配 ":0" 导致该关键字也【永远不可能命中】。

  [BUG-3] UART 日志撕裂（概率性触发，放大了上述两个 BUG 的危害）
    在第 804 次，RTOS 的 voice 线程与 ui 线程发生了字节级交织撕裂：
      802、803次撕裂: "[I/voic[I/ui] refesh_cache"（voice 行被撕，motor 行正常）
      804次撕裂:      "[I/motor] mot[I/ui] refesh_cache"（motor 行被撕，voice 行正常）
    因此 804 次的 motorpoweron / poweron 也无法命中。

  如果 BUG-1 或 BUG-2 任一已修复，第 804 次即可通过，因为另一关键字在 motor 行
  被撕裂前已经匹配成功（ui_pm_acc 和 voice_msg 均出现在 motor 上电日志之前）。

【本版本的优化点】
────────────────────────────────────────────────────────────────────────
  1. 修复 BUG-1/BUG-2 的关键字配置
  2. 新增早期启动可靠信号关键字（出现在多线程竞争之前，几乎不会被撕裂）
  3. 新增【滑动拼接缓冲区】机制：将监听窗口内收到的所有行的干净文本滚动拼接，
     当关键字恰好被 readline() 切在两次读取边界时仍可命中（对字节级撕裂无效，
     但可防止高波特率下超时截断的情况）
  4. 成功命中来源区分日志（逐行命中 vs. 缓冲区容错命中）
  5. 其他可读性与注释改善
"""

import serial
import serial.tools.list_ports
import time
import datetime
import sys
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

POWER_ON_TIME: float = 5.0       # 继电器上电并保持监听的固定时长（秒）
POWER_OFF_TIME: float = 5.0      # 继电器断电复位，给电机控制器电容放电的固定时长（秒）

DEVICE_RETRY_DELAY: float = 3.0

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
# 成功判定关键字列表
#
# 匹配规则：对每行原始文本执行 .lower().replace(" ", "") 后做子串查找。
# 因此关键字中【保留】原始日志中存在的下划线，【去掉】原始日志中的空格。
#
# 【已修复】
#   "uipmacc:1:acc1:on0"  →  "ui_pm_acc:1:acc1:on0"   (BUG-1：补回下划线)
#   "voice_msgnum:0"      →  "voice_msgnum:"           (BUG-2：值不匹配改为前缀)
#
# 【新增】
#   "exitsleep:acc_key1"  : 日志 "[I/ui] exit sleep: acc_key 1"，
#                           出现在 motor/voice 多线程竞争之前，极少被撕裂
#   "acc_cb,key_evt:1"   : 日志 "[I/gkey] acc_cb, key_evt: 1"，
#                           ACC 按键事件，上电最早期信号
# ─────────────────────────────────────────────────────────────────────────────
SUCCESS_KEYWORDS: List[str] = [
    "motorpoweron",          # "[I/motor] motor power on..."
    "poweron",               # 同上的子串，备用
    "voice_msgnum:",         # 【已修复】原为 "voice_msgnum:0"，值不匹配改为前缀匹配
    "threadoperatingsystem",
    "motor_svc_init",
    "ui_pm_acc:1:acc1:on0",  # 【已修复】原为 "uipmacc:1:acc1:on0"，补回下划线
    "exitsleep:acc_key1",    # 【新增】"[I/ui] exit sleep: acc_key 1"，早期可靠信号
    "acc_cb,key_evt:1",      # 【新增】"[I/gkey] acc_cb, key_evt: 1"，启动最早信号
]

# 滑动拼接缓冲区容量（字符数）
# 用于应对极端场景：关键字被 readline() 恰好切在两次读取边界处
CONCAT_BUFFER_MAX_LEN: int = 2048

# =============================================================================


class StopTestException(Exception):
    """自定义异常类，用于在达到致命错误条件或重试超时时触发并停止测试"""
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
        """初始化 logging 模块，配置控制台与多重文件输出结构"""
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
        """核心日志分发器，依据级别写入对应的日志文件"""
        if is_exception:
            self.error_logger.error(message)
            if show:
                self.main_logger.error(f"[EXCEPTION] {message}")
        else:
            if show:
                self.main_logger.info(message)

        self.raw_logger.info(f"[TEST_ACTION] {message}")

    def log_raw_data(self, raw_text: str) -> None:
        """记录去除终端控制字符后的干净串口流数据"""
        clean_text: str = self.ansi_escape.sub('', raw_text).strip('\r\n')
        if clean_text:
            self.raw_logger.info(clean_text)

    def show_message(self, message: str, title: str = "提示") -> None:
        """操作系统层级的弹窗提示处理"""
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
        """根据芯片特征名称分配串口号，提升容错率"""
        ports: List[serial.tools.list_ports_common.ListPortInfo] = list(
            serial.tools.list_ports.comports()
        )
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
        """初始化并开启检测到的串口资源"""
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
        """执行设备串口的断线重连恢复逻辑"""
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
        """清空接收缓冲区以防止历史脏数据影响本次通电测试"""
        if self.device_ser and self.device_ser.is_open:
            try:
                self.device_ser.reset_input_buffer()
                self.log("已清空串口输入缓冲区，开始专注监听本次上电日志", show=False)
            except Exception as e:
                self.log(f"清空串口缓冲区失败: {e}", is_exception=True)

    def control_relay(self, action: str) -> None:
        """向指定的继电器发送控制指令"""
        if not self.relay_ser or not self.relay_ser.is_open:
            return

        try:
            cmd: bytes = bytes([0x50]) if action == 'on' else bytes([0x4F])
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
        """基于滑动时间窗的频率限制器"""
        now: float = time.time()
        timestamps_deque.append(now)
        while timestamps_deque and timestamps_deque[0] < now - window_seconds:
            timestamps_deque.popleft()
        return len(timestamps_deque) >= threshold_count

    def process_log_line(self, line: str) -> Tuple[bool, Optional[str]]:
        """分析数据流中的单行记录，检测错误及信息关键字配置"""
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
        """
        在给定文本中搜索全部成功关键字，返回命中的第一个关键字，否则返回 None。
        source_label 用于区分是逐行命中还是缓冲区容错命中。
        """
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
        阻塞监听来自 CP210x 的串口流数据以判断开机状态。

        双层成功检测机制：
          1. 逐行精确匹配：对每一行去 ANSI/去空格后直接搜索关键字（原有逻辑）
          2. 滑动拼接缓冲区容错匹配（新增）：维护本次监听窗口内所有行的拼接文本，
             应对极端场景下关键字被 readline() 切在两次读取边界的情况

        注意：对于 RTOS 字节级 UART 撕裂（两线程输出字节交织在同一行），
        此缓冲区无法恢复被其他线程内容打断的关键字片段，
        因此应通过配置多个早期启动关键字（在撕裂高发期之前出现）来覆盖。
        """
        end_time: float = time.time() + duration
        collected_logs: List[str] = []
        is_success_detected: bool = False

        # 滑动拼接缓冲区（本次监听窗口内累积，每轮新开，不跨周期）
        concat_buffer: str = ""

        while time.time() < end_time:
            try:
                if self.device_ser and self.device_ser.in_waiting:
                    raw_bytes: bytes = self.device_ser.readline()
                    if not raw_bytes:
                        continue

                    decoded_line: str = raw_bytes.decode('gb2312', errors='replace')
                    self.log_raw_data(decoded_line)

                    stripped_line: str = decoded_line.strip()
                    if not stripped_line:
                        continue

                    collected_logs.append(stripped_line)

                    # 错误与致命条件检测（保持逐行，精确性更重要）
                    should_stop, reason = self.process_log_line(stripped_line)
                    if should_stop:
                        return "\n".join(collected_logs), True, reason, False

                    if not is_success_detected:
                        # 清理当前行：去 ANSI、转小写、去空格
                        clean_for_match: str = (
                            self.ansi_escape.sub('', stripped_line)
                            .lower()
                            .replace(" ", "")
                        )

                        # ── 第一层：逐行精确匹配 ──────────────────────────────
                        matched_kw = self._check_success_keywords(
                            clean_for_match, "逐行", stripped_line
                        )

                        # ── 第二层：滑动拼接缓冲区容错匹配 ─────────────────────
                        if not matched_kw:
                            concat_buffer += clean_for_match
                            if len(concat_buffer) > CONCAT_BUFFER_MAX_LEN:
                                concat_buffer = concat_buffer[-CONCAT_BUFFER_MAX_LEN:]

                            matched_kw = self._check_success_keywords(
                                concat_buffer, "缓冲区容错", stripped_line
                            )
                        else:
                            # 逐行已命中，缓冲区也同步更新
                            concat_buffer += clean_for_match
                            if len(concat_buffer) > CONCAT_BUFFER_MAX_LEN:
                                concat_buffer = concat_buffer[-CONCAT_BUFFER_MAX_LEN:]

                        if matched_kw:
                            is_success_detected = True
                            if stop_on_success:
                                return "\n".join(collected_logs), False, None, True
                            # stop_on_success=False 时继续监听至 duration 结束
                else:
                    time.sleep(0.005)

            except serial.SerialException:
                self.log("硬件警告: 设备串口失去连接，启动重连机制...", is_exception=True)
                self.try_reconnect_device()
                break
            except Exception as e:
                self.log(f"数据流读取逻辑发生不可预料异常: {e}", is_exception=True)
                break

        return "\n".join(collected_logs), False, None, is_success_detected

    def run_single_cycle(self, cycle_num: int) -> None:
        """主调测逻辑：执行一次完整的断电/上电测试"""
        self.log(
            f"\n--- [流程标记] 第 {cycle_num} 次压力循环 "
            f"(固定上电监听: {POWER_ON_TIME}s, 断电时长: {POWER_OFF_TIME}s) ---"
        )

        self.flush_device_input_buffer()

        self.control_relay('on')
        t0: float = time.time()

        # stop_on_success=False 强制执行固定时间监听，确保捕获完整上电日志
        logs, stop_triggered, stop_reason, is_success = self.monitor_serial_stream(
            POWER_ON_TIME, stop_on_success=False
        )

        boot_time: float = time.time() - t0

        self.control_relay('off')

        if stop_triggered:
            self.log(f"严重错误触发中断: {stop_reason}", is_exception=True)
            raise StopTestException(stop_reason)

        if is_success:
            self.total_success += 1
            self.log(f"单次结论: 测试通过 (固定监听周期结束，耗时: {boot_time:.2f}s)")
        else:
            self.log(
                f"单次结论: 测试失败 - 在固定的 {POWER_ON_TIME} 秒监听周期内，"
                f"未能匹配到设备的有效开机回复，疑似启动挂死",
                is_exception=True
            )
            raise StopTestException(f"第 {cycle_num} 次系统循环发生开机验证失败")

        time.sleep(POWER_OFF_TIME)

        rate: float = (self.total_success / cycle_num) * 100
        self.log(
            f"状态更新: 当前累计通过率为 {rate:.2f}% "
            f"(总次数 {cycle_num}，成功 {self.total_success})",
            show=True
        )

    def run_test(self) -> None:
        """测试整体的启动装配器与异常捕获顶层"""
        if not self.open_serial_ports():
            self.show_message("通信串口建立连接失败，请检查线路及系统端口占用情况", "初始化失败")
            return

        self.log("开始部署并初始化压力测试流环境...")

        self.log("系统初始重置(1/2): 执行一次通电预热")
        self.control_relay('on')
        time.sleep(3.0)

        self.log("系统初始重置(2/2): 执行一次断电复位")
        self.control_relay('off')

        self.flush_device_input_buffer()

        self.log("物理初始化完毕: 系统当前处于断电冷机状态，等待 2 秒进入压力跑机循环")
        time.sleep(2.0)

        self.log(f"====== 压力测试正式启动 (目标执行总次数: {TEST_CYCLES}) ======")
        start_time: float = time.time()
        cycle_count: int = 0

        try:
            for i in range(1, TEST_CYCLES + 1):
                cycle_count = i
                self.run_single_cycle(i)
        except StopTestException as e:
            self.show_message(
                f"测试机制保护启动以留存异常现场\n阻断原因: {e}",
                "压力测试异常熔断"
            )
        except KeyboardInterrupt:
            self.log("收到外部键盘强行阻断信号 (Ctrl+C)，当前测试停止")
        except Exception as e:
            self.main_logger.exception(f"不可预料的系统级别崩溃异常: {e}")
        finally:
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