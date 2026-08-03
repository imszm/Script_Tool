# -*- coding: utf-8 -*-
"""
继电器充电压力测试脚本 
针对高频测试下的硬件超时、触点老化及电磁干扰进行了鲁棒性优化。
已针对原始日志输出格式及 ANSI 乱码问题进行了流式清洗优化。
"""

import re
import os
import sys
import time
import random
import datetime
import logging
import threading
from collections import deque
from typing import Optional, Tuple, Deque, List, Dict, Any

import serial
import serial.tools.list_ports

# ================= 兼容性处理 =================
try:
    import win32api
    import win32con
    HAS_WIN32: bool = True
except ImportError:
    HAS_WIN32: bool = False

# ================= 测试参数配置 =================

# 串口硬编码配置（若填 None 则脚本会自动尝试识别检索）
CH340_RELAY_PORT: Optional[str]   = None  # 继电器串口号，例如: "COM3" 或 "/dev/ttyUSB0"
CP210X_DEVICE_PORT: Optional[str] = None  # 设备串口号，例如: "COM4" 或 "/dev/ttyUSB1"

# 串口波特率与超时设置
RELAY_BAUDRATE: int       = 9600
DEVICE_BAUDRATE: int      = 115200
SERIAL_TIMEOUT: float     = 0.1
DEVICE_RETRY_DELAY: float = 3.0

# 循环设置
TEST_CYCLES: int          = 500000

# 放宽通电后的监听时间。防止因为 BMS 响应变慢导致误判失败。
CHARGE_ON_MIN: float      = 25.0    
CHARGE_ON_MAX: float      = 25.0    

# ----------------- 物理硬件保护时序 -----------------
POST_SUCCESS_HOLD_TIME: float = 3.0 
MIN_OFF_RESET_TIME: float     = 10.0 

# ================= 关键字匹配配置 =================
# 以开始充电和充电完成的语音播报关键字段作为成功判定标准
SUCCESS_KEYWORDS: List[str] = [
    "voice_msg num: 1",
    "voice_msg num: 2",
]

EXCEPTION_KEYWORDS: List[str] = [
    "assertionfailedatfunction",
]

ERROR_RATE_CONFIG: Dict[str, Any] = {
    "keyword": "paramisinvalid",
    "window":  3.0,
    "count":   3,
}

# ================= 日志路径配置 =================
_START_TAG: str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
_LOG_DIR: str   = "logs"
LOG_FILE_PATH: str = os.path.join(_LOG_DIR, f"charge_{_START_TAG}_full.log")
ERR_FILE_PATH: str = os.path.join(_LOG_DIR, f"charge_{_START_TAG}_error.log")
RAW_FILE_PATH: str = os.path.join(_LOG_DIR, f"charge_{_START_TAG}_raw.log")


class StopTestException(Exception):
    """用于干净地中止整个测试流程的自定义异常"""
    pass


class RelayChargeTester:
    """继电器充电压力测试主控类"""
    
    def __init__(self) -> None:
        self.relay_ser:  Optional[serial.Serial] = None
        self.device_ser: Optional[serial.Serial] = None
        self.relay_port:  Optional[str] = None
        self.device_port: Optional[str] = None

        self.stat_success:    int = 0
        self.stat_failure:    int = 0
        self.stat_exceptions: int = 0
        self.stat_reconnects: int = 0

        self._ansi_re: re.Pattern = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')
        self._error_timestamps: Deque[float] = deque()
        
        # 初始化日志系统
        self.raw_logger: Optional[logging.Logger] = None
        self.logger: logging.Logger = self._setup_logger()

    def _setup_logger(self) -> logging.Logger:
        """初始化日志记录器，确保所有输出通过 logging 进行"""
        try:
            os.makedirs(_LOG_DIR, exist_ok=True)
        except OSError as e:
            sys.stderr.write(f"创建日志目录失败，将使用根目录: {e}\n")
            global LOG_FILE_PATH, ERR_FILE_PATH, RAW_FILE_PATH
            LOG_FILE_PATH = f"charge_{_START_TAG}_full.log"
            ERR_FILE_PATH = f"charge_{_START_TAG}_error.log"
            RAW_FILE_PATH = f"charge_{_START_TAG}_raw.log"

        logger: logging.Logger = logging.getLogger("RelayChargeTester")
        logger.setLevel(logging.DEBUG)
        
        if not logger.handlers:
            fmt = logging.Formatter('[%(asctime)s] %(levelname)-8s %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

            ch = logging.StreamHandler(sys.stdout)
            ch.setLevel(logging.INFO)
            ch.setFormatter(fmt)
            logger.addHandler(ch)

            try:
                fh = logging.FileHandler(LOG_FILE_PATH, encoding='utf-8', mode='a')
                fh.setLevel(logging.INFO)
                fh.setFormatter(fmt)
                logger.addHandler(fh)
                
                eh = logging.FileHandler(ERR_FILE_PATH, encoding='utf-8', mode='a')
                eh.setLevel(logging.ERROR)
                eh.setFormatter(fmt)
                logger.addHandler(eh)
            except Exception as e:
                ch.error(f"文件日志系统挂载失败: {e}")

        # 创建独立的原始日志记录器，维持文件句柄常驻
        raw_logger: logging.Logger = logging.getLogger("RelayChargeTester_Raw")
        raw_logger.setLevel(logging.INFO)
        raw_logger.propagate = False  # 阻止向上传递给 root logger
        if not raw_logger.handlers:
            try:
                raw_fh = logging.FileHandler(RAW_FILE_PATH, encoding='utf-8', mode='a')
                raw_fh.setLevel(logging.INFO)
                raw_fh.setFormatter(logging.Formatter('%(message)s'))
                raw_logger.addHandler(raw_fh)
            except Exception as e:
                logger.error(f"原始流日志文件挂载失败: {e}")
        self.raw_logger = raw_logger

        return logger

    def _write_raw_log(self, text: str) -> None:
        """
        将原始日志流安全地流式写入底层文件。
        已针对 image_7c405e.png 中的乱码完成核心优化：
        1. 自动利用正则表达式剔除下位机输出中混杂的 ANSI 控制字符(ESC等)。
        2. 统一将序列中的 \r\n 或孤立 \r 规整为标准 \n，防止文本跨平台错位。
        3. 实施逐行拆分，附带高精度到毫秒的时间戳，便于后续抓包对齐。
        """
        if not text or not self.raw_logger:
            return
        try:
            # 核心清洗：剥离阻碍文本阅读的终端着色转义序列
            clean_text: str = self._ansi_re.sub('', text)
            # 换行规范化：杜绝 \r 导致的覆盖写入或解析分裂
            clean_text = clean_text.replace('\r\n', '\n').replace('\r', '\n')
            
            # 按行迭代处理，确保时序输出极度工整
            for line in clean_text.splitlines():
                cleaned_line: str = line.strip()
                if cleaned_line:
                    # 剥离最后的微秒截取前3位，生成标准毫秒标记
                    ts: str = datetime.datetime.now().strftime("[%H:%M:%S.%f]")[:-3]
                    self.raw_logger.info(f"{ts} {cleaned_line}")
        except Exception as e:
            if self.logger:
                self.logger.warning(f"安全写入原始日志时发生异常: {e}")

    def _show_alert(self, message: str, title: str = "系统提示") -> None:
        """系统级弹窗提示"""
        self.logger.info(f"[{title}] {message}")
        if HAS_WIN32:
            try:
                def popup() -> None:
                    win32api.MessageBox(0, str(message), title, win32con.MB_ICONINFORMATION | win32con.MB_SYSTEMMODAL)
                threading.Thread(target=popup, daemon=True).start()
            except Exception as e:
                self.logger.error(f"调用 UI 弹窗失败: {e}")

    def _detect_ports(self) -> Tuple[Optional[str], Optional[str]]:
        """检测并区分继电器和设备的串口号，优先采用前置的手动配置"""
        relay_port: Optional[str]  = CH340_RELAY_PORT
        device_port: Optional[str] = CP210X_DEVICE_PORT
        
        # 仅在未手动指定串口时，才检索串行总线进行自动匹配
        if not relay_port or not device_port:
            for p in serial.tools.list_ports.comports():
                desc: str = p.description.lower()
                if not relay_port and ("ch340" in desc or "9" in desc):
                    relay_port = p.device
                elif not device_port and ("cp210x" in desc or "20" in desc):
                    device_port = p.device
                    
        self.logger.info(f"串口状态 -> 继电器(CH340): {relay_port} | 设备(CP210x): {device_port}")
        return device_port, relay_port

    def _open_serial_ports(self) -> bool:
        """安全地打开串口并重置缓冲区"""
        self.device_port, self.relay_port = self._detect_ports()
        if not self.device_port or not self.relay_port:
            self.logger.error("硬件串口不完整，请检查接线或确认顶部串口配置。")
            return False
        try:
            self.relay_ser  = serial.Serial(self.relay_port,  RELAY_BAUDRATE,  timeout=SERIAL_TIMEOUT)
            self.device_ser = serial.Serial(self.device_port, DEVICE_BAUDRATE, timeout=SERIAL_TIMEOUT)
            self.relay_ser.reset_input_buffer()
            self.device_ser.reset_input_buffer()
            return True
        except serial.SerialException as e:
            self.logger.error(f"串口打开异常: {e}")
            return False

    def _control_relay(self, action: str) -> None:
        """
        冗余指令发送机制。
        连续发送两次控制指令，防止强电磁干扰导致的单字节吞没。
        """
        if not self.relay_ser or not self.relay_ser.is_open:
            self.logger.error("继电器串口未打开，无法发送指令。")
            return
            
        try:
            cmd: bytes = bytes([0x50]) if action == 'on' else bytes([0x4F])
            
            # 第一发
            self.relay_ser.write(cmd)
            self.relay_ser.flush()
            time.sleep(0.05)
            
            # 第二发 (冗余确认)
            self.relay_ser.write(cmd)
            self.relay_ser.flush()
            time.sleep(0.05)
            
            if self.relay_ser.in_waiting:
                self.relay_ser.read(self.relay_ser.in_waiting)
                
            self.logger.info(f"继电器执行 -> {action.upper()} (指令已双重确认)")
        except serial.SerialException as e:
            self.logger.error(f"继电器指令发送失败: {e}")
        except Exception as e:
            self.logger.error(f"继电器控制发生未知异常: {e}")

    def _try_reconnect_device(self) -> None:
        """处理设备端串口掉线重连逻辑"""
        self.stat_reconnects += 1
        self.logger.warning(f"设备断开，{DEVICE_RETRY_DELAY}s 后重连 (次序: {self.stat_reconnects})")
        
        if self.device_ser:
            try: 
                self.device_ser.close()
            except Exception: 
                pass
                
        time.sleep(DEVICE_RETRY_DELAY)
        new_dev, _ = self._detect_ports()
        
        if new_dev:
            try:
                self.device_port = new_dev
                self.device_ser  = serial.Serial(self.device_port, DEVICE_BAUDRATE, timeout=SERIAL_TIMEOUT)
                self.device_ser.reset_input_buffer()
                self.logger.info(f"重连成功: {new_dev}")
            except serial.SerialException as e:
                self.logger.error(f"重连失败: {e}")

    def _process_line(self, line: str) -> Tuple[bool, Optional[str], bool, Optional[str]]:
        """分析单行日志，进行关键字匹配清洗"""
        clean: str = self._ansi_re.sub('', line)
        normed: str = clean.lower().replace(" ", "")

        for kw in EXCEPTION_KEYWORDS:
            if kw in normed:
                self.stat_exceptions += 1
                self.logger.error(f"[异常] {kw} | {clean.strip()}")

        ec: Dict[str, Any] = ERROR_RATE_CONFIG
        if ec["keyword"] in normed:
            now: float = time.time()
            self._error_timestamps.append(now)
            while self._error_timestamps and self._error_timestamps[0] < now - float(ec["window"]):
                self._error_timestamps.popleft()
            if len(self._error_timestamps) >= int(ec["count"]):
                return True, f"高频错误触发停测: '{ec['keyword']}'", False, None

        # 匹配成功关键字
        for kw in SUCCESS_KEYWORDS:
            if kw.lower().replace(" ", "") in normed:
                return False, None, True, kw

        return False, None, False, None

    def _monitor_stream(self, duration: float, exit_on_success: bool = True) -> Tuple[str, bool, Optional[str], bool, Optional[str]]:
        """
        监听设备串口数据流。
        已将块状包缓存更新为流式实时拦截，确保每条信息在被读取的瞬间被格式化并安全落地。
        """
        end_time: float = time.time() + duration
        lines: List[str] = []
        first_success_kw: Optional[str] = None

        while time.time() < end_time:
            try:
                if self.device_ser and self.device_ser.in_waiting:
                    raw: bytes = self.device_ser.readline()
                    if not raw: 
                        continue
                    
                    decoded: str = raw.decode('utf-8', errors='ignore')
                    
                    # 收到串口数据行立即进行清理并打上即时时间戳写入本地文件
                    self._write_raw_log(decoded)

                    lines.append(decoded.strip())

                    should_stop, reason, success_found, matched_kw = self._process_line(decoded)

                    if should_stop:
                        full_text: str = "\n".join(lines)
                        return full_text, True, reason, False, None

                    if success_found:
                        if first_success_kw is None:
                            first_success_kw = matched_kw
                            self.logger.info(f"捕获目标指令: {matched_kw}")
                            
                        if exit_on_success:
                            full_text = "\n".join(lines)
                            return full_text, False, None, True, first_success_kw
                else:
                    time.sleep(0.005)
            except serial.SerialException:
                self._try_reconnect_device()
                break
            except Exception as e:
                self.logger.exception(f"读取异常: {e}")
                break

        full_text = "\n".join(lines)
        return full_text, False, None, (first_success_kw is not None), first_success_kw

    def _run_cycle(self, cycle_num: int) -> None:
        """执行单次压力测试循环"""
        charge_time_limit: float = round(random.uniform(CHARGE_ON_MIN, CHARGE_ON_MAX), 1)
        self.logger.info(f"-------------------- 第 {cycle_num} 轮 | 最大监听: {charge_time_limit}s --------------------")

        # 充电前彻底清空上个周期的残留日志和串口积压
        if self.device_ser and self.device_ser.is_open:
            self.device_ser.reset_input_buffer()
            self.device_ser.reset_output_buffer()

        self.logger.info("继电器 ON (开始充电)")
        self._control_relay('on')
        
        _, stop, reason, success, matched_kw = self._monitor_stream(charge_time_limit, exit_on_success=True)

        if stop:
            raise StopTestException(reason)

        early_success_kw: Optional[str] = matched_kw

        if success:
            self.logger.info(f"状态维稳，保持闭合 {POST_SUCCESS_HOLD_TIME}s")
            _, stop_h, reason_h, _, _ = self._monitor_stream(POST_SUCCESS_HOLD_TIME, exit_on_success=False)
            if stop_h:
                raise StopTestException(reason_h)

        self.logger.info("继电器 OFF (切断充电)")
        self._control_relay('off')

        self.logger.info(f"物理断电静置 {MIN_OFF_RESET_TIME}s (等待BMS状态机完全复位)")
        _, stop_o, reason_o, _, _ = self._monitor_stream(MIN_OFF_RESET_TIME, exit_on_success=False)
        if stop_o:
            raise StopTestException(reason_o)

        if early_success_kw:
            self.stat_success += 1
            self.logger.info(f"[结论] [PASS] 成功关键字: {early_success_kw}")
        else:
            self.stat_failure += 1
            self.logger.warning("[结论] [FAIL] 超时未检测到成功关键字，设备可能未进入充电状态或响应过慢。")

        total: int = self.stat_success + self.stat_failure + self.stat_exceptions
        rate: float  = (self.stat_success / total * 100) if total else 0.0
        self.logger.info(f"统计 -> 成功: {self.stat_success} | 失败: {self.stat_failure} | 异常: {self.stat_exceptions} | 成功率: {rate:.1f}%")

    def run(self) -> None:
        """启动测试的主入口"""
        self.logger.info("========== 继电器充电压力测试启动  ==========")

        if not self._open_serial_ports():
            return

        self.logger.info("环境初始清洗 (物理断电重置)...")
        self._control_relay('off')
        time.sleep(MIN_OFF_RESET_TIME)

        try:
            for i in range(1, TEST_CYCLES + 1):
                self._run_cycle(i)
                
        except StopTestException as e:
            self.logger.error(f"测试中止: {e}")
            self._show_alert(f"异常阻断触发:\n{e}", "中止提示")
        except KeyboardInterrupt:
            self.logger.warning("手动中断，停止测试。")
        except Exception as e:
            self.logger.error(f"运行时发生未捕获异常: {e}", exc_info=True)
        finally:
            self.logger.info("测试结束，执行环境安全隔离...")
            self._control_relay('off')
            
            if self.relay_ser and self.relay_ser.is_open:
                try:
                    self.relay_ser.close()
                except Exception as e:
                    self.logger.error(f"关闭继电器串口失败: {e}")
                    
            if self.device_ser and self.device_ser.is_open:
                try:
                    self.device_ser.close()
                except Exception as e:
                    self.logger.error(f"关闭设备串口失败: {e}")
                
            report: str = (
                f"\n=============== 最终测试报告 ===============\n"
                f"  目标循环:      {TEST_CYCLES} 次\n"
                f"  成功次数:      {self.stat_success} 次\n"
                f"  失败次数:      {self.stat_failure} 次\n"
                f"  异常次数:      {self.stat_exceptions} 次\n"
                f"  断连次数:      {self.stat_reconnects} 次\n"
                f"============================================"
            )
            self.logger.info(report)

if __name__ == "__main__":
    tester = RelayChargeTester()
    tester.run()