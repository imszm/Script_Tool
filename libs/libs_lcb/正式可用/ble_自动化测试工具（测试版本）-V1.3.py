# -*- coding: utf-8 -*-
"""
BLE 协议自动化测试工具（完整版本）
============================================================
功能概述：
1) 复用原有 BLE 协议脚本：ctypes 加载 ppx_ble.dll，pyserial 串口通信，解析/组包
2) 新增自动化测试框架：
   - 从 Excel/CSV 读取用例（默认 ./testcases.xlsx 或 ./testcases.csv）
   - 执行批量测试与可选压力循环
   - 自动断言：对比 DLL 全局变量(g_ppx_ble_data.led_msg)与期望字段
   - 日志与报告：保存原始日志、结果 CSV、可读的 HTML 报告（含统计）
3) 兼容性：若未安装 pandas/openpyxl，会自动退化到 CSV 读取；可用 --make-sample 生成示例用例

使用示例：
------------------------------------------------------------
1. 直接运行（读取默认Excel用例 testcases.xlsx）：
   py -3.11 自动化BLE测试.py

2. 指定 Excel 用例与串口：
   py -3.11 自动化BLE测试.py --cases E:\\cases\\ble_led_cases.xlsx --port COM97 --baud 460800

3. 压力循环（在执行整表用例之后，再对第一条用例循环 100 次，每次间隔 0.5s）：
   py -3.11 自动化BLE测试.py --loop-count 100 --loop-delay 0.5

4. 生成一份示例Excel：
   py -3.11 自动化BLE测试.py --make-sample

Excel/CSV 用例字段说明：
------------------------------------------------------------
必需（用于设置 LED）：
- screen_on, brightness, digital, logo, rim_state, rdygo, turn_left, turn_right, ring

可选（用于断言，任意字段可留空；留空则不校验该项）：
- expect_screen_on, expect_brightness, expect_digital, expect_logo, expect_rim_state,
  expect_rdygo, expect_turn_left, expect_turn_right, expect_ring

其它可选控制：
- recv_timeout  (float, 单位秒；覆盖默认接收超时)
- delay_after   (float, 单位秒；每条用例执行后的延时)
- comment       (字符串，备注)

输出产物（自动按时间戳归档）：
------------------------------------------------------------
./reports/2025-08-15_20-xx-xx/
  - raw.log                 原始详细日志（时间戳+级别）
  - results.csv             每条用例的结果明细
  - report.html             富文本测试报告（统计+表格）

注意事项：
------------------------------------------------------------
- 请确保 DLL 路径正确且与 Python 架构匹配（x86/x64），串口参数正确。
- 若使用 Excel，建议安装：  pip install pandas openpyxl
- 若无法安装依赖，请改用 CSV 用例（UTF-8 编码）。
"""

import ctypes
from ctypes import *
import serial
import time
import platform
import sys
import os
import argparse
import datetime as _dt
from typing import List, Dict, Any, Optional, Tuple
import itertools
import chardet
import csv



# 尝试导入 pandas（用于 Excel/CSV 读写），失败则退化到 CSV 解析
try:
    import pandas as pd  # type: ignore
    _HAS_PANDAS = True
except Exception:
    _HAS_PANDAS = TRUE

# ==============================================
# 配置区域 - 根据实际情况修改这些参数
# ==============================================
BLE_PROTOCOL_DLL = r"C:\Test_ziliao\libs\libs\ppx_ble.dll"  # BLE协议DLL文件名
SERIAL_PORT = "COM54"                              # 串口端口
BAUDRATE = 460800                                   # 串口波特率
DEBUG_MODE = True                                   # 调试模式，打印更多信息
DEFAULT_RECV_TIMEOUT = 1.0                          # 串口接收默认超时（秒）

# ==============================================
# 协议常量定义
# ==============================================
# 设备ID
PPX_ID_BLE = 0x60

# 命令类型
PPX_MSG_READ = 0x01
PPX_MSG_WRITE = 0x03

# 命令类别
class PpxCmdType:
    REQ = 0  # 请求命令
    RSP = 1  # 响应命令

# 寄存器地址
PPX_BLE_LED_MSG_REG = 0x08

# 解析状态
PPX_PARSE_SUCCESS = 1
PPX_PARSE_FAILURE = 0

# ==============================================
# 结构体定义 (必须与DLL中的定义完全一致)
# ==============================================
class ppx_ble_msg_t(Structure):
    _fields_ = [
        ("id", c_uint8),       # 设备ID
        ("cmd", c_uint8),      # 写/读命令
        ("reg_addr", c_uint8), # BLE数据地址
        ("reg_nums", c_uint8)  # BLE寄存器数量
    ]

class ppx_led_msg_t(Structure):
    _fields_ = [
        ("screen_on", c_uint32, 1),      # 显示开关: 1开, 0关
        ("brightness", c_uint32, 3),     # 亮度级别 0-7
        ("blink_period", c_uint32, 4),   # 闪烁周期: N * 200ms
        ("blink_duty", c_uint32, 4),     # 闪烁占空比
        ("blink_en", c_uint32, 8),       # 闪烁使能
        ("err_flag", c_uint32, 2),       # 错误码标志
        ("err_code", c_uint32, 4),       # 错误码: 0-F
        ("digital", c_uint32, 7),        # 电池SOC: 0-100
        ("logo", c_uint32, 2),           # LOGO: 0关, 1白, 2红
        ("rim_state", c_uint32, 2),      # 护盾: 0关, 1白, 2绿
        ("rdygo", c_uint32, 2),          # Ready Go: 0关, 1白, 2红
        ("turn_left", c_uint32, 2),      # 左转向灯: 0关, 1白, 2橙
        ("turn_right", c_uint32, 2),     # 右转向灯: 0关, 1白, 2橙
        ("ring", c_uint32, 2),           # 灯环: 0关, 1蓝, 2红
        ("rsvd_data", c_uint32, 19)      # 保留
    ]

class ppx_ble_data_t(Structure):
    _fields_ = [
        ("id_num", c_uint8),                     # 设备ID号
        ("model", c_uint8 * 8),                  # 型号 (8字节)
        ("serial_num", c_uint8 * 26),            # 序列号 (26字节)
        ("hw_version", c_uint8),                 # 硬件版本
        ("sw_version", c_uint8 * 20),            # 软件版本 (20字节)
        ("status", c_uint32),                    # 状态
        ("ldr_value", c_uint16),                 # 光敏电阻亮度
        ("io_status", c_uint16),                 # IO引脚状态
        ("led_msg", ppx_led_msg_t),              # LED显示信息
        ("card_id", c_uint32),                   # NFC卡ID
        ("dat_setting", c_uint32)                # 数据设置
    ]

# ==============================================
# BLE协议通信类（复用并增强日志）
# ==============================================
class BLEProtocol:
    def __init__(self, dll_path: str, serial_port: str, baudrate: int, recv_timeout: float = DEFAULT_RECV_TIMEOUT):
        self.dll_loaded = False
        self.serial_connected = False
        self.recv_timeout = recv_timeout

        # 日志回调（由外部 AutoTester 注入），默认为打印
        self._logger = None  # type: Optional[callable]

        # 加载DLL
        self._load_dll(dll_path)

        # 初始化串口
        self._init_serial(serial_port, baudrate)

        # 检查全局变量
        self._check_global_vars()

    def set_logger(self, logger_func):
        """设置日志回调：logger_func(level:str, message:str)"""
        self._logger = logger_func

    # ---------------- 内部工具 ----------------
    def _log(self, level: str, message: str):
        if self._logger:
            self._logger(level, message)
        else:
            print(f"[{level}] {message}")

    def _debug_print(self, message: str, is_error: bool = False):
        prefix = "ERROR" if is_error else "DEBUG"
        if is_error or DEBUG_MODE:
            self._log(prefix, message)

    # ---------------- DLL/串口初始化 ----------------
    def _load_dll(self, dll_path: str):
        try:
            self.ble_lib = cdll.LoadLibrary(dll_path)
            self.dll_loaded = True
            self._debug_print(f"成功加载DLL: {dll_path}")

            # 严格按照给定的函数原型设置
            self.ble_lib.ppx_com_ble_parse.argtypes = [
                POINTER(c_uint8),    # pdata
                c_uint8,             # data_len
                POINTER(ppx_ble_msg_t)  # ble_msg
            ]
            self.ble_lib.ppx_com_ble_parse.restype = c_int

            self.ble_lib.ppx_com_ble_format.argtypes = [
                c_int,                # cmd_type
                POINTER(ppx_ble_msg_t), # ble_msg
                c_void_p              # buffer
            ]
            self.ble_lib.ppx_com_ble_format.restype = c_uint16

        except Exception as e:
            self._debug_print(f"加载DLL失败: {e}", is_error=True)
            self.dll_loaded = False

    def _init_serial(self, port: str, baudrate: int):
        try:
            self.serial_port = serial.Serial(
                port=port,
                baudrate=baudrate,
                timeout=self.recv_timeout
            )
            self.serial_connected = True
            self._debug_print(f"串口已连接: {port}, 波特率: {baudrate}")
        except Exception as e:
            self._debug_print(f"串口连接失败: {e}", is_error=True)
            self.serial_connected = False

    def _check_global_vars(self):
        try:
            self.g_ppx_ble_data = ppx_ble_data_t.in_dll(self.ble_lib, "g_ppx_ble_data")
            self._debug_print("成功获取全局变量 g_ppx_ble_data")

            if DEBUG_MODE:
                self._print_ble_data(self.g_ppx_ble_data)

        except Exception as e:
            self._debug_print(f"获取全局变量失败: {e}", is_error=True)

    # ---------------- 基本通信 ----------------
    def close(self):
        if hasattr(self, 'serial_port') and self.serial_port and self.serial_port.is_open:
            self.serial_port.close()
            self.serial_connected = False
            self._debug_print("串口已关闭")

    def send_data(self, data: bytes) -> bool:
        if not self.serial_connected:
            self._debug_print("串口未连接，无法发送数据", is_error=True)
            return False
        try:
            self._debug_print(f"发送数据: {self._bytes_to_hex(data)}")
            self.serial_port.write(data)
            return True
        except Exception as e:
            self._debug_print(f"发送数据失败: {e}", is_error=True)
            return False

    def receive_data(self, timeout: Optional[float] = None, max_length: int = 512) -> Optional[bytes]:
        if not self.serial_connected:
            self._debug_print("串口未连接，无法接收数据", is_error=True)
            return None
        try:
            start_time = time.time()
            data = bytes()
            _timeout = self.recv_timeout if timeout is None else timeout
            while time.time() - start_time < _timeout:
                waiting = self.serial_port.in_waiting
                if waiting > 0:
                    data += self.serial_port.read(waiting)
                    if len(data) >= max_length:
                        break
                else:
                    time.sleep(0.005)  # 微等，减少CPU占用
            if data:
                self._debug_print(f"接收数据: {self._bytes_to_hex(data)}")
                return data
            else:
                self._debug_print("接收数据超时")
                return None
        except Exception as e:
            self._debug_print(f"接收数据失败: {e}", is_error=True)
            return None

    def parse_data(self, data: bytes) -> Tuple[bool, Optional[ppx_ble_msg_t], Optional[int]]:
        """ 用 DLL 的解析函数解析数据包 """
        if not data:
            self._debug_print("无数据可解析", is_error=True)
            return False, None, None
        if not self.dll_loaded:
            self._debug_print("DLL未加载，无法解析数据", is_error=True)
            return False, None, None
        try:
            ble_msg = ppx_ble_msg_t()
            ble_msg.id = PPX_ID_BLE  # 关键：预填充ID

            data_len = len(data)
            data_array = (c_uint8 * data_len)(*data)

            result = self.ble_lib.ppx_com_ble_parse(
                cast(data_array, POINTER(c_uint8)),
                c_uint8(data_len),
                byref(ble_msg)
            )

            if result == PPX_PARSE_SUCCESS:
                self._debug_print("数据解析成功:")
                self._print_ble_msg(ble_msg)
                # 更新全局变量信息（DLL 内部应已更新 g_ppx_ble_data）
                self._print_ble_data_changes(ble_msg)
                # 粗略提取命令类型（若协议规定首字节低4位为类型，可根据实际更改）
                cmd_type = data[0] & 0x0F if len(data) > 0 else None
                return True, ble_msg, cmd_type
            else:
                self._debug_print(f"数据解析失败，返回码: {result}", is_error=True)
                return False, None, None
        except Exception as e:
            self._debug_print(f"解析数据时发生异常: {e}", is_error=True)
            return False, None, None

    def format_data(self, cmd_type: int, ble_msg: ppx_ble_msg_t) -> Tuple[bool, Optional[bytes]]:
        if not self.dll_loaded:
            self._debug_print("DLL未加载，无法组包数据", is_error=True)
            return False, None
        try:
            buffer_size = 256
            buffer = create_string_buffer(buffer_size)

            length = self.ble_lib.ppx_com_ble_format(
                cmd_type,
                byref(ble_msg),
                buffer
            )

            if length > 0:
                data = buffer.raw[:length]
                self._debug_print(f"组包成功，数据长度: {length} 字节")
                return True, data
            else:
                self._debug_print("组包失败，返回长度为0", is_error=True)
                return False, None
        except Exception as e:
            self._debug_print(f"组包数据时发生异常: {e}", is_error=True)
            return False, None

    # ---------------- 业务操作：LED 设置/读取 ----------------
    def set_led_display(self,
                        screen_on=1, brightness=0, digital=0,
                        logo=0, rim_state=0, rdygo=0,
                        turn_left=0, turn_right=0, ring=0,
                        recv_timeout: Optional[float] = None) -> Tuple[bool, Optional[bytes]]:
        if not self.dll_loaded:
            self._debug_print("DLL未加载，无法设置LED", is_error=True)
            return False, None
        try:
            # 修改全局变量中的LED显示信息
            self.g_ppx_ble_data.led_msg.screen_on = int(screen_on)
            self.g_ppx_ble_data.led_msg.brightness = int(brightness)
            self.g_ppx_ble_data.led_msg.digital = int(digital)
            self.g_ppx_ble_data.led_msg.logo = int(logo)
            self.g_ppx_ble_data.led_msg.rim_state = int(rim_state)
            self.g_ppx_ble_data.led_msg.rdygo = int(rdygo)
            self.g_ppx_ble_data.led_msg.turn_left = int(turn_left)
            self.g_ppx_ble_data.led_msg.turn_right = int(turn_right)
            self.g_ppx_ble_data.led_msg.ring = int(ring)

            self._debug_print("修改后的LED显示信息:")
            self._print_led_msg(self.g_ppx_ble_data.led_msg)

            # 准备消息结构体
            ble_msg = ppx_ble_msg_t()
            ble_msg.id = PPX_ID_BLE
            ble_msg.cmd = PPX_MSG_WRITE
            ble_msg.reg_addr = PPX_BLE_LED_MSG_REG
            ble_msg.reg_nums = 1

            # 组包并发送
            success, data = self.format_data(PpxCmdType.REQ, ble_msg)
            if success and data is not None:
                if self.send_data(data):
                    # 等待并接收响应
                    response = self.receive_data(timeout=recv_timeout or self.recv_timeout)
                    if response:
                        self.parse_data(response)  # 解析并更新全局变量
                        return True, response
            return False, None
        except Exception as e:
            self._debug_print(f"设置LED显示时发生异常: {e}", is_error=True)
            return False, None

    def read_led_status(self, recv_timeout: Optional[float] = None) -> Tuple[bool, Optional[bytes], Optional[ppx_led_msg_t]]:
        if not self.dll_loaded:
            self._debug_print("DLL未加载，无法读取LED状态", is_error=True)
            return False, None, None
        try:
            # 准备消息结构体
            ble_msg = ppx_ble_msg_t()
            ble_msg.id = PPX_ID_BLE
            ble_msg.cmd = PPX_MSG_READ
            ble_msg.reg_addr = PPX_BLE_LED_MSG_REG
            ble_msg.reg_nums = 1

            success, data = self.format_data(PpxCmdType.REQ, ble_msg)
            if success and data is not None:
                if self.send_data(data):
                    response = self.receive_data(timeout=recv_timeout or self.recv_timeout)
                    if response:
                        parse_success, parsed_msg, cmd_type = self.parse_data(response)
                        return parse_success, response, self.g_ppx_ble_data.led_msg if parse_success else None
            return False, None, None
        except Exception as e:
            self._debug_print(f"读取LED状态时发生异常: {e}", is_error=True)
            return False, None, None

    # ---------------- 打印/辅助 ----------------
    def _bytes_to_hex(self, data: bytes) -> str:
        return ' '.join(f'{b:02X}' for b in data)

    def _print_ble_msg(self, ble_msg: ppx_ble_msg_t):
        self._log("INFO", f"  ID: 0x{ble_msg.id:02X}\n  命令: 0x{ble_msg.cmd:02X} ({'写' if ble_msg.cmd == PPX_MSG_WRITE else '读'})\n  寄存器地址: 0x{ble_msg.reg_addr:02X}\n  寄存器数量: {ble_msg.reg_nums}")

    def _print_led_msg(self, led_msg: ppx_led_msg_t):
        self._log(
            "INFO",
            (
                f"  屏幕状态: {'开' if led_msg.screen_on else '关'}\n"
                f"  亮度级别: {led_msg.brightness}\n"
                f"  数字显示: {led_msg.digital}\n"
                f"  LOGO状态: {led_msg.logo} (0关, 1白, 2红)\n"
                f"  护盾状态: {led_msg.rim_state} (0关, 1白, 2绿)\n"
                f"  ReadyGo状态: {led_msg.rdygo} (0关, 1白, 2红)\n"
                f"  左转向灯: {led_msg.turn_left} (0关, 1白, 2橙)\n"
                f"  右转向灯: {led_msg.turn_right} (0关, 1白, 2橙)\n"
                f"  灯环状态: {led_msg.ring} (0关, 1蓝, 2红)"
            )
        )

    def _print_ble_data(self, ble_data: ppx_ble_data_t):
        self._log(
            "INFO",
            (
                "\n全局BLE数据结构体内容:\n"
                f"设备ID号: {ble_data.id_num}\n"
                f"硬件版本: {ble_data.hw_version}\n"
                f"状态: 0x{ble_data.status:08X}\n"
                f"光敏电阻亮度: {ble_data.ldr_value}\n"
                f"IO状态: 0x{ble_data.io_status:04X}\n"
                f"NFC卡ID: 0x{ble_data.card_id:08X}\n"
                f"数据设置: 0x{ble_data.dat_setting:08X}\n"
                "\nLED显示信息:"
            )
        )
        self._print_led_msg(ble_data.led_msg)

    def _print_ble_data_changes(self, ble_msg: ppx_ble_msg_t):
        if ble_msg.reg_addr == PPX_BLE_LED_MSG_REG:
            self._log("INFO", "\n全局变量中的LED显示信息已更新:")
            self._print_led_msg(self.g_ppx_ble_data.led_msg)

# ==============================================
# 自动化测试框架
# ==============================================
class FileLogger:
    """ 简单文件日志器：写入 raw.log，同时回显到控制台 """
    def __init__(self, log_path: str):
        self.log_path = log_path
        os.makedirs(os.path.dirname(self.log_path), exist_ok=True)

    def __call__(self, level: str, message: str):
        ts = _dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
        line = f"[{ts}] [{level}] {message}"
        # 控制台输出
        print(line)
        # 文件输出
        with open(self.log_path, 'a', encoding='utf-8') as f:
            f.write(line + "\n")


def _coerce_int(val: Any) -> Optional[int]:
    if val is None: return None
    try:
        if isinstance(val, str) and val.strip() == "":
            return None
        return int(float(val))
    except Exception:
        return None


def _coerce_float(val: Any) -> Optional[float]:
    if val is None: return None
    try:
        if isinstance(val, str) and val.strip() == "":
            return None
        return float(val)
    except Exception:
        return None


def _boolish_int(val: Any) -> Optional[int]:
    """ 将 1/0/True/False/"on"/"off"/"是"/"否" 等转换为 1/0 """
    if val is None: return None
    if isinstance(val, (int, float)):
        return 1 if val != 0 else 0
    s = str(val).strip().lower()
    if s in ("1", "true", "t", "yes", "y", "on", "是", "开"):
        return 1
    if s in ("0", "false", "f", "no", "n", "off", "否", "关"):
        return 0
    try:
        return 1 if float(s) != 0 else 0
    except Exception:
        return None


class AutoTester:
    def __init__(self, ble, out_dir="output"):
        self.ble = ble
        self.out_dir = out_dir
        os.makedirs(out_dir, exist_ok=True)
        self.results: List[Dict[str, Any]] = []

    # ---------------- 用例读取 ----------------
    def load_cases(self, file_path=None):
        """
        优先加载 testcases.xlsx，没有则加载 testcases.csv
        """
        if file_path is None or not os.path.exists(file_path):
            if os.path.exists("testcases.xlsx"):
                file_path = "testcases.xlsx"
            elif os.path.exists("testcases.csv"):
                file_path = "testcases.csv"
            else:
                raise FileNotFoundError("未找到 testcases.xlsx 或 testcases.csv，请先生成用例文件")

        ext = os.path.splitext(file_path)[1].lower()
        if ext in (".xlsx", ".xls"):
            cases = pd.read_excel(file_path)
        elif ext == ".csv":
            encoding = detect_encoding(file_path)
            cases = pd.read_csv(file_path, encoding=encoding)
        else:
            raise ValueError(f"不支持的用例文件格式: {ext}")

        if cases is None or cases.empty:
            raise ValueError(f"用例文件为空: {file_path}")

        return cases

    # ---------------- 执行用例 ----------------
    def run_cases(self, cases: List[Dict[str, Any]]):
        for idx, row in enumerate(cases.to_dict(orient="records"), start=1):
            start_ts = time.time()
            case_id = row.get('id', idx)
            comment = str(row.get('comment', '') or '')

            # 读取设置参数
            params = {
                'screen_on': _boolish_int(row.get('screen_on')) if _boolish_int(row.get('screen_on')) is not None else 1,
                'brightness': _coerce_int(row.get('brightness')) or 0,
                'digital': _coerce_int(row.get('digital')) or 0,
                'logo': _coerce_int(row.get('logo')) or 0,
                'rim_state': _coerce_int(row.get('rim_state')) or 0,
                'rdygo': _coerce_int(row.get('rdygo')) or 0,
                'turn_left': _coerce_int(row.get('turn_left')) or 0,
                'turn_right': _coerce_int(row.get('turn_right')) or 0,
                'ring': _coerce_int(row.get('ring')) or 0,
            }

            recv_timeout = _coerce_float(row.get('recv_timeout'))
            delay_after = _coerce_float(row.get('delay_after')) or 0.0

            self.ble._log("INFO", f"==== 执行用例 #{case_id} ====")
            self.ble._log("INFO", f"参数: {params} | 备注: {comment}")

            # 写入命令
            ok, resp = self.ble.set_led_display(**params, recv_timeout=recv_timeout)
            recv_hex = self.ble._bytes_to_hex(resp) if ok and resp else ''

            # 读取响应
            read_ok, read_resp, led_msg = self.ble.read_led_status(recv_timeout=recv_timeout)
            read_hex = self.ble._bytes_to_hex(read_resp) if read_ok and read_resp else ''

            # 判定结果
            verdict = "PASS" if ok and read_ok else "FAIL"
            elapsed = time.time() - start_ts
            self.ble._log("INFO", f"结果: {verdict} | 用时: {elapsed:.3f}s")

            # 展开 led_msg 为字典
            checks = {}
            if led_msg:
                try:
                    checks = {
                        "screen_on": led_msg.screen_on,
                        "brightness": led_msg.brightness,
                        "digital": led_msg.digital,
                        "logo": led_msg.logo,
                        "rim_state": led_msg.rim_state,
                        "rdygo": led_msg.rdygo,
                        "turn_left": led_msg.turn_left,
                        "turn_right": led_msg.turn_right,
                        "ring": led_msg.ring,
                    }
                except Exception as e:
                    checks = {"error": f"无法解析 led_msg: {e}"}

            # 回填结果
            self.results.append({
                'case_id': case_id,
                'comment': comment,
                'screen_on': params['screen_on'],
                'brightness': params['brightness'],
                'digital': params['digital'],
                'logo': params['logo'],
                'rim_state': params['rim_state'],
                'rdygo': params['rdygo'],
                'turn_left': params['turn_left'],
                'turn_right': params['turn_right'],
                'ring': params['ring'],
                'recv_timeout': recv_timeout,
                'delay_after': delay_after,
                'send_ok': "PASS" if ok else "FAIL",
                'read_ok': "PASS" if read_ok else "FAIL",
                'recv_hex': recv_hex,   # 写入响应
                'read_hex': read_hex,   # 读取响应
                'verdict': verdict,
                'checks': checks,       # ✅ 现在是 dict，而不是 <object ...>
                'elapsed_s': round(elapsed, 3),
                'timestamp': time.strftime("%Y-%m-%d %H:%M:%S"),
            })

            if delay_after > 0:
                time.sleep(delay_after)




    def _assert_led_expectations(self, row: Dict[str, Any], led_msg: Optional[ppx_led_msg_t]) -> Tuple[str, str]:
        if led_msg is None:
            return 'FAIL', '未能读取到 LED 状态'

        fields = [
            ('screen_on', _boolish_int(row.get('expect_screen_on'))),
            ('brightness', _coerce_int(row.get('expect_brightness'))),
            ('digital', _coerce_int(row.get('expect_digital'))),
            ('logo', _coerce_int(row.get('expect_logo'))),
            ('rim_state', _coerce_int(row.get('expect_rim_state'))),
            ('rdygo', _coerce_int(row.get('expect_rdygo'))),
            ('turn_left', _coerce_int(row.get('expect_turn_left'))),
            ('turn_right', _coerce_int(row.get('expect_turn_right'))),
            ('ring', _coerce_int(row.get('expect_ring'))),
        ]

        fail_msgs = []
        for name, expected in fields:
            if expected is None:
                continue  # 留空表示不检查
            actual = getattr(led_msg, name)
            if actual != expected:
                fail_msgs.append(f"{name}: 实际={actual}, 期望={expected}")

        if fail_msgs:
            return "FAIL", "; ".join(fail_msgs)
        else:
            return "PASS", "全部匹配"


    # ---------------- 压力循环 ----------------
    def loop_case(self, case: Dict[str, Any], loop_count: int = 100, delay: float = 0.5):
        params = {
            'screen_on': _boolish_int(case.get('screen_on')) if _boolish_int(case.get('screen_on')) is not None else 1,
            'brightness': _coerce_int(case.get('brightness')) or 0,
            'digital': _coerce_int(case.get('digital')) or 0,
            'logo': _coerce_int(case.get('logo')) or 0,
            'rim_state': _coerce_int(case.get('rim_state')) or 0,
            'rdygo': _coerce_int(case.get('rdygo')) or 0,
            'turn_left': _coerce_int(case.get('turn_left')) or 0,
            'turn_right': _coerce_int(case.get('turn_right')) or 0,
            'ring': _coerce_int(case.get('ring')) or 0,
        }
        self.ble._log("INFO", f"开始压力循环: 次数={loop_count}, 间隔={delay}s, 参数={params}")
        for i in range(1, loop_count + 1):
            ok, _ = self.ble.set_led_display(**params)
            self.ble._log("INFO", f"[循环 {i}/{loop_count}] 发送 {'OK' if ok else 'FAIL'}")
            time.sleep(delay)

    # ---------------- 报告输出 ----------------
    def save_results_csv(self, path: str):
        import csv
        keys = [
            'case_id','comment','screen_on','brightness','digital','logo','rim_state','rdygo','turn_left','turn_right','ring',
            'recv_timeout','delay_after','send_ok','recv_ok','read_ok','recv_hex','read_hex','verdict','checks','elapsed_s','timestamp'
        ]
        with open(path, 'w', newline='', encoding='utf-8') as f:
            w = csv.DictWriter(f, fieldnames=keys)
            w.writeheader()
            for r in self.results:
                w.writerow({k: r.get(k, '') for k in keys})

    def save_report_html(self, path: str, title: str = "BLE 自动化测试报告"):
        total = len(self.results)
        passed = sum(1 for r in self.results if r.get('verdict') == 'PASS')
        failed = total - passed
        pass_rate = (passed / total * 100) if total else 0.0

        # 简单表格
        rows_html = []
        def esc(x: Any) -> str:
            s = str(x if x is not None else '')
            return (s.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;'))
        for r in self.results:
            rows_html.append(
                "<tr>" +
                "".join([
                    f"<td>{esc(r.get('case_id'))}</td>",
                    f"<td>{esc(r.get('comment'))}</td>",
                    f"<td>{esc(r.get('screen_on'))}</td>",
                    f"<td>{esc(r.get('brightness'))}</td>",
                    f"<td>{esc(r.get('digital'))}</td>",
                    f"<td>{esc(r.get('logo'))}</td>",
                    f"<td>{esc(r.get('rim_state'))}</td>",
                    f"<td>{esc(r.get('rdygo'))}</td>",
                    f"<td>{esc(r.get('turn_left'))}</td>",
                    f"<td>{esc(r.get('turn_right'))}</td>",
                    f"<td>{esc(r.get('ring'))}</td>",
                    f"<td>{esc(r.get('send_ok'))}</td>",
                    f"<td>{esc(r.get('read_ok'))}</td>",
                    f"<td style='font-family:monospace'>{esc(r.get('recv_hex'))}</td>",
                    f"<td style='font-family:monospace'>{esc(r.get('read_hex'))}</td>",
                    f"<td><b style='color:{'green' if r.get('verdict')=='PASS' else 'red'}'>{esc(r.get('verdict'))}</b></td>",
                    f"<td>{esc(r.get('checks'))}</td>",
                    f"<td>{esc(r.get('elapsed_s'))}</td>",
                    f"<td>{esc(r.get('timestamp'))}</td>",
                ]) + "</tr>"
            )

        html = f"""
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8" />
<title>{esc(title)}</title>
<style>
body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, 'Noto Sans', 'PingFang SC', 'Microsoft YaHei', sans-serif; margin: 24px; }}
h1 {{ margin-bottom: 8px; }}
.summary {{ margin: 8px 0 16px; }}
.kpi span {{ display: inline-block; margin-right: 16px; padding: 6px 10px; border-radius: 8px; background: #f2f2f2; }}
table {{ border-collapse: collapse; width: 100%; font-size: 14px; }}
th, td {{ border: 1px solid #ddd; padding: 6px 8px; text-align: left; vertical-align: top; }}
th {{ background: #fafafa; position: sticky; top: 0; }}
tr:nth-child(even) {{ background: #fcfcfc; }}
code {{ background: #f6f8fa; padding: 2px 4px; border-radius: 4px; }}
</style>
</head>
<body>
<h1>{esc(title)}</h1>
<div class="summary">
  <div>执行时间：{esc(_dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))}</div>
  <div class="kpi">
    <span>总用例：<b>{total}</b></span>
    <span>通过：<b style="color:green">{passed}</b></span>
    <span>失败：<b style="color:red">{failed}</b></span>
    <span>通过率：<b>{pass_rate:.2f}%</b></span>
  </div>
</div>
<table>
  <thead>
    <tr>
      <th>ID</th><th>备注</th>
      <th>screen_on</th><th>brightness</th><th>digital</th><th>logo</th><th>rim_state</th><th>rdygo</th><th>turn_left</th><th>turn_right</th><th>ring</th>
      <th>发送OK</th><th>读取OK</th><th>写入响应</th><th>读取响应</th>
      <th>结论</th><th>校验详情</th><th>耗时(s)</th><th>时间戳</th>
    </tr>
  </thead>
  <tbody>
    {''.join(rows_html)}
  </tbody>
</table>
</body>
</html>
"""
        with open(path, 'w', encoding='utf-8') as f:
            f.write(html)

# ==============================================
# 实用函数：示例用例生成
# ==============================================
SAMPLE_CASES = [
    {
        'id': 1,
        'comment': '基本点亮 + 白LOGO + 左右转向橙',
        'screen_on': 1, 'brightness': 7, 'digital': 88,
        'logo': 1, 'rim_state': 1, 'rdygo': 1,
        'turn_left': 2, 'turn_right': 2, 'ring': 2,
        'expect_screen_on': 1, 'expect_brightness': 7, 'expect_digital': 88,
        'expect_logo': 1, 'expect_rim_state': 1, 'expect_rdygo': 1,
        'expect_turn_left': 2, 'expect_turn_right': 2, 'expect_ring': 2,
        'recv_timeout': 1.0, 'delay_after': 0.2,
    },
    {
        'id': 2,
        'comment': '关闭显示，校验全部为0（除 brightness）',
        'screen_on': 0, 'brightness': 3, 'digital': 0,
        'logo': 0, 'rim_state': 0, 'rdygo': 0,
        'turn_left': 0, 'turn_right': 0, 'ring': 0,
        'expect_screen_on': 0, 'expect_brightness': 3, 'expect_digital': 0,
        'expect_logo': 0, 'expect_rim_state': 0, 'expect_rdygo': 0,
        'expect_turn_left': 0, 'expect_turn_right': 0, 'expect_ring': 0,
        'recv_timeout': 1.0, 'delay_after': 0.2,
    },
]


def make_sample_cases(path: str):
    os.makedirs(os.path.dirname(path) or '.', exist_ok=True)
    if _HAS_PANDAS and path.lower().endswith(('.xlsx', '.xls')):
        df = pd.DataFrame(SAMPLE_CASES)
        df.to_excel(path, index=False)
    else:
        # 生成 CSV
        import csv
        keys = sorted({k for d in SAMPLE_CASES for k in d.keys()})
        with open(path, 'w', newline='', encoding='utf-8') as f:
            w = csv.DictWriter(f, fieldnames=keys)
            w.writeheader()
            for r in SAMPLE_CASES:
                w.writerow(r)
                
def make_combo_cases():
    fields = {
        "screen_on": [0, 1],
        "brightness": [0,1,2,3,4,5,6,7],
        "digital": [0,50,100],
        "logo": [0,1,2],
        "rim_state": [0,1,2],
        "rdygo": [0,1,2],
        "turn_left": [0,1,2],
        "turn_right": [0,1,2],
        "ring": [0,1,2],
    }
    keys = list(fields.keys())
    cases = []
    idx = 1
    for values in itertools.product(*[fields[k] for k in keys]):
        row = dict(zip(keys, values))
        # ⭐ 过滤条件：只保留 screen_on=1 且 brightness >= 5
        if not (row["screen_on"] == 1 and row["brightness"] >= 5):
            continue
        row["id"] = idx
        cases.append(row)
        idx += 1
    return cases
             

# ==============================================
# 主程序
# ==============================================

def main():

    print("\nBLE 协议自动化测试工具")
    print("=" * 60)
    print(f"Python版本: {platform.python_version()}")
    print(f"系统架构: {platform.architecture()[0]}")
    print(f"操作系统: {platform.system()} {platform.release()}")

    # -------------------- 解析参数 --------------------
    parser = argparse.ArgumentParser(description="BLE 自动化测试工具（Excel/CSV 用例 + HTML 报告）")
    parser.add_argument('--dll', default=BLE_PROTOCOL_DLL, help='ppx_ble.dll 路径')
    parser.add_argument('--port', default=SERIAL_PORT, help='串口号，如 COM44')
    parser.add_argument('--baud', type=int, default=BAUDRATE, help='波特率，如 460800')
    parser.add_argument('--cases', default='testcases.xlsx', help='用例文件（xlsx/xls/csv），默认 testcases.xlsx')
    parser.add_argument('--out', default=None, help='报告输出目录（默认 ./reports/时间戳/）')
    parser.add_argument('--loop-count', type=int, default=1, help='对第1条用例进行压力循环次数（0 表示不进行）')
    parser.add_argument('--loop-delay', type=float, default=0.5, help='压力循环间隔秒')
    parser.add_argument('--make-sample', action='store_true', help='生成示例用例（不执行测试）')
    parser.add_argument('--combo', action='store_true', help='自动生成排列组合用例')
    args = parser.parse_args()

    # 如果用户直接点 Run（没有传参数），默认开启 combo 模式
    if len(sys.argv) == 1:
        args.combo = False

    # -------------------- 输出目录 --------------------
    ts_dir = _dt.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    out_dir = args.out or os.path.join('reports', ts_dir)
    os.makedirs(out_dir, exist_ok=True)

    # -------------------- 日志器 --------------------
    logger = FileLogger(os.path.join(out_dir, 'raw.log'))

    # -------------------- 生成示例用例 --------------------
    if args.make_sample:
        out_path = args.cases
        print(f"生成示例用例: {out_path}")
        make_sample_cases(out_path)
        print("已生成。")
        return

    # -------------------- 初始化 BLE 协议 --------------------
    logger("INFO", "初始化 BLE 协议通信...")
    ble = BLEProtocol(args.dll, args.port, args.baud)
    ble.set_logger(logger)

    if not ble.dll_loaded or not ble.serial_connected:
        logger("ERROR", "初始化失败，请检查 DLL 路径/串口参数")
        ble.close()
        sys.exit(1)

    # -------------------- 初始化测试器 --------------------
    tester = AutoTester(ble, out_dir)

    # -------------------- 读取用例 --------------------
    try:
        if args.combo:
            cases = make_combo_cases()  # 生成排列组合用例
            if isinstance(cases, list):
                cases = pd.DataFrame(cases)
            logger("INFO", f"生成排列组合用例 {len(cases)} 条")
        else:
            cases = tester.load_cases(args.cases)
            if cases is None or cases.empty:
                logger("ERROR", f"用例文件为空: {args.cases}")
                ble.close()
                sys.exit(2)
            logger("INFO", f"加载用例 {len(cases)} 条 来自: {args.cases}")
    except Exception as e:
        logger("ERROR", f"读取用例失败: {e}")
        ble.close()
        sys.exit(3)

    # -------------------- 执行用例 --------------------
    try:
        tester.run_cases(cases)

        # 压力循环（可选）
        if args.loop_count > 0 and len(cases) > 0:
            tester.loop_case(cases.iloc[0], loop_count=args.loop_count, delay=args.loop_delay)

        # -------------------- 保存结果 --------------------
        csv_path = os.path.join(out_dir, 'results.csv')
        html_path = os.path.join(out_dir, 'report.html')
        tester.save_results_csv(csv_path)
        tester.save_report_html(html_path)
        logger("INFO", f"已保存结果: {csv_path}")
        logger("INFO", f"已生成报告: {html_path}")

    finally:
        ble.close()
        logger("INFO", "程序结束")


if __name__ == "__main__":
    main()

