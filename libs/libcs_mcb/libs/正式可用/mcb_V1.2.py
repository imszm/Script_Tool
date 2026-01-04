# -*- coding: utf-8 -*-

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
import csv


# 尝试导入 pandas（用于 Excel/CSV 读写），失败则退化到 CSV 解析
try:
    import pandas as pd  # type: ignore
    _HAS_PANDAS = True
except Exception:
    _HAS_PANDAS = False

# ==============================================
# 配置区域 - 根据实际情况修改这些参数
# ==============================================
REGION_PROTOCOL_DLL = r"C:\Test_ziliao\libs\libcs_mcb\libs\ppx_region"  # 控制器协议DLL路径
SERIAL_PORT = "COM4"                              # 串口端口
BAUDRATE = 115200                                   # 串口波特率
DEBUG_MODE = True                                   # 调试模式，打印更多信息
DEFAULT_RECV_TIMEOUT = 1.0                          # 串口接收默认超时（秒）

# ==============================================
# 协议常量定义（源自 ppx_region.h）
# ==============================================
# 设备ID（需与硬件协商，默认暂设为0x61，与灯板区分）
PPX_ID_REGION = 0x61

# 命令类型（读/写）
PPX_MSG_READ = 0x01
PPX_MSG_WRITE = 0x03

# 命令类别（请求/响应）
class PpxCmdType:
    REQ = 0  # 请求命令
    RSP = 1  # 响应命令

# 解析状态（与DLL保持一致）
PPX_PARSE_SUCCESS = 1
PPX_PARSE_FAILURE = 0

# 寄存器地址（源自 ppx_region_reg_t 枚举）
class PpxRegionReg:
    PPX_ID_NUM_REG = 0x00               # 设备ID号
    PPX_MODEL_REG = 0x01                # 型号
    PPX_SERIAL_NUM_REG = 0x02           # 序列号
    PPX_HW_VERSION_REG = 0x03           # 硬件版本
    PPX_SW_VESRION_REG = 0x04           # 软件版本
    PPX_RIM_STATE_REG = 0x05            # 护盾状态（下坡/座椅/碰撞等）
    PPX_MCU_ERRCODE_REG = 0x06          # MCU错误码
    PPX_CTRL_MODEL_REG = 0x07           # 控制模式
    PPX_SPEED_REF_REG = 0x08            # 速度参考值
    PPX_MOTOR_SPEED_REG = 0x09          # 电机速度
    PPX_BUS_VOLTAGE_REG = 0x0A          # 总线电压（0.1V）
    PPX_BUS_CURRENT_REG = 0x0B          # 总线电流（0.1A）
    PPX_PHASE_CUR_A_REG = 0x0C          # A相电流（0.1A）
    PPX_PHASE_CUR_B_REG = 0x0D          # B相电流（0.1A）
    PPX_PHASE_CUR_C_REG = 0x0E          # C相电流（0.1A）
    PPX_HALL_STATE_REG = 0x0F           # 霍尔状态
    PPX_PI_VQ_REG = 0x10                # PI_VQ参数
    PPX_PI_IQ_REG = 0x11                # PI_IQ参数
    PPX_BRAKE_STATE_REG = 0x12          # 刹车状态
    PPX_IMU_PITCH_REG = 0x13            # IMU俯仰角（0.1deg）
    PPX_IMU_ROLL_REG = 0x14             # IMU横滚角（0.1deg）
    PPX_BOARD_TEMP_REG = 0x15           # 板温
    PPX_BRAKE_MILEAGE_REG = 0x16        # 刹车里程（dm）
    PPX_MOTOR_ANGLE_REG = 0x17          # 电机角度
    PPX_SINGLE_MILEAGE_REG = 0x18       # 单次里程（m）
    PPX_ANGULAR_SPEED_REG = 0x19        # 角速度（0.1deg）
    PPX_RT_SETTING_REG = 0x1A           # 实时设置（LED/清错误码）
    PPX_RUN_MODE_REG = 0x1B             # 运行模式
    PPX_GEARS_REG = 0x1C                # 档位
    PPX_TARGET_SPEED_REG = 0x1D         # 目标速度（rpm）
    PPX_RATED_VOLT_REG = 0x1E           # 额定电压（0.1V）
    PPX_RATED_CUR_REG = 0x1F            # 额定电流（0.1A）
    PPX_MAX_VOLTAGE_REG = 0x20          # 最大电压（0.1V）
    PPX_MIN_VOLTAGE_REG = 0x21          # 最小电压（0.1V）
    PPX_ACCERATION_REG = 0x22           # 加速度
    PPX_DAT_SETTING_REG = 0x23          # 数据设置（IMU/校准/烧录）
    PPX_RVSD_DATA_REG = 0x24            # 保留数据

# 运行模式（源自 ppx_run_mode_t 枚举）
class PpxRunMode:
    PPX_MODE_IDLE = 0    # 空闲模式
    PPX_MODE_SET = 1     # 设置模式
    PPX_MODE_RUN = 2     # 运行模式
    PPX_MODE_LOCK = 3    # 锁定模式
    PPX_MODE_AID = 4     # 辅助模式
    PPX_MODE_BRAKE = 5   # 刹车模式
    PPX_MODE_IAP = 6     # 升级模式
    PPX_MODE_TST = 7     # 测试模式

# 实时设置位（源自 ppx_rt_setting_t 枚举）
class PpxRtSetting:
    PPX_BRAKE_LED_ON = (1 << 0)  # 刹车灯开启
    PPX_TAIL_LED_ON = (1 << 1)   # 尾灯开启
    PPX_RIGHT_LED_ON = (1 << 2)  # 右转向灯开启
    PPX_LEFT_LED_ON = (1 << 3)   # 左转向灯开启
    PPX_CLR_ERRCODE = (1 << 15)  # 清除错误码

# 数据设置位（源自 ppx_data_setting_t 枚举）
class PpxDatSetting:
    # 请求类型
    PPX_CHR_CHECK = (1 << 0)     # 特征值检查
    PPX_IMU_OPEN = (1 << 1)      # IMU开启
    PPX_IMU_CALI = (1 << 2)      # IMU校准
    PPX_IAP_MODE = (1 << 3)      # IAP升级模式
    PPX_SN_WRITE = (1 << 4)      # 写入序列号
    PPX_TST_MOTO = (1 << 5)      # 电机测试
    PPX_ACC_CALI = (1 << 6)      # 加速度校准
    # 响应状态
    PPX_CHR_CHECK_SUCC = (1 << 16)  # 特征值检查成功
    PPX_IMU_OPEN_SUCC = (1 << 17)   # IMU开启成功
    PPX_IMU_CALI_SUCC = (1 << 18)   # IMU校准成功
    PPX_IAP_MODE_FALSE = (1 << 19)  # IAP模式关闭
    PPX_ACC_CALI_SIDE = (1 << 20)   # 加速度校准侧
    PPX_ACC_CALI_SUCC = (1 << 21)   # 加速度校准成功

# 护盾状态位（源自 ppx_rim_state_t 枚举）
class PpxRimState:
    PPX_RIM_DOWN = 0x04    # 下坡
    PPX_RIM_SEAT = 0x08    # 座椅
    PPX_RIM_8DEG = 0x10    # 8度
    PPX_RIM_DUMP = 0x20    # 碰撞
    PPX_RIM_BUMP = 0x40    # 颠簸
    PPX_RIM_TURN = 0x80    # 转弯

# ==============================================
# 结构体定义 (与 ppx_region.h 完全一致)
# ==============================================
class ppx_region_excp_t(Structure):
    """异常响应结构体（源自 ppx_region_excp_t）"""
    _fields_ = [
        ("parse_status", c_uint8),  # 解析状态
        ("cmd_status", c_uint8),    # 命令状态
        ("data_status", c_uint8)    # 数据状态
    ]

class ppx_region_msg_t(Structure):
    """控制器消息结构体（源自 ppx_region_msg_t）"""
    _fields_ = [
        ("id", c_uint8),                # 设备ID
        ("cmd", c_uint8),               # 写/读命令
        ("msg_type", c_uint8),          # 主设备使用的消息类型
        ("reg_addr", c_uint8),          # 寄存器地址
        ("reg_nums", c_uint8),          # 寄存器数量
        ("reg_excp", ppx_region_excp_t) # 异常响应
    ]

class ppx_region_data_t(Structure):
    """控制器数据结构体（源自 ppx_region_data_t，注意内存对齐）"""
    _pack_ = 1  # 强制1字节对齐，与C语言#pragma pack(1)保持一致
    _fields_ = [
        # 读数据区域
        ("id_num", c_uint8),                     # 设备ID号
        ("model", c_uint8 * 8),                  # 型号（PPX_MODEL_SIZE默认8）
        ("serial_num", c_uint8 * 26),            # 序列号（PPX_SN_SIZE默认26）
        ("hw_version", c_uint16),                # 硬件版本
        ("sw_version", c_uint8 * 20),            # 软件版本（PPX_SW_VER_SIZE默认20）
        ("rim_state", c_uint8),                  # 护盾状态（PpxRimState）
        ("mcu_errcode", c_uint32),               # MCU错误码
        ("ctrl_model", c_uint8),                 # 控制模式
        ("speed_ref", c_int16),                  # 速度参考值
        ("motor_speed", c_int16),                # 电机速度（rpm）
        ("bus_voltage", c_uint16),               # 总线电压（0.1V）
        ("bus_current", c_uint16),               # 总线电流（0.1A）
        ("phase_current_a", c_int16),            # A相电流（0.1A）
        ("phase_current_b", c_int16),            # B相电流（0.1A）
        ("phase_current_c", c_int16),            # C相电流（0.1A）
        ("hall_state", c_uint8),                 # 霍尔状态
        ("pi_vq", c_int16),                      # PI_VQ参数
        ("pi_iq", c_int16),                      # PI_IQ参数
        ("brake_state", c_uint8),                # 刹车状态
        ("imu_pitch", c_int16),                  # IMU俯仰角（0.1deg）
        ("imu_roll", c_int16),                   # IMU横滚角（0.1deg）
        ("imu_acc", c_uint8),                    # IMU加速度（0.01g）
        ("board_temp", c_uint8),                 # 板温（℃）
        ("brake_mileage", c_uint8),              # 刹车里程（dm）
        ("motor_angle", c_int32),                # 电机角度
        ("single_mileage", c_uint32),            # 单次里程（m）
        ("angular_speed", c_int16),              # 角速度（0.1deg）
        # 写数据区域
        ("rt_setting", c_uint16),                # 实时设置（PpxRtSetting）
        ("run_mode", c_uint8),                   # 运行模式（PpxRunMode）
        ("gear", c_uint8),                       # 档位
        ("target_speed", c_int16),               # 目标速度（rpm）
        ("rated_voltage", c_uint16),             # 额定电压（0.1V）
        ("rated_current", c_uint16),             # 额定电流（0.1A）
        ("max_voltage", c_uint16),               # 最大电压（0.1V）
        ("min_voltage", c_uint16),               # 最小电压（0.1V）
        ("acceration", c_uint32),                # 加速度
        ("dat_setting", c_uint32),               # 数据设置（PpxDatSetting）
        ("rsvd_data", c_uint32)                  # 保留数据
    ]

# ==============================================
# 控制器协议通信类（复用灯板框架，适配Region协议）
# ==============================================
class RegionProtocol:
    def __init__(self, dll_path: str, serial_port: str, baudrate: int, recv_timeout: float = DEFAULT_RECV_TIMEOUT):
        self.dll_loaded = False
        self.serial_connected = False
        self.recv_timeout = recv_timeout

        # 日志回调（由外部 AutoTester 注入）
        self._logger = None  # type: Optional[callable]

        # 加载DLL
        self._load_dll(dll_path)

        # 初始化串口
        self._init_serial(serial_port, baudrate) 

        # 检查全局变量（DLL中的g_ppx_region_data）
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
            self.region_lib = cdll.LoadLibrary(dll_path)
            self.dll_loaded = True
            self._debug_print(f"成功加载控制器DLL: {dll_path}")

            # 配置DLL函数原型（与ppx_region.h中的函数定义一致）
            self.region_lib.ppx_com_region_parse.argtypes = [
                POINTER(c_uint8),          # pdata: 输入数据
                c_uint8,                   # data_len: 数据长度
                POINTER(ppx_region_msg_t)  # region_msg: 输出消息结构体
            ]
            self.region_lib.ppx_com_region_parse.restype = c_int  # 返回解析状态（成功/失败）

            self.region_lib.ppx_com_region_format.argtypes = [
                c_int,                     # cmd_type: 命令类型（REQ/RSP）
                POINTER(ppx_region_msg_t), # region_msg: 输入消息结构体
                c_void_p                   # buffer: 输出组包缓存
            ]
            self.region_lib.ppx_com_region_format.restype = c_uint16  # 返回组包长度

        except Exception as e:
            self._debug_print(f"加载控制器DLL失败: {e}", is_error=True)
            self.dll_loaded = False

    def _init_serial(self, port: str, baudrate: int):
        try:
            self.serial_port = serial.Serial(
                port=port,
                baudrate=baudrate,
                timeout=self.recv_timeout,
                parity=serial.PARITY_NONE,    # 默认无校验（需与硬件一致）
                stopbits=serial.STOPBITS_ONE  # 默认1位停止位（需与硬件一致）
            )
            self.serial_connected = True
            self._debug_print(f"串口已连接: {port}, 波特率: {baudrate}")
        except Exception as e:
            self._debug_print(f"串口连接失败: {e}", is_error=True)
            self.serial_connected = False

    def _check_global_vars(self):
        """检查DLL中的全局变量g_ppx_region_data"""
        try:
            self.g_ppx_region_data = ppx_region_data_t.in_dll(self.region_lib, "g_ppx_region_data")
            self._debug_print("成功获取全局变量 g_ppx_region_data")

            if DEBUG_MODE:
                self._print_region_data(self.g_ppx_region_data)
        except Exception as e:
            self._debug_print(f"获取全局变量失败: {e}", is_error=True)

    # ---------------- 基本通信 ----------------
    def close(self):
        """关闭串口连接"""
        if hasattr(self, 'serial_port') and self.serial_port and self.serial_port.is_open:
            self.serial_port.close()
            self.serial_connected = False
            self._debug_print("串口已关闭")

    def send_data(self, data: bytes) -> bool:
        """发送字节数据到串口"""
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
        """从串口接收数据"""
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

    def parse_data(self, data: bytes) -> Tuple[bool, Optional[ppx_region_msg_t], Optional[int], Optional[int]]:
        """用DLL解析数据包"""
        if not data:
            self._debug_print("无数据可解析", is_error=True)
            return False, None, None, None
        if not self.dll_loaded:
            self._debug_print("DLL未加载，无法解析数据", is_error=True)
            return False, None, None, None
        try:
            region_msg = ppx_region_msg_t()
            region_msg.id = PPX_ID_REGION  # 预填充设备ID

            data_len = len(data)
            data_array = (c_uint8 * data_len)(*data)

            # 调用DLL解析函数
            parse_result = self.region_lib.ppx_com_region_parse(
                cast(data_array, POINTER(c_uint8)),
                c_uint8(data_len),
                byref(region_msg)
            )

            if parse_result == PPX_PARSE_SUCCESS:
                self._debug_print("数据解析成功:")
                self._print_region_msg(region_msg)
                # 更新全局变量信息
                self._print_region_data_changes(region_msg)
                # 提取命令类型（首字节低4位）
                cmd_type = data[0] & 0x0F if len(data) > 0 else None
                return True, region_msg, cmd_type, int(parse_result)
            else:
                self._debug_print(f"数据解析失败，返回码: {parse_result}", is_error=True)
                return False, None, None, int(parse_result)
        except Exception as e:
            self._debug_print(f"解析数据时发生异常: {e}", is_error=True)
            return False, None, None, None

    def format_data(self, cmd_type: int, region_msg: ppx_region_msg_t) -> Tuple[bool, Optional[bytes]]:
        """用DLL组包数据"""
        if not self.dll_loaded:
            self._debug_print("DLL未加载，无法组包数据", is_error=True)
            return False, None
        try:
            buffer_size = 256  # 足够大的缓存
            buffer = create_string_buffer(buffer_size)

            length = self.region_lib.ppx_com_region_format(
                cmd_type,
                byref(region_msg),
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

    # ---------------- 业务操作：控制器设置/读取 ----------------
    def set_run_mode(self, run_mode: int, 
                    gear: int = 0, 
                    target_speed: int = 0,
                    recv_timeout: Optional[float] = None) -> Tuple[bool, Optional[bytes], Optional[int]]:
        """设置运行模式、档位和目标速度"""
        if not self.dll_loaded:
            self._debug_print("DLL未加载，无法设置运行模式", is_error=True)
            return False, None, None
        try:
            # 验证运行模式有效性
            if run_mode < 0 or run_mode > 7:
                self._debug_print(f"无效的运行模式: {run_mode}（必须0-7）", is_error=True)
                return False, None, None

            # 修改全局变量中的运行参数
            self.g_ppx_region_data.run_mode = int(run_mode)
            self.g_ppx_region_data.gear = int(gear)
            self.g_ppx_region_data.target_speed = int(target_speed)

            self._debug_print("修改后的运行参数:")
            self._log("INFO", 
                     f"  运行模式: {run_mode} ({self._get_run_mode_str(run_mode)})\n"
                     f"  档位: {gear}\n"
                     f"  目标速度: {target_speed} rpm")

            # 准备消息结构体
            region_msg = ppx_region_msg_t()
            region_msg.id = PPX_ID_REGION
            region_msg.cmd = PPX_MSG_WRITE
            region_msg.reg_addr = PpxRegionReg.PPX_RUN_MODE_REG
            region_msg.reg_nums = 3  # 连续写入3个寄存器（模式+档位+速度）

            # 组包并发送
            success, data = self.format_data(PpxCmdType.REQ, region_msg)
            if success and data is not None:
                if self.send_data(data):
                    # 等待响应
                    response = self.receive_data(timeout=recv_timeout or self.recv_timeout)
                    if response:
                        parse_success, parsed_msg, cmd_type, parse_result = self.parse_data(response)
                        return True, response, parse_result
            return False, None, None
        except Exception as e:
            self._debug_print(f"设置运行模式时发生异常: {e}", is_error=True)
            return False, None, None

    def set_rt_setting(self, 
                      brake_led: bool = False, 
                      tail_led: bool = False, 
                      right_led: bool = False, 
                      left_led: bool = False,
                      clear_err: bool = False,
                      recv_timeout: Optional[float] = None) -> Tuple[bool, Optional[bytes], Optional[int]]:
        """设置实时参数（灯光控制、清除错误码）"""
        if not self.dll_loaded:
            self._debug_print("DLL未加载，无法设置实时参数", is_error=True)
            return False, None, None
        try:
            # 计算rt_setting位掩码
            rt_setting = 0
            if brake_led:
                rt_setting |= PpxRtSetting.PPX_BRAKE_LED_ON
            if tail_led:
                rt_setting |= PpxRtSetting.PPX_TAIL_LED_ON
            if right_led:
                rt_setting |= PpxRtSetting.PPX_RIGHT_LED_ON
            if left_led:
                rt_setting |= PpxRtSetting.PPX_LEFT_LED_ON
            if clear_err:
                rt_setting |= PpxRtSetting.PPX_CLR_ERRCODE

            # 修改全局变量
            self.g_ppx_region_data.rt_setting = rt_setting

            self._debug_print("修改后的实时设置:")
            self._log("INFO", 
                     f"  刹车灯: {'开' if brake_led else '关'}\n"
                     f"  尾灯: {'开' if tail_led else '关'}\n"
                     f"  右转向灯: {'开' if right_led else '关'}\n"
                     f"  左转向灯: {'开' if left_led else '关'}\n"
                     f"  清除错误码: {'是' if clear_err else '否'}")

            # 准备消息结构体
            region_msg = ppx_region_msg_t()
            region_msg.id = PPX_ID_REGION
            region_msg.cmd = PPX_MSG_WRITE
            region_msg.reg_addr = PpxRegionReg.PPX_RT_SETTING_REG
            region_msg.reg_nums = 1

            # 组包并发送
            success, data = self.format_data(PpxCmdType.REQ, region_msg)
            if success and data is not None:
                if self.send_data(data):
                    response = self.receive_data(timeout=recv_timeout or self.recv_timeout)
                    if response:
                        parse_success, parsed_msg, cmd_type, parse_result = self.parse_data(response)
                        return True, response, parse_result
            return False, None, None
        except Exception as e:
            self._debug_print(f"设置实时参数时发生异常: {e}", is_error=True)
            return False, None, None

    def read_vehicle_state(self, recv_timeout: Optional[float] = None) -> Tuple[bool, Optional[bytes], Optional[ppx_region_data_t], Optional[int]]:
        """读取车辆状态（电压、电流、速度等）"""
        if not self.dll_loaded:
            self._debug_print("DLL未加载，无法读取车辆状态", is_error=True)
            return False, None, None, None
        try:
            # 准备消息结构体（读取多个关键寄存器）
            region_msg = ppx_region_msg_t()
            region_msg.id = PPX_ID_REGION
            region_msg.cmd = PPX_MSG_READ
            region_msg.reg_addr = PpxRegionReg.PPX_BUS_VOLTAGE_REG
            region_msg.reg_nums = 10  # 连续读取10个寄存器（电压到角速度）

            success, data = self.format_data(PpxCmdType.REQ, region_msg)
            if success and data is not None:
                if self.send_data(data):
                    response = self.receive_data(timeout=recv_timeout or self.recv_timeout)
                    if response:
                        parse_success, parsed_msg, cmd_type, parse_result = self.parse_data(response)
                        return parse_success, response, self.g_ppx_region_data if parse_success else None, parse_result
            return False, None, None, None
        except Exception as e:
            self._debug_print(f"读取车辆状态时发生异常: {e}", is_error=True)
            return False, None, None, None

    def start_imu_calibration(self, recv_timeout: Optional[float] = None) -> Tuple[bool, Optional[bytes], Optional[int]]:
        """启动IMU校准"""
        if not self.dll_loaded:
            self._debug_print("DLL未加载，无法启动IMU校准", is_error=True)
            return False, None, None
        try:
            # 设置数据位：IMU校准
            self.g_ppx_region_data.dat_setting = PpxDatSetting.PPX_IMU_CALI

            self._debug_print("启动IMU校准...")

            # 准备消息结构体
            region_msg = ppx_region_msg_t()
            region_msg.id = PPX_ID_REGION
            region_msg.cmd = PPX_MSG_WRITE
            region_msg.reg_addr = PpxRegionReg.PPX_DAT_SETTING_REG
            region_msg.reg_nums = 1

            # 组包并发送
            success, data = self.format_data(PpxCmdType.REQ, region_msg)
            if success and data is not None:
                if self.send_data(data):
                    response = self.receive_data(timeout=recv_timeout or 5.0)  # 校准可能需要更长时间
                    if response:
                        parse_success, parsed_msg, cmd_type, parse_result = self.parse_data(response)
                        # 检查校准结果
                        if parse_success and (self.g_ppx_region_data.dat_setting & PpxDatSetting.PPX_IMU_CALI_SUCC):
                            self._log("INFO", "IMU校准成功")
                        else:
                            self._log("WARNING", "IMU校准可能未成功")
                        return True, response, parse_result
            return False, None, None
        except Exception as e:
            self._debug_print(f"启动IMU校准时发生异常: {e}", is_error=True)
            return False, None, None

    # ---------------- 打印/辅助 ----------------
    def _bytes_to_hex(self, data: bytes) -> str:
        """字节数组转十六进制字符串"""
        return ' '.join(f'{b:02X}' for b in data)

    def _get_run_mode_str(self, mode: int) -> str:
        """运行模式数值转描述字符串"""
        modes = {
            0: "空闲模式",
            1: "设置模式",
            2: "运行模式",
            3: "锁定模式",
            4: "辅助模式",
            5: "刹车模式",
            6: "升级模式",
            7: "测试模式"
        }
        return modes.get(mode, f"未知模式({mode})")

    def _print_region_msg(self, region_msg: ppx_region_msg_t):
        """打印控制器消息结构体"""
        self._log("INFO", 
                 f"  ID: 0x{region_msg.id:02X}\n"
                 f"  命令: 0x{region_msg.cmd:02X} ({'写' if region_msg.cmd == PPX_MSG_WRITE else '读'})\n"
                 f"  消息类型: 0x{region_msg.msg_type:02X}\n"
                 f"  寄存器地址: 0x{region_msg.reg_addr:02X}\n"
                 f"  寄存器数量: {region_msg.reg_nums}\n"
                 f"  异常状态: 解析={region_msg.reg_excp.parse_status}, 命令={region_msg.reg_excp.cmd_status}, 数据={region_msg.reg_excp.data_status}")

    def _print_region_data(self, region_data: ppx_region_data_t):
        """打印控制器数据结构体"""
        # 字符串类型字段需要解码
        model_str = bytes(region_data.model).decode('utf-8', errors='replace').strip('\x00')
        sn_str = bytes(region_data.serial_num).decode('utf-8', errors='replace').strip('\x00')
        sw_ver_str = bytes(region_data.sw_version).decode('utf-8', errors='replace').strip('\x00')

        self._log("INFO",
                 (f"\n全局控制器数据结构体内容:\n"
                  f"设备ID号: {region_data.id_num}\n"
                  f"型号: {model_str}\n"
                  f"序列号: {sn_str}\n"
                  f"硬件版本: 0x{region_data.hw_version:04X}\n"
                  f"软件版本: {sw_ver_str}\n"
                  f"护盾状态: 0x{region_data.rim_state:02X} {self._get_rim_state_str(region_data.rim_state)}\n"
                  f"MCU错误码: 0x{region_data.mcu_errcode:08X}\n"
                  f"控制模式: {region_data.ctrl_model}\n"
                  f"速度参考值: {region_data.speed_ref}\n"
                  f"电机速度: {region_data.motor_speed} rpm\n"
                  f"总线电压: {region_data.bus_voltage * 0.1} V\n"
                  f"总线电流: {region_data.bus_current * 0.1} A\n"
                  f"A相电流: {region_data.phase_current_a * 0.1} A\n"
                  f"B相电流: {region_data.phase_current_b * 0.1} A\n"
                  f"C相电流: {region_data.phase_current_c * 0.1} A\n"
                  f"霍尔状态: 0x{region_data.hall_state:02X}\n"
                  f"PI_VQ参数: {region_data.pi_vq}\n"
                  f"PI_IQ参数: {region_data.pi_iq}\n"
                  f"刹车状态: {region_data.brake_state}\n"
                  f"IMU俯仰角: {region_data.imu_pitch * 0.1} °\n"
                  f"IMU横滚角: {region_data.imu_roll * 0.1} °\n"
                  f"IMU加速度: {region_data.imu_acc * 0.01} g\n"
                  f"板温: {region_data.board_temp} °C\n"
                  f"刹车里程: {region_data.brake_mileage} dm\n"
                  f"电机角度: {region_data.motor_angle}\n"
                  f"单次里程: {region_data.single_mileage} m\n"
                  f"角速度: {region_data.angular_speed * 0.1} °/s\n"
                  f"实时设置: 0x{region_data.rt_setting:04X} {self._get_rt_setting_str(region_data.rt_setting)}\n"
                  f"运行模式: {region_data.run_mode} {self._get_run_mode_str(region_data.run_mode)}\n"
                  f"档位: {region_data.gear}\n"
                  f"目标速度: {region_data.target_speed} rpm\n"
                  f"额定电压: {region_data.rated_voltage * 0.1} V\n"
                  f"额定电流: {region_data.rated_current * 0.1} A\n"
                  f"最大电压: {region_data.max_voltage * 0.1} V\n"
                  f"最小电压: {region_data.min_voltage * 0.1} V\n"
                  f"加速度: {region_data.acceration}\n"
                  f"数据设置: 0x{region_data.dat_setting:08X} {self._get_dat_setting_str(region_data.dat_setting)}\n"
                  f"保留数据: 0x{region_data.rsvd_data:08X}"))

    def _print_region_data_changes(self, region_msg: ppx_region_msg_t):
        """打印数据变化（根据寄存器地址过滤）"""
        if region_msg.reg_addr == PpxRegionReg.PPX_RUN_MODE_REG:
            self._log("INFO", "\n全局变量中的运行参数已更新:")
            self._log("INFO", 
                     f"  运行模式: {self.g_ppx_region_data.run_mode} {self._get_run_mode_str(self.g_ppx_region_data.run_mode)}\n"
                     f"  档位: {self.g_ppx_region_data.gear}\n"
                     f"  目标速度: {self.g_ppx_region_data.target_speed} rpm")
        elif region_msg.reg_addr == PpxRegionReg.PPX_RT_SETTING_REG:
            self._log("INFO", "\n全局变量中的实时设置已更新:")
            self._log("INFO", f"  {self._get_rt_setting_str(self.g_ppx_region_data.rt_setting)}")
        elif region_msg.reg_addr == PpxRegionReg.PPX_BUS_VOLTAGE_REG:
            self._log("INFO", "\n全局变量中的车辆状态已更新:")
            self._log("INFO", 
                     f"  总线电压: {self.g_ppx_region_data.bus_voltage * 0.1} V\n"
                     f"  总线电流: {self.g_ppx_region_data.bus_current * 0.1} A\n"
                     f"  电机速度: {self.g_ppx_region_data.motor_speed} rpm")

    def _get_rim_state_str(self, rim_state: int) -> str:
        """护盾状态位转描述字符串"""
        states = []
        if rim_state & PpxRimState.PPX_RIM_DOWN:
            states.append("下坡")
        if rim_state & PpxRimState.PPX_RIM_SEAT:
            states.append("座椅")
        if rim_state & PpxRimState.PPX_RIM_8DEG:
            states.append("8度")
        if rim_state & PpxRimState.PPX_RIM_DUMP:
            states.append("碰撞")
        if rim_state & PpxRimState.PPX_RIM_BUMP:
            states.append("颠簸")
        if rim_state & PpxRimState.PPX_RIM_TURN:
            states.append("转弯")
        return f"({','.join(states)})" if states else "(无状态)"

    def _get_rt_setting_str(self, rt_setting: int) -> str:
        """实时设置位转描述字符串"""
        settings = []
        if rt_setting & PpxRtSetting.PPX_BRAKE_LED_ON:
            settings.append("刹车灯开")
        if rt_setting & PpxRtSetting.PPX_TAIL_LED_ON:
            settings.append("尾灯开")
        if rt_setting & PpxRtSetting.PPX_RIGHT_LED_ON:
            settings.append("右转向灯开")
        if rt_setting & PpxRtSetting.PPX_LEFT_LED_ON:
            settings.append("左转向灯开")
        if rt_setting & PpxRtSetting.PPX_CLR_ERRCODE:
            settings.append("清除错误码")
        return f"({','.join(settings)})" if settings else "(无设置)"

    def _get_dat_setting_str(self, dat_setting: int) -> str:
        """数据设置位转描述字符串"""
        settings = []
        # 请求类型
        if dat_setting & PpxDatSetting.PPX_CHR_CHECK:
            settings.append("特征值检查")
        if dat_setting & PpxDatSetting.PPX_IMU_OPEN:
            settings.append("IMU开启")
        if dat_setting & PpxDatSetting.PPX_IMU_CALI:
            settings.append("IMU校准")
        if dat_setting & PpxDatSetting.PPX_IAP_MODE:
            settings.append("IAP升级模式")
        if dat_setting & PpxDatSetting.PPX_SN_WRITE:
            settings.append("写入序列号")
        if dat_setting & PpxDatSetting.PPX_TST_MOTO:
            settings.append("电机测试")
        if dat_setting & PpxDatSetting.PPX_ACC_CALI:
            settings.append("加速度校准")
        # 响应状态
        if dat_setting & PpxDatSetting.PPX_CHR_CHECK_SUCC:
            settings.append("特征值检查成功")
        if dat_setting & PpxDatSetting.PPX_IMU_OPEN_SUCC:
            settings.append("IMU开启成功")
        if dat_setting & PpxDatSetting.PPX_IMU_CALI_SUCC:
            settings.append("IMU校准成功")
        if dat_setting & PpxDatSetting.PPX_IAP_MODE_FALSE:
            settings.append("IAP模式关闭")
        if dat_setting & PpxDatSetting.PPX_ACC_CALI_SIDE:
            settings.append("加速度校准侧")
        if dat_setting & PpxDatSetting.PPX_ACC_CALI_SUCC:
            settings.append("加速度校准成功")
        return f"({','.join(settings)})" if settings else "(无设置)"

# ==============================================
# 自动化测试框架
# ==============================================
class FileLogger:
    """简单文件日志器：写入 raw.log，同时回显到控制台"""
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
    """将 1/0/True/False/"on"/"off"/"是"/"否" 等转换为 1/0 """
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
    def __init__(self, region, out_dir="output"):
        self.region = region
        self.out_dir = out_dir
        os.makedirs(out_dir, exist_ok=True)
        self.results: List[Dict[str, Any]] = []

    # ---------------- 用例读取 ----------------
    def load_cases(self, file_path=None):
        """优先加载 testcases.xlsx，没有则加载 testcases.csv"""
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
            # 检测CSV编码
            with open(file_path, 'rb') as f:
                raw_data = f.read()
            result = chardet.detect(raw_data)
            encoding = result['encoding'] or 'utf-8'
            cases = pd.read_csv(file_path, encoding=encoding)
        else:
            raise ValueError(f"不支持的用例文件格式: {ext}")

        if cases is None or cases.empty:
            raise ValueError(f"用例文件为空: {file_path}")

        return cases

    # ---------------- 执行用例 ----------------
    def run_cases(self, cases):
        for idx, row in enumerate(cases.to_dict(orient="records"), start=1):
            start_ts = time.time()
            case_id = row.get('id', idx)
            comment = str(row.get('comment', '') or '')
            test_type = row.get('test_type', 'set_run_mode')  # 测试类型：set_run_mode/set_rt_setting等

            # 公共参数
            recv_timeout = _coerce_float(row.get('recv_timeout'))
            delay_after = _coerce_float(row.get('delay_after')) or 0.0

            self.region._log("INFO", f"==== 执行用例 #{case_id} ====")
            self.region._log("INFO", f"测试类型: {test_type} | 备注: {comment}")

            # 根据测试类型执行不同操作
            ok = False
            resp = None
            parse_result = None
            read_ok = False
            read_resp = None
            region_data = None
            parse_read_result = None

            try:
                if test_type == 'set_run_mode':
                    # 设置运行模式
                    params = {
                        'run_mode': _coerce_int(row.get('run_mode')) or 0,
                        'gear': _coerce_int(row.get('gear')) or 0,
                        'target_speed': _coerce_int(row.get('target_speed')) or 0,
                    }
                    self.region._log("INFO", f"参数: {params}")
                    ok, resp, parse_result = self.region.set_run_mode(**params, recv_timeout=recv_timeout)

                elif test_type == 'set_rt_setting':
                    # 设置实时参数
                    params = {
                        'brake_led': _boolish_int(row.get('brake_led')) or False,
                        'tail_led': _boolish_int(row.get('tail_led')) or False,
                        'right_led': _boolish_int(row.get('right_led')) or False,
                        'left_led': _boolish_int(row.get('left_led')) or False,
                        'clear_err': _boolish_int(row.get('clear_err')) or False,
                    }
                    self.region._log("INFO", f"参数: {params}")
                    ok, resp, parse_result = self.region.set_rt_setting(** params, recv_timeout=recv_timeout)

                elif test_type == 'start_imu_calibration':
                    # 启动IMU校准
                    self.region._log("INFO", "启动IMU校准测试")
                    ok, resp, parse_result = self.region.start_imu_calibration(recv_timeout=recv_timeout or 10.0)

                else:
                    self.region._log("ERROR", f"未知测试类型: {test_type}")

                # 读取状态进行验证
                read_ok, read_resp, region_data, parse_read_result = self.region.read_vehicle_state(recv_timeout=recv_timeout)

            except Exception as e:
                self.region._debug_print(f"用例执行异常: {e}", is_error=True)

            # 处理结果
            recv_hex = self.region._bytes_to_hex(resp) if ok and resp else ''
            read_hex = self.region._bytes_to_hex(read_resp) if read_ok and read_resp else ''
            verdict = "PASS" if ok and read_ok else "FAIL"
            elapsed = time.time() - start_ts
            self.region._log("INFO", f"结果: {verdict} | 用时: {elapsed:.3f}s")

            # 提取校验信息
            checks = {}
            if region_data:
                try:
                    checks = {
                        "run_mode": region_data.run_mode,
                        "gear": region_data.gear,
                        "target_speed": region_data.target_speed,
                        "bus_voltage": region_data.bus_voltage * 0.1,
                        "bus_current": region_data.bus_current * 0.1,
                        "motor_speed": region_data.motor_speed,
                        "rt_setting": region_data.rt_setting,
                        "mcu_errcode": region_data.mcu_errcode,
                    }
                except Exception as e:
                    checks = {"error": f"无法解析控制器数据: {e}"}

            # 保存结果
            self.results.append({
                'case_id': case_id,
                'comment': comment,
                'test_type': test_type,
                'run_mode': _coerce_int(row.get('run_mode')),
                'gear': _coerce_int(row.get('gear')),
                'target_speed': _coerce_int(row.get('target_speed')),
                'brake_led': _boolish_int(row.get('brake_led')),
                'tail_led': _boolish_int(row.get('tail_led')),
                'right_led': _boolish_int(row.get('right_led')),
                'left_led': _boolish_int(row.get('left_led')),
                'clear_err': _boolish_int(row.get('clear_err')),
                'recv_timeout': recv_timeout,
                'delay_after': delay_after,
                'send_ok': "PASS" if ok else "FAIL",
                'read_ok': "PASS" if read_ok else "FAIL",
                'recv_hex': recv_hex,
                'read_hex': read_hex,
                'parse_send': parse_result,
                'parse_read': parse_read_result,
                'verdict': verdict,
                'checks': checks,
                'elapsed_s': round(elapsed, 3),
                'timestamp': time.strftime("%Y-%m-%d %H:%M:%S"),
            })

            if delay_after > 0:
                time.sleep(delay_after)

    # ---------------- 压力循环 ----------------
    def loop_case(self, case: Dict[str, Any], loop_count: int = 100, delay: float = 0.5):
        """对单个用例进行压力测试"""
        test_type = case.get('test_type', 'set_run_mode')
        self.region._log("INFO", f"开始压力循环: 次数={loop_count}, 间隔={delay}s, 类型={test_type}")

        for i in range(1, loop_count + 1):
            try:
                if test_type == 'set_run_mode':
                    params = {
                        'run_mode': _coerce_int(case.get('run_mode')) or 0,
                        'gear': _coerce_int(case.get('gear')) or 0,
                        'target_speed': _coerce_int(case.get('target_speed')) or 0,
                    }
                    ok, _, _ = self.region.set_run_mode(**params)
                elif test_type == 'set_rt_setting':
                    params = {
                        'brake_led': _boolish_int(case.get('brake_led')) or False,
                        'tail_led': _boolish_int(case.get('tail_led')) or False,
                        'right_led': _boolish_int(case.get('right_led')) or False,
                        'left_led': _boolish_int(case.get('left_led')) or False,
                    }
                    ok, _, _ = self.region.set_rt_setting(** params)
                else:
                    ok = False

                self.region._log("INFO", f"[循环 {i}/{loop_count}] 发送 {'OK' if ok else 'FAIL'}")
            except Exception as e:
                self.region._debug_print(f"循环 {i} 异常: {e}", is_error=True)
            time.sleep(delay)

    # ---------------- 报告输出 ----------------
    def save_results_csv(self, path: str):
        """保存结果到CSV"""
        keys = [
            'case_id','comment','test_type','run_mode','gear','target_speed',
            'brake_led','tail_led','right_led','left_led','clear_err',
            'recv_timeout','delay_after','send_ok','read_ok','recv_hex','read_hex',
            'parse_send','parse_read','verdict','checks','elapsed_s','timestamp'
        ]
        with open(path, 'w', newline='', encoding='utf-8') as f:
            w = csv.DictWriter(f, fieldnames=keys)
            w.writeheader()
            for r in self.results:
                w.writerow({k: r.get(k, '') for k in keys})

    def save_report_html(self, path: str, title: str = "控制器自动化测试报告"):
        """生成HTML报告"""
        total = len(self.results)
        passed = sum(1 for r in self.results if r.get('verdict') == 'PASS')
        failed = total - passed
        pass_rate = (passed / total * 100) if total else 0.0

        # 生成表格行HTML
        rows_html = []
        def esc(x: Any) -> str:
            s = str(x if x is not None else '')
            return s.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
        
        for idx, r in enumerate(self.results):
            rows_html.append(
                f"<tr data-row-index='{idx}'>" +
                "".join([
                f"<td>{esc(r.get('case_id'))}</td>",
                f"<td>{esc(r.get('comment'))}</td>",
                f"<td>{esc(r.get('test_type'))}</td>",
                f"<td>{esc(r.get('run_mode'))}</td>",
                f"<td>{esc(r.get('gear'))}</td>",
                f"<td>{esc(r.get('target_speed'))}</td>",
                f"<td>{'√' if r.get('brake_led') else '-'}</td>",
                f"<td>{'√' if r.get('tail_led') else '-'}</td>",
                f"<td>{'√' if r.get('right_led') else '-'}</td>",
                f"<td>{'√' if r.get('left_led') else '-'}</td>",
                f"<td>{'√' if r.get('clear_err') else '-'}</td>",
                f"<td>{esc(r.get('send_ok'))}</td>",
                f"<td>{esc(r.get('read_ok'))}</td>",
                f"<td style='font-family:monospace'>{esc(r.get('recv_hex'))}</td>",
                f"<td style='font-family:monospace'>{esc(r.get('read_hex'))}</td>",
                f"<td>{esc(r.get('parse_send'))}</td>",
                f"<td>{esc(r.get('parse_read'))}</td>",
                f"<td><b style='color:{'green' if r.get('verdict')=='PASS' else 'red'}'>{esc(r.get('verdict'))}</b></td>",
                f"<td>{esc(r.get('checks'))}</td>",
                f"<td>{esc(r.get('elapsed_s'))}</td>",
                f"<td>{esc(r.get('timestamp'))}</td>",
                ]) + "</tr>"
            )

        # HTML主体
        html = f"""
    <!DOCTYPE html>
    <html lang="zh-CN">
    <head>
    <meta charset="utf-8" />
    <title>{esc(title)}</title>
    <style>
    body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; margin: 24px; }}
    h1 {{ margin-bottom: 8px; }}
    .summary {{ margin: 8px 0 16px; }}
    .kpi span {{ display: inline-block; margin-right: 16px; padding: 6px 10px; border-radius: 8px; background: #f2f2f2; }}
    .filter-area {{ margin: 16px 0; padding: 12px; border: 1px solid #eee; border-radius: 8px; background: #fafafa; }}
    .filter-row {{ display: flex; flex-wrap: wrap; gap: 12px; margin-bottom: 8px; align-items: center; }}
    .filter-label {{ font-weight: 500; color: #666; min-width: 80px; }}
    .filter-input {{ padding: 6px 8px; border: 1px solid #ddd; border-radius: 4px; width: 180px; }}
    .filter-btn {{ padding: 6px 16px; border: none; border-radius: 4px; background: #4285f4; color: white; cursor: pointer; }}
    .filter-btn:hover {{ background: #3367d6; }}
    table {{ border-collapse: collapse; width: 100%; font-size: 14px; margin-top: 16px; }}
    th, td {{ border: 1px solid #ddd; padding: 6px 8px; text-align: left; vertical-align: top; }}
    th {{ background: #fafafa; position: sticky; top: 0; z-index: 1; }}
    tr:nth-child(even) {{ background: #fcfcfc; }}
    tr.hidden {{ display: none; }}
    code {{ background: #f6f8fa; padding: 2px 4px; border-radius: 4px; }}
    @media (max-width: 1200px) {{
        .filter-row {{ flex-direction: column; align-items: flex-start; }}
        .filter-label {{ min-width: auto; margin-bottom: 4px; }}
        .filter-input {{ width: 100%; max-width: 300px; }}
    }}
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

    <div class="filter-area">
      <h3>表格筛选</h3>
      <div class="filter-row">
        <div class="filter-label">测试类型：</div>
        <select id="filter-test-type" class="filter-input">
          <option value="">全部</option>
          <option value="set_run_mode">设置运行模式</option>
          <option value="set_rt_setting">设置实时参数</option>
          <option value="start_imu_calibration">IMU校准</option>
        </select>
      </div>
      <div class="filter-row">
        <div class="filter-label">用例结论：</div>
        <select id="filter-verdict" class="filter-input">
          <option value="">全部</option>
          <option value="PASS">通过（PASS）</option>
          <option value="FAIL">失败（FAIL）</option>
        </select>
      </div>
      <div class="filter-row">
        <div class="filter-label">关键词搜索：</div>
        <input type="text" id="filter-keyword" placeholder="搜索备注、响应数据等" class="filter-input">
        <button id="filter-btn" class="filter-btn">执行筛选</button>
        <button id="reset-btn" class="filter-btn" style="background:#9aa0a6">重置筛选</button>
      </div>
      <div class="filter-row">
        <div class="filter-label">筛选结果：</div>
        <div id="filter-result">当前显示 <b>{total}</b> 条，共 <b>{total}</b> 条</div>
      </div>
    </div>

    <table id="result-table">
      <thead>
        <tr>
          <th>ID</th><th>备注</th><th>测试类型</th>
          <th>运行模式</th><th>档位</th><th>目标速度</th>
          <th>刹车灯</th><th>尾灯</th><th>右转向</th><th>左转向</th><th>清错误</th>
          <th>发送OK</th><th>读取OK</th>
          <th>写入响应</th><th>读取响应</th>
          <th>写解析码</th><th>读解析码</th>
          <th>结论</th><th>校验详情</th><th>耗时(s)</th><th>时间戳</th>
        </tr>
      </thead>
      <tbody id="table-body">
        {''.join(rows_html)}
      </tbody>
    </table>

    <script>
    const tableBody = document.getElementById('table-body');
    const rows = tableBody.querySelectorAll('tr');
    const filterTestType = document.getElementById('filter-test-type');
    const filterVerdict = document.getElementById('filter-verdict');
    const filterKeyword = document.getElementById('filter-keyword');
    const filterBtn = document.getElementById('filter-btn');
    const resetBtn = document.getElementById('reset-btn');
    const filterResult = document.getElementById('filter-result');

    function applyFilter() {{
      const testType = filterTestType.value.trim();
      const verdict = filterVerdict.value.trim();
      const keyword = filterKeyword.value.trim().toLowerCase();
      let visibleCount = 0;

      rows.forEach(row => {{
        const cells = row.querySelectorAll('td');
        const rowTestType = cells[2].textContent.trim(); // 测试类型列
        const rowVerdict = cells[17].textContent.trim(); // 结论列
        const rowAllText = row.textContent.trim().toLowerCase();

        const typeMatch = !testType || rowTestType === testType;
        const verdictMatch = !verdict || rowVerdict === verdict;
        const keywordMatch = !keyword || rowAllText.includes(keyword);

        if (typeMatch && verdictMatch && keywordMatch) {{
          row.classList.remove('hidden');
          visibleCount++;
        }} else {{
          row.classList.add('hidden');
        }}
      }});

      filterResult.innerHTML = `当前显示 <b>${{visibleCount}}</b> 条，共 <b>${{rows.length}}</b> 条`;
    }}

    filterBtn.addEventListener('click', applyFilter);
    resetBtn.addEventListener('click', () => {{
      filterTestType.value = '';
      filterVerdict.value = '';
      filterKeyword.value = '';
      rows.forEach(row => row.classList.remove('hidden'));
      filterResult.innerHTML = `当前显示 <b>${{rows.length}}</b> 条，共 <b>${{rows.length}}</b> 条`;
    }});

    filterKeyword.addEventListener('keydown', (e) => {{
      if (e.key === 'Enter') applyFilter();
    }});

    window.addEventListener('load', applyFilter);
    </script>
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
        'test_type': 'set_run_mode',
        'comment': '设置为运行模式，1档，速度50rpm',
        'run_mode': 2,
        'gear': 1,
        'target_speed': 50,
        'recv_timeout': 1.0,
        'delay_after': 0.5,
    },
    {
        'id': 2,
        'test_type': 'set_rt_setting',
        'comment': '开启刹车灯和右转向灯，清除错误码',
        'brake_led': 1,
        'right_led': 1,
        'clear_err': 1,
        'recv_timeout': 1.0,
        'delay_after': 0.5,
    },
    {
        'id': 3,
        'test_type': 'start_imu_calibration',
        'comment': '启动IMU校准（超时设为10秒）',
        'recv_timeout': 10.0,
        'delay_after': 2.0,
    },
]


def make_sample_cases(path: str):
    """生成示例用例文件"""
    os.makedirs(os.path.dirname(path) or '.', exist_ok=True)
    if _HAS_PANDAS and path.lower().endswith(('.xlsx', '.xls')):
        df = pd.DataFrame(SAMPLE_CASES)
        df.to_excel(path, index=False)
    else:
        # 生成CSV
        keys = sorted({k for d in SAMPLE_CASES for k in d.keys()})
        with open(path, 'w', newline='', encoding='utf-8') as f:
            w = csv.DictWriter(f, fieldnames=keys)
            w.writeheader()
            for r in SAMPLE_CASES:
                w.writerow(r)


def make_combo_cases():
    """生成排列组合用例"""
    # 运行模式测试用例
    run_mode_cases = []
    for run_mode in [0, 1, 2, 3, 5]:  # 测试关键模式
        for gear in [0, 1, 2]:
            for speed in [0, 30, 60, 90]:
                run_mode_cases.append({
                    'test_type': 'set_run_mode',
                    'run_mode': run_mode,
                    'gear': gear,
                    'target_speed': speed,
                    'comment': f"模式{run_mode}({PpxRunMode(run_mode).name})，档位{gear}，速度{speed}rpm"
                })

    # 灯光控制用例
    light_cases = []
    for brake in [0, 1]:
        for tail in [0, 1]:
            for right in [0, 1]:
                for left in [0, 1]:
                    if brake + tail + right + left == 0:
                        continue  # 跳过全关（无意义）
                    light_cases.append({
                        'test_type': 'set_rt_setting',
                        'brake_led': brake,
                        'tail_led': tail,
                        'right_led': right,
                        'left_led': left,
                        'comment': f"灯光组合: 刹车{brake} 尾{tail} 右{right} 左{left}"
                    })

    # 合并用例并添加ID
    all_cases = run_mode_cases + light_cases
    for idx, case in enumerate(all_cases, 1):
        case['id'] = idx
        case['recv_timeout'] = 1.0
        case['delay_after'] = 0.3

    return all_cases

# ==============================================
# 主程序
# ==============================================

def main():
    print("\n控制器（Region）自动化测试工具")
    print("=" * 60)
    print(f"Python版本: {platform.python_version()}")
    print(f"系统架构: {platform.architecture()[0]}")
    print(f"操作系统: {platform.system()} {platform.release()}")

    # 解析参数
    parser = argparse.ArgumentParser(description="控制器自动化测试工具（Excel/CSV 用例 + HTML 报告）")
    parser.add_argument('--dll', default=REGION_PROTOCOL_DLL, help='ppx_region.dll 路径')
    parser.add_argument('--port', default=SERIAL_PORT, help='串口号，如 COM44')
    parser.add_argument('--baud', type=int, default=BAUDRATE, help='波特率，如 460800')
    parser.add_argument('--cases', default='testcases.xlsx', help='用例文件（xlsx/xls/csv），默认 testcases.xlsx')
    parser.add_argument('--out', default=None, help='报告输出目录（默认 ./reports/时间戳/）')
    parser.add_argument('--loop-count', type=int, default=1, help='对第1条用例进行压力循环次数（0 表示不进行）')
    parser.add_argument('--loop-delay', type=float, default=0.5, help='压力循环间隔秒')
    parser.add_argument('--make-sample', action='store_true', help='生成示例用例（不执行测试）')
    parser.add_argument('--combo', action='store_true', help='自动生成排列组合用例')
    args = parser.parse_args()

    # 如果无参数，默认开启组合用例模式
    if len(sys.argv) == 1:
        args.combo = True

    # 输出目录
    ts_dir = _dt.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    out_dir = args.out or os.path.join('reports', ts_dir)
    os.makedirs(out_dir, exist_ok=True)

    # 日志器
    logger = FileLogger(os.path.join(out_dir, 'raw.log'))

    # 生成示例用例
    if args.make_sample:
        out_path = args.cases
        print(f"生成示例用例: {out_path}")
        make_sample_cases(out_path)
        print("已生成。")
        return

    # 初始化控制器协议
    logger("INFO", "初始化控制器协议通信...")
    region = RegionProtocol(args.dll, args.port, args.baud)
    region.set_logger(logger)

    if not region.dll_loaded or not region.serial_connected:
        logger("ERROR", "初始化失败，请检查 DLL 路径/串口参数")
        region.close()
        sys.exit(1)

    # 初始化测试器
    tester = AutoTester(region, out_dir)

    # 读取用例
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
                region.close()
                sys.exit(2)
            logger("INFO", f"加载用例 {len(cases)} 条 来自: {args.cases}")
    except Exception as e:
        logger("ERROR", f"读取用例失败: {e}")
        region.close()
        sys.exit(3)

    # 执行用例
    try:
        tester.run_cases(cases)

        # 压力循环（可选）
        if args.loop_count > 0 and len(cases) > 0:
            tester.loop_case(cases.iloc[0], loop_count=args.loop_count, delay=args.loop_delay)

        # 保存结果
        csv_path = os.path.join(out_dir, 'results.csv')
        html_path = os.path.join(out_dir, 'report.html')
        tester.save_results_csv(csv_path)
        tester.save_report_html(html_path)
        logger("INFO", f"已保存结果: {csv_path}")
        logger("INFO", f"已生成报告: {html_path}")

    except Exception as e:
        logger("ERROR", f"测试执行失败: {e}")
    finally:
        region.close()
        logger("INFO", "程序结束")


if __name__ == "__main__":
    main()
