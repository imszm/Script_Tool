# -*- coding: utf-8 -*-
"""
完整示例：解析 MCB Region 数据包并打印 msg_out
"""

import ctypes
from ctypes import *
from typing import Optional

# ---------------- 常量 ----------------
PPX_SW_VER_SIZE = 20
PPX_MODEL_SIZE = 8
PPX_SN_SIZE = 26

PPX_ID_MCB = 0x20
PPX_MSG_MULTREAD = 0x02
PPX_PARSE_SUCCESS = 1  # DLL 成功返回值

# 软件版本寄存器
PPX_SW_VERSION_REG = 4

# ---------------- 结构体定义 ----------------
class ppx_region_excp_t(Structure):
    _fields_ = [
        ("parse_status", c_uint8),
        ("cmd_status", c_uint8),
        ("data_status", c_uint8),
    ]

class ppx_region_msg_t(Structure):
    _fields_ = [
        ("id", c_uint8),
        ("cmd", c_uint8),
        ("msg_type", c_uint8),
        ("reg_addr", c_uint8),
        ("reg_nums", c_uint8),
        ("reg_excp", ppx_region_excp_t),
    ]

class ppx_region_data_t(Structure):
    _pack_ = 1
    _fields_ = [
        ("id_num", c_uint8),
        ("model", c_uint8 * PPX_MODEL_SIZE),
        ("serial_num", c_uint8 * PPX_SN_SIZE),
        ("hw_version", c_uint16),
        ("sw_version", c_uint8 * PPX_SW_VER_SIZE),
        ("rim_state", c_uint8),
        ("mcu_errcode", c_uint32),
        ("ctrl_model", c_uint8),
        ("speed_ref", c_int16),
        ("motor_speed", c_int16),
        ("bus_voltage", c_uint16),
        ("bus_current", c_uint16),
        ("phase_current_a", c_int16),
        ("phase_current_b", c_int16),
        ("phase_current_c", c_int16),
        ("hall_state", c_uint8),
        ("pi_vq", c_int16),
        ("pi_iq", c_int16),
        ("brake_state", c_uint8),
        ("imu_pitch", c_int16),
        ("imu_roll", c_int16),
        ("imu_acc", c_uint8),
        ("brake_mileage", c_uint8),
        ("motor_angle", c_int32),
        ("single_mileage", c_uint32),
        ("angular_speed", c_int16),
        ("rt_setting", c_uint16),
        ("run_mode", c_uint8),
        ("gear", c_uint8),
        ("target_speed", c_int16),
        ("rated_voltage", c_uint16),
        ("rated_current", c_uint16),
        ("max_voltage", c_uint16),
        ("min_voltage", c_uint16),
        ("acceration", c_uint32),
        ("dat_setting", c_uint32),
        ("rsvd_data", c_uint32),
    ]

# 全局数据对象
g_ppx_region_data = ppx_region_data_t()

# ---------------- 加载 DLL ----------------
dll = ctypes.CDLL(r".\ppx_region.dll")
ppx_com_region_parse = dll.ppx_com_region_parse
ppx_com_region_parse.argtypes = [POINTER(c_uint8), c_uint8, POINTER(ppx_region_msg_t)]
ppx_com_region_parse.restype = c_int

# ---------------- 输入帧 ----------------
data_hex = "a5 20 df a5 02 0c 38 38 33 33 33 33 33 33 33 33 33 33 08 55"
data_bytes = bytes.fromhex(data_hex)
data_len = len(data_bytes)
pdata = (c_uint8 * data_len)(*data_bytes)

# ---------------- 调用解析函数 ----------------
msg_out = ppx_region_msg_t()
ret = ppx_com_region_parse(pdata, data_len, byref(msg_out))

# ---------------- 打印解析结果 ----------------
print(f"解析返回值: {ret}")
print(f"msg_out.id       = {msg_out.id}")
print(f"msg_out.cmd      = {msg_out.cmd}")
print(f"msg_out.msg_type = {msg_out.msg_type}")
print(f"msg_out.reg_addr = {msg_out.reg_addr}")
print(f"msg_out.reg_nums = {msg_out.reg_nums}")
print("msg_out.reg_excp:")
print(f"  parse_status = {msg_out.reg_excp.parse_status}")
print(f"  cmd_status   = {msg_out.reg_excp.cmd_status}")
print(f"  data_status  = {msg_out.reg_excp.data_status}")

# ---------------- 打印 g_ppx_region_data 的软件版本 ----------------
sw_ver_bytes = bytes(g_ppx_region_data.sw_version).split(b'\x00', 1)[0]
sw_ver_str = sw_ver_bytes.decode('ascii', errors='ignore')
print(f"MCB 软件版本号: {sw_ver_str}")
