# -*- coding: utf-8 -*-
"""
MCB 硬件诊断工具 V2.9 (霍尔传感器体检版)
============================================================
目的：排查 0x040000/0x240000 故障的物理根源。
原理：不给电机通电，仅读取霍尔传感器状态 (Hall State)。

使用方法：
1. 运行脚本。
2. 看到 "开始监听" 后，【用手用力转动电机轮子】。
3. 观察屏幕上的 [Hall State] 数值变化。

判断标准：
- 正常：数值在 1, 2, 3, 4, 5, 6 之间快速跳变。
- 故障：数值一直卡在 0 或 7，或者不动。
  -> 0/7 代表霍尔插头松脱或断线。
  -> 不动代表霍尔损坏。
"""

import ctypes
from ctypes import *
import serial
import time
import os
import threading
import sys

# ==============================================
# 基础配置
# ==============================================
DLL_PATH = os.path.join(os.path.dirname(__file__), "ppx_region.dll")
SERIAL_PORT = "COM9"
BAUDRATE = 460800
MCB_DEV_ID = 0x20

# 寄存器
REG_HW_VERSION = 3
REG_HALL_STATE = 15  # [15] 霍尔状态 (0-7)
REG_BUS_VOLT = 10


# 结构体定义
class ppx_region_excp_t(Structure):
    _fields_ = [("parse_status", c_uint8), ("cmd_status", c_uint8), ("data_status", c_uint8)]


class ppx_region_msg_t(Structure):
    _fields_ = [("id", c_uint8), ("cmd", c_uint8), ("msg_type", c_uint8), ("reg_addr", c_uint8), ("reg_nums", c_uint8),
                ("reg_excp", ppx_region_excp_t)]


class ppx_region_data_t(Structure):
    _pack_ = 1
    _fields_ = [("id_num", c_uint8), ("model", c_uint8 * 8), ("serial_num", c_uint8 * 26),
                ("hw_version", c_uint16), ("sw_version", c_uint8 * 20),
                ("rim_state", c_uint8), ("mcu_errcode", c_uint32),
                ("ctrl_model", c_uint8), ("speed_ref", c_int16),
                ("motor_speed", c_int16),
                ("bus_voltage", c_uint16), ("bus_current", c_uint16),
                ("phase_current_a", c_int16), ("phase_current_b", c_int16), ("phase_current_c", c_int16),
                ("hall_state", c_uint8),  # Reg 15
                ("pi_vq", c_int16), ("pi_iq", c_int16), ("brake_state", c_uint8), ("imu_pitch", c_int16),
                ("imu_roll", c_int16), ("imu_acc", c_uint8), ("brake_mileage", c_uint8), ("motor_angle", c_int32),
                ("single_mileage", c_uint32), ("angular_speed", c_int16), ("rt_setting", c_uint16),
                ("run_mode", c_uint8), ("gear", c_uint8), ("target_speed", c_int16), ("rated_voltage", c_uint16),
                ("rated_current", c_uint16), ("max_voltage", c_uint16), ("min_voltage", c_uint16),
                ("acceration", c_uint32), ("dat_setting", c_uint32), ("rsvd_data", c_uint32)]


class TestEngine:
    def __init__(self):
        self.ser = None
        self.lib = None
        self.g_data = None

    def setup(self):
        if not os.path.exists(DLL_PATH): print("DLL缺失"); return False
        try:
            self.lib = cdll.LoadLibrary(DLL_PATH)
            self.lib.ppx_com_region_format.argtypes = [c_int, POINTER(ppx_region_msg_t), c_void_p]
            self.lib.ppx_com_region_format.restype = c_uint16
            self.lib.ppx_com_region_parse.argtypes = [POINTER(c_uint8), c_uint8, POINTER(ppx_region_msg_t)]
            self.lib.ppx_com_region_parse.restype = c_int
            self.g_data = ppx_region_data_t.in_dll(self.lib, "g_ppx_region_data")
            self.ser = serial.Serial(SERIAL_PORT, BAUDRATE, timeout=0.1)
            return True
        except Exception as e:
            print(f"连接失败: {e}"); return False

    def read_reg(self, reg):
        self.ser.reset_input_buffer()
        msg = ppx_region_msg_t()
        msg.id, msg.cmd, msg.reg_addr, msg.reg_nums = MCB_DEV_ID, 0x01, reg, 1
        buf = create_string_buffer(256)
        length = self.lib.ppx_com_region_format(0, byref(msg), buf)
        self.ser.write(buf.raw[:length])
        time.sleep(0.05)
        if self.ser.in_waiting:
            recv = self.ser.read(self.ser.in_waiting)
            msg_res = ppx_region_msg_t()
            msg_res.id = MCB_DEV_ID
            if self.lib.ppx_com_region_parse((c_uint8 * len(recv))(*recv), len(recv), byref(msg_res)) == 1:
                if reg == REG_HALL_STATE: return self.g_data.hall_state
                if reg == REG_BUS_VOLT: return self.g_data.bus_voltage
        return None


def main():
    eng = TestEngine()
    if not eng.setup(): return

    print("=" * 50)
    print("      MCB 霍尔传感器体检工具 V2.9")
    print("=" * 50)

    # 检查电压
    v = eng.read_reg(REG_BUS_VOLT)
    if v:
        print(f"✅ 当前电压: {v * 0.1:.1f}V (硬件供电正常)")
    else:
        print("❌ 无法读取电压，通信中断"); return

    print("\n🎧 开始监听霍尔信号 (持续 20秒)...")
    print("👉 请现在【用手用力转动】电机轮子！")
    print("-" * 30)
    print("时间(s) | 霍尔状态 (Hall State)")
    print("-" * 30)

    start_t = time.time()
    last_hall = -1
    change_count = 0

    while time.time() - start_t < 20:
        h = eng.read_reg(REG_HALL_STATE)

        if h is not None:
            status_str = f"{h} "
            if h == 0 or h == 7:
                status_str += "❌ (异常:断线/非法)"
            else:
                status_str += "✅ (正常)"

            # 只有变化时才打印，避免刷屏
            if h != last_hall:
                print(f"{time.time() - start_t:5.1f}s | {status_str}")
                last_hall = h
                if 1 <= h <= 6: change_count += 1

        time.sleep(0.1)

    print("-" * 30)
    print("🔍 诊断结果：")
    if change_count > 5:
        print("✅ 霍尔传感器工作正常！(检测到状态跳变)")
        print("👉 结论：硬件连接没问题。故障原因是【相线线序错误】。")
        print("👉 建议：请调换黄/绿/蓝粗线的接线顺序，再次尝试 V2.8 脚本。")
    else:
        print("❌ 霍尔传感器无反应！")
        print("👉 结论：霍尔线(细线)没插好，或者传感器已损坏。")

    eng.ser.close()


if __name__ == "__main__":
    main()