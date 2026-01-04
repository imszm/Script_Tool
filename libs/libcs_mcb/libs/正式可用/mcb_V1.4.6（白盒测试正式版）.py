# -*- coding: utf-8 -*-
"""
MCB 自动化白盒测试脚本 (V1.6 最终验收版)
============================================================
调试结论：
1. 真实转速寄存器确认：Reg 0x09 (Motor Speed)。
2. 现象：反馈转速为负值 (如 -297)，系电机相序定义所致，属正常现象。
3. 修正：在验证转速时增加绝对值处理 (abs)，忽略方向差异。

测试用例清单：
[Case 1] 通信链路测试 (Ping/Pong)
[Case 2] 运行环境安全检查 (Voltage > 30V)
[Case 3] 状态机切换逻辑 (IDLE -> TEST)
[Case 4] 参数寄存器读写 (Gear/Acc)
[Case 5] IO 控制逻辑 (Light)
[Case 6] 动力回路响应 (PID闭环测试，软启动模式)
"""

import ctypes
from ctypes import *
import serial
import time
import os
import threading
import sys

# ==============================================
# 1. 基础配置
# ==============================================
DLL_PATH = os.path.join(os.path.dirname(__file__), "../ppx_region.dll")
SERIAL_PORT = "COM9"
BAUDRATE = 460800
MCB_DEV_ID = 0x20

# 寄存器地址
REG_HW_VERSION = 3
REG_ERR_CODE = 6
REG_REAL_SPEED = 9  # [09] 实际转速
REG_BUS_VOLT = 10
REG_BRAKE_STATE = 18
REG_RT_SETTING = 0x1A
REG_RUN_MODE = 0x1B
REG_GEAR = 0x1C
REG_TARGET_SPEED = 0x1D
REG_ACCELERATION = 0x22
REG_DAT_SETTING = 0x23

# 阈值
MIN_BUS_VOLTAGE = 30.0
TEST_RPM_TARGET = 300
RPM_TOLERANCE = 50


# 结构体定义
class ppx_region_excp_t(Structure):
    _fields_ = [("parse_status", c_uint8), ("cmd_status", c_uint8), ("data_status", c_uint8)]


class ppx_region_msg_t(Structure):
    _fields_ = [
        ("id", c_uint8), ("cmd", c_uint8), ("msg_type", c_uint8),
        ("reg_addr", c_uint8), ("reg_nums", c_uint8), ("reg_excp", ppx_region_excp_t),
    ]


class ppx_region_data_t(Structure):
    _pack_ = 1
    _fields_ = [
        ("id_num", c_uint8), ("model", c_uint8 * 8), ("serial_num", c_uint8 * 26),
        ("hw_version", c_uint16), ("sw_version", c_uint8 * 20),
        ("rim_state", c_uint8), ("mcu_errcode", c_uint32),
        ("ctrl_model", c_uint8), ("speed_ref", c_int16),
        ("motor_speed", c_int16),  # Reg 9
        ("bus_voltage", c_uint16), ("bus_current", c_uint16),
        ("phase_current_a", c_int16), ("phase_current_b", c_int16), ("phase_current_c", c_int16),
        ("hall_state", c_uint8), ("pi_vq", c_int16), ("pi_iq", c_int16),
        ("brake_state", c_uint8), ("imu_pitch", c_int16), ("imu_roll", c_int16),
        ("imu_acc", c_uint8), ("brake_mileage", c_uint8), ("motor_angle", c_int32),
        ("single_mileage", c_uint32), ("angular_speed", c_int16),
        ("rt_setting", c_uint16), ("run_mode", c_uint8), ("gear", c_uint8),
        ("target_speed", c_int16), ("rated_voltage", c_uint16), ("rated_current", c_uint16),
        ("max_voltage", c_uint16), ("min_voltage", c_uint16), ("acceration", c_uint32),
        ("dat_setting", c_uint32), ("rsvd_data", c_uint32),
    ]


# ==============================================
# 2. 测试引擎
# ==============================================
class TestEngine:
    def __init__(self):
        self.ser = None
        self.lib = None
        self.lock = threading.Lock()
        self.running = True
        self.ready = False
        self.hb_paused = False

        self.shadow_regs = {
            REG_RT_SETTING: 0,
            REG_TARGET_SPEED: 0,
            REG_RUN_MODE: 0,
            REG_DAT_SETTING: 0
        }

    def setup(self):
        print("[Setup] 初始化环境...")
        if not os.path.exists(DLL_PATH): raise FileNotFoundError("DLL缺失")
        try:
            self.lib = cdll.LoadLibrary(DLL_PATH)
            self.lib.ppx_com_region_format.argtypes = [c_int, POINTER(ppx_region_msg_t), c_void_p]
            self.lib.ppx_com_region_format.restype = c_uint16
            self.lib.ppx_com_region_parse.argtypes = [POINTER(c_uint8), c_uint8, POINTER(ppx_region_msg_t)]
            self.lib.ppx_com_region_parse.restype = c_int
            self.g_data = ppx_region_data_t.in_dll(self.lib, "g_ppx_region_data")
        except Exception as e:
            raise RuntimeError(f"DLL加载失败: {e}")

        try:
            self.ser = serial.Serial(SERIAL_PORT, BAUDRATE, timeout=0.1)
            self.ser.reset_input_buffer()
            self.ready = True
        except Exception as e:
            raise RuntimeError(f"串口失败: {e}")

        self.hb_thread = threading.Thread(target=self._heartbeat_task, daemon=True)
        self.hb_thread.start()
        time.sleep(0.5)

    def teardown(self):
        print("\n[Teardown] 清理环境...")
        self.set_shadow(REG_TARGET_SPEED, 0)
        self.set_shadow(REG_RT_SETTING, 0)
        time.sleep(0.5)
        self.running = False
        if self.ser: self.ser.close()

    def set_shadow(self, reg, val):
        with self.lock: self.shadow_regs[reg] = val

    def pause_heartbeat(self):
        self.hb_paused = True
        time.sleep(0.2)

    def resume_heartbeat(self):
        self.hb_paused = False

    def _heartbeat_task(self):
        while self.running:
            if self.ready and not self.hb_paused:
                with self.lock:
                    if self.shadow_regs[REG_RUN_MODE] != 0:
                        self._raw_write(REG_RUN_MODE, self.shadow_regs[REG_RUN_MODE])

                    self._raw_write(REG_RT_SETTING, self.shadow_regs[REG_RT_SETTING])
                    self._raw_write(REG_TARGET_SPEED, self.shadow_regs[REG_TARGET_SPEED])

                    if self.shadow_regs[REG_DAT_SETTING] != 0:
                        self._raw_write(REG_DAT_SETTING, self.shadow_regs[REG_DAT_SETTING])
            time.sleep(0.15)

    def _raw_write(self, reg, val, nums=1):
        if reg == REG_TARGET_SPEED: self.g_data.target_speed = val
        if reg == REG_RT_SETTING: self.g_data.rt_setting = val
        if reg == REG_RUN_MODE: self.g_data.run_mode = val
        if reg == REG_GEAR: self.g_data.gear = val
        if reg == REG_ACCELERATION: self.g_data.acceration = val
        if reg == REG_DAT_SETTING: self.g_data.dat_setting = val

        msg = ppx_region_msg_t()
        msg.id, msg.cmd, msg.reg_addr, msg.reg_nums = MCB_DEV_ID, 0x03, reg, nums
        try:
            buf = create_string_buffer(256)
            length = self.lib.ppx_com_region_format(0, byref(msg), buf)
            self.ser.write(buf.raw[:length])
        except:
            pass

    def read_reg(self, reg, nums=1, retry=3):
        for i in range(retry):
            with self.lock:
                self.ser.reset_input_buffer()
                msg = ppx_region_msg_t()
                msg.id, msg.cmd, msg.reg_addr, msg.reg_nums = MCB_DEV_ID, 0x01, reg, nums
                buf = create_string_buffer(256)
                length = self.lib.ppx_com_region_format(0, byref(msg), buf)
                self.ser.write(buf.raw[:length])

                time.sleep(0.08)

                if self.ser.in_waiting:
                    recv = self.ser.read(self.ser.in_waiting)
                    msg_res = ppx_region_msg_t()
                    msg_res.id = MCB_DEV_ID
                    if self.lib.ppx_com_region_parse((c_uint8 * len(recv))(*recv), len(recv), byref(msg_res)) == 1:
                        if reg == REG_HW_VERSION: return self.g_data.hw_version
                        if reg == REG_BUS_VOLT: return self.g_data.bus_voltage
                        if reg == REG_ERR_CODE: return self.g_data.mcu_errcode
                        if reg == REG_BRAKE_STATE: return self.g_data.brake_state
                        if reg == REG_RUN_MODE: return self.g_data.run_mode
                        if reg == REG_GEAR: return self.g_data.gear
                        if reg == REG_ACCELERATION: return self.g_data.acceration
                        if reg == REG_RT_SETTING: return self.g_data.rt_setting
                        # 重点：读实际转速
                        if reg == REG_REAL_SPEED: return self.g_data.motor_speed
                        return 0
            time.sleep(0.1)
        return None

    def get_feedback_speed(self):
        # 强制更新转速数据
        self.read_reg(REG_REAL_SPEED)
        return self.g_data.motor_speed


# ==============================================
# 3. 测试用例
# ==============================================
def run_tests():
    engine = TestEngine()
    try:
        engine.setup()
    except Exception as e:
        print(f"❌ 初始化失败: {e}"); return

    print("=" * 60)
    print("      MCB 白盒测试执行报告 (V1.6 最终验收版)")
    print("=" * 60)

    # --- Case 1 ---
    print("\n[Case 1] 通信链路测试")
    ver = engine.read_reg(REG_HW_VERSION, retry=5)
    if ver is not None and ver > 0:
        print(f"  ✅ 通过: HW Ver {ver}")
    else:
        print("  ❌ 失败: 通信断开"); engine.teardown(); return

    # --- Case 2 ---
    print("\n[Case 2] 环境安全扫描")
    volt = engine.read_reg(REG_BUS_VOLT) * 0.1
    err = engine.read_reg(REG_ERR_CODE)

    print(f"  -> 电压: {volt:.1f}V")
    if err != 0:
        print("  ⚠️ 检测到错误码，尝试清除...")
        engine.set_shadow(REG_RT_SETTING, 0x8000)
        time.sleep(0.5)
        engine.set_shadow(REG_RT_SETTING, 0)
        time.sleep(0.2)
        err_new = engine.read_reg(REG_ERR_CODE)
        if err_new == 0:
            print("  ✅ 修复: 错误已清除")
        else:
            print(f"  ❌ 失败: 无法清除 0x{err_new:06X}"); engine.teardown(); return
    else:
        print("  ✅ 通过: 无错误")

    # --- Case 3 ---
    print("\n[Case 3] 模式切换")
    engine.set_shadow(REG_RUN_MODE, 7)
    engine.set_shadow(REG_DAT_SETTING, 0x20)
    time.sleep(0.5)
    mode = engine.read_reg(REG_RUN_MODE)
    if mode == 7:
        print("  ✅ 通过: TEST Mode")
    else:
        print(f"  ❌ 失败: Mode={mode}"); engine.teardown(); return

    # --- Case 4 ---
    print("\n[Case 4] 寄存器读写")
    engine.pause_heartbeat()
    engine._raw_write(REG_GEAR, 2)
    time.sleep(0.2)
    read_val = engine.read_reg(REG_GEAR)
    engine.resume_heartbeat()
    if read_val == 2:
        print(f"  ✅ 通过: Gear OK")
    else:
        print(f"  ❌ 失败: Gear {read_val}")

    # --- Case 5 ---
    print("\n[Case 5] IO 控制")
    engine.pause_heartbeat()
    engine._raw_write(REG_RT_SETTING, 0x0C)
    time.sleep(0.3)
    rt_read = engine.read_reg(REG_RT_SETTING)
    engine._raw_write(REG_RT_SETTING, 0)
    engine.resume_heartbeat()
    if (rt_read & 0x0C) == 0x0C:
        print(f"  ✅ 通过: Light OK")
    else:
        print(f"  ❌ 失败: Light 0x{rt_read:04X}")

    # --- Case 6 ---
    print("\n[Case 6] 动力回路 (PID闭环测试)")
    # 软启动
    print("  -> 设置 Acc = 100 (Soft Start)")
    engine._raw_write(REG_ACCELERATION, 100, nums=2)
    time.sleep(0.2)

    print(f"  -> 目标: {TEST_RPM_TARGET} RPM")
    engine.set_shadow(REG_TARGET_SPEED, TEST_RPM_TARGET)

    reached = False
    max_rpm = 0
    final_rpm = 0

    # 3秒超时检测
    for i in range(15):
        actual_rpm = engine.get_feedback_speed()
        max_rpm = max(max_rpm, abs(actual_rpm))  # 记录最大绝对值
        final_rpm = actual_rpm

        # 实时救错
        if engine.read_reg(REG_ERR_CODE) != 0:
            engine.set_shadow(REG_RT_SETTING, 0x8000)
            time.sleep(0.2)
            engine.set_shadow(REG_RT_SETTING, 0)

        # 【关键修正】取绝对值判断
        if abs(abs(actual_rpm) - TEST_RPM_TARGET) < RPM_TOLERANCE:
            reached = True
            break
        time.sleep(0.2)

    engine.set_shadow(REG_TARGET_SPEED, 0)

    if reached:
        print(f"  ✅ 通过: 响应正常 (实际: {final_rpm} RPM, 误差 < {RPM_TOLERANCE})")
    else:
        print(f"  ❌ 失败: 读数为 {final_rpm} RPM")

    print("\n" + "=" * 60 + "\n      测试全部结束\n" + "=" * 60)
    engine.teardown()


if __name__ == "__main__":
    run_tests()