# -*- coding: utf-8 -*-
"""
MCB 智能控制终端 (V4.0 旗舰版)
============================================================
功能亮点：
1. [核心控制] 完美集成了波特率修正、模式切换、心跳守护。
2. [实时监控] 新增读取电压(V)、电流(A)、错误码的功能。
3. [高级特效] 新增 SOS 求救信号、警用爆闪模式。
4. [交互升级] 实时刷新状态显示。

使用说明：
- 确保 ppx_region.dll 在同级目录。
- 运行后自动初始化，无需手动干预。
"""

import ctypes
from ctypes import *
import serial
import time
import os
import threading
import random

# ==============================================
# 核心配置
# ==============================================
DLL_PATH = os.path.join(os.path.dirname(__file__), "../ppx_region.dll")
SERIAL_PORT = "COM6"
BAUDRATE = 460800
MCB_DEV_ID = 0x20

# 寄存器地址映射 (参考 ppx_region.h)
REG_HW_VERSION = 3
REG_BUS_VOLT = 10  # 母线电压 (0.1V)
REG_BUS_CURR = 11  # 母线电流 (0.1A)
REG_ERR_CODE = 6  # 错误码
REG_RT_SETTING = 0x1A  # 灯光控制
REG_RUN_MODE = 0x1B  # 运行模式
REG_DAT_SETTING = 0x23  # 权限位

MODE_TST = 7  # 测试模式


# ==============================================
# 结构体定义 (Standard)
# ==============================================
class ppx_region_excp_t(Structure):
    _fields_ = [("parse_status", c_uint8), ("cmd_status", c_uint8), ("data_status", c_uint8)]


class ppx_region_msg_t(Structure):
    _fields_ = [
        ("id", c_uint8), ("cmd", c_uint8), ("msg_type", c_uint8),
        ("reg_addr", c_uint8), ("reg_nums", c_uint8), ("reg_excp", ppx_region_excp_t),
    ]


# 完整数据结构映射
class ppx_region_data_t(Structure):
    _pack_ = 1
    _fields_ = [
        ("id_num", c_uint8), ("model", c_uint8 * 8), ("serial_num", c_uint8 * 26),
        ("hw_version", c_uint16), ("sw_version", c_uint8 * 20),
        ("rim_state", c_uint8), ("mcu_errcode", c_uint32),
        ("ctrl_model", c_uint8), ("speed_ref", c_int16), ("motor_speed", c_int16),
        ("bus_voltage", c_uint16), ("bus_current", c_uint16),  # [10, 11]
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
# 智能控制器类
# ==============================================
class SmartMCB(object):
    def __init__(self, dll_path, port, baudrate):
        self.ready = False
        self.ser = None
        self.lock = threading.Lock()  # 串口互斥锁

        # 状态变量
        self.target_left = 0
        self.target_right = 0
        self.running = True
        self.monitor_data = {"volt": 0.0, "curr": 0.0, "err": 0}

        # 1. 加载 DLL
        if not os.path.exists(dll_path):
            print(f"❌ 错误: 找不到 {dll_path}")
            return
        try:
            self.lib = cdll.LoadLibrary(dll_path)
            self.lib.ppx_com_region_format.argtypes = [c_int, POINTER(ppx_region_msg_t), c_void_p]
            self.lib.ppx_com_region_format.restype = c_uint16
            self.lib.ppx_com_region_parse.argtypes = [POINTER(c_uint8), c_uint8, POINTER(ppx_region_msg_t)]
            self.lib.ppx_com_region_parse.restype = c_int
            self.g_data = ppx_region_data_t.in_dll(self.lib, "g_ppx_region_data")
        except Exception as e:
            print(f"❌ DLL 加载失败: {e}")
            return

        # 2. 连接串口
        try:
            self.ser = serial.Serial(port, baudrate, timeout=0.1)
            print(f"✅ 串口已连接: {port}")
            self.ready = True
        except Exception as e:
            print(f"❌ 串口打开失败: {e}")
            return

        # 3. 启动后台线程
        self.tx_thread = threading.Thread(target=self._heartbeat_loop, daemon=True)
        self.tx_thread.start()

        self.rx_thread = threading.Thread(target=self._monitor_loop, daemon=True)
        self.rx_thread.start()

    def close(self):
        self.running = False
        time.sleep(0.5)
        if self.ser and self.ser.is_open:
            self.ser.close()

    # --- 线程1: 心跳发送 (负责控制) ---
    def _heartbeat_loop(self):
        while self.running:
            if self.ready:
                # 组装控制值
                val = 0
                if self.target_left: val |= 0x08
                if self.target_right: val |= 0x04

                with self.lock:  # 串口加锁
                    self._send_cmd(0x03, REG_RT_SETTING, val, wait_resp=False)
            time.sleep(0.15)  # 150ms 刷新一次控制

    # --- 线程2: 状态监控 (负责读取) ---
    def _monitor_loop(self):
        while self.running:
            if self.ready:
                with self.lock:  # 串口加锁
                    # 读取电压电流 (Reg 10, len=2)
                    # 注意：这里为了简单演示，每次读一个关键值，轮询
                    if self._send_cmd(0x01, REG_BUS_VOLT, 0, wait_resp=True):
                        self.monitor_data["volt"] = self.g_data.bus_voltage * 0.1
                        self.monitor_data["curr"] = self.g_data.bus_current * 0.1

                    time.sleep(0.05)
                    if self._send_cmd(0x01, REG_ERR_CODE, 0, wait_resp=True):
                        self.monitor_data["err"] = self.g_data.mcu_errcode
            time.sleep(1.0)  # 1秒刷新一次状态

    def _send_cmd(self, cmd, reg, val, wait_resp=False):
        # 更新 DLL 数据
        if cmd == 0x03:
            if reg == REG_RT_SETTING: self.g_data.rt_setting = val
            if reg == REG_RUN_MODE: self.g_data.run_mode = val
            if reg == REG_DAT_SETTING: self.g_data.dat_setting = val

        msg = ppx_region_msg_t()
        msg.id, msg.cmd, msg.reg_addr, msg.reg_nums = MCB_DEV_ID, cmd, reg, 1

        try:
            buf = create_string_buffer(256)
            length = self.lib.ppx_com_region_format(0, byref(msg), buf)
            self.ser.write(buf.raw[:length])

            if wait_resp:
                time.sleep(0.05)
                if self.ser.in_waiting:
                    recv = self.ser.read(self.ser.in_waiting)
                    msg_res = ppx_region_msg_t()
                    msg_res.id = MCB_DEV_ID
                    return self.lib.ppx_com_region_parse((c_uint8 * len(recv))(*recv), len(recv), byref(msg_res)) == 1
        except:
            pass
        return False

    # --- 初始化 ---
    def initialize(self):
        print("🔄 初始化系统...")
        with self.lock:
            # 1. 读版本
            if not self._send_cmd(0x01, REG_HW_VERSION, 0, True):
                print("❌ 无法读取版本，请检查连接")
                return False
            print(f"   HW Ver: {self.g_data.hw_version}")

            # 2. 切模式
            if not self._send_cmd(0x01, REG_RUN_MODE, 0, True): return False
            if self.g_data.run_mode != MODE_TST:
                print("   切换至测试模式...")
                self._send_cmd(0x03, REG_RUN_MODE, MODE_TST, True)

            # 3. 开权限
            print("   激活控制权限...")
            self._send_cmd(0x03, REG_DAT_SETTING, 0x20, True)

        print("✅ 就绪！")
        return True

    # --- 控制接口 ---
    def set_light(self, left, right):
        self.target_left = left
        self.target_right = right

    def get_status_str(self):
        return f"电压: {self.monitor_data['volt']:.1f}V | 电流: {self.monitor_data['curr']:.1f}A | 错误码: {self.monitor_data['err']}"

    # --- 特效模式 ---
    def mode_sos(self):
        print("\n🆘 正在发送 SOS 信号...")
        pattern = [0.2] * 3 + [0.6] * 3 + [0.2] * 3  # 三短 三长 三短
        for duration in pattern:
            self.set_light(1, 1)
            time.sleep(duration)
            self.set_light(0, 0)
            time.sleep(0.2)
        time.sleep(1)

    def mode_strobe(self):
        print("\n🚨 警用爆闪模式...")
        for _ in range(5):
            # 左闪3下
            for _ in range(3):
                self.set_light(1, 0)
                time.sleep(0.08)
                self.set_light(0, 0)
                time.sleep(0.08)
            # 右闪3下
            for _ in range(3):
                self.set_light(0, 1)
                time.sleep(0.08)
                self.set_light(0, 0)
                time.sleep(0.08)


# ==============================================
# 主界面
# ==============================================
def main():
    mcb = SmartMCB(DLL_PATH, SERIAL_PORT, BAUDRATE)
    if not mcb.ready: return
    if not mcb.initialize(): return

    while True:
        os.system('cls' if os.name == 'nt' else 'clear')  # 刷屏
        print("=" * 50)
        print("      MCB 智能控制终端 V4.0 (旗舰版)")
        print("=" * 50)
        print(f"📊 实时状态: {mcb.get_status_str()}")
        print("-" * 50)
        print(" [1] 左转灯 (常亮)    [5] 流水灯演示")
        print(" [2] 右转灯 (常亮)    [6] 警用爆闪 (Strobe)")
        print(" [3] 双闪警示 (Hazard)  [7] SOS 求救信号")
        print(" [4] 关灯 (OFF)       [0] 退出程序")
        print("=" * 50)

        choice = input("请输入指令 > ")

        if choice == '1':
            mcb.set_light(1, 0)
        elif choice == '2':
            mcb.set_light(0, 1)
        elif choice == '3':
            mcb.set_light(1, 1)
        elif choice == '4':
            mcb.set_light(0, 0)
        elif choice == '5':
            print("🌊 流水灯演示中...")
            for _ in range(3):
                mcb.set_light(1, 0);
                time.sleep(0.3)
                mcb.set_light(0, 0);
                time.sleep(0.1)
                mcb.set_light(0, 1);
                time.sleep(0.3)
                mcb.set_light(0, 0);
                time.sleep(0.1)
        elif choice == '6':
            mcb.mode_strobe()
        elif choice == '7':
            mcb.mode_sos()
        elif choice == '0':
            break
        else:
            pass

        # 简单延时防止刷屏太快
        if choice in ['5', '6', '7']: input("\n按回车键继续...")

    mcb.close()
    print("程序已退出")


if __name__ == "__main__":
    main()