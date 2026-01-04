# -*- coding: utf-8 -*-
"""
MCB 动力控制与诊断系统 (V6.0 全能诊断版)
============================================================
针对“灯亮但电机不转”问题的专项修复。

核心升级：
1. [电压监控] 实时检测母线电压，如果电压过低(0V)，会在界面高亮报警。
2. [故障修复] 新增 'c' 指令：发送清除错误码 (Clear Error) 命令。
3. [状态诊断] 新增刹车状态(Brake)检测，排除刹车断电保护的干扰。
4. [错误解码] 将错误码 Err 显示为十六进制 (Hex)，方便查表。

使用步骤：
1. 运行脚本，观察电压(Volt)是否为 0.0V。如果是，请检查电源连接。
2. 如果电压正常但有错误码，输入 'c' 清除错误。
3. 观察刹车(Brake)状态是否为 0 (正常)。
4. 再尝试加转速。
"""

import ctypes
from ctypes import *
import serial
import time
import os
import threading

# ==============================================
# 核心配置
# ==============================================
DLL_PATH = os.path.join(os.path.dirname(__file__), "../ppx_region.dll")
SERIAL_PORT = "COM9"
BAUDRATE = 460800
MCB_DEV_ID = 0x20

# 寄存器地址
REG_HW_VERSION = 3
REG_ERR_CODE = 6  # 错误码
REG_BUS_VOLT = 10  # 电压
REG_BUS_CURR = 11  # 电流
REG_BRAKE_STATE = 18  # [18] 刹车状态 (重要!)
REG_RT_SETTING = 0x1A  # 灯光/清错
REG_RUN_MODE = 0x1B  # 模式
REG_GEAR = 0x1C  # 档位
REG_TARGET_SPEED = 0x1D  # 速度
REG_ACCELERATION = 0x22  # 加速度
REG_DAT_SETTING = 0x23  # 权限

MODE_TST = 7

# 命令位定义
PPX_CLR_ERRCODE = (1 << 15)  # 0x8000


# ==============================================
# 结构体定义
# ==============================================
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
        ("ctrl_model", c_uint8), ("speed_ref", c_int16), ("motor_speed", c_int16),
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
# 智能控制器类
# ==============================================
class MotorController(object):
    def __init__(self, dll_path, port, baudrate):
        self.ready = False
        self.ser = None
        self.lock = threading.Lock()

        self.target_left = 0
        self.target_right = 0
        self.target_speed = 0
        self.do_clear_err = False  # 触发清除错误标志

        self.running = True
        self.monitor_data = {"volt": 0.0, "curr": 0.0, "err": 0, "brake": 0}

        # 加载DLL
        if not os.path.exists(dll_path): return
        try:
            self.lib = cdll.LoadLibrary(dll_path)
            self.lib.ppx_com_region_format.argtypes = [c_int, POINTER(ppx_region_msg_t), c_void_p]
            self.lib.ppx_com_region_format.restype = c_uint16
            self.lib.ppx_com_region_parse.argtypes = [POINTER(c_uint8), c_uint8, POINTER(ppx_region_msg_t)]
            self.lib.ppx_com_region_parse.restype = c_int
            self.g_data = ppx_region_data_t.in_dll(self.lib, "g_ppx_region_data")
        except:
            return

        # 连接串口
        try:
            self.ser = serial.Serial(port, baudrate, timeout=0.1)
            print(f"✅ 串口已连接: {port}");
            self.ready = True
        except:
            return

        # 启动线程
        threading.Thread(target=self._heartbeat_loop, daemon=True).start()
        threading.Thread(target=self._monitor_loop, daemon=True).start()

    def close(self):
        self.running = False
        self.target_speed = 0
        time.sleep(0.5)
        if self.ser and self.ser.is_open: self.ser.close()

    # --- 线程1: 心跳与控制 ---
    def _heartbeat_loop(self):
        while self.running:
            if self.ready:
                with self.lock:
                    # 1. 组装 RT_SETTING (灯光 + 清错)
                    rt_val = 0
                    if self.target_left: rt_val |= 0x08
                    if self.target_right: rt_val |= 0x04

                    # 如果用户按了 'c'，发送清除错误位
                    if self.do_clear_err:
                        rt_val |= PPX_CLR_ERRCODE  # 0x8000
                        # 仅发送一次高电平脉冲，下次循环自动清零
                        self.do_clear_err = False

                    self._send_cmd(0x03, REG_RT_SETTING, rt_val, nums=1)

                    time.sleep(0.05)

                    # 2. 发送速度
                    self._send_cmd(0x03, REG_TARGET_SPEED, self.target_speed, nums=1)

            time.sleep(0.15)

            # --- 线程2: 状态监控 ---

    def _monitor_loop(self):
        while self.running:
            if self.ready:
                with self.lock:
                    # 电压电流
                    if self._send_cmd(0x01, REG_BUS_VOLT, 0, nums=2, wait_resp=True):
                        self.monitor_data["volt"] = self.g_data.bus_voltage * 0.1
                        self.monitor_data["curr"] = self.g_data.bus_current * 0.1

                    time.sleep(0.05)
                    # 错误码 + 刹车状态
                    if self._send_cmd(0x01, REG_ERR_CODE, 0, nums=1, wait_resp=True):
                        self.monitor_data["err"] = self.g_data.mcu_errcode

                    time.sleep(0.05)
                    if self._send_cmd(0x01, REG_BRAKE_STATE, 0, nums=1, wait_resp=True):
                        self.monitor_data["brake"] = self.g_data.brake_state

            time.sleep(1.0)

    def _send_cmd(self, cmd, reg, val, nums=1, wait_resp=False):
        if cmd == 0x03:
            if reg == REG_RT_SETTING: self.g_data.rt_setting = val
            if reg == REG_RUN_MODE: self.g_data.run_mode = val
            if reg == REG_DAT_SETTING: self.g_data.dat_setting = val
            if reg == REG_TARGET_SPEED: self.g_data.target_speed = val
            if reg == REG_GEAR: self.g_data.gear = val
            if reg == REG_ACCELERATION: self.g_data.acceration = val

        msg = ppx_region_msg_t()
        msg.id, msg.cmd, msg.reg_addr, msg.reg_nums = MCB_DEV_ID, cmd, reg, nums

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

    def initialize(self):
        print("🔄 初始化诊断...")
        with self.lock:
            if not self._send_cmd(0x01, REG_HW_VERSION, 0, wait_resp=True):
                print("❌ 连接失败");
                return False
            print(f"   HW Ver: {self.g_data.hw_version}")

            # 模式设置
            self._send_cmd(0x03, REG_RUN_MODE, MODE_TST, wait_resp=False)
            self._send_cmd(0x03, REG_DAT_SETTING, 0x20, wait_resp=False)
            # 参数设置
            self._send_cmd(0x03, REG_GEAR, 1, nums=1)
            self._send_cmd(0x03, REG_ACCELERATION, 1000, nums=2)  # 增加加速度到1000

        print("✅ 诊断就绪！")
        return True

    def clear_error(self):
        print("   -> 发送清除错误指令...")
        self.do_clear_err = True  # 通知线程发送

    def set_speed(self, rpm):
        self.target_speed = int(rpm)

    def set_light(self, l, r):
        self.target_left = l;
        self.target_right = r

    def get_status(self):
        return self.monitor_data


# ==============================================
# 主交互
# ==============================================
def main():
    mcb = MotorController(DLL_PATH, SERIAL_PORT, BAUDRATE)
    if not mcb.ready: return
    if not mcb.initialize(): return

    speed = 0
    while True:
        os.system('cls' if os.name == 'nt' else 'clear')
        st = mcb.get_status()

        # 状态分析
        volt_status = "❌ 异常 (0V)" if st['volt'] < 5.0 else "✅ 正常"
        brake_status = "🔴 刹车中 (电机锁定)" if st['brake'] > 0 else "🟢 松开"
        err_hex = f"0x{st['err']:06X}"
        err_msg = "✅ 无故障" if st['err'] == 0 else f"⚠️ 故障码 {err_hex} (可能需清除)"

        if st['err'] == 0x200000 and st['volt'] < 5.0:
            err_msg += " -> [欠压保护]"

        print("=" * 60)
        print("      MCB 全能诊断台 V6.0")
        print("=" * 60)
        print(f"🔋 电压: {st['volt']:.1f}V [{volt_status}]  | ⚡ 电流: {st['curr']:.1f}A")
        print(f"🛑 刹车: {st['brake']} [{brake_status}]")
        print(f"🔧 错误: {st['err']} ({err_hex}) -> {err_msg}")
        print("-" * 60)
        print(f"⚙️  目标转速: {mcb.target_speed} RPM")
        print("=" * 60)
        print(" [c] 清除错误 (Clear Error)  <-- 如果有错误码，请先按这个")
        print(" [w/s] 加/减速 (+100/-100)")
        print(" [e] 设定转速")
        print(" [SPACE] 急停")
        print(" [1/2/3/4] 灯光控制")
        print(" [0] 退出")
        print("=" * 60)

        choice = input("指令 > ").lower()

        if choice == 'c':
            mcb.clear_error()
        elif choice == '1':
            mcb.set_light(1, 0)
        elif choice == '2':
            mcb.set_light(0, 1)
        elif choice == '3':
            mcb.set_light(1, 1)
        elif choice == '4':
            mcb.set_light(0, 0)
        elif choice == 'w':
            speed += 100
            mcb.set_speed(speed)
        elif choice == 's':
            speed -= 100
            mcb.set_speed(speed)
        elif choice == ' ':
            speed = 0
            mcb.set_speed(0)
        elif choice == 'e':
            try:
                mcb.set_speed(int(input("RPM: ")))
            except:
                pass
        elif choice == '0':
            break

    mcb.close()


if __name__ == "__main__":
    main()