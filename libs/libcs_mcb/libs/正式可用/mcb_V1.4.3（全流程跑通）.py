# -*- coding: utf-8 -*-
import ctypes
from ctypes import *
import serial
import platform
import time

print("MCB 左转向灯控制工具（dat_setting控制版）")
print("=" * 60)
print(f"Python版本: {platform.python_version()}")
print(f"操作系统: {platform.system()} {platform.release()}")

# ---------------- 数据结构 ----------------
class ppx_region_excp_t(Structure):
    _fields_ = [
        ("parse_status", c_uint8),
        ("cmd_status",   c_uint8),
        ("data_status",  c_uint8),
    ]

class ppx_region_msg_t(Structure):
    _fields_ = [
        ("id",        c_uint8),
        ("cmd",       c_uint8),
        ("msg_type",  c_uint8),
        ("reg_addr",  c_uint8),
        ("reg_nums",  c_uint8),
        ("reg_excp",  ppx_region_excp_t),
    ]

# ---------------- 控制类 ----------------
class RegionController:
    def __init__(self, dll_path, port="COM4", baudrate=406800):
        self.ser = serial.Serial(port, baudrate, timeout=1)
        print(f"[DEBUG] 串口已连接: {port}, 波特率: {baudrate}")
        self.dll = ctypes.CDLL(dll_path)
        print(f"[DEBUG] 成功加载 DLL: {dll_path}")

        self.dll.ppx_com_region_format.argtypes = [c_int, POINTER(ppx_region_msg_t), c_void_p]
        self.dll.ppx_com_region_format.restype  = c_uint16
        self.dll.ppx_com_region_parse.argtypes  = [POINTER(c_uint8), c_uint8, POINTER(ppx_region_msg_t)]
        self.dll.ppx_com_region_parse.restype   = c_int
        print("[DEBUG] DLL 函数原型配置完成")

    # ---------------- 写入 dat_setting ----------------
    def write_dat_setting(self, value):
        msg = ppx_region_msg_t()
        msg.id = 0x20
        msg.cmd = 0x02  # 写命令
        msg.msg_type = 0
        msg.reg_addr = 0x23
        msg.reg_nums = 2

        data = (c_uint8 * 2)(value & 0xFF, (value >> 8) & 0xFF)
        buf = (c_uint8 * 128)()
        length = self.dll.ppx_com_region_format(0, byref(msg), buf)
        # 替换数据区
        for i in range(2):
            buf[6 + i] = data[i]
        frame = bytes(buf[:length])

        print(f"[INFO] 写入 dat_setting=0x{value:04X} ({value})")
        print(f"[DEBUG] 最终发送帧 ({length}字节):", frame.hex(" "))
        self.ser.write(frame)
        time.sleep(0.2)
        rsp = self.ser.read(128)
        if rsp:
            print("[DEBUG] 接收到数据:", rsp.hex(" "))
        else:
            print("[WARN] 未收到响应")

    # ---------------- 读取 dat_setting ----------------
    def read_dat_setting(self):
        msg = ppx_region_msg_t()
        msg.id = 0x20
        msg.cmd = 0x01
        msg.msg_type = 0
        msg.reg_addr = 0x23
        msg.reg_nums = 2

        buf = (c_uint8 * 128)()
        length = self.dll.ppx_com_region_format(0, byref(msg), buf)
        frame = bytes(buf[:length])

        print(f"[INFO] 读取 dat_setting 帧 ({length}字节):", frame.hex(" "))
        self.ser.write(frame)
        time.sleep(0.2)
        rsp = self.ser.read(128)
        if rsp:
            print("[DEBUG] 接收到数据:", rsp.hex(" "))
        else:
            print("[WARN] 未收到响应")

    def close(self):
        self.ser.close()
        print("[DEBUG] 串口已关闭")

# ---------------- 主程序 ----------------
if __name__ == "__main__":
    dll_path = r"C:\Users\szm21\Downloads\Script\libs\libcs_mcb\libs\ppx_region.dll"
    ctl = RegionController(dll_path)

    print("\n[INFO] 左转向灯亮起 -> 熄灭 -> 再亮起 测试开始 ...")
    LEFT_LIGHT_BIT = 0x0008

    ctl.write_dat_setting(LEFT_LIGHT_BIT)
    ctl.read_dat_setting()

    time.sleep(1)
    ctl.write_dat_setting(0x0000)
    ctl.read_dat_setting()

    time.sleep(1)
    ctl.write_dat_setting(LEFT_LIGHT_BIT)
    ctl.read_dat_setting()

    print("[INFO] 测试完成: 左转向灯亮起→熄灭→再亮起")
    ctl.close()
