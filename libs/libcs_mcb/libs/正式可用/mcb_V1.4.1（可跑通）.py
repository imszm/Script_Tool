# -*- coding: utf-8 -*-
import ctypes
from ctypes import *
import serial
import platform
import time

print("控制器（Region）最小测试工具")
print("=" * 60)
print(f"Python版本: {platform.python_version()}")
print(f"操作系统: {platform.system()} {platform.release()}")

# ---------------- 常量 ----------------
PPX_SW_VER_SIZE = 20
PPX_MODEL_SIZE  = 8
PPX_SN_SIZE     = 26

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

class ppx_region_data_t(Structure):
    _pack_ = 1
    _fields_ = [
        ("id_num",        c_uint8),
        ("model",         c_uint8 * PPX_MODEL_SIZE),
        ("serial_num",    c_uint8 * PPX_SN_SIZE),
        ("hw_version",    c_uint16),
        ("sw_version",    c_uint8 * PPX_SW_VER_SIZE),
        ("rim_state",     c_uint8),
        ("mcu_errcode",   c_uint32),
        ("ctrl_model",    c_uint8),
        ("speed_ref",     c_int16),
        ("motor_speed",   c_int16),
        ("bus_voltage",   c_uint16),
        ("bus_current",   c_uint16),
        ("phase_current_a", c_int16),
        ("phase_current_b", c_int16),
        ("phase_current_c", c_int16),
        ("hall_state",    c_uint8),
        ("pi_vq",         c_int16),
        ("pi_iq",         c_int16),
        ("brake_state",   c_uint8),
        ("imu_pitch",     c_int16),
        ("imu_roll",      c_int16),
        ("imu_acc",       c_uint8),
        ("brake_mileage", c_uint8),
        ("motor_angle",   c_int32),
        ("single_mileage", c_uint32),
        ("angular_speed", c_int16),
        ("rt_setting",    c_uint16),
        ("run_mode",      c_uint8),
        ("gear",          c_uint8),
        ("target_speed",  c_int16),
        ("rated_voltage", c_uint16),
        ("rated_current", c_uint16),
        ("max_voltage",   c_uint16),
        ("min_voltage",   c_uint16),
        ("acceration",    c_uint32),
        ("dat_setting",   c_uint32),
        ("rsvd_data",     c_uint32),
    ]

# ---------------- 控制类 ----------------
class RegionController:
    REG_MAP = {
        "model":       (0x01, 8,   "str"),
        "serial_num":  (0x02, 26,  "str"),
        "hw_version":  (0x03, 2,   "uint16"),
        "sw_version":  (0x04, 20,  "str"),
        "rim_state":   (0x05, 1,   "uint8"),
        "motor_speed": (0x09, 2,   "int16"),
        "bus_voltage": (0x0A, 2,   "uint16"),
        "bus_current": (0x0B, 2,   "uint16"),
        "run_mode":    (0x1B, 1,   "uint8"),
    }

    def __init__(self, dll_path, port="COM4", baudrate=115200):
        self.ser = serial.Serial(port, baudrate, timeout=1)
        print(f"[DEBUG] 串口已连接: {port}, 波特率: {baudrate}")

        self.dll = ctypes.CDLL(dll_path)
        print(f"[DEBUG] 成功加载 DLL: {dll_path}")

        # 配置函数原型
        self.dll.ppx_com_region_format.argtypes = [c_int, POINTER(ppx_region_msg_t), c_void_p]
        self.dll.ppx_com_region_format.restype  = c_uint16
        self.dll.ppx_com_region_parse.argtypes  = [POINTER(c_uint8), c_uint8, POINTER(ppx_region_msg_t)]
        self.dll.ppx_com_region_parse.restype   = c_int
        print("[DEBUG] DLL 函数原型配置完成")

        # DLL 全局变量
        self.g_region_data = ppx_region_data_t.in_dll(self.dll, "g_ppx_region_data")
        print("[DEBUG] 成功映射 g_ppx_region_data")

    def send_and_recv(self, frame, timeout=1):
        self.ser.write(frame)
        time.sleep(0.1)
        return self.ser.read(128)

    def read_register(self, name):
        if name not in self.REG_MAP:
            print(f"[ERROR] 未知寄存器 {name}")
            return

        reg_addr, reg_nums, dtype = self.REG_MAP[name]

        msg = ppx_region_msg_t()
        msg.id = 0x20   # MCB ID
        msg.cmd = 0x01  # READ
        msg.msg_type = 0
        msg.reg_addr = reg_addr
        msg.reg_nums = reg_nums

        buf = (c_uint8 * 128)()
        length = self.dll.ppx_com_region_format(0, byref(msg), buf)
        frame = bytes(buf[:length])
        print(f"[DEBUG] 发送读取 {name} (0x{reg_addr:02X}, {reg_nums}字节) 帧:", frame.hex(" "))

        rsp = self.send_and_recv(frame)
        if not rsp:
            print("[ERROR] 未收到响应")
            return

        print("[DEBUG] 接收到数据:", rsp.hex(" "))
        arr = (c_uint8 * len(rsp))(*rsp)
        ret = self.dll.ppx_com_region_parse(arr, len(rsp), byref(msg))
        print(f"[DEBUG] ppx_com_region_parse 返回: {ret}")

        if ret != 1:
            print("[ERROR] 解析失败")
            return

        # 根据类型解析字段
        val = None
        if name == "model":
            raw = bytes(self.g_region_data.model)
            val = raw.split(b'\x00', 1)[0].decode("ascii", errors="ignore")
        elif name == "serial_num":
            raw = bytes(self.g_region_data.serial_num)
            val = raw.split(b'\x00', 1)[0].decode("ascii", errors="ignore")
        elif name == "hw_version":
            val = self.g_region_data.hw_version
        elif name == "sw_version":
            raw = bytes(self.g_region_data.sw_version)
            val = raw.split(b'\x00', 1)[0].decode("ascii", errors="ignore")
        elif name == "rim_state":
            val = self.g_region_data.rim_state
        elif name == "motor_speed":
            val = self.g_region_data.motor_speed
        elif name == "bus_voltage":
            val = self.g_region_data.bus_voltage
        elif name == "bus_current":
            val = self.g_region_data.bus_current
        elif name == "run_mode":
            val = self.g_region_data.run_mode

        print(f"[INFO] {name} = {val}")

    def close(self):
        self.ser.close()
        print("[DEBUG] 串口已关闭")

# ---------------- main ----------------
if __name__ == "__main__":
    dll_path = r"C:\Users\szm21\Downloads\Script\libs\libcs_mcb\libs\ppx_region.dll"
    ctl = RegionController(dll_path)

    for reg in ["model", "serial_num", "hw_version", "sw_version",
                "rim_state", "motor_speed", "bus_voltage", "bus_current", "run_mode"]:
        print(f"\n[INFO] 读取 {reg} ...")
        ctl.read_register(reg)

    ctl.close()
