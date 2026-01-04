# -*- coding: utf-8 -*-
import ctypes
from ctypes import *
import serial
import platform
import time

# ---------------- 配置 ----------------
REGION_PROTOCOL_DLL = r"C:\Test_ziliao\libs\libcs_mcb\libs\ppx_region"
SERIAL_PORT = "COM8"
BAUDRATE = 115200
DEBUG_MODE = True

# ---------------- 常量 ----------------
PPX_MSG_WRITE = 0x03
PPX_RUN_MODE_REG = 0x1B  # run_mode 寄存器地址
PPX_PARSE_SUCCESS = 1    # DLL 返回成功值（ppx_packet_status_t PPX_TRUE == 1）

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
        ("model", c_uint8 * 8),
        ("serial_num", c_uint8 * 26),
        ("hw_version", c_uint16),
        ("sw_version", c_uint8 * 20),
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

# ---------------- 全局变量 ----------------
g_ppx_region_data = ppx_region_data_t()

# ---------------- 协议类 ----------------
class RegionProtocol:
    def __init__(self, dll_path, serial_port, baudrate):
        self.dll_loaded = False
        self.serial_connected = False

        try:
            self.region_lib = cdll.LoadLibrary(dll_path)
            self.dll_loaded = True
            if DEBUG_MODE:
                print(f"[DEBUG] 成功加载控制器DLL: {dll_path}")
        except Exception as e:
            print(f"[ERROR] 加载DLL失败: {e}")

        try:
            self.serial_port = serial.Serial(
                port=serial_port,
                baudrate=baudrate,
                timeout=0.1,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
            )
            self.serial_connected = True
            if DEBUG_MODE:
                print(f"[DEBUG] 串口已连接: {serial_port}, 波特率: {baudrate}")
        except Exception as e:
            print(f"[ERROR] 串口连接失败: {e}")

        if self.dll_loaded:
            self.region_lib.ppx_com_region_format.argtypes = [
                c_int, POINTER(ppx_region_msg_t), c_void_p
            ]
            self.region_lib.ppx_com_region_format.restype = c_uint16

            self.region_lib.ppx_com_region_parse.argtypes = [
                POINTER(c_uint8), c_uint8, POINTER(ppx_region_msg_t)
            ]
            self.region_lib.ppx_com_region_parse.restype = c_int

            if DEBUG_MODE:
                print("[DEBUG] DLL 函数原型配置完成")

    def __del__(self):
        self.close()

    def close(self):
        if hasattr(self, "serial_port") and self.serial_port.is_open:
            self.serial_port.close()
            if DEBUG_MODE:
                print("[DEBUG] 串口已关闭")

    def _debug_print(self, msg: str, is_error: bool = False):
        if not DEBUG_MODE:
            return
        if is_error:
            print(f"[ERROR] {msg}")
        else:
            print(f"[DEBUG] {msg}")

    def _print_region_msg(self, msg: ppx_region_msg_t):
        print(
            f"  ID=0x{msg.id:02X}, CMD=0x{msg.cmd:02X}, MSG_TYPE=0x{msg.msg_type:02X}, "
            f"REG_ADDR=0x{msg.reg_addr:02X}, REG_NUMS={msg.reg_nums}"
        )
        print(
            f"  异常状态: parse={msg.reg_excp.parse_status}, "
            f"cmd={msg.reg_excp.cmd_status}, data={msg.reg_excp.data_status}"
        )

    @staticmethod
    def extract_frames(buf: bytes):
        """从原始缓冲区提取所有合法帧"""
        frames = []
        start = 0
        while True:
            try:
                start = buf.index(0xA5, start)
                end = buf.index(0x55, start + 1)
                frame = buf[start:end + 1]
                frames.append(frame)
                start = end + 1
            except ValueError:
                break
        return frames

    def _receive_and_parse(self, timeout=3):
        if not self.serial_connected:
            self._debug_print("串口未连接", is_error=True)
            return

        start = time.time()
        buf = bytearray()

        while time.time() - start < timeout:
            data = self.serial_port.read(64)
            if data:
                buf.extend(data)
                start = time.time()
            else:
                break

        if not buf:
            self._debug_print("超时未接收到 MCU 响应", is_error=True)
            return

        self._debug_print(f"接收到 {len(buf)} 字节: {buf.hex(' ')}")

        parsed = False
        for frame in self.extract_frames(buf):
            msg_out = ppx_region_msg_t()
            msg_out.id = 0x20  # 自动设置请求 ID 进行解析
            ret = self.region_lib.ppx_com_region_parse((c_uint8 * len(frame))(*frame), len(frame), byref(msg_out))
            self._debug_print(f"ppx_com_region_parse 返回: {ret}")
            if ret == PPX_PARSE_SUCCESS:
                print("[INFO] MCU 响应解析成功")
                self._print_region_msg(msg_out)
                parsed = True
                break

        if not parsed:
            self._debug_print("未找到有效帧，解析失败", is_error=True)

    def print_region_data(self):
        d = g_ppx_region_data
        print("\n全局控制器数据:")
        print(f"  设备ID号: {d.id_num}")
        print(f"  电机速度: {d.motor_speed} rpm")
        print(f"  总线电压: {d.bus_voltage * 0.1:.1f} V")
        print(f"  总线电流: {d.bus_current * 0.1:.1f} A")
        print(f"  角速度: {d.angular_speed * 0.1:.1f} deg/s")
        print(f"  累计里程: {d.single_mileage} m")

    def set_run_mode(self, run_mode: int, gear: int = 0, target_speed: int = 0):
        if not self.dll_loaded:
            self._debug_print("DLL未加载", is_error=True)
            return False

        g_ppx_region_data.run_mode = run_mode
        g_ppx_region_data.gear = gear
        g_ppx_region_data.target_speed = target_speed

        msg = ppx_region_msg_t()
        msg.id = 0x20
        msg.cmd = PPX_MSG_WRITE
        msg.reg_addr = PPX_RUN_MODE_REG
        msg.reg_nums = 3

        buffer = create_string_buffer(256)
        length = self.region_lib.ppx_com_region_format(0, byref(msg), buffer)
        if length > 0:
            data = buffer.raw[:length]
            self._debug_print(f"组包成功，长度={length} 数据={data.hex(' ')}")

            # --- 自解析生成的数据（调试用） ---
            msg_out = ppx_region_msg_t()
            msg_out.id = 0x20
            ret = self.region_lib.ppx_com_region_parse((c_uint8 * length)(*data), length, byref(msg_out))
            self._debug_print(f"ppx_com_region_parse (自解析) 返回: {ret}")
            if ret == PPX_PARSE_SUCCESS:
                self._debug_print("自解析成功")
                self._print_region_msg(msg_out)

            # 发送到 MCU
            if self.serial_connected:
                self.serial_port.reset_input_buffer()
                self.serial_port.write(data)
                self._debug_print("已发送到串口")
                self._receive_and_parse()
            return True
        else:
            self._debug_print("组包失败", is_error=True)
            return False

# ---------------- 主函数 ----------------
def main():
    print("\n控制器（Region）最小测试工具")
    print("=" * 60)
    print(f"Python版本: {platform.python_version()}")
    print(f"操作系统: {platform.system()} {platform.release()}")

    region = RegionProtocol(REGION_PROTOCOL_DLL, SERIAL_PORT, BAUDRATE)
    if region.dll_loaded:
        region.print_region_data()
        print("\n[INFO] 调用 set_run_mode 示例...")
        region.set_run_mode(2, gear=1, target_speed=50)

    region.close()

if __name__ == "__main__":
    main()
