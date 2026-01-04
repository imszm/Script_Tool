import ctypes
from ctypes import c_uint8, c_uint16, c_int16, c_uint32, c_int32, c_int, Structure

# ==== 对应 C 结构体 ====
class PpxRegionExcp(Structure):
    _fields_ = [
        ("parse_status", c_uint8),
        ("cmd_status", c_uint8),
        ("data_status", c_uint8),
    ]

class PpxRegionMsg(Structure):
    _fields_ = [
        ("id", c_uint8),          # dev id
        ("cmd", c_uint8),         # write/read cmd
        ("msg_type", c_uint8),    # master 用
        ("reg_addr", c_uint8),    # region data addr
        ("reg_nums", c_uint8),    # region data number
        ("reg_excp", PpxRegionExcp),
    ]

class PpxRegionData(Structure):
    _pack_ = 1
    _fields_ = [
        ("id_num", c_uint8),
        ("model", c_uint8 * 16),        # 假设 PPX_MODEL_SIZE = 16
        ("serial_num", c_uint8 * 16),   # 假设 PPX_SN_SIZE = 16
        ("hw_version", c_uint16),
        ("sw_version", c_uint8 * 8),    # 假设 PPX_SW_VER_SIZE = 8
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

# ==== 加载 DLL ====
dll_path = r"C:\Test_ziliao\libs\libcs_mcb\libs\ppx_region.dll"
dll = ctypes.CDLL(dll_path)

# ==== 函数原型 ====
dll.ppx_com_region_parse.argtypes = [ctypes.POINTER(c_uint8), c_uint8, ctypes.POINTER(PpxRegionMsg)]
dll.ppx_com_region_parse.restype = c_int

dll.ppx_com_region_format.argtypes = [c_uint16, ctypes.POINTER(PpxRegionMsg), ctypes.c_void_p]
dll.ppx_com_region_format.restype = c_uint16

# ==== 全局变量 g_ppx_region_data ====
g_ppx_region_data = PpxRegionData.in_dll(dll, "g_ppx_region_data")

print("✅ DLL 初始化完成")
