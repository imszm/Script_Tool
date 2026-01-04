import ctypes
from ctypes import *
import serial
import time
import platform
import sys

# ==============================================
# 配置区域 - 根据实际情况修改这些参数
# ==============================================
BLE_PROTOCOL_DLL = r"E:\AutoTest\ppx_ble.dll"  # BLE协议DLL文件名
SERIAL_PORT = "COM97"                      # 串口端口
BAUDRATE = 460800                         # 串口波特率
DEBUG_MODE = True                         # 调试模式，打印更多信息

# ==============================================
# 协议常量定义
# ==============================================
# 设备ID
PPX_ID_BLE = 0x60

# 命令类型
PPX_MSG_READ = 0x01
PPX_MSG_WRITE = 0x03

# 命令类别
class PpxCmdType:
    REQ = 0  # 请求命令
    RSP = 1  # 响应命令

# 寄存器地址
PPX_BLE_LED_MSG_REG = 0x08

# 解析状态
PPX_PARSE_SUCCESS = 1
PPX_PARSE_FAILURE = 0

# ==============================================
# 结构体定义 (必须与DLL中的定义完全一致)
# ==============================================
class ppx_ble_msg_t(Structure):
    _fields_ = [
        ("id", c_uint8),      # 设备ID
        ("cmd", c_uint8),     # 写/读命令
        ("reg_addr", c_uint8), # BLE数据地址
        ("reg_nums", c_uint8)  # BLE寄存器数量
    ]

class ppx_led_msg_t(Structure):
    _fields_ = [
        ("screen_on", c_uint32, 1),      # 显示开关: 1开, 0关
        ("brightness", c_uint32, 3),     # 亮度级别 0-7
        ("blink_period", c_uint32, 4),   # 闪烁周期: N * 200ms
        ("blink_duty", c_uint32, 4),     # 闪烁占空比
        ("blink_en", c_uint32, 8),       # 闪烁使能
        ("err_flag", c_uint32, 2),       # 错误码标志
        ("err_code", c_uint32, 4),       # 错误码: 0-F
        ("digital", c_uint32, 7),        # 电池SOC: 0-100
        ("logo", c_uint32, 2),           # LOGO: 0关, 1白, 2红
        ("rim_state", c_uint32, 2),      # 护盾: 0关, 1白, 2绿
        ("rdygo", c_uint32, 2),          # Ready Go: 0关, 1白, 2红
        ("turn_left", c_uint32, 2),      # 左转向灯: 0关, 1白, 2橙
        ("turn_right", c_uint32, 2),     # 右转向灯: 0关, 1白, 2橙
        ("ring", c_uint32, 2),           # 灯环: 0关, 1蓝, 2红
        ("rsvd_data", c_uint32, 19)      # 保留
    ]

class ppx_ble_data_t(Structure):
    _fields_ = [
        ("id_num", c_uint8),                     # 设备ID号
        ("model", c_uint8 * 8),                  # 型号 (8字节)
        ("serial_num", c_uint8 * 26),            # 序列号 (26字节)
        ("hw_version", c_uint8),                 # 硬件版本
        ("sw_version", c_uint8 * 20),            # 软件版本 (20字节)
        ("status", c_uint32),                    # 状态
        ("ldr_value", c_uint16),                 # 光敏电阻亮度
        ("io_status", c_uint16),                 # IO引脚状态
        ("led_msg", ppx_led_msg_t),              # LED显示信息
        ("card_id", c_uint32),                   # NFC卡ID
        ("dat_setting", c_uint32)                # 数据设置
    ]

# ==============================================
# BLE协议通信类
# ==============================================
class BLEProtocol:
    def __init__(self, dll_path, serial_port, baudrate):
        self.dll_loaded = False
        self.serial_connected = False
        
        # 加载DLL
        self._load_dll(dll_path)
        
        # 初始化串口
        self._init_serial(serial_port, baudrate)
        
        # 检查全局变量
        self._check_global_vars()
    
    def _load_dll(self, dll_path):
        try:
            self.ble_lib = cdll.LoadLibrary(dll_path)
            self.dll_loaded = True
            self._debug_print(f"成功加载DLL: {dll_path}")
            
            # 严格按照给定的函数原型设置
            self.ble_lib.ppx_com_ble_parse.argtypes = [
                POINTER(c_uint8),    # pdata
                c_uint8,             # data_len
                POINTER(ppx_ble_msg_t)  # ble_msg
            ]
            self.ble_lib.ppx_com_ble_parse.restype = c_int
            
            self.ble_lib.ppx_com_ble_format.argtypes = [
                c_int,                # cmd_type
                POINTER(ppx_ble_msg_t), # ble_msg
                c_void_p              # buffer
            ]
            self.ble_lib.ppx_com_ble_format.restype = c_uint16
            
        except Exception as e:
            self._debug_print(f"加载DLL失败: {e}", is_error=True)
            self.dll_loaded = False
    
    def _init_serial(self, port, baudrate):
        try:
            self.serial_port = serial.Serial(
                port=port,
                baudrate=baudrate,
                timeout=1
            )
            self.serial_connected = True
            self._debug_print(f"串口已连接: {port}, 波特率: {baudrate}")
        except Exception as e:
            self._debug_print(f"串口连接失败: {e}", is_error=True)
            self.serial_connected = False
    
    def _check_global_vars(self):
        try:
            self.g_ppx_ble_data = ppx_ble_data_t.in_dll(self.ble_lib, "g_ppx_ble_data")
            self._debug_print("成功获取全局变量 g_ppx_ble_data")
            
            if DEBUG_MODE:
                self._print_ble_data(self.g_ppx_ble_data)
                
        except Exception as e:
            self._debug_print(f"获取全局变量失败: {e}", is_error=True)
    
    def close(self):
        if hasattr(self, 'serial_port') and self.serial_port.is_open:
            self.serial_port.close()
            self.serial_connected = False
            self._debug_print("串口已关闭")
    
    def send_data(self, data):
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
    
    def receive_data(self, timeout=1.0, max_length=256):
        if not self.serial_connected:
            self._debug_print("串口未连接，无法接收数据", is_error=True)
            return None
        
        try:
            start_time = time.time()
            data = bytes()
            
            while time.time() - start_time < timeout:
                if self.serial_port.in_waiting > 0:
                    data += self.serial_port.read(self.serial_port.in_waiting)
                    if len(data) >= max_length:
                        break
            
            if data:
                self._debug_print(f"接收数据: {self._bytes_to_hex(data)}")
                return data
            else:
                self._debug_print("接收数据超时")
                return None
                
        except Exception as e:
            self._debug_print(f"接收数据失败: {e}", is_error=True)
            return None
    
    def parse_data(self, data):
        """ 修改后的解析方法 """
        if not data:
            self._debug_print("无数据可解析", is_error=True)
            return False, None, None
        
        if not self.dll_loaded:
            self._debug_print("DLL未加载，无法解析数据", is_error=True)
            return False, None, None
        
        try:
            ble_msg = ppx_ble_msg_t()
            ble_msg.id = PPX_ID_BLE  # 关键修改：预填充ID
            
            data_len = len(data)
            data_array = (c_uint8 * data_len)(*data)
            
            # 调用解析函数
            result = self.ble_lib.ppx_com_ble_parse(
                cast(data_array, POINTER(c_uint8)),  # pdata
                c_uint8(data_len),                  # data_len
                byref(ble_msg)                       # ble_msg
            )
            
            if result == PPX_PARSE_SUCCESS:
                self._debug_print("数据解析成功:")
                self._print_ble_msg(ble_msg)
                
                # 更新全局变量信息
                self._print_ble_data_changes(ble_msg)
                
                # 解析命令类型 (第一个字节的低4位)
                cmd_type = data[0] & 0x0F
                return True, ble_msg, cmd_type
            else:
                self._debug_print(f"数据解析失败，返回码: {result}", is_error=True)
                return False, None, None
                
        except Exception as e:
            self._debug_print(f"解析数据时发生异常: {e}", is_error=True)
            return False, None, None
    
    def format_data(self, cmd_type, ble_msg):
        if not self.dll_loaded:
            self._debug_print("DLL未加载，无法组包数据", is_error=True)
            return False, None
        
        try:
            buffer_size = 256
            buffer = create_string_buffer(buffer_size)
            
            length = self.ble_lib.ppx_com_ble_format(
                cmd_type,
                byref(ble_msg),
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
    
    def set_led_display(self, screen_on=1, brightness=0, digital=0, 
                        logo=0, rim_state=0, rdygo=0, 
                        turn_left=0, turn_right=0, ring=0):
        if not self.dll_loaded:
            self._debug_print("DLL未加载，无法设置LED", is_error=True)
            return False, None
        
        try:
            # 修改全局变量中的LED显示信息
            self.g_ppx_ble_data.led_msg.screen_on = screen_on
            self.g_ppx_ble_data.led_msg.brightness = brightness
            self.g_ppx_ble_data.led_msg.digital = digital
            self.g_ppx_ble_data.led_msg.logo = logo
            self.g_ppx_ble_data.led_msg.rim_state = rim_state
            self.g_ppx_ble_data.led_msg.rdygo = rdygo
            self.g_ppx_ble_data.led_msg.turn_left = turn_left
            self.g_ppx_ble_data.led_msg.turn_right = turn_right
            self.g_ppx_ble_data.led_msg.ring = ring
            
            self._debug_print("修改后的LED显示信息:")
            self._print_led_msg(self.g_ppx_ble_data.led_msg)
            
            # 准备消息结构体
            ble_msg = ppx_ble_msg_t()
            ble_msg.id = PPX_ID_BLE
            ble_msg.cmd = PPX_MSG_WRITE
            ble_msg.reg_addr = PPX_BLE_LED_MSG_REG
            ble_msg.reg_nums = 1
            
            # 组包并发送
            success, data = self.format_data(PpxCmdType.REQ, ble_msg)
            if success:
                if self.send_data(data):
                    # 等待并接收响应
                    response = self.receive_data()
                    if response:
                        parse_success, parsed_msg, cmd_type = self.parse_data(response)
                        return parse_success, response
            return False, None
                
        except Exception as e:
            self._debug_print(f"设置LED显示时发生异常: {e}", is_error=True)
            return False, None
    
    def read_led_status(self):
        if not self.dll_loaded:
            self._debug_print("DLL未加载，无法读取LED状态", is_error=True)
            return False, None, None
        
        try:
            # 准备消息结构体
            ble_msg = ppx_ble_msg_t()
            ble_msg.id = PPX_ID_BLE
            ble_msg.cmd = PPX_MSG_READ
            ble_msg.reg_addr = PPX_BLE_LED_MSG_REG
            ble_msg.reg_nums = 1
            
            # 组包并发送
            success, data = self.format_data(PpxCmdType.REQ, ble_msg)
            if success:
                if self.send_data(data):
                    # 等待并接收响应
                    response = self.receive_data()
                    if response:
                        parse_success, parsed_msg, cmd_type = self.parse_data(response)
                        return parse_success, response, self.g_ppx_ble_data.led_msg
            return False, None, None
                
        except Exception as e:
            self._debug_print(f"读取LED状态时发生异常: {e}", is_error=True)
            return False, None, None
    
    # ==============================================
    # 辅助方法
    # ==============================================
    def _debug_print(self, message, is_error=False):
        prefix = "[ERROR] " if is_error else "[DEBUG] "
        if is_error or DEBUG_MODE:
            print(prefix + message)
    
    def _bytes_to_hex(self, data):
        return ' '.join(f'{b:02X}' for b in data)
    
    def _print_ble_msg(self, ble_msg):
        print(f"  ID: 0x{ble_msg.id:02X}")
        print(f"  命令: 0x{ble_msg.cmd:02X} ({'写' if ble_msg.cmd == PPX_MSG_WRITE else '读'})")
        print(f"  寄存器地址: 0x{ble_msg.reg_addr:02X}")
        print(f"  寄存器数量: {ble_msg.reg_nums}")
    
    def _print_led_msg(self, led_msg):
        print(f"  屏幕状态: {'开' if led_msg.screen_on else '关'}")
        print(f"  亮度级别: {led_msg.brightness}")
        print(f"  数字显示: {led_msg.digital}")
        print(f"  LOGO状态: {led_msg.logo} (0关, 1白, 2红)")
        print(f"  护盾状态: {led_msg.rim_state} (0关, 1白, 2绿)")
        print(f"  ReadyGo状态: {led_msg.rdygo} (0关, 1白, 2红)")
        print(f"  左转向灯: {led_msg.turn_left} (0关, 1白, 2橙)")
        print(f"  右转向灯: {led_msg.turn_right} (0关, 1白, 2橙)")
        print(f"  灯环状态: {led_msg.ring} (0关, 1蓝, 2红)")
    
    def _print_ble_data(self, ble_data):
        print("\n全局BLE数据结构体内容:")
        print(f"设备ID号: {ble_data.id_num}")
        print(f"硬件版本: {ble_data.hw_version}")
        print(f"状态: 0x{ble_data.status:08X}")
        print(f"光敏电阻亮度: {ble_data.ldr_value}")
        print(f"IO状态: 0x{ble_data.io_status:04X}")
        print(f"NFC卡ID: 0x{ble_data.card_id:08X}")
        print(f"数据设置: 0x{ble_data.dat_setting:08X}")
        print("\nLED显示信息:")
        self._print_led_msg(ble_data.led_msg)
    
    def _print_ble_data_changes(self, ble_msg):
        if ble_msg.reg_addr == PPX_BLE_LED_MSG_REG:
            print("\n全局变量中的LED显示信息已更新:")
            self._print_led_msg(self.g_ppx_ble_data.led_msg)

# ==============================================
# 主程序
# ==============================================
def main():
    print("\nBLE协议通信测试程序")
    print("=" * 50)
    
    # 显示系统信息
    print(f"Python版本: {platform.python_version()}")
    print(f"系统架构: {platform.architecture()[0]}")
    print(f"操作系统: {platform.system()} {platform.release()}")
    
    # 初始化BLE协议通信
    print("\n初始化BLE协议通信...")
    ble = BLEProtocol(BLE_PROTOCOL_DLL, SERIAL_PORT, BAUDRATE)
    
    if not ble.dll_loaded or not ble.serial_connected:
        print("初始化失败，请检查错误信息")
        ble.close()
        sys.exit(1)
    
    try:
        # 示例1: 设置LED显示
        print("\n测试1: 设置LED显示")
        print("=" * 30)
        success, response = ble.set_led_display(
            screen_on=1, 
            brightness=7, 
            digital=88,
            logo=2, 
            rim_state=1, 
            rdygo=1,
            turn_left=2,
            turn_right=2,
            ring=2
        )
        if success:
            print("LED设置成功")
            print(f"响应数据: {ble._bytes_to_hex(response)}")
            
        else:
            print("LED设置失败")
        
        # 示例2: 读取LED状态
        print("\n测试2: 读取LED状态")
        print("=" * 30)
        success, response, led_msg = ble.read_led_status()
        if success:
            print("LED状态读取成功")
            print(f"响应数据: {ble._bytes_to_hex(response)}")
            print("\n当前LED状态:")
            ble._print_led_msg(led_msg)
        else:
            print("LED状态读取失败")
        
    finally:
        ble.close()
        print("\n程序结束")

if __name__ == "__main__":
    main()
