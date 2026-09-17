#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import ctypes
import serial
import time
import sys
import os
import socket
import json
import select
import traceback
import random
import gc
import threading
from collections import deque

# ========== 配置区 ==========
MAX_CYCLES = 1000
TESTED_COM_PORT = "COM28"   # 被测电机串口（原被测脚本）
LOAD_COM_PORT = "COM11"     # 负载电机串口（原负载脚本）
BAUD_RATE = 460800

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DLL_PATH = os.path.join(SCRIPT_DIR, "ppx_region.dll")
LOG_DIR = os.path.join(SCRIPT_DIR, "log")

TESTED_RUN_SPEED = -1000    # 被测电机骑行模式转速（与负载反向对拖）
LOAD_RUN_SPEED = 1000       # 负载电机推行模式转速
LOAD_RUN_CURRENT = 20       # 负载电机推行电流(A)
MOS_TEMP_LIMIT = 85         # MOSFET 超温阈值(℃)
MOTOR_TEMP_LIMIT = 150      # 电机超温阈值(℃)
MOS_TEMP_RECOVER = 70       # MOSFET 恢复阈值(℃，回差防止抖动)
MOTOR_TEMP_RECOVER = 120    # 电机恢复阈值(℃)

# ========== 协议常量 ==========
PPX_ID_MCB = 0x20
PPX_CMD_REQ = 0x00
PPX_MSG_WRITE = 0x03
PPX_MSG_MULTIWR = 0x04
PPX_MSG_READ = 0x02

PPX_RUN_MODE_REG = 25
PPX_RT_SETTING_REG = 24
PPX_BRAKE_STATE_REG = 22
PPX_TEMP_REG = 12
PPX_TARGET_SPEED_REG = 27
PPX_TARGET_ACCEL_REG = 28
PPX_TARGET_CUR_REG = 29
PPX_MCU_ERRCODE_REG = 5
PPX_MOTOR_SPEED_REG = 6

PPX_MODE_IDLE = 0
PPX_MODE_PWR_PUSH = 4
PPX_MODE_RUNNING = 2
PPX_MODE_CLUTCH_OPEN = 12
PPX_MODE_LOCK = 3

PPX_CLR_ERRCODE = (1 << 15)
PPX_BRAKE_CLOSED = 0
PPX_BRAKE_OPENING = 1
PPX_BRAKE_OPENED = 2

# ========== IPC通信 ==========
IPC_HOST = '127.0.0.1'
IPC_PORT = 12345

# ========== 结构体 ==========
class PpxRegionExcp(ctypes.Structure):
    _fields_ = [("parse_status", ctypes.c_uint8), ("cmd_status", ctypes.c_uint8), ("data_status", ctypes.c_uint8)]

class PpxRegionMsg(ctypes.Structure):
    _fields_ = [("id", ctypes.c_uint8), ("cmd", ctypes.c_uint8), ("msg_type", ctypes.c_uint8),
                ("reg_addr", ctypes.c_uint8), ("reg_nums", ctypes.c_uint8), ("reg_excp", PpxRegionExcp)]

PPX_MODEL_SIZE = 8
PPX_SN_SIZE = 26
PPX_SW_VER_SIZE = 32

class PpxRegionData(ctypes.Structure):
    _pack_ = 1
    _fields_ = [
        ("id_num", ctypes.c_uint8), ("model", ctypes.c_uint8 * PPX_MODEL_SIZE),
        ("serial_num", ctypes.c_uint8 * PPX_SN_SIZE), ("hw_version", ctypes.c_uint16),
        ("sw_version", ctypes.c_uint8 * PPX_SW_VER_SIZE), ("mcu_errcode", ctypes.c_uint32),
        ("motor_speed", ctypes.c_int16), ("bus_voltage", ctypes.c_uint16), ("bus_current", ctypes.c_uint16),
        ("rim_state", ctypes.c_uint8), ("ctrl_model", ctypes.c_uint8), ("speed_ref", ctypes.c_int16),
        ("mosfet_temp", ctypes.c_int16), ("motor_temp", ctypes.c_int16), ("motor_limit_flg", ctypes.c_uint8),
        ("motor_vq", ctypes.c_int32), ("motor_iq", ctypes.c_int16), ("motor_cali_state", ctypes.c_uint8),
        ("motor_cali_res", ctypes.c_uint16), ("motor_cali_ld", ctypes.c_uint16), ("motor_cali_lq", ctypes.c_uint16),
        ("motor_cali_bemf", ctypes.c_uint16), ("brake_state", ctypes.c_uint8), ("single_mileage", ctypes.c_uint32),
        ("rt_setting", ctypes.c_uint16), ("run_mode", ctypes.c_uint8), ("road_cond", ctypes.c_uint8),
        ("target_speed", ctypes.c_int16), ("target_accel", ctypes.c_uint16), ("target_current", ctypes.c_int16),
        ("rated_voltage", ctypes.c_uint16), ("rated_current", ctypes.c_uint16), ("dat_setting", ctypes.c_uint32),
        ("reserved_data", ctypes.c_uint32)
    ]

class PpxRegionCtrl(ctypes.Structure):
    _fields_ = [("msg", PpxRegionMsg), ("data", PpxRegionData)]

# ========== 电机控制类（两个脚本共用，label 用于区分打印） ==========
class MotorController:
    def __init__(self, com_port, dll_path, label):
        self.com_port = com_port
        self.label = label
        self.ser = None
        try:
            self.dll = ctypes.CDLL(dll_path, winmode=0)
        except OSError as e:
            raise RuntimeError(f"{label}加载DLL失败: {dll_path} ({e})")
        self.dll.ppx_com_region_format.argtypes = [ctypes.c_int, ctypes.POINTER(PpxRegionCtrl), ctypes.c_void_p]
        self.dll.ppx_com_region_format.restype = ctypes.c_uint16
        self.dll.ppx_com_region_parse.argtypes = [ctypes.POINTER(ctypes.c_uint8), ctypes.c_uint16, ctypes.POINTER(PpxRegionCtrl)]
        self.dll.ppx_com_region_parse.restype = ctypes.c_int
        self._open_serial()

    def _open_serial(self):
        for attempt in range(3):
            try:
                if self.ser and self.ser.is_open:
                    self.ser.close()
                self.ser = serial.Serial(self.com_port, BAUD_RATE, timeout=0.3)
                print(f"{self.label}串口 {self.com_port} 已打开")
                return True
            except Exception as e:
                print(f"{self.label}打开串口 {self.com_port} 失败 (尝试 {attempt+1}/3): {e}")
                time.sleep(0.5)
        return False

    def _send_recv(self, ctrl):
        for retry in range(3):
            try:
                if self.ser is None or not self.ser.is_open:
                    if not self._open_serial():
                        return None
                buf = (ctypes.c_uint8 * 512)()
                length = self.dll.ppx_com_region_format(PPX_CMD_REQ, ctypes.byref(ctrl), buf)
                if length <= 0:
                    return None
                self.ser.reset_input_buffer()
                self.ser.reset_output_buffer()
                self.ser.write(bytes(buf[:length]))
                self.ser.flush()
                time.sleep(0.05)
                if self.ser.in_waiting == 0:
                    return None
                raw = self.ser.read(self.ser.in_waiting)
                ret = self.dll.ppx_com_region_parse((ctypes.c_uint8 * len(raw)).from_buffer_copy(raw), len(raw), ctypes.byref(ctrl))
                return ctrl if ret == 1 else None
            except (serial.SerialException, PermissionError, OSError) as e:
                print(f"{self.label}串口通信异常 (重试 {retry+1}/3): {e}")
                self.ser = None
                time.sleep(0.2)
        return None

    def clear_error(self):
        ctrl = PpxRegionCtrl()
        ctrl.msg.id = PPX_ID_MCB
        ctrl.msg.cmd = PPX_MSG_WRITE
        ctrl.msg.reg_addr = PPX_RT_SETTING_REG
        ctrl.msg.reg_nums = 1
        ctrl.data.rt_setting = PPX_CLR_ERRCODE
        return self._send_recv(ctrl) is not None

    def set_run_mode(self, mode):
        ctrl = PpxRegionCtrl()
        ctrl.msg.id = PPX_ID_MCB
        ctrl.msg.cmd = PPX_MSG_WRITE
        ctrl.msg.reg_addr = PPX_RUN_MODE_REG
        ctrl.msg.reg_nums = 1
        ctrl.data.run_mode = mode
        return self._send_recv(ctrl) is not None

    def set_speed_params(self, speed, accel, current):
        ctrl = PpxRegionCtrl()
        ctrl.msg.id = PPX_ID_MCB
        ctrl.msg.cmd = PPX_MSG_MULTIWR
        ctrl.msg.reg_addr = PPX_TARGET_SPEED_REG
        ctrl.msg.reg_nums = 3
        ctrl.data.target_speed = speed
        ctrl.data.target_accel = accel
        ctrl.data.target_current = current
        return self._send_recv(ctrl) is not None

    def read_temperature(self):
        ctrl = PpxRegionCtrl()
        ctrl.msg.id = PPX_ID_MCB
        ctrl.msg.cmd = PPX_MSG_READ
        ctrl.msg.reg_addr = PPX_TEMP_REG
        ctrl.msg.reg_nums = 2
        result = self._send_recv(ctrl)
        if result:
            return result.data.mosfet_temp, result.data.motor_temp
        return None, None

    def read_error_code(self):
        ctrl = PpxRegionCtrl()
        ctrl.msg.id = PPX_ID_MCB
        ctrl.msg.cmd = PPX_MSG_READ
        ctrl.msg.reg_addr = PPX_MCU_ERRCODE_REG
        ctrl.msg.reg_nums = 1
        result = self._send_recv(ctrl)
        if result:
            return result.data.mcu_errcode
        return None

    def read_motor_speed(self):
        ctrl = PpxRegionCtrl()
        ctrl.msg.id = PPX_ID_MCB
        ctrl.msg.cmd = PPX_MSG_READ
        ctrl.msg.reg_addr = PPX_MOTOR_SPEED_REG
        ctrl.msg.reg_nums = 1
        result = self._send_recv(ctrl)
        if result:
            return result.data.motor_speed
        return None

    def read_brake_state(self):
        ctrl = PpxRegionCtrl()
        ctrl.msg.id = PPX_ID_MCB
        ctrl.msg.cmd = PPX_MSG_READ
        ctrl.msg.reg_addr = PPX_BRAKE_STATE_REG
        ctrl.msg.reg_nums = 1
        result = self._send_recv(ctrl)
        return result.data.brake_state if result else -1

    def close(self):
        if self.ser and self.ser.is_open:
            self.ser.close()
            print(f"{self.label}串口已关闭")

# ========== 日志：控制台输出同时写入日志文件（文件内每行带时间戳） ==========
class Tee:
    def __init__(self, filename):
        self.file = open(filename, 'w', encoding='utf-8')
        self.stdout = sys.__stdout__
        self._lock = threading.Lock()
        self._line_buf = ""

    def write(self, data):
        # 加锁避免两个线程的输出在字符级撕裂
        with self._lock:
            try:
                self.stdout.write(data)
            except Exception:
                pass
            try:
                self._line_buf += data
                while '\n' in self._line_buf:
                    line, self._line_buf = self._line_buf.split('\n', 1)
                    ts = time.strftime('%H:%M:%S')
                    self.file.write(f"[{ts}] {line}\n")
                self.file.flush()
            except Exception:
                pass

    def flush(self):
        with self._lock:
            try:
                self.stdout.flush()
            except Exception:
                pass
            try:
                if self._line_buf:
                    ts = time.strftime('%H:%M:%S')
                    self.file.write(f"[{ts}] {self._line_buf}")
                    self._line_buf = ""
                self.file.flush()
            except Exception:
                pass

    def close(self):
        self.flush()
        try:
            self.file.close()
        except Exception:
            pass

# ========== 双端公共基类：IPC收发、温度、超温、串口指令重试 ==========
class PeerRunnerBase:
    def __init__(self, server_ready, peer_done):
        self.server_ready = server_ready
        self.peer_done = peer_done  # 对端线程结束标志（任一方结束后另一方同步退出）
        self.own_mos_temp = 0
        self.own_motor_temp = 0
        self.other_mos_temp = 0
        self.other_motor_temp = 0
        self.other_temp_seen = False   # 是否收到过对端温度心跳（冷却恢复判断前提）
        self.overheat_stopped = False
        self.stop_received = False
        self.skip_round = False
        self.motor = None
        self.sock = None            # 服务端为已接受的客户端连接，客户端为已连接套接字
        self._recv_buf = b""        # TCP拆包缓冲
        self._payload_queue = deque()  # 一次recv可能含多条消息，逐个返回
        self.fault_rounds = 0       # 错误码清除后仍存在的连续轮数（锁存故障检测）

    def send_status(self, status_dict):
        if self.sock is None:
            return
        try:
            status_dict['mos_temp'] = self.own_mos_temp
            status_dict['motor_temp'] = self.own_motor_temp
            data = json.dumps(status_dict).encode('utf-8') + b'\n'
            self.sock.sendall(data)
        except Exception:
            pass

    def send_status_multiple(self, status_dict, count=3, interval=0.1):
        for _ in range(count):
            self.send_status(status_dict)
            time.sleep(interval)

    def recv_status(self, timeout=0.01):
        """返回下一条完整消息；带拆包缓冲与消息队列，不丢包"""
        if self._payload_queue:
            return self._payload_queue.popleft()
        if self.sock is None:
            return None
        try:
            ready, _, _ = select.select([self.sock], [], [], timeout)
            if ready:
                data = self.sock.recv(4096)
                if not data:
                    print("对端连接已断开")
                    self.skip_round = True
                    return None
                self._recv_buf += data
                while b'\n' in self._recv_buf:
                    line, self._recv_buf = self._recv_buf.split(b'\n', 1)
                    if not line.strip():
                        continue
                    try:
                        payload = json.loads(line.decode('utf-8'))
                    except Exception:
                        continue
                    if payload.get('stop', False):
                        self.stop_received = True
                        self.skip_round = True
                        self.stop_motor_now()  # 立即下发停转指令，防止对端已停止而本端电机继续运转
                        print("收到停止信号，本轮将跳过")
                        continue
                    if payload.get('overheat', False):
                        if not self.overheat_stopped:
                            print("收到超温信号，立即停止电机")
                        self.overheat_stopped = True
                        self.stop_motor_now()  # 立即下发停转指令
                    self.other_mos_temp = payload.get('mos_temp', 0)
                    self.other_motor_temp = payload.get('motor_temp', 0)
                    self.other_temp_seen = True
                    self._payload_queue.append(payload)
        except Exception:
            pass
        if self._payload_queue:
            return self._payload_queue.popleft()
        return None

    def update_own_temperature(self, extra=None):
        """读取本机温度并心跳上报，成功返回True；extra 中的附加字段（如超温标志）一并上报"""
        if self.motor is None:
            return False
        mos, mot = self.motor.read_temperature()
        if mos is not None and mot is not None:
            self.own_mos_temp = mos
            self.own_motor_temp = mot
            self.send_status(dict(extra) if extra else {})
            return True
        return False

    def check_overheat(self):
        if (self.own_mos_temp > MOS_TEMP_LIMIT or self.own_motor_temp > MOTOR_TEMP_LIMIT
                or self.other_mos_temp > MOS_TEMP_LIMIT or self.other_motor_temp > MOTOR_TEMP_LIMIT):
            if not self.overheat_stopped:
                print(f"超温！本机 MOSFET:{self.own_mos_temp}℃ 电机:{self.own_motor_temp}℃ | "
                      f"对方 MOSFET:{self.other_mos_temp}℃ 电机:{self.other_motor_temp}℃，发送超温信号并立即停止电机")
                self.send_status({'overheat': True})
                self.overheat_stopped = True
                self.stop_motor_safely()  # 检测到超温立即停止电机，不等本轮流程走完
            return True
        return False

    def send_cmd_with_retry(self, cmd_func, desc, retries=5):
        for i in range(retries):
            try:
                if cmd_func():
                    print(f"{desc} 成功")
                    return True
            except Exception as e:
                print(f"{desc} 执行异常: {e}")
            print(f"{desc} 失败，重试 {i+1}/{retries}")
            time.sleep(0.5)
        return False

    def close_motor(self):
        if self.motor:
            self.motor.close()
            self.motor = None

    def setup_motor(self, com_port, label):
        """每轮初始化电机并清错。错误码连续3轮清除后仍存在时判定为设备锁存故障
        （如过温保护锁存，串口清错无法解除），提示断电重启并返回False终止测试。"""
        self.motor = MotorController(com_port, DLL_PATH, label)
        self.motor.clear_error()
        time.sleep(0.3)
        err = self.motor.read_error_code()
        if err is not None and err != 0:
            print(f"[警告] {label}初始化错误码 0x{err:X}，尝试清除")
            self.motor.clear_error()
            time.sleep(0.3)
            err = self.motor.read_error_code()
        if err:
            self.fault_rounds += 1
            print(f"[警告] {label}错误码 0x{err:X} 清除后仍存在（连续 {self.fault_rounds} 轮）")
            if self.fault_rounds >= 3:
                print(f"[错误] {label}错误码持续无法清除，疑似设备锁存故障（如过温保护锁存）")
                print(f"[错误] 请将{label}电机控制器断电重启后重新运行脚本，测试终止")
                self.send_status_multiple({'stop': True}, count=3, interval=0.1)  # 多次重发确保对端收到
                return False
        else:
            self.fault_rounds = 0
        return True

    def stop_motor_now(self):
        """立即下发一次停转指令（单次、尽力而为），用于收到停止/超温信号时的快速止血；
        完整停转（多次重发）由 stop_motor_safely 负责"""
        try:
            if self.motor:
                self.motor.set_speed_params(speed=0, accel=50, current=0)
        except BaseException:
            pass

    def clear_stale_messages(self):
        """丢弃 IPC 接收缓冲中残留的旧消息，防止上一轮残留指令串扰到下一轮"""
        self._payload_queue.clear()
        self._recv_buf = b""

    def stop_motor_safely(self):
        """尽力停止电机输出（不关闭串口），用于跳过轮次/超温时防止电机保持运转"""
        try:
            if self.motor:
                for _ in range(3):
                    self.motor.set_speed_params(speed=0, accel=50, current=0)
                    time.sleep(0.1)
        except BaseException:
            pass

    def wait_for_cooldown(self, tag):
        """超温后立即停止电机并暂停测试，等待双方温度降到恢复阈值以下后自动恢复继续。
        冷却期间每2秒读温度并心跳给对方（温度未恢复时心跳附带超温标志，防止对端漏收
        首次超温信号而继续运行）；温度连续读取失败（串口异常）达10次时终止测试。
        返回 False 表示对端线程已结束或无法继续监控温度，应退出主循环。"""
        print(f">>> [{tag}] 超温暂停：电机已停止，等待双方温度恢复 "
              f"(MOSFET<{MOS_TEMP_RECOVER}℃ 且 电机<{MOTOR_TEMP_RECOVER}℃)...")
        self.stop_motor_safely()
        last_read = 0
        last_print = 0
        read_fail = 0
        while not self.peer_done.is_set():
            now = time.time()
            if now - last_read >= 2.0:
                # 温度（本方视角）仍未恢复时持续附带超温标志，确保对端进入/保持暂停
                still_hot = (self.own_mos_temp >= MOS_TEMP_RECOVER or self.own_motor_temp >= MOTOR_TEMP_RECOVER
                             or self.other_mos_temp >= MOS_TEMP_RECOVER or self.other_motor_temp >= MOTOR_TEMP_RECOVER)
                if self.update_own_temperature({'overheat': True} if still_hot else None):
                    read_fail = 0
                else:
                    read_fail += 1
                    if read_fail >= 10:
                        print(f"[{tag}] 温度连续读取失败（串口异常），无法监控温度，终止测试")
                        return False
                last_read = now
            self.recv_status(0.1)  # 收对方温度心跳
            if now - last_print >= 10.0:
                print(f"[{tag}] 冷却中 | 本机 MOSFET:{self.own_mos_temp}℃ 电机:{self.own_motor_temp}℃ | "
                      f"对方 MOSFET:{self.other_mos_temp}℃ 电机:{self.other_motor_temp}℃")
                last_print = now
            own_valid = self.own_mos_temp > 0 or self.own_motor_temp > 0  # 确保至少读到过一次有效温度
            if (own_valid and self.other_temp_seen
                    and self.own_mos_temp < MOS_TEMP_RECOVER and self.own_motor_temp < MOTOR_TEMP_RECOVER
                    and self.other_mos_temp < MOS_TEMP_RECOVER and self.other_motor_temp < MOTOR_TEMP_RECOVER):
                print(f"[{tag}] 温度已恢复，继续测试")
                self.overheat_stopped = False
                self.stop_received = False
                self.skip_round = False
                if hasattr(self, '_error_count'):
                    self._error_count = 0
                self.clear_stale_messages()  # 丢弃超温暂停前残留的旧指令，防止跨轮串扰
                return True
        return False

    def emergency_stop(self):
        """Ctrl+C 紧急处理：尽力停止电机并释放串口（捕获一切中断，确保两侧都能执行到）"""
        try:
            if self.motor:
                self.motor.set_speed_params(speed=0, accel=50, current=0)
        except BaseException:
            pass
        try:
            self.close_motor()
        except BaseException:
            pass

    def round_aborted(self):
        """本轮是否因停止/超温/对端线程退出等原因需要放弃"""
        return self.skip_round or self.stop_received or self.overheat_stopped or self.peer_done.is_set()

# ========== 被测电机（原 被测电机_解锁上锁成功率测试.py） ==========
class TestedMotorRunner(PeerRunnerBase):
    """IPC服务器：响应解锁/上锁请求，统计成功率"""

    def __init__(self, server_ready, peer_done):
        super().__init__(server_ready, peer_done)
        # 统计变量
        self.unlock_attempts = 0
        self.unlock_success = 0
        self.lock_attempts = 0
        self.lock_success = 0

    def init_ipc_server(self):
        server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        server.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        server.bind((IPC_HOST, IPC_PORT))
        server.listen(1)
        print(f"IPC服务器启动 {IPC_HOST}:{IPC_PORT}")
        server.setblocking(False)
        self.server_ready.set()  # 通知负载线程：服务器已就绪，可以连接
        start = time.time()
        while time.time() - start < 30:
            try:
                self.sock, addr = server.accept()
                self.sock.setblocking(False)
                print(f"负载电机已连接: {addr}")
                server.close()
                return True
            except BlockingIOError:
                time.sleep(0.1)
        print("等待连接超时")
        server.close()
        return False

    def wait_for_unlock_success(self, timeout=1.0, interval=0.1):
        start = time.time()
        while time.time() - start < timeout:
            if self.stop_received or self.skip_round:
                return False
            err = self.motor.read_error_code()
            brake = self.motor.read_brake_state()
            err_text = f"0x{err:X}" if err is not None else "读取失败"
            print(f"轮询: 错误码={err_text}, 刹车状态={brake}")
            if err is not None and err == 0 and brake == PPX_BRAKE_OPENED:
                return True
            time.sleep(interval)
        return False

    def run(self):
        try:
            self._run_impl()
        finally:
            self.peer_done.set()  # 无论正常结束还是异常退出，都通知对端同步退出

    def _run_impl(self):
        print("=" * 60)
        print(f"被测电机 - 响应解锁/上锁，统计成功率，循环 {MAX_CYCLES} 次")
        print("Ctrl+C 退出")
        print("=" * 60)

        if not self.init_ipc_server():
            return

        cycle_count = 0
        while cycle_count < MAX_CYCLES:
            if self.peer_done.is_set():
                print(">>> 负载线程已结束，被测同步退出")
                break
            try:
                if self.overheat_stopped:
                    if not self.wait_for_cooldown("被测"):
                        break
                    continue
                if self.skip_round:
                    print(">>> 跳过本轮（收到停止信号）")
                    self.skip_round = False
                    self.stop_received = False
                    self.stop_motor_safely()
                    self.close_motor()
                    time.sleep(0.5)
                    self.clear_stale_messages()  # 丢弃上一轮残留指令，防止跨轮串扰
                    continue

                cycle_count += 1
                print(f"\n{'='*40} 第 {cycle_count}/{MAX_CYCLES} 轮开始 {'='*40}")

                if self.motor is not None:
                    self.close_motor()
                    gc.collect()
                    time.sleep(1.0)

                if not self.setup_motor(TESTED_COM_PORT, "被测"):
                    break

                print(">>> 等待解锁请求...")
                unlock_received = False
                start_wait = time.time()
                last_temp_read = 0
                while time.time() - start_wait < 30:
                    if self.peer_done.is_set():
                        print(">>> 负载线程已结束，停止等待")
                        break
                    if self.stop_received or self.skip_round:
                        print("收到停止信号，退出本轮")
                        break
                    payload = self.recv_status(0.1)
                    if payload and payload.get('cmd') == 'unlock_req':
                        unlock_received = True
                        print("收到解锁请求")
                        break
                    if self.overheat_stopped:
                        break
                    now = time.time()
                    if now - last_temp_read >= 0.2:
                        self.update_own_temperature()
                        last_temp_read = now
                    if self.check_overheat():
                        break
                if self.peer_done.is_set() and not unlock_received:
                    break
                if self.round_aborted():
                    print(">>> 本轮因停止信号或超温退出")
                    continue
                if not unlock_received:
                    print("等待解锁请求超时")
                    self.send_status({'stop': True})
                    self.skip_round = True
                    continue

                print(">>> 执行解锁（打开刹车）")
                self.motor.clear_error()
                time.sleep(0.2)
                if not self.send_cmd_with_retry(lambda: self.motor.set_run_mode(PPX_MODE_CLUTCH_OPEN), "打开刹车"):
                    print("打开刹车失败")
                    self.send_status({'cmd': 'unlock_res', 'success': False})
                    self.unlock_attempts += 1
                    continue

                print(">>> 检测解锁状态（轮询100ms，超时1秒）")
                unlock_ok = self.wait_for_unlock_success(timeout=1.0, interval=0.1)
                self.unlock_attempts += 1
                if unlock_ok:
                    self.unlock_success += 1
                    print("解锁成功（无错误码，刹车已打开）")
                else:
                    err = self.motor.read_error_code()
                    brake = self.motor.read_brake_state()
                    err_text = f"0x{err:X}" if err is not None else "读取失败"
                    print(f"解锁失败：错误码={err_text}, 刹车状态={brake}")
                    self.send_status({'cmd': 'unlock_res', 'success': False})
                    continue

                print(f">>> 切换到骑行模式，目标转速{TESTED_RUN_SPEED}")
                self.motor.clear_error()
                time.sleep(0.2)
                if not self.send_cmd_with_retry(lambda: self.motor.set_run_mode(PPX_MODE_RUNNING), "切换到骑行模式"):
                    print("切换骑行模式失败")
                    self.send_status({'cmd': 'unlock_res', 'success': False})
                    continue
                self.motor.set_speed_params(speed=TESTED_RUN_SPEED, accel=50, current=10)
                self.send_status({'cmd': 'unlock_res', 'success': True})
                print("解锁成功，骑行模式，持续下发")

                print(">>> 等待上锁请求...")
                lock_received = False
                start_wait = time.time()
                last_send = 0
                last_temp_read = 0
                while time.time() - start_wait < 30:
                    now = time.time()
                    if self.peer_done.is_set():
                        print(">>> 负载线程已结束，立即停止电机")
                        self.stop_motor_safely()
                        break
                    if self.stop_received or self.skip_round:
                        print("收到停止信号，退出本轮")
                        break
                    if self.overheat_stopped:
                        break
                    if now - last_send >= 0.2:
                        self.motor.set_speed_params(speed=TESTED_RUN_SPEED, accel=50, current=10)
                        last_send = now
                    payload = self.recv_status(0.1)
                    if payload and payload.get('cmd') == 'lock_req':
                        lock_received = True
                        print("收到上锁请求")
                        break
                    if now - last_temp_read >= 0.2:
                        self.update_own_temperature()
                        last_temp_read = now
                    if self.check_overheat():
                        break
                if self.round_aborted():
                    continue
                if not lock_received:
                    print("等待上锁请求超时")
                    self.send_status({'stop': True})
                    self.skip_round = True
                    continue

                print(">>> 执行上锁")
                self.motor.clear_error()
                time.sleep(0.2)
                if not self.send_cmd_with_retry(lambda: self.motor.set_run_mode(PPX_MODE_LOCK), "上锁"):
                    print("上锁失败")
                    self.send_status({'cmd': 'lock_res', 'success': False})
                    self.lock_attempts += 1
                    continue
                time.sleep(0.5)
                brake_state = self.motor.read_brake_state()
                self.lock_attempts += 1
                if brake_state == PPX_BRAKE_CLOSED:
                    self.lock_success += 1
                    print("刹车已闭合，上锁成功")
                    self.send_status({'cmd': 'lock_res', 'success': True})
                else:
                    print(f"刹车状态异常: {brake_state}，上锁失败")
                    self.send_status({'cmd': 'lock_res', 'success': False})

                print(f">>> 第 {cycle_count} 轮完成")
                print(f"解锁: {self.unlock_success}/{self.unlock_attempts} 成功 ({self.unlock_success/self.unlock_attempts*100:.1f}%)")
                if self.lock_attempts > 0:
                    print(f"上锁: {self.lock_success}/{self.lock_attempts} 成功 ({self.lock_success/self.lock_attempts*100:.1f}%)")
                else:
                    print("上锁: 0/0")

                self.close_motor()
                print("等待3秒进入下一轮...")
                time.sleep(3.0)

            except Exception as e:
                print(f"第 {cycle_count} 轮异常: {e}")
                traceback.print_exc()
                self.send_status({'stop': True})
                self.skip_round = True
                continue

        print(f"\n全部 {cycle_count} 轮执行完毕（或程序被终止）")
        print("=" * 40)
        print("最终统计结果（由被测电机统计）：")
        print(f"解锁: {self.unlock_success}/{self.unlock_attempts} 成功 ({self.unlock_success/self.unlock_attempts*100 if self.unlock_attempts > 0 else 0:.1f}%)")
        if self.lock_attempts > 0:
            print(f"上锁: {self.lock_success}/{self.lock_attempts} 成功 ({self.lock_success/self.lock_attempts*100:.1f}%)")
        else:
            print("上锁: 0/0")
        print("=" * 40)

        self.close_motor()
        if self.sock:
            self.sock.close()
        print("被测程序结束")

# ========== 负载电机（原 负载电机_解锁上锁成功率测试.py） ==========
class LoadMotorRunner(PeerRunnerBase):
    """IPC客户端：推行模式运行，发起解锁/上锁请求，不统计成功率"""

    def connect_ipc_server(self):
        # 等待被测线程的IPC服务器就绪后再连接，保持原"先服务器后客户端"的启动顺序
        if not self.server_ready.wait(timeout=35):
            print("等待IPC服务器就绪超时")
            return False
        for attempt in range(3):
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            try:
                s.connect((IPC_HOST, IPC_PORT))
                s.setblocking(False)
                self.sock = s
                print(f"负载已连接IPC服务器 {IPC_HOST}:{IPC_PORT}")
                return True
            except Exception as e:
                print(f"连接IPC服务器失败 (尝试 {attempt+1}/3): {e}")
                s.close()
                time.sleep(0.5)
        print("连接IPC服务器失败")
        return False

    def wait_for_motor_stop(self, timeout=3.0):
        start = time.time()
        while time.time() - start < timeout:
            try:
                speed = self.motor.read_motor_speed()
                if speed is not None and abs(speed) < 10:
                    return True
            except Exception:
                pass
            time.sleep(0.1)
        print("电机未在超时内停止")
        return False

    def send_stop_and_continue(self):
        """发送停止信号3次，间隔100ms，然后等待0.3秒确保送达，再返回"""
        print(">>> 发送停止信号（连续3次，间隔100ms）")
        for _ in range(3):
            self.send_status({'stop': True})
            time.sleep(0.1)
        time.sleep(0.3)  # 等待对端处理
        self.close_motor()
        self.skip_round = True  # 标记本轮放弃，避免后续重复发送停止信号
        print(">>> 停止信号已发送，将跳过本轮")

    def keep_running(self, desc="运行"):
        """推行模式保活：周期下发转速电流，自动清错；返回False表示连续错误需中止本轮"""
        if not hasattr(self, '_error_count'):
            self._error_count = 0
        self.motor.set_speed_params(speed=LOAD_RUN_SPEED, accel=50, current=LOAD_RUN_CURRENT)
        err = self.motor.read_error_code()
        if err is not None and err != 0:
            print(f"{desc}中错误码 0x{err:X}，自动清除")
            self.motor.clear_error()
            time.sleep(0.1)
            self.motor.set_speed_params(speed=LOAD_RUN_SPEED, accel=50, current=LOAD_RUN_CURRENT)
            self._error_count += 1
            if self._error_count >= 3:
                print("错误码连续出现3次，停止运行并跳过本轮")
                self.send_stop_and_continue()
                self._error_count = 0
                return False
        else:
            self._error_count = 0
        return True

    def run(self):
        try:
            self._run_impl()
        finally:
            self.peer_done.set()  # 无论正常结束还是异常退出，都通知对端同步退出

    def _run_impl(self):
        print("=" * 60)
        print(f"负载电机 - 推行模式（电流模式），目标转速{LOAD_RUN_SPEED}，电流{LOAD_RUN_CURRENT}A")
        print(f"循环 {MAX_CYCLES} 次，发送解锁/上锁请求，不统计成功率")
        print("不操作电磁刹车，自动清除错误")
        print("Ctrl+C 退出")
        print("=" * 60)

        if not self.connect_ipc_server():
            return

        cycle_count = 0
        while cycle_count < MAX_CYCLES:
            if self.peer_done.is_set():
                print(">>> 被测线程已结束，负载同步退出")
                break
            try:
                if self.overheat_stopped:
                    if not self.wait_for_cooldown("负载"):
                        break
                    continue
                if self.skip_round:
                    print(">>> 跳过本轮（收到停止信号）")
                    self.skip_round = False
                    self.stop_received = False
                    self.stop_motor_safely()
                    self.close_motor()
                    time.sleep(0.5)
                    self.clear_stale_messages()  # 丢弃上一轮残留指令，防止跨轮串扰
                    continue

                cycle_count += 1
                print(f"\n{'='*40} 第 {cycle_count}/{MAX_CYCLES} 轮开始 {'='*40}")

                if self.motor is not None:
                    self.close_motor()
                    gc.collect()
                    time.sleep(0.5)

                if not self.setup_motor(LOAD_COM_PORT, "负载"):
                    break

                print(">>> 切换到推行模式")
                if not self.send_cmd_with_retry(lambda: self.motor.set_run_mode(PPX_MODE_PWR_PUSH), "切换推行模式"):
                    print("无法切换到推行模式，终止本轮")
                    self.close_motor()
                    self.send_stop_and_continue()
                    continue

                print(f">>> 启动运行（速度{LOAD_RUN_SPEED}，电流{LOAD_RUN_CURRENT}A）")
                if not self.send_cmd_with_retry(
                        lambda: self.motor.set_speed_params(speed=LOAD_RUN_SPEED, accel=50, current=LOAD_RUN_CURRENT),
                        "设置参数"):
                    print("设置参数失败")
                    self.close_motor()
                    self.send_stop_and_continue()
                    continue

                print(">>> 运行0.5秒...")
                start_time = time.time()
                last_send = 0
                last_temp_read = 0
                last_temp_print = 0
                while time.time() - start_time < 0.5:
                    now = time.time()
                    if self.round_aborted():
                        break
                    if now - last_send >= 0.1:
                        last_send = now
                        if not self.keep_running():
                            break
                    if now - last_temp_read >= 0.1:
                        self.update_own_temperature()
                        last_temp_read = now
                    if now - last_temp_print >= 1.0:
                        print(f"[负载] 本机 MOSFET:{self.own_mos_temp}℃ 电机:{self.own_motor_temp}℃ | 对方 MOSFET:{self.other_mos_temp}℃ 电机:{self.other_motor_temp}℃")
                        last_temp_print = now
                    if self.check_overheat():
                        break
                    time.sleep(0.02)
                if self.round_aborted():
                    continue

                print(">>> 发送解锁请求（连续3次，负载继续运行）")
                self.send_status_multiple({'cmd': 'unlock_req'}, count=3, interval=0.1)
                time.sleep(0.05)

                # 等待解锁响应，超时2秒
                unlock_success = False
                start_wait = time.time()
                wait_timeout = 2.0
                last_send = 0
                while time.time() - start_wait < wait_timeout:
                    now = time.time()
                    if self.round_aborted():
                        break
                    if now - last_send >= 0.2:
                        last_send = now
                        if not self.keep_running():
                            break
                    payload = self.recv_status(0.05)
                    if payload and payload.get('cmd') == 'unlock_res':
                        success = payload.get('success', False)
                        if success:
                            unlock_success = True
                            print("被测解锁成功")
                        else:
                            print("被测解锁失败")
                        break
                    if now - last_temp_read >= 0.2:
                        self.update_own_temperature()
                        last_temp_read = now
                    if self.check_overheat():
                        break
                    time.sleep(0.01)
                else:
                    print("等待解锁响应超时（2秒）")

                if not unlock_success:
                    if not self.round_aborted():
                        print(">>> 解锁失败，发送停止信号并跳过本轮")
                        self.send_stop_and_continue()
                    continue

                print(">>> 解锁成功，继续运行...")
                run_time = random.uniform(1.0, 2.0)
                print(f">>> 继续运行 {run_time:.1f} 秒（电流{LOAD_RUN_CURRENT}A，速度{LOAD_RUN_SPEED}）")
                start_time = time.time()
                last_send = 0
                while time.time() - start_time < run_time:
                    now = time.time()
                    if self.round_aborted():
                        break
                    if now - last_send >= 0.2:
                        last_send = now
                        if not self.keep_running():
                            break
                    if now - last_temp_read >= 0.2:
                        self.update_own_temperature()
                        last_temp_read = now
                    if now - last_temp_print >= 1.0:
                        print(f"[负载] 本机 MOSFET:{self.own_mos_temp}℃ 电机:{self.own_motor_temp}℃ | 对方 MOSFET:{self.other_mos_temp}℃ 电机:{self.other_motor_temp}℃")
                        last_temp_print = now
                    if self.check_overheat():
                        break
                    time.sleep(0.02)
                if self.round_aborted():
                    continue

                print(">>> 减速停止")
                self.motor.set_speed_params(speed=0, accel=50, current=0)
                if not self.wait_for_motor_stop(timeout=3.0):
                    print("电机未完全停止，继续尝试上锁")
                for _ in range(3):
                    self.motor.set_speed_params(speed=0, accel=50, current=0)
                    time.sleep(0.2)

                print(">>> 发送上锁请求（连续3次）")
                self.send_status_multiple({'cmd': 'lock_req'}, count=3, interval=0.1)
                time.sleep(0.05)

                lock_success = False
                start_wait = time.time()
                last_send = 0
                while time.time() - start_wait < 5.0:
                    now = time.time()
                    if self.round_aborted():
                        break
                    if now - last_send >= 0.2:
                        self.motor.set_speed_params(speed=0, accel=50, current=0)
                        last_send = now
                    payload = self.recv_status(0.1)
                    if payload and payload.get('cmd') == 'lock_res':
                        success = payload.get('success', False)
                        if success:
                            lock_success = True
                            print("被测上锁成功")
                        else:
                            print("被测上锁失败")
                        break
                    if now - last_temp_read >= 0.2:
                        self.update_own_temperature()
                        last_temp_read = now
                    if now - last_temp_print >= 1.0:
                        print(f"[负载] 本机 MOSFET:{self.own_mos_temp}℃ 电机:{self.own_motor_temp}℃ | 对方 MOSFET:{self.other_mos_temp}℃ 电机:{self.other_motor_temp}℃")
                        last_temp_print = now
                    if self.check_overheat():
                        break
                    time.sleep(0.02)
                else:
                    print("等待上锁结果超时")
                    self.send_stop_and_continue()
                    continue

                if self.round_aborted():
                    continue

                print(f">>> 第 {cycle_count} 轮完成")
                self.close_motor()
                print("等待3秒进入下一轮...")
                time.sleep(3.0)

            except Exception as e:
                print(f"第 {cycle_count} 轮异常: {e}")
                traceback.print_exc()
                self.close_motor()
                self.send_stop_and_continue()
                continue

        print(f"\n全部 {cycle_count} 轮执行完毕（或程序被终止）")
        self.close_motor()
        if self.sock:
            self.sock.close()
        print("负载程序结束")

# ========== 主程序 ==========
def main():
    # 全程完整日志：控制台输出同时写入 log 文件夹
    os.makedirs(LOG_DIR, exist_ok=True)
    log_path = os.path.join(LOG_DIR, f"解锁上锁成功率测试_{time.strftime('%Y%m%d_%H%M%S')}.log")
    tee = Tee(log_path)
    sys.stdout = tee
    sys.stderr = tee

    print("=" * 60)
    print("解锁上锁成功率测试（合并版）")
    print("被测电机（IPC服务器）+ 负载电机（IPC客户端）双线程同时运行")
    print(f"被测串口: {TESTED_COM_PORT} | 负载串口: {LOAD_COM_PORT}")
    print(f"日志文件: {log_path}")
    print("=" * 60)

    server_ready = threading.Event()
    peer_done = threading.Event()
    tested = TestedMotorRunner(server_ready, peer_done)
    load = LoadMotorRunner(server_ready, peer_done)

    t_tested = threading.Thread(target=tested.run, name="TestedMotor", daemon=True)
    t_load = threading.Thread(target=load.run, name="LoadMotor", daemon=True)

    t_tested.start()  # 先启动被测电机（IPC服务器）
    t_load.start()    # 再启动负载电机（IPC客户端，等待服务器就绪后自动连接）

    try:
        while t_tested.is_alive() or t_load.is_alive():
            t_tested.join(timeout=0.5)
            t_load.join(timeout=0.5)
    except KeyboardInterrupt:
        print("\n用户中断，正在停止电机...")
        for runner in (load, tested):
            try:
                runner.emergency_stop()
            except BaseException:
                pass  # 二次 Ctrl+C 也不影响另一侧的紧急停止
        tee.close()
        os._exit(0)

    tee.close()

if __name__ == "__main__":
    main()
