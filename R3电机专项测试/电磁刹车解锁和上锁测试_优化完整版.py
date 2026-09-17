#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import ctypes
import json
import os
import select
import socket
import sys
import threading
import time
import traceback
import logging
from collections import Counter, deque
from typing import Optional, Tuple, Dict, Any, Callable

import serial

# ============================================================
# 配置区
# ============================================================
MAX_CYCLES: int = 1000

TESTED_COM_PORT: str = "COM28"       # 被测电机
LOAD_COM_PORT: str = "COM11"         # 负载电机
BAUD_RATE: int = 460800

SCRIPT_DIR: str = os.path.dirname(os.path.abspath(__file__))
DLL_PATH: str = os.path.join(SCRIPT_DIR, "ppx_region.dll")
LOG_DIR: str = os.path.join(SCRIPT_DIR, "log")

# -------------------- 电机运行参数 --------------------
TESTED_RUN_SPEED: int = -1000
TESTED_RUN_CURRENT: int = 10

LOAD_RUN_SPEED: int = 1000
LOAD_RUN_CURRENT: int = 20

# -------------------- 时序参数 --------------------
PRE_RUN_TIME: float = 0.5
RUN_TIME_MIN: float = 1.0
RUN_TIME_MAX: float = 2.0

UNLOCK_WAIT_TIMEOUT: float = 2.0
TESTED_UNLOCK_CHECK_TIMEOUT: float = 1.0
LOCK_WAIT_TIMEOUT: float = 5.0
ROUND_INTERVAL: float = 3.0

# -------------------- 错误码策略 --------------------
CLEAR_ERROR_EACH_ROUND: bool = True
CLEAR_ERROR_ON_START: bool = True

# -------------------- IPC --------------------
IPC_HOST: str = "127.0.0.1"
IPC_PORT: int = 12345

# ============================================================
# 协议常量
# ============================================================
PPX_ID_MCB: int = 0x20
PPX_CMD_REQ: int = 0x00
PPX_MSG_WRITE: int = 0x03
PPX_MSG_MULTIWR: int = 0x04
PPX_MSG_READ: int = 0x02

PPX_RUN_MODE_REG: int = 25
PPX_RT_SETTING_REG: int = 24
PPX_BRAKE_STATE_REG: int = 22
PPX_TEMP_REG: int = 12
PPX_TARGET_SPEED_REG: int = 27
PPX_TARGET_ACCEL_REG: int = 28
PPX_TARGET_CUR_REG: int = 29
PPX_MCU_ERRCODE_REG: int = 5
PPX_MOTOR_SPEED_REG: int = 6
PPX_BUS_VOLTAGE_REG: int = 8
PPX_BUS_CURRENT_REG: int = 9

PPX_MODE_IDLE: int = 0
PPX_MODE_PWR_PUSH: int = 4
PPX_MODE_RUNNING: int = 2
PPX_MODE_LOCK: int = 3

PPX_CLR_ERRCODE: int = (1 << 15)

PPX_BRAKE_CLOSED: int = 0
PPX_BRAKE_OPENING: int = 1
PPX_BRAKE_OPENED: int = 2

# ============================================================
# ctypes 结构体
# ============================================================
class PpxRegionExcp(ctypes.Structure):
    _fields_ = [
        ("parse_status", ctypes.c_uint8),
        ("cmd_status", ctypes.c_uint8),
        ("data_status", ctypes.c_uint8),
    ]

class PpxRegionMsg(ctypes.Structure):
    _fields_ = [
        ("id", ctypes.c_uint8),
        ("cmd", ctypes.c_uint8),
        ("msg_type", ctypes.c_uint8),
        ("reg_addr", ctypes.c_uint8),
        ("reg_nums", ctypes.c_uint8),
        ("reg_excp", PpxRegionExcp),
    ]

class PpxRegionData(ctypes.Structure):
    _pack_ = 1
    _fields_ = [
        ("id_num", ctypes.c_uint8),
        ("model", ctypes.c_uint8 * 8),
        ("serial_num", ctypes.c_uint8 * 26),
        ("hw_version", ctypes.c_uint16),
        ("sw_version", ctypes.c_uint8 * 32),
        ("mcu_errcode", ctypes.c_uint32),
        ("motor_speed", ctypes.c_int16),
        ("bus_voltage", ctypes.c_uint16),
        ("bus_current", ctypes.c_uint16),
        ("rim_state", ctypes.c_uint8),
        ("ctrl_model", ctypes.c_uint8),
        ("speed_ref", ctypes.c_int16),
        ("mosfet_temp", ctypes.c_int16),
        ("motor_temp", ctypes.c_int16),
        ("motor_limit_flg", ctypes.c_uint8),
        ("motor_vq", ctypes.c_int32),
        ("motor_iq", ctypes.c_int16),
        ("motor_cali_state", ctypes.c_uint8),
        ("motor_cali_res", ctypes.c_uint16),
        ("motor_cali_ld", ctypes.c_uint16),
        ("motor_cali_lq", ctypes.c_uint16),
        ("motor_cali_bemf", ctypes.c_uint16),
        ("brake_state", ctypes.c_uint8),
        ("single_mileage", ctypes.c_uint32),
        ("rt_setting", ctypes.c_uint16),
        ("run_mode", ctypes.c_uint8),
        ("road_cond", ctypes.c_uint8),
        ("target_speed", ctypes.c_int16),
        ("target_accel", ctypes.c_uint16),
        ("target_current", ctypes.c_int16),
        ("rated_voltage", ctypes.c_uint16),
        ("rated_current", ctypes.c_uint16),
        ("dat_setting", ctypes.c_uint32),
        ("reserved_data", ctypes.c_uint32),
    ]

class PpxRegionCtrl(ctypes.Structure):
    _fields_ = [
        ("msg", PpxRegionMsg),
        ("data", PpxRegionData),
    ]

# ============================================================
# 控制器封装
# ============================================================
class MotorController:
    def __init__(self, com_port: str, dll_path: str, label: str):
        self.com_port = com_port
        self.label = label
        self.dll_path = dll_path
        self.ser: Optional[serial.Serial] = None

        try:
            self.dll = ctypes.CDLL(dll_path, winmode=0)
        except OSError as exc:
            logging.error(f"{label}加载DLL失败: {dll_path} ({exc})")
            raise RuntimeError(f"DLL加载失败") from exc

        self.dll.ppx_com_region_format.argtypes = [
            ctypes.c_int,
            ctypes.POINTER(PpxRegionCtrl),
            ctypes.c_void_p,
        ]
        self.dll.ppx_com_region_format.restype = ctypes.c_uint16

        self.dll.ppx_com_region_parse.argtypes = [
            ctypes.POINTER(ctypes.c_uint8),
            ctypes.c_uint16,
            ctypes.POINTER(PpxRegionCtrl),
        ]
        self.dll.ppx_com_region_parse.restype = ctypes.c_int

        if not self._open_serial():
            raise RuntimeError(f"{label}串口 {com_port} 打开失败")

    def _open_serial(self) -> bool:
        for attempt in range(1, 4):
            try:
                if self.ser is not None:
                    try:
                        if self.ser.is_open:
                            self.ser.close()
                    except Exception:
                        pass

                self.ser = serial.Serial(
                    port=self.com_port,
                    baudrate=BAUD_RATE,
                    timeout=0.02,
                    write_timeout=0.5,
                )
                logging.info(f"{self.label}串口 {self.com_port} 已打开")
                return True
            except Exception as exc:
                logging.warning(f"{self.label}打开串口 {self.com_port} 失败 ({attempt}/3): {exc}")
                self.ser = None
                time.sleep(0.5)
        return False

    def _ensure_serial(self) -> bool:
        if self.ser is not None and self.ser.is_open:
            return True
        return self._open_serial()

    def _send_recv(self, ctrl: PpxRegionCtrl, timeout: float = 0.30, retry_count: int = 3) -> Optional[PpxRegionCtrl]:
        for retry in range(1, retry_count + 1):
            try:
                if not self._ensure_serial():
                    return None

                buf = (ctypes.c_uint8 * 512)()
                length = self.dll.ppx_com_region_format(PPX_CMD_REQ, ctypes.byref(ctrl), buf)
                if length <= 0:
                    logging.error(f"{self.label}协议组帧失败")
                    return None

                # 建立本地引用，防御多线程下对象被置空
                current_ser = self.ser
                if current_ser is None or not current_ser.is_open:
                    return None

                try:
                    current_ser.reset_input_buffer()
                except Exception:
                    pass

                current_ser.write(bytes(buf[:length]))
                current_ser.flush()

                deadline = time.monotonic() + timeout
                raw = bytearray()

                while time.monotonic() < deadline:
                    # 动态检查 self.ser 是否在其他线程已被释放
                    if self.ser is None:
                        break

                    try:
                        waiting = current_ser.in_waiting
                    except (AttributeError, OSError, serial.SerialException):
                        break

                    if waiting > 0:
                        raw.extend(current_ser.read(waiting))
                        if raw:
                            arr = (ctypes.c_uint8 * len(raw)).from_buffer_copy(bytes(raw))
                            ret = self.dll.ppx_com_region_parse(arr, len(raw), ctypes.byref(ctrl))
                            if ret == 1:
                                return ctrl
                    time.sleep(0.005)

                if retry < retry_count:
                    time.sleep(0.03)

            except (serial.SerialException, PermissionError, OSError) as exc:
                logging.warning(f"{self.label}串口通信异常（第{retry}/{retry_count}次）: {exc}")
                self.close()
                time.sleep(0.10)
            except Exception as exc:
                logging.error(f"{self.label}协议处理异常（第{retry}/{retry_count}次）: {exc}")
                self.close()
                time.sleep(0.05)

        return None

    def _write_reg(self, reg_addr: int, value: int, msg_type: int = PPX_MSG_WRITE) -> bool:
        ctrl = PpxRegionCtrl()
        ctrl.msg.id = PPX_ID_MCB
        ctrl.msg.cmd = msg_type
        ctrl.msg.reg_addr = reg_addr
        ctrl.msg.reg_nums = 1

        if msg_type == PPX_MSG_WRITE and reg_addr == PPX_RT_SETTING_REG:
            ctrl.data.rt_setting = int(value)
        elif reg_addr == PPX_RUN_MODE_REG:
            ctrl.data.run_mode = int(value)
        elif reg_addr == PPX_BRAKE_STATE_REG:
            ctrl.data.brake_state = int(value)
        else:
            raise ValueError(f"不支持的写寄存器: reg={reg_addr}")

        return self._send_recv(ctrl) is not None

    def clear_error(self) -> bool:
        return self._write_reg(PPX_RT_SETTING_REG, PPX_CLR_ERRCODE)

    def set_run_mode(self, mode: int) -> bool:
        return self._write_reg(PPX_RUN_MODE_REG, mode)

    def set_speed_params(self, speed: int, accel: int, current: int) -> bool:
        ctrl = PpxRegionCtrl()
        ctrl.msg.id = PPX_ID_MCB
        ctrl.msg.cmd = PPX_MSG_MULTIWR
        ctrl.msg.reg_addr = PPX_TARGET_SPEED_REG
        ctrl.msg.reg_nums = 3
        ctrl.data.target_speed = speed
        ctrl.data.target_accel = accel
        ctrl.data.target_current = current
        return self._send_recv(ctrl) is not None

    def read_error_code(self) -> Optional[int]:
        ctrl = PpxRegionCtrl()
        ctrl.msg.id = PPX_ID_MCB
        ctrl.msg.cmd = PPX_MSG_READ
        ctrl.msg.reg_addr = PPX_MCU_ERRCODE_REG
        ctrl.msg.reg_nums = 1
        res = self._send_recv(ctrl)
        return res.data.mcu_errcode if res else None

    def read_brake_state(self) -> int:
        ctrl = PpxRegionCtrl()
        ctrl.msg.id = PPX_ID_MCB
        ctrl.msg.cmd = PPX_MSG_READ
        ctrl.msg.reg_addr = PPX_BRAKE_STATE_REG
        ctrl.msg.reg_nums = 1
        res = self._send_recv(ctrl)
        return res.data.brake_state if res else -1

    def close(self) -> None:
        try:
            if self.ser is not None and self.ser.is_open:
                self.ser.close()
                logging.info(f"{self.label}串口已关闭")
        except Exception as exc:
            logging.error(f"{self.label}关闭串口异常: {exc}")
        finally:
            self.ser = None

# ============================================================
# 日志初始化
# ============================================================
def setup_logging() -> None:
    os.makedirs(LOG_DIR, exist_ok=True)
    log_file = os.path.join(LOG_DIR, f"解锁上锁成功率测试_{time.strftime('%Y%m%d_%H%M%S')}.log")
    
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    
    formatter = logging.Formatter('[%(asctime)s] %(message)s', datefmt='%H:%M:%S')
    
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(formatter)
    logger.addHandler(ch)
    
    fh = logging.FileHandler(log_file, mode='w', encoding='utf-8')
    fh.setFormatter(formatter)
    logger.addHandler(fh)

# ============================================================
# 公共 Runner
# ============================================================
class PeerRunnerBase(threading.Thread):
    def __init__(self, peer_done: threading.Event, label: str):
        super().__init__()
        self.peer_done = peer_done
        self.label = label
        self.motor: Optional[MotorController] = None
        self.sock: Optional[socket.socket] = None
        self._recv_buf = b""
        self._payload_queue = deque()
        self._socket_lock = threading.Lock()
        self.current_token: Optional[str] = None

    def send_status(self, payload: Dict[str, Any]) -> bool:
        if self.sock is None:
            return False
        if self.current_token and "round_token" not in payload:
            payload["round_token"] = self.current_token
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8") + b"\n"
        try:
            with self._socket_lock:
                self.sock.sendall(data)
            return True
        except Exception as exc:
            logging.error(f"[{self.label}] IPC发送失败: {exc}")
            return False

    def send_status_multiple(self, payload: Dict[str, Any], count: int = 3) -> bool:
        ok = False
        for _ in range(count):
            if self.send_status(payload):
                ok = True
            time.sleep(0.1)
        return ok

    def recv_status(self, timeout: float = 0.05) -> Optional[Dict[str, Any]]:
        if self._payload_queue:
            return self._payload_queue.popleft()
        if self.sock is None:
            return None
        try:
            ready, _, _ = select.select([self.sock], [], [], timeout)
            if not ready:
                return None
            data = self.sock.recv(4096)
            if not data:
                return None
            self._recv_buf += data
            while b"\n" in self._recv_buf:
                line, self._recv_buf = self._recv_buf.split(b"\n", 1)
                if not line.strip(): continue
                self._payload_queue.append(json.loads(line.decode("utf-8")))
        except Exception:
            pass
        return self._payload_queue.popleft() if self._payload_queue else None

    def send_cmd_with_retry(self, cmd_func: Callable[[], bool], desc: str, retries: int = 3) -> bool:
        for i in range(retries):
            if cmd_func():
                logging.info(f"{desc} 成功")
                return True
            logging.warning(f"{desc} 失败，重试 {i+1}/{retries}")
            time.sleep(0.1)
        return False

# ============================================================
# 测试执行类
# ============================================================
class TestedMotorRunner(PeerRunnerBase):
    def run(self) -> None:
        try:
            self.motor = MotorController(TESTED_COM_PORT, DLL_PATH, "被测电机")
            if CLEAR_ERROR_ON_START:
                self.motor.clear_error()
                logging.info(">>> TestedMotorRunner 启动时清除一次错误码")
                
            while not self.peer_done.is_set():
                msg = self.recv_status(timeout=0.1)
                if not msg:
                    continue

                cmd = msg.get("cmd")
                token = msg.get("round_token")

                if cmd == "unlock":
                    self.current_token = token
                    logging.info(f"收到解锁请求 token={token}")
                    
                    # 【核心优化】检测并清除历史错误码，确保刹车能够响应解锁指令
                    err = self.motor.read_error_code()
                    if err and err != 0:
                        logging.warning(f"检测到历史错误码 0x{err:X}，执行自动清除以防止解锁失败")
                        self.motor.clear_error()
                        time.sleep(0.1)

                    logging.info(">>> 执行解锁（打开刹车）")
                    self.send_cmd_with_retry(lambda: self.motor._write_reg(PPX_BRAKE_STATE_REG, PPX_BRAKE_OPENED), "打开刹车")
                    
                    logging.info(">>> 检测解锁状态（100ms轮询，超时1秒）")
                    start_poll = time.time()
                    success = False
                    
                    while time.time() - start_poll < TESTED_UNLOCK_CHECK_TIMEOUT:
                        err = self.motor.read_error_code()
                        brk = self.motor.read_brake_state()
                        err_str = f"0x{err:X}" if err is not None else "None"
                        logging.info(f"解锁轮询: 错误码={err_str}, 刹车状态={brk}")
                        
                        if err == 0 and brk == PPX_BRAKE_OPENED:
                            logging.info("解锁成功（错误码=0，刹车已打开）")
                            success = True
                            break
                        time.sleep(0.1)
                        
                    if success:
                        logging.info(">>> 切换到骑行模式，目标转速 -1000")
                        self.send_cmd_with_retry(lambda: self.motor.set_run_mode(PPX_MODE_RUNNING), "切换到骑行模式")
                        self.send_cmd_with_retry(lambda: self.motor.set_speed_params(TESTED_RUN_SPEED, 100, TESTED_RUN_CURRENT), "设置参数")
                        logging.info("解锁成功，进入骑行模式")
                        self.send_status_multiple({"status": "unlocked", "round_token": token})
                    else:
                        logging.error(f"解锁失败：错误码={err_str}, 刹车状态={brk}")
                        self.send_status_multiple({"status": "failed", "round_token": token})

                elif cmd == "lock":
                    logging.info(f"收到上锁请求 token={token}")
                    logging.info(">>> 减速停止")
                    self.send_cmd_with_retry(lambda: self.motor.set_speed_params(0, 100, 0), "停机指令")
                    time.sleep(1.0)
                    
                    logging.info(">>> 执行上锁")
                    self.send_cmd_with_retry(lambda: self.motor._write_reg(PPX_BRAKE_STATE_REG, PPX_BRAKE_CLOSED), "上锁")
                    logging.info("刹车已闭合，上锁成功")
                    self.send_status_multiple({"status": "locked", "round_token": token})

        except Exception as exc:
            logging.error(f"被测电机线程异常: {exc}")
            traceback.print_exc()
        finally:
            if self.motor:
                self.motor.close()
            if self.sock:
                self.sock.close()

class LoadMotorRunner(PeerRunnerBase):
    def run(self) -> None:
        try:
            self.motor = MotorController(LOAD_COM_PORT, DLL_PATH, "负载电机")
            if CLEAR_ERROR_ON_START:
                self.motor.clear_error()
                logging.info(">>> LoadMotorRunner 启动时清除一次错误码")

            for cycle in range(1, MAX_CYCLES + 1):
                if self.peer_done.is_set():
                    break
                    
                token = f"{cycle:04d}-{int(time.time()*1e9)}"
                self.current_token = token
                logging.info(f"\n======================================== 第 {cycle}/{MAX_CYCLES} 轮开始 token={token} ========================================")
                
                if CLEAR_ERROR_EACH_ROUND:
                    self.motor.clear_error()

                logging.info(">>> 切换到推行模式")
                self.send_cmd_with_retry(lambda: self.motor.set_run_mode(PPX_MODE_PWR_PUSH), "切换推行模式")
                
                logging.info(f">>> 启动运行（速度{LOAD_RUN_SPEED}，电流{LOAD_RUN_CURRENT}A）")
                self.send_cmd_with_retry(lambda: self.motor.set_speed_params(LOAD_RUN_SPEED, 100, LOAD_RUN_CURRENT), "设置参数")
                
                logging.info(f">>> 预运行 {PRE_RUN_TIME}s")
                time.sleep(PRE_RUN_TIME)
                
                logging.info(f">>> 继续运行 {PRE_RUN_TIME}s （负载速度={LOAD_RUN_SPEED}，电流={LOAD_RUN_CURRENT}A）")
                time.sleep(PRE_RUN_TIME)

                logging.info(f">>> 发送解锁请求 token={token}（连续3次，负载继续运行）")
                self.send_status_multiple({"cmd": "unlock", "round_token": token})

                # 等待被测解锁
                unlocked = False
                start_wait = time.time()
                while time.time() - start_wait < UNLOCK_WAIT_TIMEOUT:
                    msg = self.recv_status()
                    if msg and msg.get("round_token") == token:
                        if msg.get("status") == "unlocked":
                            logging.info(f"被测解锁成功 token={token}")
                            unlocked = True
                            break
                        elif msg.get("status") == "failed":
                            logging.error(f"被测解锁失败 token={token}")
                            break
                    time.sleep(0.05)

                if not unlocked:
                    logging.warning(">>> 停止当前轮次：解锁失败/超时")
                    self.send_cmd_with_retry(lambda: self.motor.set_speed_params(0, 100, 0), "停止运行")
                    time.sleep(ROUND_INTERVAL)
                    continue

                logging.info(f">>> 继续运行 {RUN_TIME_MIN}s （负载速度={LOAD_RUN_SPEED}，电流={LOAD_RUN_CURRENT}A）")
                time.sleep(RUN_TIME_MIN)

                # 发送上锁请求
                logging.info(f">>> 发送上锁请求 token={token}（连续3次）")
                self.send_status_multiple({"cmd": "lock", "round_token": token})

                locked = False
                start_wait = time.time()
                while time.time() - start_wait < LOCK_WAIT_TIMEOUT:
                    msg = self.recv_status()
                    if msg and msg.get("status") == "locked" and msg.get("round_token") == token:
                        logging.info(f"被测上锁成功 token={token}")
                        locked = True
                        break
                    time.sleep(0.05)

                self.send_cmd_with_retry(lambda: self.motor.set_speed_params(0, 100, 0), "负载停止")
                logging.info(f">>> 第 {cycle} 轮完成")
                logging.info(f"等待 {ROUND_INTERVAL}s 进入下一轮...")
                time.sleep(ROUND_INTERVAL)
                
        except Exception as exc:
            logging.error(f"负载电机线程异常: {exc}")
            traceback.print_exc()
        finally:
            self.peer_done.set()
            if self.motor:
                self.motor.close()
            if self.sock:
                self.sock.close()

# ============================================================
# 主控逻辑
# ============================================================
def main():
    setup_logging()
    logging.info("电磁刹车解锁 / 上锁成功率测试 - 优化完整版启动")
    
    # 建立简单的 IPC 本地 Server
    server_sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    server_sock.bind((IPC_HOST, IPC_PORT))
    server_sock.listen(1)
    logging.info(f"IPC服务器启动 {IPC_HOST}:{IPC_PORT}")

    peer_done = threading.Event()
    
    tested_runner = TestedMotorRunner(peer_done, "被测")
    load_runner = LoadMotorRunner(peer_done, "负载")

    def accept_thread():
        try:
            conn, _ = server_sock.accept()
            tested_runner.sock = conn
        except Exception:
            pass

    acc_th = threading.Thread(target=accept_thread, daemon=True)
    acc_th.start()

    time.sleep(0.5)
    
    # 负载作为 Client 连入
    client_sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        client_sock.connect((IPC_HOST, IPC_PORT))
        logging.info(f"负载已连接IPC服务器 {IPC_HOST}:{IPC_PORT}")
        load_runner.sock = client_sock
    except Exception as e:
        logging.error(f"客户端连接IPC失败: {e}")
        return

    tested_runner.start()
    load_runner.start()

    try:
        while not peer_done.is_set():
            time.sleep(1)
    except KeyboardInterrupt:
        logging.info("用户中断，正在执行双端紧急停机...")
        peer_done.set()
    finally:
        server_sock.close()
        tested_runner.join(timeout=2.0)
        load_runner.join(timeout=2.0)
        logging.info("所有线程已结束，测试退出。")

if __name__ == "__main__":
    main()
