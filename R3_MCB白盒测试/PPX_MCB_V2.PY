import os
import sys
import ctypes
import logging
import serial
import time
import argparse
import threading
from typing import Optional, Dict, Any, List
from datetime import datetime

# ==========================================
# 配置标准日志记录（含完整日志文件）
# ==========================================
LOG_FILENAME = f"motor_ctrl_full_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILENAME, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

# ==========================================
# 协议常量定义
# ==========================================
PPX_ID_MCB: int = 0x20
PPX_CMD_REQ: int = 0x00
PPX_CMD_RSP: int = 0x80
PPX_MSG_READ: int = 0x01
PPX_MSG_MULTIRD: int = 0x02
PPX_MSG_WRITE: int = 0x03
PPX_MSG_MULTIWR: int = 0x04
PPX_CMD_EXCP: int = 0xC0

PPX_ID_CCB: int = 0x10

PPX_MODE_IDLE: int = 0
PPX_MODE_SETTING: int = 1
PPX_MODE_RUNNING: int = 2
PPX_MODE_LOCK: int = 3
PPX_MODE_PWR_PUSH: int = 4
PPX_MODE_BRK_EMER: int = 5
PPX_MODE_REHAB_WALK: int = 8
PPX_MODE_REHAB_TRAIN: int = 9
PPX_MODE_BRAKE: int = 10
PPX_MODE_BRK_ECND: int = 11
PPX_MODE_CLUTCH_OPEN: int = 12

PPX_MCU_ERRCODE_REG: int = 5
PPX_MOTOR_SPEED_REG: int = 6
PPX_BUS_VOLTAGE_REG: int = 7
PPX_BUS_CURRENT_REG: int = 8
PPX_RIM_STATE_REG: int = 9
PPX_CTRL_MODEL_REG: int = 10
PPX_SPEED_REF_REG: int = 11
PPX_BRAKE_STATE_REG: int = 22
PPX_RT_SETTING_REG: int = 24
PPX_RUN_MODE_REG: int = 25
PPX_ROAD_COND_REG: int = 26
PPX_TARGET_SPEED_REG: int = 27
PPX_TARGET_ACCEL_REG: int = 28
PPX_TARGET_CUR_REG: int = 29

PPX_CLR_ERRCODE: int = (1 << 15)
PPX_MOTOR_STOP: int = (1 << 10)
PPX_PARA_LEARN: int = (1 << 11)

PPX_BRAKE_CLOSED: int = 0
PPX_BRAKE_OPENING: int = 1
PPX_BRAKE_OPENED: int = 2

# ==========================================
# 稳定性测试常量（可调整）
# ==========================================
STABILITY_LONG_RUN_CYCLES: int = 100       # 长时间运行循环次数
STABILITY_BURST_COUNT: int = 200           # 指令洪泛发包数
STABILITY_IDLE_RECOVERY_WAIT: float = 5.0  # 长时间空闲后恢复等待秒数
STABILITY_HEARTBEAT_INTERVAL: float = 0.5  # 心跳检测间隔秒数
STABILITY_RAMP_CYCLES: int = 5             # 斜坡往复次数
STABILITY_MODE_SWITCH_ROUNDS: int = 10     # 模式反复切换轮数

# ==========================================
# 结构体定义（链路层、寄存器层）
# ==========================================
class PpxPacketData(ctypes.Structure):
    _fields_: List[tuple] = [
        ("id", ctypes.c_uint8),
        ("cmd", ctypes.c_uint8),
        ("data_len", ctypes.c_uint16),
        ("data", ctypes.c_uint8 * 416)
    ]

class PpxRegionExcp(ctypes.Structure):
    _fields_: List[tuple] = [
        ("parse_status", ctypes.c_uint8),
        ("cmd_status", ctypes.c_uint8),
        ("data_status", ctypes.c_uint8)
    ]

class PpxRegionMsg(ctypes.Structure):
    _fields_: List[tuple] = [
        ("id", ctypes.c_uint8),
        ("cmd", ctypes.c_uint8),
        ("msg_type", ctypes.c_uint8),
        ("reg_addr", ctypes.c_uint8),
        ("reg_nums", ctypes.c_uint8),
        ("reg_excp", PpxRegionExcp)
    ]

class PpxRegionData(ctypes.Structure):
    _pack_: int = 1
    _fields_: List[tuple] = [
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
        ("reserved_data", ctypes.c_uint32)
    ]

class PpxRegionCtrl(ctypes.Structure):
    _fields_: List[tuple] = [
        ("msg", PpxRegionMsg),
        ("data", PpxRegionData)
    ]

# ==========================================
# 电机控制驱动类（含稳定性深度测试）
# ==========================================
class MCBMotorController:
    def __init__(self, com_port: str, dll_path: str) -> None:
        self.com_port: str = com_port
        self.baudrate: int = 460800
        self.dll_path: str = os.path.abspath(dll_path)
        self.serial_conn: Optional[serial.Serial] = None
        self.dll: Optional[ctypes.CDLL] = None
        self._lock = threading.Lock()
        self.test_results: List[Dict[str, Any]] = []
        # 稳定性测试统计
        self._crash_detected: bool = False
        self._consecutive_failures: int = 0
        self._total_stress_commands: int = 0
        self._stress_failures: int = 0

    # ---------- 环境初始化 ----------
    def init_environment(self) -> bool:
        try:
            dll_dir = os.path.dirname(self.dll_path)
            if sys.platform == "win32" and sys.version_info >= (3, 8):
                os.add_dll_directory(dll_dir)

            self.dll = ctypes.CDLL(self.dll_path, winmode=0)

            self.dll.ppx_com_region_format.argtypes = [ctypes.c_int, ctypes.POINTER(PpxRegionCtrl), ctypes.c_void_p]
            self.dll.ppx_com_region_format.restype = ctypes.c_uint16

            self.dll.ppx_com_packet_parse.argtypes = [ctypes.POINTER(ctypes.c_uint8), ctypes.c_uint16, ctypes.POINTER(PpxPacketData)]
            self.dll.ppx_com_packet_parse.restype = ctypes.c_int

            self.dll.ppx_com_region_parse.argtypes = [ctypes.POINTER(ctypes.c_uint8), ctypes.c_uint16, ctypes.POINTER(PpxRegionCtrl)]
            self.dll.ppx_com_region_parse.restype = ctypes.c_int

            self.serial_conn = serial.Serial(
                port=self.com_port,
                baudrate=self.baudrate,
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                timeout=0.3
            )
            logging.info("电机控制驱动环境初始化成功。")
            return True
        except Exception as e:
            logging.error(f"环境构建失败: {e}")
            return False

    # ---------- 通用收发 ----------
    def _send_and_receive(self, ctrl_frame: PpxRegionCtrl, label: str = "") -> Optional[bytes]:
        if not self.serial_conn or not self.dll:
            logging.error("串口或DLL未初始化。")
            return None

        with self._lock:
            try:
                tx_buffer = (ctypes.c_uint8 * 512)()
                packet_len = self.dll.ppx_com_region_format(PPX_CMD_REQ, ctypes.byref(ctrl_frame), tx_buffer)
                if packet_len <= 0:
                    logging.error(f"[{label}] 组包失败。")
                    return None

                raw_tx_bytes = bytes(tx_buffer[:packet_len])
                logging.info(f"TX [{label}] (HEX): {raw_tx_bytes.hex().upper()}")

                self.serial_conn.reset_input_buffer()
                self.serial_conn.write(raw_tx_bytes)
                self.serial_conn.flush()

                for retry in range(2):
                    timeout_counter = 0
                    while self.serial_conn.in_waiting == 0 and timeout_counter < 15:
                        time.sleep(0.02)
                        timeout_counter += 1

                    in_waiting = self.serial_conn.in_waiting
                    if in_waiting == 0:
                        if retry == 0:
                            logging.warning(f"[{label}] 物理超时，重试...")
                            continue
                        else:
                            logging.warning(f"[{label}] 重试后仍超时。")
                            return None

                    rx_bytes = self.serial_conn.read(in_waiting)
                    logging.info(f"RX [{label}] (HEX): {rx_bytes.hex().upper()}")

                    if len(rx_bytes) >= 9 and rx_bytes[0] == 0xA5 and rx_bytes[-1] == 0x55:
                        return rx_bytes
                    else:
                        if retry == 0:
                            logging.warning(f"[{label}] 帧不完整，重试...")
                            time.sleep(0.05)
                        else:
                            logging.warning(f"[{label}] 重试后帧仍不完整，尝试解析。")
                            return rx_bytes
                return rx_bytes
            except Exception as e:
                logging.error(f"[{label}] 收发异常: {e}")
                return None

    def _parse_response(self, rx_bytes: bytes) -> Optional[PpxRegionCtrl]:
        if not rx_bytes or not self.dll:
            return None
        try:
            raw_rx_ptr = (ctypes.c_uint8 * len(rx_bytes))(*rx_bytes)
            packet_layer_output = PpxPacketData()
            packet_status = self.dll.ppx_com_packet_parse(raw_rx_ptr, len(rx_bytes), ctypes.byref(packet_layer_output))
            if packet_status != 1:
                logging.error(f"链路解析失败（返回码: {packet_status}）。")
                return None

            logging.info(f"--- 链路层 ID: {hex(packet_layer_output.id)} | CMD: {hex(packet_layer_output.cmd)} ---")

            if (packet_layer_output.cmd & PPX_CMD_EXCP) == PPX_CMD_EXCP:
                logging.error("MCB 返回 Exception 命令。")
                return None

            rx_ctrl = PpxRegionCtrl()
            rx_ctrl.msg.id = packet_layer_output.id
            rx_ctrl.msg.cmd = packet_layer_output.cmd

            pure_data_ptr = packet_layer_output.data
            pure_data_len = packet_layer_output.data_len
            region_status = self.dll.ppx_com_region_parse(pure_data_ptr, pure_data_len, ctypes.byref(rx_ctrl))
            if region_status != 1:
                region_status = self.dll.ppx_com_region_parse(raw_rx_ptr, len(rx_bytes), ctypes.byref(rx_ctrl))
            if region_status != 1:
                logging.error(f"寄存器解析失败（返回码: {region_status}）。")
                return None

            return rx_ctrl
        except Exception as e:
            logging.error(f"响应解析异常: {e}")
            return None

    # ---------- 读/写基础方法 ----------
    def _read_register(self, reg_addr: int, reg_nums: int, label: str = "") -> Optional[PpxRegionCtrl]:
        ctrl_frame = PpxRegionCtrl()
        ctrl_frame.msg.id = PPX_ID_MCB
        ctrl_frame.msg.cmd = PPX_MSG_READ if reg_nums <= 1 else PPX_MSG_MULTIRD
        ctrl_frame.msg.reg_addr = reg_addr
        ctrl_frame.msg.reg_nums = reg_nums

        rx_bytes = self._send_and_receive(ctrl_frame, label)
        if rx_bytes is None:
            return None
        return self._parse_response(rx_bytes)

    def _write_register(self, reg_addr: int, reg_nums: int, data_fields: Dict[str, Any],
                        label: str = "", safe_bits: bool = False) -> bool:
        ctrl_frame = PpxRegionCtrl()
        ctrl_frame.msg.id = PPX_ID_MCB
        ctrl_frame.msg.cmd = PPX_MSG_WRITE if reg_nums <= 1 else PPX_MSG_MULTIWR
        ctrl_frame.msg.reg_addr = reg_addr
        ctrl_frame.msg.reg_nums = reg_nums

        if safe_bits:
            current = self._read_register(reg_addr, reg_nums, f"{label}_读前")
            if current is not None:
                for field_name, value in data_fields.items():
                    if hasattr(current.data, field_name):
                        cur = getattr(current.data, field_name)
                        if isinstance(value, int) and isinstance(cur, int):
                            merged = cur | value
                            setattr(ctrl_frame.data, field_name, merged)
                            logging.info(f"[{label}] 位合并: {field_name} {hex(cur)} | {hex(value)} -> {hex(merged)}")
                        else:
                            setattr(ctrl_frame.data, field_name, value)
            else:
                logging.warning(f"[{label}] 读前失败，直接写入。")
                for field_name, value in data_fields.items():
                    if hasattr(ctrl_frame.data, field_name):
                        setattr(ctrl_frame.data, field_name, value)
        else:
            for field_name, value in data_fields.items():
                if hasattr(ctrl_frame.data, field_name):
                    setattr(ctrl_frame.data, field_name, value)

        rx_bytes = self._send_and_receive(ctrl_frame, label)
        if rx_bytes is None:
            return False
        result = self._parse_response(rx_bytes)
        if result is None:
            return False
        logging.info(f"[{label}] 写入成功，响应CMD: {hex(result.msg.cmd)}")
        return True

    # ---------- 业务指令 ----------
    def clear_error_code(self) -> bool:
        logging.info("--- 清除故障码 ---")
        return self._write_register(PPX_RT_SETTING_REG, 1, {'rt_setting': PPX_CLR_ERRCODE}, "清除故障码")

    def read_motor_status(self) -> bool:
        logging.info("--- 读取电机状态 ---")
        result = self._read_register(PPX_MCU_ERRCODE_REG, 4, "读取状态")
        if result is None:
            return False
        logging.info("--- 当前电机状态 ---")
        logging.info(f"故障码: {result.data.mcu_errcode}")
        logging.info(f"转速: {result.data.motor_speed} rpm")
        logging.info(f"母线电压: {result.data.bus_voltage * 0.1:.1f} V")
        logging.info(f"母线电流: {result.data.bus_current * 0.1:.1f} A")
        return True

    def read_brake_state(self) -> Optional[int]:
        result = self._read_register(PPX_BRAKE_STATE_REG, 1, "读取刹车")
        if result is None:
            return None
        state = result.data.brake_state
        state_map = {0: "闭合(锁定)", 1: "解锁中", 2: "已解锁"}
        logging.info(f"电磁刹状态: {state} -> {state_map.get(state, '未知')}")
        return state

    def set_run_mode(self, mode: int) -> bool:
        mode_names = {0: "空闲", 1: "设置", 2: "骑行", 3: "锁车", 4: "助力推行",
                      5: "紧急刹车", 8: "康复助行", 9: "康复训练", 10: "延迟关刹",
                      11: "立即关刹", 12: "电磁刹解锁"}
        label = f"设置模式({mode_names.get(mode, f'未知{mode}')})"
        logging.info(f"--- {label} ---")
        return self._write_register(PPX_RUN_MODE_REG, 1, {'run_mode': mode}, label)

    def set_target_speed(self, speed_rpm: int) -> bool:
        logging.info(f"--- 设置目标速度: {speed_rpm} rpm ---")
        return self._write_register(PPX_TARGET_SPEED_REG, 3, {
            'target_speed': speed_rpm, 'target_accel': 500, 'target_current': 0
        }, f"速度{speed_rpm}rpm")

    def set_target_current(self, current_01a: int, speed_rpm: int = 0) -> bool:
        logging.info(f"--- 设置电流: {current_01a * 0.1:.1f}A ---")
        return self._write_register(PPX_TARGET_SPEED_REG, 3, {
            'target_speed': speed_rpm, 'target_accel': 500, 'target_current': current_01a
        }, f"电流{current_01a * 0.1:.1f}A")

    def emergency_stop(self) -> bool:
        logging.info("--- 紧急停止 ---")
        return self.set_run_mode(PPX_MODE_IDLE)

    def start_motor_speed_mode(self, target_speed_rpm: int) -> bool:
        logging.info("=" * 50)
        logging.info(f"=== 启动电机（速度模式）目标: {target_speed_rpm} rpm ===")
        logging.info("=" * 50)
        self.clear_error_code()
        time.sleep(0.1)
        self.read_motor_status()
        time.sleep(0.1)
        self.set_run_mode(PPX_MODE_IDLE)
        time.sleep(0.1)
        self._write_register(PPX_TARGET_SPEED_REG, 3, {
            'target_speed': target_speed_rpm, 'target_accel': 500, 'target_current': 0
        }, "预设速度")
        time.sleep(0.1)
        if not self.set_run_mode(PPX_MODE_RUNNING):
            logging.error("切换骑行模式失败。")
            return False
        time.sleep(0.3)
        self.read_motor_status()
        return True

    def stop_motor(self) -> bool:
        logging.info("=== 停止电机 ===")
        self._write_register(PPX_TARGET_SPEED_REG, 3, {
            'target_speed': 0, 'target_accel': 500, 'target_current': 0
        }, "速度归零")
        time.sleep(0.1)
        self.set_run_mode(PPX_MODE_IDLE)
        time.sleep(0.2)
        self.read_motor_status()
        logging.info("=== 电机已停止 ===")
        return True

    def ramp_test(self, max_speed: int = 500, step: int = 50, dwell: float = 1.0) -> bool:
        logging.info(f"=== 斜坡测试: 0->{max_speed} rpm (步进{step}, 停留{dwell}s) ===")
        self.clear_error_code()
        time.sleep(0.1)
        self._write_register(PPX_TARGET_SPEED_REG, 3, {
            'target_speed': 0, 'target_accel': 500, 'target_current': 0
        }, "初始参数")
        time.sleep(0.1)
        if not self.set_run_mode(PPX_MODE_RUNNING):
            logging.error("切入骑行失败。")
            return False

        speeds = list(range(0, max_speed + 1, step))
        if speeds[-1] != max_speed:
            speeds.append(max_speed)
        for s in speeds:
            logging.info(f">>> 速度: {s} rpm")
            self._write_register(PPX_TARGET_SPEED_REG, 1, {'target_speed': s}, f"速度{s}")
            time.sleep(dwell)

        speeds_desc = list(range(max_speed - step, -1, -step))
        if speeds_desc[-1] != 0:
            speeds_desc.append(0)
        for s in speeds_desc:
            logging.info(f">>> 速度: {s} rpm")
            self._write_register(PPX_TARGET_SPEED_REG, 1, {'target_speed': s}, f"速度{s}")
            time.sleep(dwell)

        self.set_run_mode(PPX_MODE_IDLE)
        time.sleep(0.2)
        self.read_motor_status()
        logging.info("=== 斜坡测试完成 ===")
        return True

    # ---------- 测试结果记录 ----------
    def _record_test_result(self, test_name: str, passed: bool, details: str = "") -> None:
        self.test_results.append({
            "name": test_name,
            "status": "PASS" if passed else "FAIL",
            "details": details,
            "timestamp": datetime.now().strftime("%H:%M:%S.%f")[:-3]
        })
        logging.info(f"[TEST RESULT] {'PASS' if passed else 'FAIL'} | {test_name} | {details}")

    # ---------- 基础功能测试辅助方法 ----------
    def _test_read_state(self) -> bool:
        success = self.read_motor_status()
        self._record_test_result("读取电机状态", success, "寄存器5~9")
        return success

    def _test_clear_error(self) -> bool:
        success = self.clear_error_code()
        self._record_test_result("清除故障码", success, "写入CLR_ERRCODE")
        return success

    def _test_brake_state(self) -> bool:
        state = self.read_brake_state()
        success = state is not None
        self._record_test_result("读取电磁刹状态", success, f"状态码:{state}")
        return success

    def _test_mode_switch(self, mode: int, mode_name: str) -> bool:
        logging.info(f"=== 测试: 切换{mode_name} ===")
        success = self.set_run_mode(mode)
        time.sleep(0.1)
        verify = self._read_register(PPX_RUN_MODE_REG, 1, f"验证{mode_name}")
        if verify and verify.data.run_mode == mode:
            self._record_test_result(f"模式切换:{mode_name}", True, f"回读={mode}")
            return True
        else:
            actual = verify.data.run_mode if verify else "N/A"
            self._record_test_result(f"模式切换:{mode_name}", False, f"期望{mode},实际{actual}")
            return False

    def _test_all_modes(self) -> bool:
        modes = [(0, "空闲"), (3, "锁车"), (4, "助力推行"), (10, "延迟关刹"),
                 (11, "立即关刹"), (12, "电磁刹解锁"), (0, "恢复空闲")]
        all_ok = True
        for m, n in modes:
            if not self._test_mode_switch(m, n):
                all_ok = False
            time.sleep(0.15)
        return all_ok

    def _test_speed_control(self, speed: int) -> bool:
        logging.info(f"=== 测试: 速度控制 {speed}rpm ===")
        self.set_run_mode(PPX_MODE_RUNNING)
        time.sleep(0.1)
        self.set_target_speed(speed)
        time.sleep(0.3)
        readback = self._read_register(PPX_TARGET_SPEED_REG, 2, f"验证速度{speed}")
        if readback:
            actual = readback.data.target_speed
            if abs(actual - speed) <= max(10, speed * 0.1):
                self._record_test_result(f"速度控制{speed}rpm", True, f"回读={actual}")
                return True
            else:
                self._record_test_result(f"速度控制{speed}rpm", False, f"期望{speed},实际{actual}")
                return False
        else:
            self._record_test_result(f"速度控制{speed}rpm", False, "回读失败")
            return False

    def _test_current_control(self, current_01a: int) -> bool:
        logging.info(f"=== 测试: 电流控制 {current_01a * 0.1:.1f}A ===")
        self.set_run_mode(PPX_MODE_RUNNING)
        time.sleep(0.1)
        self.set_target_current(current_01a)
        time.sleep(0.2)
        readback = self._read_register(PPX_TARGET_CUR_REG, 1, f"验证电流{current_01a*0.1:.1f}A")
        if readback:
            actual = readback.data.target_current
            if abs(actual - current_01a) <= 5:
                self._record_test_result(f"电流控制{current_01a*0.1:.1f}A", True, f"回读={actual*0.1:.1f}A")
                return True
            else:
                self._record_test_result(f"电流控制{current_01a*0.1:.1f}A", False, f"期望{current_01a},实际{actual}")
                return False
        else:
            self._record_test_result(f"电流控制{current_01a*0.1:.1f}A", False, "回读失败")
            return False

    def _test_ramp_profile(self) -> bool:
        try:
            success = self.ramp_test(max_speed=300, step=100, dwell=0.8)
            self._record_test_result("斜坡测试", success, "0->300->0")
            return success
        except Exception as e:
            self._record_test_result("斜坡测试", False, str(e))
            return False

    def _test_param_boundary(self) -> bool:
        all_ok = True
        if self.set_target_speed(-100):
            self._record_test_result("边界:负速度-100", True, "设置成功")
        else:
            self._record_test_result("边界:负速度-100", False, "拒绝或失败")
            all_ok = False
        time.sleep(0.1)
        self.set_run_mode(PPX_MODE_RUNNING)
        time.sleep(0.1)
        if self.set_target_current(0):
            self._record_test_result("边界:零电流", True, "设置成功")
        else:
            self._record_test_result("边界:零电流", False, "设置失败")
            all_ok = False
        time.sleep(0.1)
        self.set_run_mode(PPX_MODE_IDLE)
        return all_ok

    # ---------- 白盒边界测试 ----------
    def _test_read_invalid_register(self) -> bool:
        logging.info("=== 测试: 读取非法寄存器 (0xFF) ===")
        result = self._read_register(0xFF, 1, "非法读")
        passed = result is None
        self._record_test_result("读取非法寄存器", passed, "预期异常" if passed else "意外返回数据")
        return passed

    def _test_write_invalid_register(self) -> bool:
        logging.info("=== 测试: 写入非法寄存器 (0x00) ===")
        ctrl_frame = PpxRegionCtrl()
        ctrl_frame.msg.id = PPX_ID_MCB
        ctrl_frame.msg.cmd = PPX_MSG_WRITE
        ctrl_frame.msg.reg_addr = 0x00
        ctrl_frame.msg.reg_nums = 1
        ctrl_frame.data.id_num = 0x12
        rx_bytes = self._send_and_receive(ctrl_frame, "非法写")
        if rx_bytes is None:
            self._record_test_result("写入非法寄存器", False, "通信失败")
            return False
        try:
            raw_rx_ptr = (ctypes.c_uint8 * len(rx_bytes))(*rx_bytes)
            packet_layer_output = PpxPacketData()
            packet_status = self.dll.ppx_com_packet_parse(raw_rx_ptr, len(rx_bytes), ctypes.byref(packet_layer_output))
            if packet_status != 1:
                self._record_test_result("写入非法寄存器", False, "链路解析失败")
                return False
            if (packet_layer_output.cmd & PPX_CMD_EXCP) == PPX_CMD_EXCP:
                self._record_test_result("写入非法寄存器", True, "MCB返回异常帧，写入被拒绝")
                return True
            else:
                self._record_test_result("写入非法寄存器", False, "MCB未拒绝(固件缺陷)")
                return False
        except Exception as e:
            self._record_test_result("写入非法寄存器", False, f"解析异常:{e}")
            return False

    def _test_write_readonly_reg(self) -> bool:
        logging.info("=== 测试: 写入只读寄存器 (motor_speed) ===")
        ctrl_frame = PpxRegionCtrl()
        ctrl_frame.msg.id = PPX_ID_MCB
        ctrl_frame.msg.cmd = PPX_MSG_WRITE
        ctrl_frame.msg.reg_addr = PPX_MOTOR_SPEED_REG
        ctrl_frame.msg.reg_nums = 1
        ctrl_frame.data.motor_speed = 1234
        rx_bytes = self._send_and_receive(ctrl_frame, "写只读")
        if rx_bytes is None:
            self._record_test_result("写入只读寄存器", False, "通信失败")
            return False
        try:
            raw_rx_ptr = (ctypes.c_uint8 * len(rx_bytes))(*rx_bytes)
            packet_layer_output = PpxPacketData()
            packet_status = self.dll.ppx_com_packet_parse(raw_rx_ptr, len(rx_bytes), ctypes.byref(packet_layer_output))
            if packet_status != 1:
                self._record_test_result("写入只读寄存器", False, "链路解析失败")
                return False
            if (packet_layer_output.cmd & PPX_CMD_EXCP) == PPX_CMD_EXCP:
                self._record_test_result("写入只读寄存器", True, "MCB返回异常帧，写入被拒绝")
                return True
            else:
                self._record_test_result("写入只读寄存器", False, "MCB未拒绝(固件缺陷)")
                return False
        except Exception as e:
            self._record_test_result("写入只读寄存器", False, f"解析异常:{e}")
            return False

    def _test_invalid_mode_number(self) -> bool:
        logging.info("=== 测试: 无效运行模式 (13) ===")
        ctrl_frame = PpxRegionCtrl()
        ctrl_frame.msg.id = PPX_ID_MCB
        ctrl_frame.msg.cmd = PPX_MSG_WRITE
        ctrl_frame.msg.reg_addr = PPX_RUN_MODE_REG
        ctrl_frame.msg.reg_nums = 1
        ctrl_frame.data.run_mode = 13
        rx_bytes = self._send_and_receive(ctrl_frame, "设置模式(未知13)")
        if rx_bytes is None:
            self._record_test_result("无效运行模式", False, "通信失败")
            return False
        try:
            raw_rx_ptr = (ctypes.c_uint8 * len(rx_bytes))(*rx_bytes)
            packet_layer_output = PpxPacketData()
            packet_status = self.dll.ppx_com_packet_parse(raw_rx_ptr, len(rx_bytes), ctypes.byref(packet_layer_output))
            if packet_status != 1:
                self._record_test_result("无效运行模式", False, "链路解析失败")
                return False
            if (packet_layer_output.cmd & PPX_CMD_EXCP) == PPX_CMD_EXCP:
                self._record_test_result("无效运行模式", True, "MCB返回异常帧，模式被拒绝")
                return True
            else:
                self._record_test_result("无效运行模式", False, "MCB未拒绝(固件缺陷)")
                return False
        except Exception as e:
            self._record_test_result("无效运行模式", False, f"解析异常:{e}")
            return False

    def _test_flag_motor_stop(self) -> bool:
        logging.info("=== 测试: 电机停止标志位 ===")
        self.set_run_mode(PPX_MODE_RUNNING)
        time.sleep(0.1)
        self.set_target_speed(200)
        time.sleep(0.3)
        read_before = self._read_register(PPX_MOTOR_SPEED_REG, 1, "停止前速度")
        if read_before is None:
            self.stop_motor()
            self._record_test_result("电机停止标志位", False, "无法读取速度")
            return False
        speed_before = read_before.data.motor_speed
        logging.info(f"停止前速度: {speed_before} rpm")
        self._write_register(PPX_RT_SETTING_REG, 1, {'rt_setting': PPX_MOTOR_STOP}, "MOTOR_STOP")
        time.sleep(0.3)
        read_after = self._read_register(PPX_MOTOR_SPEED_REG, 1, "停止后速度")
        if read_after is None:
            self.stop_motor()
            self._record_test_result("电机停止标志位", False, "无法读取速度")
            return False
        speed_after = read_after.data.motor_speed
        logging.info(f"停止后速度: {speed_after} rpm")
        speed_reduced = abs(speed_after) < max(10, abs(speed_before) * 0.5)
        self.stop_motor()
        self._record_test_result("电机停止标志位", speed_reduced, f"速度从{speed_before}->{speed_after}")
        return speed_reduced

    def _test_rapid_commands(self) -> bool:
        logging.info("=== 测试: 指令洪泛稳定性 ===")
        self.set_run_mode(PPX_MODE_RUNNING)
        time.sleep(0.1)
        all_ok = True
        for i in range(10):
            s = 50 + (i * 30) % 500
            ok = self.set_target_speed(s)
            if not ok:
                all_ok = False
            time.sleep(0.05)
        final_ok = self.read_motor_status()
        self.set_run_mode(PPX_MODE_IDLE)
        all_ok = all_ok and final_ok
        self._record_test_result("指令洪泛稳定性", all_ok, "连续10次快速写" if all_ok else "通信失败")
        return all_ok

    def _test_read_all_status_registers(self) -> bool:
        logging.info("=== 测试: 全状态寄存器扫描 ===")
        result = self._read_register(PPX_MCU_ERRCODE_REG, 25, "全状态")
        if result is None:
            self._record_test_result("全状态寄存器扫描", False, "读取失败")
            return False
        data = result.data
        checks = [
            0 <= data.bus_voltage * 0.1 <= 60,
            0 <= data.bus_current * 0.1 <= 30,
            -10000 <= data.target_speed <= 10000,
            0 <= data.target_accel <= 20000,
            -10000 <= data.motor_speed <= 10000
        ]
        reasonable = all(checks)
        self._record_test_result("全状态寄存器扫描", reasonable, "字段合理" if reasonable else "字段异常")
        return reasonable

    def _test_response_time(self) -> bool:
        logging.info("=== 测试: 通信响应时间 ===")
        start = time.time()
        result = self._read_register(PPX_MCU_ERRCODE_REG, 1, "计时读")
        elapsed = (time.time() - start) * 1000
        if result is None:
            self._record_test_result("通信响应时间", False, "读失败")
            return False
        passed = elapsed < 200
        self._record_test_result("通信响应时间", passed, f"{elapsed:.1f} ms" + ("(超时)" if not passed else ""))
        return passed

    # ---------- 稳定性专项测试 ----------
    def _health_check(self) -> bool:
        """单次健康检查：读故障码寄存器，若连续失败则标记疑似死机"""
        result = self._read_register(PPX_MCU_ERRCODE_REG, 1, "健康检查")
        if result is None:
            self._consecutive_failures += 1
            logging.warning(f"健康检查失败，连续失败次数: {self._consecutive_failures}")
            if self._consecutive_failures >= 5:
                self._crash_detected = True
                logging.critical("!!! 疑似 MCB 死机/无响应（连续5次健康检查失败）!!!")
            return False
        self._consecutive_failures = 0
        return True

    def _reset_stability_stats(self) -> None:
        self._crash_detected = False
        self._consecutive_failures = 0
        self._total_stress_commands = 0
        self._stress_failures = 0

    def _stress_write(self, label: str, reg_addr: int, data_fields: Dict[str, Any]) -> bool:
        """记录统计的写操作"""
        self._total_stress_commands += 1
        result = self._write_register(reg_addr, 1, data_fields, label)
        if not result:
            self._stress_failures += 1
        return result

    def _stress_read(self, label: str, reg_addr: int) -> Optional[PpxRegionCtrl]:
        """记录统计的读操作"""
        self._total_stress_commands += 1
        result = self._read_register(reg_addr, 1, label)
        if result is None:
            self._stress_failures += 1
        return result

    # ---------- 稳定性测试用例 ----------
    def _test_stability_long_run(self) -> bool:
        """长时间运行稳定性：固定中速连续运行，周期性读取状态"""
        cycles = STABILITY_LONG_RUN_CYCLES
        logging.info(f"=== 稳定性测试: 长时间运行 {cycles} 个周期 ===")
        self.stop_motor()
        time.sleep(0.3)
        self.clear_error_code()
        time.sleep(0.1)
        self._reset_stability_stats()

        self._write_register(PPX_TARGET_SPEED_REG, 3, {
            'target_speed': 250, 'target_accel': 500, 'target_current': 0
        }, "长跑预设")
        time.sleep(0.1)

        if not self.set_run_mode(PPX_MODE_RUNNING):
            self._record_test_result("长时间运行稳定性", False, "无法进入骑行模式")
            self.stop_motor()
            return False

        failures_in_loop = 0
        for cycle in range(1, cycles + 1):
            if self._crash_detected:
                logging.critical(f"第 {cycle}/{cycles} 周期检测到死机，终止测试。")
                break
            ok = self._health_check()
            if not ok:
                failures_in_loop += 1
            if cycle % 20 == 0:
                logging.info(f"长时间运行进度: {cycle}/{cycles} 周期, 当前累计失败: {failures_in_loop}")
            time.sleep(STABILITY_HEARTBEAT_INTERVAL)

        self.stop_motor()
        time.sleep(0.2)

        crashed = self._crash_detected
        passed = (not crashed) and (failures_in_loop <= cycles * 0.05)
        details = (f"{cycles}周期, 失败{failures_in_loop}次" +
                   (" | 检测到死机" if crashed else ""))
        self._record_test_result("长时间运行稳定性", passed, details)
        return passed

    def _test_stability_burst_commands(self) -> bool:
        """指令洪泛压力测试：短时间内连续发送大量读写指令"""
        count = STABILITY_BURST_COUNT
        logging.info(f"=== 稳定性测试: 指令洪泛压力 {count} 次 ===")
        self.stop_motor()
        time.sleep(0.3)
        self.clear_error_code()
        time.sleep(0.1)
        self._reset_stability_stats()

        self.set_run_mode(PPX_MODE_RUNNING)
        time.sleep(0.1)

        for i in range(count):
            if self._crash_detected:
                break
            if i % 3 == 0:
                speed = (i % 600) + 50
                self._stress_write(f"洪泛速度{speed}", PPX_TARGET_SPEED_REG, {'target_speed': speed})
            elif i % 3 == 1:
                self._stress_read("洪泛读状态", PPX_MCU_ERRCODE_REG)
            else:
                self._health_check()
            time.sleep(0.01)

        self.stop_motor()
        time.sleep(0.2)

        crashed = self._crash_detected
        failure_rate = (self._stress_failures / max(1, self._total_stress_commands)) * 100
        passed = (not crashed) and (failure_rate <= 5.0)
        details = (f"{self._total_stress_commands}次, 失败{self._stress_failures}({failure_rate:.1f}%)" +
                   (" | 死机" if crashed else ""))
        self._record_test_result("指令洪泛压力", passed, details)
        return passed

    def _test_stability_idle_recovery(self) -> bool:
        """长时间空闲后恢复：空闲等待，然后测试通信是否正常"""
        logging.info(f"=== 稳定性测试: 空闲恢复 (等待 {STABILITY_IDLE_RECOVERY_WAIT}s) ===")
        self.stop_motor()
        time.sleep(0.3)
        self.clear_error_code()
        time.sleep(0.1)

        logging.info(f"进入空闲状态，等待 {STABILITY_IDLE_RECOVERY_WAIT} 秒...")
        time.sleep(STABILITY_IDLE_RECOVERY_WAIT)

        self._reset_stability_stats()
        ok1 = self._health_check()
        ok2 = self._health_check()
        ok3 = self._stress_read("空闲后读全状态", PPX_MCU_ERRCODE_REG)

        passed = ok1 and ok2 and (ok3 is not None)
        self._record_test_result("长时间空闲后恢复", passed,
                                 "正常" if passed else f"健康检查失败(ok1={ok1},ok2={ok2},ok3={ok3 is not None})")
        return passed

    def _test_stability_mode_switch_stress(self) -> bool:
        """反复模式切换压力：在空闲和骑行之间反复切换"""
        rounds = STABILITY_MODE_SWITCH_ROUNDS
        logging.info(f"=== 稳定性测试: 模式反复切换 {rounds} 轮 ===")
        self.stop_motor()
        time.sleep(0.3)
        self.clear_error_code()
        time.sleep(0.1)
        self._reset_stability_stats()

        modes = [(PPX_MODE_IDLE, "空闲"), (PPX_MODE_RUNNING, "骑行"),
                 (PPX_MODE_LOCK, "锁车"), (PPX_MODE_PWR_PUSH, "推行")]
        failures = 0
        for rnd in range(rounds):
            if self._crash_detected:
                break
            for mode, name in modes:
                ok = self.set_run_mode(mode)
                if not ok:
                    failures += 1
                time.sleep(0.08)
                if not self._health_check():
                    failures += 1
                time.sleep(0.05)
            if (rnd + 1) % 5 == 0:
                logging.info(f"模式切换进度: {rnd + 1}/{rounds} 轮, 失败: {failures}")

        self.stop_motor()
        time.sleep(0.2)

        crashed = self._crash_detected
        passed = (not crashed) and (failures <= rounds * 0.1)
        details = f"{rounds}轮共{rounds * len(modes)}次切换, 失败{failures}次" + (" | 死机" if crashed else "")
        self._record_test_result("模式反复切换压力", passed, details)
        return passed

    def _test_stability_ramp_repeat(self) -> bool:
        """斜坡往复循环：多次上下斜坡，检测运行中是否失联"""
        cycles = STABILITY_RAMP_CYCLES
        logging.info(f"=== 稳定性测试: 斜坡往复 {cycles} 次 ===")
        self.stop_motor()
        time.sleep(0.3)
        self.clear_error_code()
        time.sleep(0.1)
        self._reset_stability_stats()

        failures = 0
        for cyc in range(1, cycles + 1):
            if self._crash_detected:
                break
            logging.info(f"斜坡往复 第 {cyc}/{cycles} 次")
            try:
                ok = self.ramp_test(max_speed=300, step=100, dwell=0.3)
                if not ok:
                    failures += 1
            except Exception as e:
                failures += 1
                logging.error(f"斜坡往复异常: {e}")
            time.sleep(0.2)
            if not self._health_check():
                failures += 1
            time.sleep(0.1)

        self.stop_motor()
        time.sleep(0.2)

        crashed = self._crash_detected
        passed = (not crashed) and (failures <= cycles * 0.2)
        details = f"{cycles}次往复, 失败{failures}次" + (" | 死机" if crashed else "")
        self._record_test_result("斜坡往复循环", passed, details)
        return passed

    def _test_stability_boundary_speed_jump(self) -> bool:
        """边界速度突变：反复在极低和极高速度间跳变"""
        logging.info("=== 稳定性测试: 边界速度跳变 ===")
        self.stop_motor()
        time.sleep(0.3)
        self.clear_error_code()
        time.sleep(0.1)
        self._reset_stability_stats()

        if not self.set_run_mode(PPX_MODE_RUNNING):
            self._record_test_result("边界速度跳变", False, "无法进入骑行")
            return False
        time.sleep(0.1)

        speed_pairs = [(0, 600), (-200, 200), (0, 800), (-100, 500)]
        failures = 0
        for low, high in speed_pairs:
            if self._crash_detected:
                break
            for _ in range(10):
                if not self._stress_write(f"跳变{high}", PPX_TARGET_SPEED_REG, {'target_speed': high}):
                    failures += 1
                time.sleep(0.03)
                if not self._stress_write(f"跳变{low}", PPX_TARGET_SPEED_REG, {'target_speed': low}):
                    failures += 1
                time.sleep(0.03)
            self._health_check()

        self.stop_motor()
        time.sleep(0.2)

        crashed = self._crash_detected
        total_ops = len(speed_pairs) * 20
        passed = (not crashed) and (failures <= total_ops * 0.05)
        details = f"{total_ops}次跳变, 失败{failures}次" + (" | 死机" if crashed else "")
        self._record_test_result("边界速度跳变", passed, details)
        return passed

    def _test_stability_noise_communication(self) -> bool:
        """通信噪声容限：发送不完整/异常帧后检查MCB是否仍可正常通信"""
        logging.info("=== 稳定性测试: 通信噪声容限 ===")
        self.stop_motor()
        time.sleep(0.3)
        self.clear_error_code()
        time.sleep(0.1)

        if not self.serial_conn or not self.serial_conn.is_open:
            self._record_test_result("通信噪声容限", False, "串口未打开")
            return False

        # 发送垃圾数据
        garbage_packets = [
            bytes([0x00, 0x00, 0x00]),
            bytes([0xA5, 0x00, 0x00, 0x00, 0x55]),
            bytes([0xFF] * 20),
            bytes([0xA5, 0x20, 0x80, 0x00, 0x00, 0x55]),
        ]

        all_recovered = True
        for idx, garbage in enumerate(garbage_packets):
            try:
                self.serial_conn.write(garbage)
                self.serial_conn.flush()
                time.sleep(0.1)
            except Exception as e:
                logging.error(f"发送噪声包失败: {e}")

            # 检查是否仍能正常读取
            recovered = self._health_check()
            if not recovered:
                logging.warning(f"噪声包 #{idx} 后通信异常")
                all_recovered = False
            time.sleep(0.2)

        # 最终确认恢复
        final_ok = self.read_motor_status()
        passed = all_recovered and final_ok
        self._record_test_result("通信噪声容限", passed,
                                 "MCB正常恢复" if passed else "MCB通信异常/死机")
        return passed

    def _test_stability_continuous_health_monitor(self) -> bool:
        """持续心跳监控：后台持续读取状态，统计长时间内的响应成功率"""
        duration = 30  # 秒
        logging.info(f"=== 稳定性测试: 持续心跳监控 {duration}s ===")
        self.stop_motor()
        time.sleep(0.3)
        self._reset_stability_stats()

        self.set_run_mode(PPX_MODE_RUNNING)
        time.sleep(0.1)
        self._stress_write("心跳速度200", PPX_TARGET_SPEED_REG, {'target_speed': 200})

        start = time.time()
        heartbeats_sent = 0
        heartbeats_failed = 0
        while time.time() - start < duration:
            if self._crash_detected:
                break
            ok = self._health_check()
            heartbeats_sent += 1
            if not ok:
                heartbeats_failed += 1
            time.sleep(STABILITY_HEARTBEAT_INTERVAL)

        self.stop_motor()
        time.sleep(0.2)

        crashed = self._crash_detected
        fail_rate = (heartbeats_failed / max(1, heartbeats_sent)) * 100
        passed = (not crashed) and (fail_rate <= 2.0)
        details = (f"{duration}s内{heartbeats_sent}次心跳, 失败{heartbeats_failed}({fail_rate:.1f}%)" +
                   (" | 死机" if crashed else ""))
        self._record_test_result("持续心跳监控", passed, details)
        return passed

    # ---------- 全自动稳定性测试主入口 ----------
    def run_automated_tests(self) -> bool:
        logging.info("=" * 60)
        logging.info("=== 全自动白盒稳定性测试 ===")
        logging.info(f"日志文件: {LOG_FILENAME}")
        logging.info("=" * 60)
        self.test_results = []
        start_total = time.time()

        # ---- 阶段一：基础连接与功能快速检查 ----
        logging.info(">>> 阶段一: 基础连接与功能快速检查 <<<")
        if not self.read_motor_status():
            logging.critical("基础连接失败，终止测试。")
            self._record_test_result("基础连接验证", False, "无法读取状态")
            return False
        self._record_test_result("基础连接验证", True, "OK")
        time.sleep(0.2)

        self._test_clear_error()
        time.sleep(0.1)
        self._test_read_state()
        self._test_brake_state()
        time.sleep(0.2)

        self._test_all_modes()
        time.sleep(0.2)

        for sp in [100, 300, 600]:
            self._test_speed_control(sp)
            time.sleep(0.2)

        for cur in [50, 100]:
            self._test_current_control(cur)
            time.sleep(0.2)

        self._test_ramp_profile()
        time.sleep(0.2)
        self._test_param_boundary()
        time.sleep(0.2)

        self._test_read_invalid_register()
        self._test_write_invalid_register()
        self._test_write_readonly_reg()
        self._test_invalid_mode_number()
        self._test_flag_motor_stop()
        self._test_rapid_commands()
        self._test_read_all_status_registers()
        self._test_response_time()

        # ---- 阶段二：深度稳定性压力测试 ----
        logging.info(">>> 阶段二: 深度稳定性压力测试 <<<")
        logging.info(f"配置: 长跑{STABILITY_LONG_RUN_CYCLES}周期 | 洪泛{STABILITY_BURST_COUNT}次 | 空闲{STABILITY_IDLE_RECOVERY_WAIT}s | 斜坡{STABILITY_RAMP_CYCLES}轮 | 模式{STABILITY_MODE_SWITCH_ROUNDS}轮")

        self._test_stability_long_run()
        time.sleep(0.3)

        self._test_stability_burst_commands()
        time.sleep(0.3)

        self._test_stability_idle_recovery()
        time.sleep(0.3)

        self._test_stability_mode_switch_stress()
        time.sleep(0.3)

        self._test_stability_ramp_repeat()
        time.sleep(0.3)

        self._test_stability_boundary_speed_jump()
        time.sleep(0.3)

        self._test_stability_noise_communication()
        time.sleep(0.3)

        self._test_stability_continuous_health_monitor()
        time.sleep(0.3)

        # ---- 最终收尾 ----
        logging.info(">>> 最终安全停止 <<<")
        self.stop_motor()
        time.sleep(0.3)
        final_status = self.read_motor_status()
        self._record_test_result("最终安全停止", final_status, "空闲状态")

        # ---- 报告汇总 ----
        elapsed = time.time() - start_total
        total = len(self.test_results)
        passed = sum(1 for r in self.test_results if r["status"] == "PASS")
        failed = total - passed

        logging.info("=" * 60)
        logging.info("=== 稳定性测试报告 ===")
        logging.info(f"总耗时: {elapsed:.2f}s | 用例总数: {total} | 通过: {passed} | 失败: {failed}")
        pass_rate = (passed / total * 100) if total else 0
        logging.info(f"通过率: {pass_rate:.1f}%")
        for i, r in enumerate(self.test_results, 1):
            status_marker = "PASS" if r["status"] == "PASS" else "FAIL"
            logging.info(f"  [{status_marker}] #{i:02d} {r['name']} - {r['details']}")

        if failed:
            logging.warning(f"存在 {failed} 个失败用例，请检查日志: {LOG_FILENAME}")
            logging.info("--- 失败用例详细列表 ---")
            failed_cases = [r for r in self.test_results if r["status"] == "FAIL"]
            for case in failed_cases:
                logging.info(f"  * {case['name']} | {case['details']} | {case['timestamp']}")
            logging.info("------------------------")
        else:
            logging.info("所有用例通过，MCB稳定性良好。")
        logging.info("=" * 60)
        logging.info(f"完整日志已保存至: {LOG_FILENAME}")
        logging.info("=" * 60)

        return failed == 0

    def close(self) -> None:
        if self.serial_conn and self.serial_conn.is_open:
            self.serial_conn.close()
            logging.info("串口已关闭。")

# ==========================================
# 交互式菜单
# ==========================================
def print_menu() -> None:
    menu = (
        "\n==================================================\n"
        "  MCB 电机控制菜单 (四轮代步车/轮椅测试)\n"
        "==================================================\n"
        "  0. 读取电机状态\n"
        "  1. 启动电机 (低速 100rpm)\n"
        "  2. 启动电机 (中速 300rpm)\n"
        "  3. 启动电机 (自定义速度)\n"
        "  4. 停止电机\n"
        "  5. 速度斜坡测试\n"
        "  6. 设置目标电流\n"
        "  7. 清除故障码\n"
        "  8. 切换运行模式\n"
        "  9. 读取电磁刹状态\n"
        "  a. 执行全自动稳定性测试\n"
        "  q. 退出\n"
        "--------------------------------------------------"
    )
    logging.info(menu)

def interactive_mode(controller: MCBMotorController) -> None:
    while True:
        print_menu()
        try:
            ch = input("请选择操作: ").strip().lower()
            if ch == 'q':
                logging.info("退出程序。")
                break
            elif ch == '0':
                controller.read_motor_status()
            elif ch == '1':
                controller.start_motor_speed_mode(100)
            elif ch == '2':
                controller.start_motor_speed_mode(300)
            elif ch == '3':
                speed = int(input("目标速度 (rpm): "))
                controller.start_motor_speed_mode(speed)
            elif ch == '4':
                controller.stop_motor()
            elif ch == '5':
                mx = input("最大速度 (默认500): ")
                mx = int(mx) if mx else 500
                st = input("步进 (默认50): ")
                st = int(st) if st else 50
                dw = input("停留秒数 (默认1.0): ")
                dw = float(dw) if dw else 1.0
                controller.ramp_test(mx, st, dw)
            elif ch == '6':
                cur_a = float(input("目标电流 (A): "))
                controller.set_target_current(int(cur_a * 10))
            elif ch == '7':
                controller.clear_error_code()
            elif ch == '8':
                logging.info("\n运行模式: 0-空闲 2-骑行 4-推行 10-延迟关刹 11-立即关刹 12-解锁")
                mode = int(input("模式编号: "))
                controller.set_run_mode(mode)
            elif ch == '9':
                controller.read_brake_state()
            elif ch == 'a':
                controller.run_automated_tests()
            else:
                logging.warning("未知选项。")
        except ValueError as e:
            logging.error(f"输入错误: {e}")
        except Exception as e:
            logging.error(f"交互异常: {e}")

# ==========================================
# 主入口
# ==========================================
def main() -> None:
    TARGET_COM_PORT = "COM27"
    TARGET_DLL_PATH = r"C:\Users\szm21\Downloads\Script\R3_MCB白盒测试\ppx_region.dll"

    parser = argparse.ArgumentParser(description="MCB电机控制 - 全自动稳定性测试")
    parser.add_argument("--port", default=TARGET_COM_PORT, help="串口号")
    parser.add_argument("--dll", default=TARGET_DLL_PATH, help="DLL路径")
    parser.add_argument("--interactive", action="store_true", default=True, help="交互模式(默认)")
    parser.add_argument("--auto", type=int, nargs='?', const=200, metavar="SPEED", help="启动到指定rpm")
    parser.add_argument("--ramp", type=int, nargs='?', const=500, metavar="MAX", help="斜坡测试")
    parser.add_argument("--stop", action="store_true", help="停止电机")
    parser.add_argument("--status", action="store_true", help="读取状态")
    parser.add_argument("--fulltest", action="store_true", help="执行全自动稳定性测试")
    parser.add_argument("--stress", action="store_true", help="仅执行稳定性压力测试(跳过基础检查)")
    args = parser.parse_args()

    controller = MCBMotorController(com_port=args.port, dll_path=args.dll)
    if not controller.init_environment():
        logging.error("初始化失败，退出。")
        return

    try:
        controller.clear_error_code()
        time.sleep(0.1)

        if args.stop:
            controller.stop_motor()
        elif args.status:
            controller.read_motor_status()
        elif args.stress:
            logging.info("=== 仅执行稳定性压力测试 ===")
            controller.test_results = []
            controller._test_stability_long_run()
            time.sleep(0.3)
            controller._test_stability_burst_commands()
            time.sleep(0.3)
            controller._test_stability_idle_recovery()
            time.sleep(0.3)
            controller._test_stability_mode_switch_stress()
            time.sleep(0.3)
            controller._test_stability_ramp_repeat()
            time.sleep(0.3)
            controller._test_stability_boundary_speed_jump()
            time.sleep(0.3)
            controller._test_stability_noise_communication()
            time.sleep(0.3)
            controller._test_stability_continuous_health_monitor()
            controller.stop_motor()
            total = len(controller.test_results)
            passed = sum(1 for r in controller.test_results if r["status"] == "PASS")
            logging.info(f"=== 稳定性压力测试完成: {passed}/{total} 通过 ===")
        elif args.fulltest:
            controller.run_automated_tests()
        elif args.auto is not None:
            controller.start_motor_speed_mode(args.auto)
        elif args.ramp is not None:
            controller.ramp_test(args.ramp)
        else:
            interactive_mode(controller)
    except KeyboardInterrupt:
        logging.info("用户中断，紧急停止...")
        controller.stop_motor()
    except Exception as e:
        logging.error(f"主流程异常: {e}")
    finally:
        controller.close()

if __name__ == "__main__":
    main()