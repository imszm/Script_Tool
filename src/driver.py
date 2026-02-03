import serial
import time
import logging

class SerialDriver:
    """
    通用串口驱动：既能控制继电器，也能监听设备日志
    """
    def __init__(self, port, baudrate, name="SerialDev"):
        self.port = port
        self.baudrate = baudrate
        self.name = name
        self.ser = None
        self.logger = logging.getLogger(f"AutoTest.Driver.{name}")

    def connect(self):
        try:
            # timeout设置短一点，方便读取循环不卡死
            self.ser = serial.Serial(self.port, self.baudrate, timeout=0.1)
            self.logger.info(f"[{self.name}] 串口 {self.port} 连接成功")
            return True
        except Exception as e:
            self.logger.error(f"[{self.name}] 连接失败: {e}")
            return False

    def send_bytes(self, cmd_bytes, desc=""):
        """发送指令 (Write)"""
        if self.ser and self.ser.is_open:
            try:
                self.ser.write(cmd_bytes)
                # self.logger.debug(f"发送[{desc}]")
                return True
            except Exception as e:
                self.logger.error(f"发送失败: {e}")
        return False

    def read_line(self):
        """读取一行日志 (Read) - 用于开关机检测"""
        if self.ser and self.ser.is_open and self.ser.in_waiting:
            try:
                # 忽略解码错误，防止乱码导致崩溃
                return self.ser.readline().decode('utf-8', errors='ignore').strip()
            except Exception:
                pass
        return None

    def close(self):
        if self.ser:
            self.ser.close()