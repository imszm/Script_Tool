# src/driver.py
import serial
import time
import logging

class SerialDriver:
    """
    通用串口驱动：支持继电器控制(Write)和设备日志监听(Read)
    """
    def __init__(self, port, baudrate, name="SerialDev"):
        self.port = port
        self.baudrate = baudrate
        self.name = name
        self.ser = None
        self.logger = logging.getLogger(f"AutoTest.Driver.{name}")

    def connect(self):
        try:
            # timeout=0.1 保证读取时不卡死
            self.ser = serial.Serial(self.port, self.baudrate, timeout=0.1)
            self.logger.info(f"[{self.name}] 串口 {self.port} 连接成功")
            return True
        except Exception as e:
            self.logger.error(f"[{self.name}] 连接失败: {e}")
            return False

    def send_bytes(self, cmd_bytes, desc=""):
        """发送指令"""
        if self.ser and self.ser.is_open:
            try:
                self.ser.write(cmd_bytes)
                # self.logger.debug(f"发送[{desc}]")
                return True
            except Exception as e:
                self.logger.error(f"发送失败: {e}")
        return False

    def read_line(self):
        """读取一行 (用于W3测试监听关键字)"""
        if self.ser and self.ser.is_open and self.ser.in_waiting:
            try:
                # 忽略解码错误
                return self.ser.readline().decode('utf-8', errors='ignore').strip()
            except:
                pass
        return None

    def read_buffer(self):
        """一次性读取缓冲区所有内容 (用于充电测试)"""
        if self.ser and self.ser.is_open:
            try:
                if self.ser.in_waiting > 0:
                    raw = self.ser.read(self.ser.in_waiting)
                    # 尝试多种解码
                    try:
                        return raw.decode("utf-8", errors="ignore")
                    except:
                        return raw.decode("latin1", errors="ignore")
            except Exception as e:
                self.logger.error(f"读取缓冲异常: {e}")
        return ""

    def close(self):
        if self.ser:
            self.ser.close()