# src/driver.py
import serial
import time
import logging

class RelayController:
    def __init__(self, port, baudrate):
        self.port = port
        self.baudrate = baudrate
        self.ser = None
        # 获取 logger 实例，名字设为 AutoTest.Driver
        self.logger = logging.getLogger("AutoTest.Driver")

    def connect(self):
        """连接设备"""
        try:
            self.ser = serial.Serial(self.port, self.baudrate, timeout=1)
            self.logger.info(f"串口 {self.port} 连接成功")
            return True
        except serial.SerialException as e:
            self.logger.error(f"串口连接失败: {e}")
            return False

    def send_cmd(self, cmd_bytes, description=""):
        """
        发送指令的通用方法
        :param cmd_bytes: 16进制指令
        :param description: 指令描述（用于日志记录）
        """
        if self.ser and self.ser.is_open:
            try:
                self.ser.write(cmd_bytes)
                # 记录日志，而不是 print
                # self.logger.debug(f"发送指令: {description}") # 如果嫌这句太吵可以注释掉
                return True
            except Exception as e:
                self.logger.error(f"发送指令失败 [{description}]: {e}")
                return False
        else:
            self.logger.warning("设备未连接，无法发送指令")
            return False

    def close(self):
        """安全关闭"""
        if self.ser and self.ser.is_open:
            self.ser.close()
            self.logger.info("串口已关闭")