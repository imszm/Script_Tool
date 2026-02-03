# src/tests/test_nfc_servo.py
import time
from src import config
from src.tests.base_test import BaseTest


class NfcServoTest(BaseTest):
    def run(self, loops):
        servo = self.drivers.get('servo')  # 我们需要在 main.py 里注册这个驱动
        device = self.drivers.get('device')  # 监听设备日志

        if not servo or not device:
            self.logger.error("NFC测试需要 舵机 和 设备 两个驱动！")
            return

        self.logger.info("初始化舵机位置...")
        servo.send_bytes(config.SERVO_CMDS['NFC_HIGH'], "舵机抬起")
        time.sleep(2)

        success_count = 0

        for i in range(1, loops + 1):
            self.logger.info(f"--- NFC 测试循环 {i} ---")

            # 1. 舵机下压 (刷卡)
            servo.send_bytes(config.SERVO_CMDS['NFC_LOW'], "NFC下压")

            # 2. 监听日志 (检查是否识别到卡/开机)
            # 我们给它 3 秒钟时间去检测日志
            check_start = time.time()
            found_signal = False

            while time.time() - check_start < 3.0:
                line = device.read_line()
                if line:
                    # 简单判断：只要出现 "nfc 1:on" 就算成功
                    if config.NFC_KEYWORDS['ON'] in line:
                        found_signal = True
                        self.logger.info(f"检测到NFC触发: {line}")
                        break

            if found_signal:
                success_count += 1
            else:
                self.logger.warning("本轮未检测到NFC日志")

            # 3. 舵机抬起
            servo.send_bytes(config.SERVO_CMDS['NFC_HIGH'], "NFC抬起")
            time.sleep(1.5)  # 等待机械动作完成

        self.logger.info(f"NFC测试结束，成功: {success_count}/{loops}")