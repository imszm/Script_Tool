# src/tests/test_power.py
import time
from src import config
from src.tests.base_test import BaseTest


class PowerCycleTest(BaseTest):
    def run(self, loops):
        relay = self.drivers.get('relay')
        device = self.drivers.get('device')  # 需要用到第二个串口

        if not relay or not device:
            self.logger.error("开关机测试需要 继电器 和 设备 两个串口！")
            return

        success_count = 0

        for i in range(1, loops + 1):
            self.logger.info(f"--- 循环 {i}: 开机测试中 ---")

            # 1. 动作：继电器上电
            relay.send_bytes(config.COMMANDS['K2_ON'], "上电")  # 假设K2控制电源
            start_time = time.time()
            boot_success = False

            # 2. 监听：等待设备打印 "motorpoweron"
            # 设置最大等待时间 15秒
            while time.time() - start_time < 15:
                line = device.read_line()  # 从设备口读取
                if line:
                    # 检查是否包含成功关键字
                    for kw in config.POWER_TEST_KEYWORDS['SUCCESS']:
                        if kw in line.lower().replace(" ", ""):
                            boot_success = True
                            self.logger.info(f"捕获关键字: {kw} (耗时 {time.time() - start_time:.2f}s)")
                            break
                if boot_success:
                    break

            if boot_success:
                success_count += 1
            else:
                self.logger.error("失败: 开机超时，未检测到关键字")

            # 3. 动作：继电器断电
            relay.send_bytes(config.COMMANDS['K2_OFF'], "断电")
            time.sleep(2.0)  # 等待电容放电

        self.logger.info(f"测试结束。成功率: {success_count}/{loops}")