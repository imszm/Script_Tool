# src/tests/test_horn.py
import time
from src import config
from src.tests.base_test import BaseTest


class HornTest(BaseTest):
    def run(self, loops):
        relay = self.drivers.get('relay')
        if not relay:
            self.logger.error("喇叭测试需要继电器驱动！")
            return

        self.logger.info(f"=== 开始喇叭压力测试，目标: {loops} 次 ===")

        for i in range(1, loops + 1):
            # 1. 按压
            relay.send_bytes(config.COMMANDS['HORN_PRESS'], "喇叭响")
            time.sleep(0.2)

            # 2. 松开
            relay.send_bytes(config.COMMANDS['HORN_RELEASE'], "喇叭停")
            time.sleep(1.0)

            if i % 10 == 0:
                self.logger.info(f"进度: {i}/{loops}")

        self.logger.info("喇叭测试完成")