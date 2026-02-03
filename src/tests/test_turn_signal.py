# src/tests/test_turn_signal.py
import time
from src import config
from src.tests.base_test import BaseTest


class TurnSignalTest(BaseTest):
    """
    转向灯压力测试 (整合左/右灯)
    """

    def run(self, loops, side="left"):
        """
        :param loops: 循环次数
        :param side: "left" 或 "right"
        """
        relay = self.drivers.get('relay')
        if not relay:
            self.logger.error("缺少继电器驱动！")
            return

        # 根据传入的 side 决定用哪个指令
        cmd_on = config.COMMANDS['LEFT_TURN_ON'] if side == "left" else config.COMMANDS['RIGHT_TURN_ON']
        cmd_off = config.COMMANDS['TURN_OFF']

        self.logger.info(f"=== 开始 {side} 转向灯测试，目标 {loops} 次 ===")

        for i in range(1, loops + 1):
            if self.stop_flag: break

            # 1. 亮灯
            relay.send_bytes(cmd_off, "复位")  # 先复位保险
            time.sleep(0.1)
            relay.send_bytes(cmd_on, f"{side} 灯亮")

            # 2. 保持
            time.sleep(2.0)  # 原脚本是 sleep(2)

            # 3. 灭灯
            relay.send_bytes(cmd_off, "灭灯")
            time.sleep(0.1)

            if i % 10 == 0:
                self.logger.info(f"进度: {i}/{loops}")

        self.logger.info(f"{side} 转向灯测试完成")