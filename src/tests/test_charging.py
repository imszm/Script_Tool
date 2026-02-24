import time
import random
from src import config
from src.tests.base_test import BaseTest


class ChargingTest(BaseTest):
    """
    对应原脚本: 继电器充电压力测试.py
    逻辑: 随机时间充电 -> 读日志 -> 断电 -> 读日志
    """

    def run(self, loops):
        relay = self.drivers['relay']
        dut = self.drivers['device']

        success_cnt = 0

        # 初始化
        relay.send_bytes(config.COMMANDS['HEX_ENABLE'])
        time.sleep(1)

        for i in range(1, loops + 1):
            if self.stop_flag: break
            self.logger.info(f"--- 充电测试 Cycle {i} ---")

            # 1. 开启充电 (0x4F)
            relay.send_bytes(config.COMMANDS['HEX_RELEASE_OFF'], "开始充电")
            # 随机充电 3-5秒
            time.sleep(random.uniform(3, 5))

            # 读取日志 (阶段1)
            log1 = dut.read_buffer()

            # 2. 关闭充电 (0x50)
            relay.send_bytes(config.COMMANDS['HEX_PRESS_ON'], "断开充电")
            time.sleep(5)  # 关机等待

            # 读取日志 (阶段2)
            log2 = dut.read_buffer()

            # 3. 合并分析
            full_log = (log1 + log2).replace(" ", "").lower()

            # 判定
            found_success = any(k in full_log for k in config.KEYWORDS['CHARGE_SUCCESS'])
            found_error = any(k in full_log for k in config.KEYWORDS['CHARGE_ERROR'])

            if found_error:
                self.logger.error("发现严重错误(Assertion Failed)!")
            elif found_success:
                success_cnt += 1
                self.logger.info("结果: 成功 (检测到语音消息)")
            else:
                self.logger.warning("结果: 失败 (未检测到关键字)")

        self.logger.info(f"充电测试结束，成功率: {success_cnt}/{loops}")