import time
from src import config
from src.tests.base_test import BaseTest


class W3PowerTest(BaseTest):
    """
    对应原脚本: W3继电器开关机压力测试.py
    逻辑: 长按开机 -> 监控日志 -> 短按关机
    """

    def relay_action(self, duration, action_name):
        """模拟按键: 按下(0x50) -> 保持 -> 松开(0x4F)"""
        relay = self.drivers['relay']
        relay.send_bytes(config.COMMANDS['HEX_PRESS_ON'], f"{action_name}按下")

        start = time.time()
        while time.time() - start < duration:
            if self.stop_flag: return
            time.sleep(0.1)

        relay.send_bytes(config.COMMANDS['HEX_RELEASE_OFF'], f"{action_name}松开")

    def run(self, loops):
        dut = self.drivers.get('device')
        if not dut:
            self.logger.error("W3测试需要连接设备串口(device)！")
            return

        self.logger.info("初始化继电器...")
        self.drivers['relay'].send_bytes(config.COMMANDS['HEX_RELEASE_OFF'], "复位松开")
        self.drivers['relay'].send_bytes(config.COMMANDS['HEX_ENABLE'], "使能")
        time.sleep(1)

        for i in range(1, loops + 1):
            if self.stop_flag: break
            self.logger.info(f"=== W3 Cycle {i} ===")

            # 1. 开机 (长按2.5s)
            self.relay_action(2.5, "开机")

            # 2. 监控开机日志 (10秒)
            check_start = time.time()
            while time.time() - check_start < 10.0:
                line = dut.read_line()
                if line:
                    if config.KEYWORDS["W3_STOP"] in line:
                        self.logger.error(f"检测到停止关键字: {line}")
                        return  # 停止测试
                    if config.KEYWORDS["W3_ERROR"] in line:
                        self.logger.error(f"通信丢失: {line}")
                        return

            # 3. 关机 (短按1.0s)
            self.relay_action(1.0, "关机")
            time.sleep(3)  # 等待关机