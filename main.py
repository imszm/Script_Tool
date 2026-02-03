# main.py
import time
from src import config
from src.logger import setup_logging
from src.driver import SerialDriver
# 导入我们写好的测试类
from src.tests.test_horn import HornTest
from src.tests.test_power import PowerCycleTest


def main():
    logger = setup_logging()
    logger.info(">>> 自动化测试框架启动 <<<")

    # 1. 统一初始化所有硬件
    drivers = {}

    # 连接继电器
    relay = SerialDriver(config.RELAY_PORT, config.BAUDRATE_RELAY, "RelayBox")
    if relay.connect():
        drivers['relay'] = relay

    # 连接被测设备 (如果跑开关机测试需要这个)
    dut = SerialDriver(config.DEVICE_PORT, config.BAUDRATE_DEVICE, "CarDevice")
    if dut.connect():
        drivers['device'] = dut

    try:
        # ==========================================
        # 在这里选择你要跑的任务！
        # ==========================================

        # 选项 A: 跑喇叭测试
        #current_test = HornTest(drivers)

        # 选项 B: 跑开关机测试
        current_test = PowerCycleTest(drivers)

        # 开始运行
        current_test.setup()
        current_test.run(loops=10)  # 运行10次看看
        current_test.teardown()

    except KeyboardInterrupt:
        logger.warning("用户强制停止")
    except Exception as e:
        logger.error(f"发生未处理异常: {e}")
    finally:
        # 结束时自动关闭所有串口
        for d in drivers.values():
            d.close()


if __name__ == "__main__":
    main()