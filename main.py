# main.py
import sys
from src import config
from src.logger import setup_logging
from src.driver import SerialDriver

# 导入所有测试类
from src.tests.test_w3_power import W3PowerTest
from src.tests.test_charging import ChargingTest
from src.tests.test_pc_upgrade import PcUpgradeTest
from src.tests.test_ccb_smt import CcbSmtTest


def main():
    logger = setup_logging()
    logger.info(">>> 自动化测试框架 V2.0 启动 <<<")

    print("\n请选择要执行的测试：")
    print("1. W3 继电器开关机测试 (需连设备串口)")
    print("2. 继电器充电测试 (需连设备串口)")
    print("3. PC 升级工具测试 (Pywinauto)")
    print("4. CCB SMT 自动化测试 (COM12 + 像素识别)")

    choice = input("请输入数字: ").strip()
    loops = int(input("请输入循环次数: ").strip() or 10)

    # 1. 动态初始化驱动 (按需连接，避免端口冲突)
    drivers = {}

    # 继电器连接逻辑
    if choice == '4':
        # CCB 测试用的是 COM12
        relay = SerialDriver(config.RELAY_CCB_PORT, config.BAUDRATE_RELAY, "RelayCCB")
    else:
        # 其他测试用的是 COM4
        relay = SerialDriver(config.RELAY_PORT, config.BAUDRATE_RELAY, "RelayGen")

    if choice in ['1', '2', '4']:  # 这些都需要继电器
        if relay.connect(): drivers['relay'] = relay

    # 设备连接逻辑
    if choice in ['1', '2']:  # 只有W3和充电需要监听设备串口
        dut = SerialDriver(config.DEVICE_PORT, config.BAUDRATE_DEVICE, "Device")
        if dut.connect(): drivers['device'] = dut

    # 2. 任务分发
    task = None
    if choice == '1':
        task = W3PowerTest(drivers)
    elif choice == '2':
        task = ChargingTest(drivers)
    elif choice == '3':
        task = PcUpgradeTest(drivers)
    elif choice == '4':
        task = CcbSmtTest(drivers)

    # 3. 执行
    if task:
        try:
            task.setup()
            task.run(loops)
            task.teardown()
        except KeyboardInterrupt:
            logger.warning("用户停止")
        finally:
            for d in drivers.values(): d.close()


if __name__ == "__main__":
    main()