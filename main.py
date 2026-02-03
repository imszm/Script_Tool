# main.py
from src import config
from src.logger import setup_logging
from src.driver import SerialDriver

# 导入所有测试类
from src.tests.test_horn import HornTest
from src.tests.test_power import PowerCycleTest
from src.tests.test_turn_signal import TurnSignalTest  # 新增
from src.tests.test_nfc_servo import NfcServoTest  # 新增
from src.tests.test_pc_tool import UpgradeToolTest  # 新增


def main():
    logger = setup_logging()

    # 1. 硬件连接 (按需开启)
    drivers = {}

    # 连接继电器 (绝大多数测试都需要)
    relay = SerialDriver(config.RELAY_PORT, config.BAUDRATE_RELAY, "Relay")
    if relay.connect(): drivers['relay'] = relay

    # 连接车机设备 (开关机、NFC需要)
    dut = SerialDriver(config.DEVICE_PORT, config.BAUDRATE_DEVICE, "Device")
    if dut.connect(): drivers['device'] = dut

    # 连接舵机 (只有NFC测试需要，连不上也没关系，只要不跑NFC测试就行)
    #servo = SerialDriver(config.SERVO_PORT, config.SERVO_BAUD, "Servo")
    #if servo.connect(): drivers['servo'] = servo

    # 2. 菜单选择 (这里为了演示，你可以手动解开注释)

    # === 任务 A: 喇叭 ===
    task = HornTest(drivers)
    task.run(100)

    # === 任务 B: 转向灯 (左) ===
    # task = TurnSignalTest(drivers)
    # task.run(50, side="left")

    # === 任务 C: NFC 压力测试 (用到 servo 和 device) ===
    # if 'servo' in drivers and 'device' in drivers:
    #     task = NfcServoTest(drivers)
    #     task.run(200)
    # else:
    #     logger.error("无法运行NFC测试：缺少舵机或设备串口")

    # === 任务 D: PC软件测试 (不需要串口) ===
    task = UpgradeToolTest(drivers)  # drivers传进去也没事，它不用
    # try:
    #     task.setup()
    #     task.run(10)
    # except Exception:
    #     pass

    # 3. 清理
    for d in drivers.values():
        d.close()


if __name__ == "__main__":
    main()