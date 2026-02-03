# src/config.py

# ================= 硬件连接配置 =================
RELAY_PORT = 'COM4'  # 继电器串口
DEVICE_PORT = 'COM25'  # 某些测试需要的设备串口（如开关机测试）
SERVO_PORT = 'COM6'      # 你的舵机串口号

BAUDRATE_RELAY = 9600
BAUDRATE_DEVICE = 115200
SERVO_BAUD = 115200

# ================= 指令集 (统一管理) =================
COMMANDS = {
    # 继电器基础指令
    "K2_ON": bytes.fromhex("A0 02 01 A3"),
    "K2_OFF": bytes.fromhex("A0 02 00 A2"),
    "K3_ON": bytes.fromhex("A0 03 01 A4"),
    "K3_OFF": bytes.fromhex("A0 03 00 A3"),

    # 喇叭测试专用 (参考自 tets_HornSpecificationIntervals.py)
    "HORN_PRESS": bytes([0x4F]),
    "HORN_RELEASE": bytes([0x50]),

    # 转向灯专用
    "LEFT_TURN_ON":   bytes([0x42]), # 左灯开
    "RIGHT_TURN_ON":  bytes([0x4F]), # 右灯开
    "TURN_OFF":       bytes([0x50]), # 关闭 (通用复位)

    # 大灯专用 (参考自 test_HeadlightsSpecificationIntervals.py)
    "HEADLIGHT_ON": bytes([0b10100000, 0b00000011, 0b00000001, 0b10100100]),
    "HEADLIGHT_OFF": bytes([0b10100000, 0b00000011, 0b00000000, 0b10100011]),

    # === 新增：舵机指令 (ASCII字符串转bytes) ===
    "NFC_LOW":  b"#000P2500T1000!",  # 下压刷卡
    "NFC_HIGH": b"#000P0500T1000!",  # 抬起归位

}

# ================= 业务参数 =================
# 开关机测试判定关键字 (参考自 继电器开关机压力测试.py)
POWER_TEST_KEYWORDS = {
    "SUCCESS": ["motorpoweron", "ui_pm_acc:1:acc1:on0"],
    "ERROR": ["assertion failed", "reg_addr(00)isunviald"],
    "ON":  "ui_pm_acc: 0:nfc 1:on 0",
    "OFF": "ui_pm_acc: 0:nfc 0:off 1"
}