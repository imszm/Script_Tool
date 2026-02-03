# src/config.py
import logging

# ================= 硬件配置 =================
# 可以在这里写死，或者后续写自动扫描逻辑
RELAY_PORT = 'COM4'  # 继电器串口
DUT_PORT = 'COM25'  # 被测设备(车机)串口
BAUDRATE_RELAY = 9600
BAUDRATE_DUT = 115200

# ================= 继电器指令集 (16进制) =================
# 建议用字典分类，清晰明了
RELAY_CMDS = {
    # 原始指令示例
    "K2_ON": bytes.fromhex("A0 02 01 A3"),
    "K2_OFF": bytes.fromhex("A0 02 00 A2"),
    "K3_ON": bytes.fromhex("A0 03 01 A4"),
    "K3_OFF": bytes.fromhex("A0 03 00 A3"),

    # 喇叭指令 (参考你的喇叭脚本)
    "HORN_PRESS": bytes([0x4F]),
    "HORN_RELEASE": bytes([0x50]),

    # 转向灯指令 (参考你的转向灯脚本)
    "LEFT_SIGNAL": bytes([0x42]),
    "RESET_SIGNAL": bytes([0x50]),

    # 你的大灯脚本里有复杂的二进制指令，建议封装成常量
    "HEADLIGHT_RIGHT_ON": bytes([0b10100000, 0b00000011, 0b00000001, 0b10100100]),
    "HEADLIGHT_OFF": bytes([0b10100000, 0b00000011, 0b00000000, 0b10100011]),
}

# ================= 判定关键字 (用于开关机测试) =================
KEYWORDS = {
    "SUCCESS": ["motorpoweron", "ui_pm_acc:1:acc1:on0"],
    "CRITICAL_ERROR": ["assertion failed", "reg_addr(00)isunviald"],
}

# ================= 全局测试参数 =================
DEFAULT_TEST_COUNT = 50
LOG_LEVEL = logging.INFO