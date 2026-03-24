# src/config.py

# ================= 硬件连接配置 =================
# 根据实际情况修改默认端口
RELAY_PORT = 'COM4'  # 通用继电器 (W3, 充电)
DEVICE_PORT = 'COM25'  # 车机设备通信口
RELAY_CCB_PORT = 'COM12'  # CCB测试专用的继电器口

# 舵机串口 (NFC 自动化)
SERVO_PORT = 'COM6'
BAUDRATE_SERVO = 115200

# 波特率
BAUDRATE_RELAY = 9600
BAUDRATE_DEVICE = 115200

# ================= 指令集 (统一管理) =================
COMMANDS = {
    # [通用/W3/充电] 16进制指令
    # 逻辑：W3中 0x50是按下(导通), 0x4F是松开
    # 逻辑：充电中 0x4F是开, 0x50是关
    "HEX_PRESS_ON": bytes([0x50]),  # 按下/关/导通
    "HEX_RELEASE_OFF": bytes([0x4F]),  # 松开/开/断开
    "HEX_ENABLE": bytes([0x51]),  # 初始化使能

    # [CCB SMT项目] ASCII 指令
    "ASCII_OFF": b'P',  # 继电器断电
    "ASCII_ON": b'O',  # 继电器上电

    # [转向灯] ASCII 指令 (继电器多路)
    "TURN_OFF": b'P',
    "LEFT_TURN_ON": b'R',   # 左灯(第2路)
    "RIGHT_TURN_ON": b'T',  # 右灯(第3路)

    # [喇叭] 16进制按压/松开 (沿用历史脚本节奏)
    "HORN_PRESS": bytes([0x4F]),
    "HORN_RELEASE": bytes([0x50]),
}

# ================= 判定关键字 =================
KEYWORDS = {
    # W3 停止条件
    "W3_STOP": "voice_msg num: 6",
    "W3_ERROR": "communication loss",

    # 充电测试判定
    "CHARGE_SUCCESS": ["voice_msgnum:9", "voice_msgnum:10"],
    "CHARGE_ERROR": ["assertionfailedatfunction"],
}

# ================= 舵机/NFC 配置 =================
# 舵机动作（示例：#000P0500T1000! 抬起，#000P2500T1000! 下压）
SERVO_CMDS = {
    "NFC_LOW": b"#000P2500T1000!",
    "NFC_HIGH": b"#000P0500T1000!",
}

NFC_KEYWORDS = {
    # 仅用于简单日志判断（与 Tool/ 总线舵机NFC压力测试.py 保持一致）
    "ON": "ui_pm_acc: 0:nfc 1:on 0",
    "OFF": "ui_pm_acc: 0:nfc 0:off 1",
}

# ================= PC工具自动化配置 (Pywinauto) =================
PC_TOOL_CONFIG = {
    # 1. 升级工具配置
    "UPGRADE_APP_TITLE": "L5 PCTOOL V3.9.00",
    "UPGRADE_BTN_ID": "Widget.buttonUpgrade",
    "UPGRADE_LOG_ID": "Widget.textEditLog",
    "UPGRADE_WAIT_TIME": 170,  # 升级等待时间(秒)

    # 2. CCB SMT 测试配置
    "CCB_TITLE_REGEX": "CCB 测试 V3.2.00.*",
    "CCB_SERIAL_PREFIX": "2010007005R615GD00590",
    "CCB_COORDS": {
        "PASS_LIGHT": (1165, 224),
        "PASS_HORN": (1171, 274)
    },
    "CCB_CHECK_POINTS": [(1701, 820), (1846, 812)],  # 像素颜色检测点
}