import serial
import time

ser = serial.Serial('COM4', 9600, timeout=0.5)

RELAY_CMDS = {
    "K2_ON":  bytes.fromhex("A0 02 01 A3"),
    "K2_OFF": bytes.fromhex("A0 02 00 A2"),
    "K3_ON":  bytes.fromhex("A0 03 01 A4"),
    "K3_OFF": bytes.fromhex("A0 03 00 A3"),
}

def send_cmd(cmd):
    """安全发送指令 + 延时"""
    ser.write(cmd)
    time.sleep(0.1)  # 给模块处理时间

def press_button(channel, press_time=0.3):
    """模拟按压按钮（按下+松开）"""
    send_cmd(RELAY_CMDS[f"K{channel}_ON"])
    time.sleep(press_time)
    send_cmd(RELAY_CMDS[f"K{channel}_OFF"])
    print(f"[OK] K{channel} 按压完成")

def left_right_cycle():
    """完整一次左->右"""
    print("\n=== 模拟左转向灯 ===")
    press_button(2, 0.4)
    time.sleep(0.6)   # 左右之间的缓冲时间

    print("=== 模拟右转向灯 ===")
    press_button(3, 0.4)
    time.sleep(1.0)   # 一轮结束等待

for i in range(10):
    print(f"\n第 {i+1} 轮开始")
    left_right_cycle()

ser.close()
