"""
单继电器模拟左/右转向灯（回弹按钮版）
--------------------------------------------------
特点：
1. 串口 COM4, 波特率 9600
2. 模拟人手按压回弹按钮
3. 按压切换状态，松开回弹
4. 左右灯互斥（K2 左，K3 右）
5. 启动前复位所有继电器
6. 输出总循环次数、成功次数和成功率
"""

import serial
import time

# ================= 配置 =================
SERIAL_PORT = 'COM4'
BAUDRATE = 9600
TEST_COUNT = 50                  # 总循环次数，可修改
PRESS_TIME = 0.6                 # 按压按钮时间
RELEASE_TIME = 0.2               # 松开按钮回弹时间
INTERVAL_BETWEEN_SWITCH = 0.5    # 左右灯切换间隔
LIGHT_ON_TIME = 0.8              # 灯亮持续时间

# ================= ICSE 指令定义 =================
CMD = {
    "K2_ON":  bytes.fromhex("A0 02 01 A3"),
    "K2_OFF": bytes.fromhex("A0 02 00 A2"),
    "K3_ON":  bytes.fromhex("A0 03 01 A4"),
    "K3_OFF": bytes.fromhex("A0 03 00 A3"),
}

# ================= 测试逻辑 =================
def run_test():
    success_count = 0
    failure_count = 0

    try:
        ser = serial.Serial(SERIAL_PORT, BAUDRATE, timeout=1)
        print(f"[初始化] 串口已打开 {SERIAL_PORT}")

        # ================= 初始化继电器 =================
        print("[系统初始化] 关闭 K2 / K3")
        ser.write(CMD["K2_OFF"])
        ser.write(CMD["K3_OFF"])
        time.sleep(0.5)

        state = 'LEFT'  # 当前控制灯：LEFT 或 RIGHT

        # ================= 循环测试 =================
        for i in range(1, TEST_COUNT + 1):
            try:
                print(f"[测试] 第 {i} 次循环 - 控制 {state} 灯")

                # 每次循环前复位
                ser.write(CMD["K2_OFF"])
                ser.write(CMD["K3_OFF"])
                time.sleep(0.3)

                # 模拟按压按钮
                print("[动作] 按压按钮")
                time.sleep(PRESS_TIME)

                # 打开对应方向的灯
                if state == 'LEFT':
                    ser.write(CMD["K2_ON"])
                    ser.write(CMD["K3_OFF"])
                    print("→ 左转灯（K2）打开")
                else:
                    ser.write(CMD["K3_ON"])
                    ser.write(CMD["K2_OFF"])
                    print("→ 右转灯（K3）打开")

                time.sleep(LIGHT_ON_TIME)

                # 松开按钮回弹
                print("[动作] 松开按钮回弹")
                time.sleep(RELEASE_TIME)

                # 熄灭所有灯（模拟再次按压关闭）
                ser.write(CMD["K2_OFF"])
                ser.write(CMD["K3_OFF"])
                print("→ 灯全部关闭")
                time.sleep(LIGHT_ON_TIME)

                time.sleep(RELEASE_TIME + INTERVAL_BETWEEN_SWITCH)

                # 切换左右灯
                state = 'RIGHT' if state == 'LEFT' else 'LEFT'
                success_count += 1

            except Exception as e:
                print(f"[错误] 第 {i} 次循环失败: {e}")
                failure_count += 1

    except serial.SerialException as e:
        print(f"[串口错误] {e}")

    finally:
        # 结束前关闭所有继电器
        if 'ser' in locals() and ser.is_open:
            ser.write(CMD["K2_OFF"])
            ser.write(CMD["K3_OFF"])
            time.sleep(0.3)
            ser.close()
            print(f"[结束] 串口已关闭 {SERIAL_PORT}")

        # ================= 统计 =================
        total = success_count + failure_count
        success_rate = (success_count / total) * 100 if total > 0 else 0
        print("======================================")
        print(f"总循环次数: {total}")
        print(f"成功次数: {success_count}")
        print(f"失败次数: {failure_count}")
        print(f"成功率: {success_rate:.2f}%")
        print("======================================")


# ================= 主入口 =================
if __name__ == "__main__":
    run_test()
