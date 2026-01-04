import serial
import time

# ---------- 配置参数 ----------
COM_PORT = "COM4"        # 串口号
BAUD_RATE = 9600         # 波特率
LOOP_COUNT = 10000       # 循环次数
LEFT_ON_TIME = 0.5       # 左灯亮时间
RIGHT_ON_TIME = 0.5      # 右灯亮时间
SWITCH_DELAY = 1         # 左右灯切换或下一轮延迟

# ---------- 打开串口 ----------
ser = serial.Serial(COM_PORT, BAUD_RATE, timeout=0.5)
time.sleep(0.5)  # 给继电器模块准备时间

# ---------- 初始化继电器：唤醒 + 全部 OFF ----------
ser.write(bytes([0x50]))  # 唤醒继电器模块
ser.flush()
time.sleep(0.5)

ser.write(b'P')            # 关闭第2路和第3路继电器
ser.flush()
time.sleep(0.5)

print("继电器初始化完成（已唤醒并全部关闭），开始左右灯循环...")

# ---------- 循环左右灯 ----------
success_count = 0

for i in range(1, LOOP_COUNT + 1):
    try:
        # 左灯（第2路）亮
        ser.write(b'R')
        ser.flush()
        time.sleep(LEFT_ON_TIME)

        # 左灯熄灭
        ser.write(b'P')
        ser.flush()
        time.sleep(SWITCH_DELAY)

        # 右灯（第3路）亮
        ser.write(b'T')
        ser.flush()
        time.sleep(RIGHT_ON_TIME)

        # 右灯熄灭
        ser.write(b'P')
        ser.flush()
        time.sleep(SWITCH_DELAY)

        # 成功计数
        success_count += 1

        # 打印当前执行次数和成功率
        print(f"已执行次数: {i}, 当前成功率: {success_count / i * 100:.2f}%")

    except Exception as e:
        print(f"第{i}次执行失败: {e}")

# ---------- 关闭串口 ----------
ser.close()
print(f"\n循环结束，总执行次数: {LOOP_COUNT}, 成功次数: {success_count}, 成功率: {success_count / LOOP_COUNT * 100:.2f}%")
