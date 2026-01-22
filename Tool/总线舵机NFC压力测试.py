# coding=utf-8
# NFC开机关机压力测试（内存优化+双日志版）
import time
import serial
import datetime
import random
import threading
from collections import deque
import sys
import os

# ========== 基础配置 ==========
SERVO_PORT = "COM6"  # 舵机串口
SERVO_BAUD = 115200  # 舵机波特率
LOG_PORT = "COM25"  # 日志串口
LOG_BAUD = 115200  # 日志波特率
TOTAL_TESTS = 10000  # 总测试次数
INIT_DEVICE_STATUS = "关机"  # 初始状态

# 状态关键字
STATUS_KEYS = {
    "开机": "ui_pm_acc: 0:nfc 1:on 0",
    "关机": "ui_pm_acc: 0:nfc 0:off 1"
}

# 动作时间配置
NFC_LOW_STAY = 1.0  # NFC最低点停留时间（秒）
NFC_HIGH_STAY = 2.0  # NFC最高点停留时间（秒）

# ========== 日志路径配置（动态生成） ==========
# 获取桌面路径
DESKTOP_PATH = os.path.join(os.path.expanduser("~"), "Desktop")
# 生成本次运行的专属文件夹名称，例如：NFC_Test_20260121_143000
CURRENT_RUN_DIR = os.path.join(DESKTOP_PATH, f"NFC_Test_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}")

# 定义两个日志文件的完整路径
# 1. 结果日志：只存测试结论
RESULT_FILE = os.path.join(CURRENT_RUN_DIR, "test_result.txt")
# 2. 原始日志：存串口吐出的所有内容（防崩溃/防内存溢出）
RAW_LOG_FILE = os.path.join(CURRENT_RUN_DIR, "raw_stream.log")

# ========== 全局变量 ==========
# 内存中的日志缓冲区，仅用于逻辑判断，不需要存太久，设置较小长度防止内存溢出
# 实际的全量日志已经实时写入磁盘了，这里只需要存最近几十秒的数据供分析即可
log_buffer = deque(maxlen=20000)
log_listener_running = False
log_serial = None
log_lock = threading.Lock()

# 统计变量
current_status = INIT_DEVICE_STATUS
success_cnt = 0
total_cnt = 0


# ========== 核心工具函数 ==========
def ensure_dir_exists():
    """确保日志文件夹存在"""
    if not os.path.exists(CURRENT_RUN_DIR):
        os.makedirs(CURRENT_RUN_DIR)


def safe_str(s):
    """仅做必要空值防护"""
    return "" if s is None else str(s)


def get_display_time():
    """获取显示时间：2025-12-11 18:31:24"""
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def get_ms_ts():
    """毫秒级时间戳"""
    return datetime.datetime.now().timestamp() * 1000


def write_result(content):
    """写入结果日志（追加模式）"""
    try:
        # print同时输出到控制台
        # print(content) # 如果觉得控制台太乱，可以注释掉这行，只看文件
        with open(RESULT_FILE, "a", encoding="utf-8") as f:
            f.write(content + "\n")
    except Exception as e:
        print(f"结果写入异常：{e}", file=sys.stderr)


# ========== 日志监听（核心优化：流式写入） ==========
def log_listener():
    """
    日志监听线程
    功能：
    1. 实时抓取串口数据
    2. 立即写入 raw_stream.log 文件（防崩溃丢失）
    3. 放入 log_buffer 供主线程分析
    """
    global log_serial, log_listener_running
    log_listener_running = True

    # 打开原始日志文件，使用 'a' 追加模式
    # buffering=1 表示行缓冲，确保数据尽快写入磁盘
    try:
        raw_file = open(RAW_LOG_FILE, "a", encoding="utf-8", buffering=1)
    except Exception as e:
        print(f"无法创建原始日志文件: {e}")
        return

    try:
        log_serial = serial.Serial(
            port=LOG_PORT,
            baudrate=LOG_BAUD,
            timeout=0.001,
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE
        )
        log_serial.flushInput()

        while log_listener_running:
            if log_serial.in_waiting > 0:
                try:
                    # 读取数据
                    raw_data = log_serial.read(log_serial.in_waiting)
                    if raw_data:
                        # 解码
                        raw_log = safe_str(raw_data.decode('utf-8', errors='replace'))
                        timestamp = get_ms_ts()

                        # 1. 【核心优化】立即写入磁盘并刷新
                        # 加上时间戳方便排查问题
                        raw_file.write(f"[{get_display_time()}] {raw_log}")
                        raw_file.flush()  # 强制落盘，防止脚本崩溃数据丢失

                        # 2. 放入内存缓冲区（供逻辑分析使用）
                        with log_lock:
                            log_buffer.append((timestamp, raw_log))
                except Exception as read_e:
                    print(f"读取或写入异常: {read_e}")

            time.sleep(0.0001)

    except Exception as e:
        err_msg = f"日志监听串口异常：{e}"
        print(err_msg, file=sys.stderr)
        write_result(err_msg)
    finally:
        if log_serial and log_serial.is_open:
            log_serial.close()
        if raw_file:
            raw_file.close()
        log_listener_running = False


# ========== 舵机控制 ==========
def servo_send(cmd):
    """极简舵机指令发送"""
    for _ in range(2):
        try:
            with serial.Serial(SERVO_PORT, SERVO_BAUD, timeout=0.2, write_timeout=1.0) as ser:
                ser.write(cmd.encode('ascii'))
            return True
        except:
            time.sleep(0.1)
    return False


def nfc_move(side):
    """NFC移动逻辑"""
    if side == "low":
        servo_send("#000P2500T1000!")
        time.sleep(1.0 + NFC_LOW_STAY)
    else:
        servo_send("#000P0500T1000!")
        time.sleep(1.0 + NFC_HIGH_STAY)


# ========== 核心分析逻辑 ==========
def get_round_logs(start_ts, end_ts):
    """
    提取指定时间区间的日志
    注意：这里只从内存 log_buffer 提取用于判断成功失败
    """
    round_logs = []
    # 使用副本进行迭代，避免锁定时间过长
    with log_lock:
        # 转换为列表进行切片分析
        current_buffer = list(log_buffer)

    # 筛选时间窗口内的日志
    for ts, log in current_buffer:
        if start_ts - 500 <= ts <= end_ts + 2000:  # 稍微放宽结束时间，防止日志延迟
            round_logs.append(log)

    # 仅按行拆分，不做复杂处理
    raw_log_str = "".join(round_logs)
    return [line for line in raw_log_str.splitlines() if line.strip()]


def get_final_status(log_lines):
    """状态判定逻辑"""
    target_status = "开机" if current_status == "关机" else "关机"
    pos = {"开机": -1, "关机": -1}

    # 倒序查找，提高效率，找到最后出现的关键字即可
    # 优化：优先匹配最后的状态
    for idx, line in enumerate(log_lines):
        line_clean = line.replace(" ", "").replace(":", "")
        # 去除标点和空格后匹配
        key_on = STATUS_KEYS["开机"].replace(" ", "").replace(":", "")
        key_off = STATUS_KEYS["关机"].replace(" ", "").replace(":", "")

        if key_on in line_clean:
            pos["开机"] = idx
        if key_off in line_clean:
            pos["关机"] = idx

    # 判定逻辑
    final_status = target_status  # 默认为没变

    # 如果两个状态都搜不到，保持原状
    if pos["开机"] == -1 and pos["关机"] == -1:
        return False, current_status

    if pos["开机"] > pos["关机"]:
        final_status = "开机"
    elif pos["关机"] > pos["开机"]:
        final_status = "关机"

    # 只要变成了目标状态，就算成功
    is_success = (final_status == target_status)
    return is_success, final_status


# ========== 单次测试 ==========
def run_test(test_num):
    """单次测试流程"""
    global current_status, success_cnt, total_cnt
    total_cnt += 1

    # 1. 打印表头
    print("-" * 30)
    header = f"Test No.{test_num} | {get_display_time()}"
    print(header)
    write_result(header)

    # 2. 执行动作
    target_status = "开机" if current_status == "关机" else "关机"
    print(f"当前: {current_status} -> 目标: {target_status}")

    start_ts = get_ms_ts()

    # 下压扫NFC
    nfc_move("low")
    # 抬起
    nfc_move("high")

    # 额外等待一点时间让日志吐完
    time.sleep(1.0)
    end_ts = get_ms_ts()

    # 3. 获取分析日志（不再打印所有原始日志到控制台，太乱了，只存文件）
    round_logs = get_round_logs(start_ts, end_ts)

    # 4. 判定结果
    is_success, final_status = get_final_status(round_logs)

    if is_success:
        success_cnt += 1
        current_status = final_status
        res_msg = f"结果: 成功 (设备已{current_status})"
    else:
        # 失败时，打印一下相关日志方便调试
        res_msg = f"结果: 失败 (期望{target_status}，实际{final_status})"
        write_result("--- 失败时段日志片段 Start ---")
        for l in round_logs[-10:]:  # 仅记录最后10行
            write_result(l)
        write_result("--- 失败时段日志片段 End ---")

    print(res_msg)
    write_result(res_msg)

    # 5. 统计信息
    fail_cnt = total_cnt - success_cnt
    rate = (success_cnt / total_cnt) * 100
    stat_msg = f"统计: 总{total_cnt} | 成功{success_cnt} | 失败{fail_cnt} | 率{rate:.2f}%"
    print(stat_msg)
    write_result(stat_msg + "\n")


# ========== 主函数 ==========
def main():
    # 1. 确保目录存在
    ensure_dir_exists()

    print(f"本次测试日志存放于: {CURRENT_RUN_DIR}")
    write_result(f"测试开始时间: {get_display_time()}")
    write_result(f"原始全量日志路径: {RAW_LOG_FILE}")

    # 2. 启动后台日志监听（守护线程）
    t = threading.Thread(target=log_listener, daemon=True)
    t.start()

    # 等待串口初始化
    time.sleep(2.0)

    # 归位
    print("初始化舵机位置...")
    nfc_move("high")

    try:
        for i in range(1, TOTAL_TESTS + 1):
            run_test(i)
            # 随机间隔，模拟真实操作
            time.sleep(random.uniform(0.5, 1.5))

    except KeyboardInterrupt:
        print("\n用户手动停止测试")
        write_result("用户手动停止测试")
    except Exception as e:
        print(f"\n主程序异常: {e}")
        write_result(f"主程序异常: {e}")
    finally:
        end_msg = f"\n测试结束: {get_display_time()}"
        print(end_msg)
        write_result(end_msg)


if __name__ == "__main__":
    main()