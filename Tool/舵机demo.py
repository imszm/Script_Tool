# coding=utf-8
import time
import serial
import datetime
import random
import threading
from collections import deque
import sys
import os
import configparser
import re

# ========== 基础配置 ==========
# 参考你的 Demo，舵机串口改为 COM7，如果实际是 COM6 请修改此处
SERVO_PORT = "COM7"
SERVO_BAUD = 115200

# 日志抓取板的串口
LOG_PORT = "COM25"
LOG_BAUD = 115200

# 测试参数
TOTAL_TESTS = 10000
INIT_DEVICE_STATUS = "关机"

# INI 配置文件路径 (直接使用 Demo 中的路径)
INI_FILE_PATH = 'C:/Users/szm21/Downloads/Test_Tool/007 Zide/test.ini'

# 动作组映射 (对应 INI 文件中的 [group] Gxxxx)
# 假设 G0001 是下压动作，G0002 是抬起/归位动作
KEY_LOW = "G0001"
KEY_HIGH = "G0002"

# 状态关键字
STATUS_KEYS = {
    "开机": "ui_pm_acc: 0:nfc 1:on 0",
    "关机": "ui_pm_acc: 0:nfc 0:off 1"
}

# 额外停留时间配置 (除了舵机运动本身的时间外，额外停留多久)
NFC_LOW_STAY = 1.0  # 下压到位后，额外停留秒数
NFC_HIGH_STAY = 2.0  # 抬起位停留秒数

# ========== 日志路径配置 ==========
DESKTOP_PATH = os.path.join(os.path.expanduser("~"), "Desktop")
CURRENT_RUN_DIR = os.path.join(DESKTOP_PATH, f"NFC_Test_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}")
RESULT_FILE = os.path.join(CURRENT_RUN_DIR, "test_result.txt")
RAW_LOG_FILE = os.path.join(CURRENT_RUN_DIR, "raw_stream.log")

# ========== 全局变量 ==========
log_buffer = deque(maxlen=20000)
log_listener_running = False
log_serial = None
log_lock = threading.Lock()

current_status = INIT_DEVICE_STATUS
success_cnt = 0
total_cnt = 0

# 全局配置对象，用于存储读取到的 INI 内容
global_config = None


# ========== INI 解析工具函数 (移植自 Demo) ==========
def parse_ini_command(raw_string):
    """
    解析 INI 中的原始字符串，提取指令和运动时间
    """
    # 1. 去掉花括号
    clean_str = raw_string.strip().replace('{', '').replace('}', '')

    # 2. 提取所有的 T数值 (时间)，用于计算延时
    time_values = re.findall(r'T(\d+)', clean_str)

    # 转换成整数并找到最大值，如果没有找到则默认 1000ms
    if time_values:
        max_time = max(map(int, time_values))
    else:
        max_time = 1000

    # 3. 去掉开头的 Gxxxx 标记
    hash_index = clean_str.find('#')
    if hash_index != -1:
        cmd_str = clean_str[hash_index:]
    else:
        cmd_str = clean_str

    return cmd_str, max_time


def load_ini_config():
    """读取并解析配置文件"""
    global global_config
    if not os.path.exists(INI_FILE_PATH):
        print(f"错误: 找不到配置文件 {INI_FILE_PATH}")
        return False

    config = configparser.ConfigParser()
    try:
        config.read(INI_FILE_PATH, encoding='utf-8')
    except:
        print("UTF-8 读取失败，尝试 GBK...")
        try:
            config.read(INI_FILE_PATH, encoding='gbk')
        except Exception as e:
            print(f"配置文件读取严重错误: {e}")
            return False

    # 检查必要的组是否存在
    if 'group' not in config:
        print("错误: INI 文件中缺少 [group] 字段")
        return False

    global_config = config
    print(f"配置文件已加载: {INI_FILE_PATH}")
    return True


# ========== 核心工具函数 ==========
def ensure_dir_exists():
    if not os.path.exists(CURRENT_RUN_DIR):
        os.makedirs(CURRENT_RUN_DIR)


def safe_str(s):
    return "" if s is None else str(s)


def get_display_time():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def get_ms_ts():
    return datetime.datetime.now().timestamp() * 1000


def write_result(content):
    try:
        with open(RESULT_FILE, "a", encoding="utf-8") as f:
            f.write(content + "\n")
    except Exception as e:
        print(f"结果写入异常：{e}", file=sys.stderr)


# ========== 日志监听 ==========
def log_listener():
    global log_serial, log_listener_running
    log_listener_running = True

    try:
        raw_file = open(RAW_LOG_FILE, "a", encoding="utf-8", buffering=1)
    except Exception as e:
        print(f"无法创建原始日志文件: {e}")
        return

    try:
        log_serial = serial.Serial(
            port=LOG_PORT,
            baudrate=LOG_BAUD,
            timeout=0.001
        )
        log_serial.flushInput()

        while log_listener_running:
            if log_serial.in_waiting > 0:
                try:
                    raw_data = log_serial.read(log_serial.in_waiting)
                    if raw_data:
                        raw_log = safe_str(raw_data.decode('utf-8', errors='replace'))
                        timestamp = get_ms_ts()

                        raw_file.write(f"[{get_display_time()}] {raw_log}")
                        raw_file.flush()

                        with log_lock:
                            log_buffer.append((timestamp, raw_log))
                except Exception:
                    pass
            time.sleep(0.0001)

    except Exception as e:
        err_msg = f"日志监听串口异常：{e}"
        print(err_msg)
        write_result(err_msg)
    finally:
        if log_serial and log_serial.is_open:
            log_serial.close()
        if raw_file:
            raw_file.close()
        log_listener_running = False


# ========== 舵机控制 (已修改为读取 INI) ==========
def servo_send_raw(cmd_str):
    """发送原始指令到底层串口"""
    for _ in range(2):  # 简单的重试机制
        try:
            with serial.Serial(SERVO_PORT, SERVO_BAUD, timeout=0.2, write_timeout=1.0) as ser:
                ser.write(cmd_str.encode('ascii'))
            return True
        except Exception:
            time.sleep(0.1)
    return False


def nfc_move(side):
    """
    NFC移动逻辑 - 修改版
    side: "low" 或 "high"
    """
    if global_config is None:
        print("错误: 配置未加载，无法移动")
        return

    # 1. 确定要读取的 Key
    if side == "low":
        target_key = KEY_LOW
        extra_stay = NFC_LOW_STAY
    else:
        target_key = KEY_HIGH
        extra_stay = NFC_HIGH_STAY

    # 2. 从配置中读取指令
    if target_key in global_config['group']:
        raw_data = global_config['group'][target_key]

        # 3. 解析指令和时间
        cmd_str, motion_time_ms = parse_ini_command(raw_data)

        # 4. 发送指令
        success = servo_send_raw(cmd_str)
        if not success:
            print(f"警告: 舵机指令发送失败 ({target_key})")
            return

        # 5. 计算休眠时间
        # 总等待 = 舵机运动耗时 + 额外的稳定停留时间 + 缓冲(50ms)
        sleep_time = (motion_time_ms / 1000.0) + extra_stay + 0.05
        time.sleep(sleep_time)

    else:
        print(f"警告: INI 文件中找不到键值 {target_key}")
        # 如果找不到配置，不做动作或抛出异常，这里选择只打印警告


# ========== 核心分析逻辑 ==========
def get_round_logs(start_ts, end_ts):
    round_logs = []
    with log_lock:
        current_buffer = list(log_buffer)

    for ts, log in current_buffer:
        if start_ts - 500 <= ts <= end_ts + 2000:
            round_logs.append(log)

    raw_log_str = "".join(round_logs)
    return [line for line in raw_log_str.splitlines() if line.strip()]


def get_final_status(log_lines):
    target_status = "开机" if current_status == "关机" else "关机"
    pos = {"开机": -1, "关机": -1}

    for idx, line in enumerate(log_lines):
        line_clean = line.replace(" ", "").replace(":", "")
        key_on = STATUS_KEYS["开机"].replace(" ", "").replace(":", "")
        key_off = STATUS_KEYS["关机"].replace(" ", "").replace(":", "")

        if key_on in line_clean:
            pos["开机"] = idx
        if key_off in line_clean:
            pos["关机"] = idx

    final_status = target_status
    if pos["开机"] == -1 and pos["关机"] == -1:
        return False, current_status

    if pos["开机"] > pos["关机"]:
        final_status = "开机"
    elif pos["关机"] > pos["开机"]:
        final_status = "关机"

    is_success = (final_status == target_status)
    return is_success, final_status


# ========== 单次测试 ==========
def run_test(test_num):
    global current_status, success_cnt, total_cnt
    total_cnt += 1

    header = f"Test No.{test_num} | {get_display_time()}"
    print("-" * 30)
    print(header)
    write_result(header)

    target_status = "开机" if current_status == "关机" else "关机"
    print(f"当前: {current_status} -> 目标: {target_status}")

    start_ts = get_ms_ts()

    # 执行动作 (现在会去读取 INI 中的 G0001 和 G0002)
    nfc_move("low")  # 下压
    nfc_move("high")  # 抬起

    time.sleep(1.0)
    end_ts = get_ms_ts()

    round_logs = get_round_logs(start_ts, end_ts)
    is_success, final_status = get_final_status(round_logs)

    if is_success:
        success_cnt += 1
        current_status = final_status
        res_msg = f"结果: 成功 (设备已{current_status})"
    else:
        res_msg = f"结果: 失败 (期望{target_status}，实际{final_status})"
        write_result("--- 失败时段日志片段 Start ---")
        for l in round_logs[-10:]:
            write_result(l)
        write_result("--- 失败时段日志片段 End ---")

    print(res_msg)
    write_result(res_msg)

    fail_cnt = total_cnt - success_cnt
    if total_cnt > 0:
        rate = (success_cnt / total_cnt) * 100
    else:
        rate = 0.0
    stat_msg = f"统计: 总{total_cnt} | 成功{success_cnt} | 失败{fail_cnt} | 率{rate:.2f}%"
    print(stat_msg)
    write_result(stat_msg + "\n")


# ========== 主函数 ==========
def main():
    ensure_dir_exists()

    print(f"日志存放: {CURRENT_RUN_DIR}")

    # 1. 先加载 INI 配置
    if not load_ini_config():
        print("程序终止：无法读取配置文件。")
        return

    # 2. 启动日志线程
    t = threading.Thread(target=log_listener, daemon=True)
    t.start()

    time.sleep(2.0)

    # 3. 初始归位
    print("初始化舵机位置 (执行 G0002)...")
    nfc_move("high")

    try:
        for i in range(1, TOTAL_TESTS + 1):
            run_test(i)
            time.sleep(random.uniform(0.5, 1.5))

    except KeyboardInterrupt:
        print("\n用户手动停止")
        write_result("用户手动停止")
    except Exception as e:
        print(f"\n主程序异常: {e}")
        write_result(f"主程序异常: {e}")
    finally:
        end_msg = f"\n测试结束: {get_display_time()}"
        print(end_msg)
        write_result(end_msg)


if __name__ == "__main__":
    main()