# coding=utf-8
# NFC开机关机压力测试（修复日志串流 + 致命异常捕获 + 标准日志版）
import time
import serial
import datetime
import random
import re
import threading
import logging
import sys
import os
from typing import Dict, List, Tuple, Optional

# ========== 基础配置 ==========
SERVO_PORT: str = "COM9"   # 舵机串口
SERVO_BAUD: int = 115200   # 舵机波特率
LOG_PORT: str   = "COM12"  # 日志串口
LOG_BAUD: int   = 115200   # 日志波特率
TOTAL_TESTS: int = 10000   # 总测试次数
INIT_DEVICE_STATUS: str = "关机"  # 初始状态

# 编译正则表达式，用于匹配并过滤掉设备日志中附带的终端颜色等 ANSI 转义码
ANSI_ESCAPE: re.Pattern = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')

# 状态匹配关键字集合
STATUS_KEYS: Dict[str, List[str]] = {
    "开机": ["ui_pm_acc:nfc-on,screen-off", "ui_pm_acc:acc-on,screen-off", "Thread Operating System"],
    "关机": ["ui_pm_acc:nfc-off,screen-on", "ui_pm_acc:acc-off,screen-on", "power_off"]
}

# 需要触发停止测试并保留现场的致命异常关键字（原始字符串，仅用于日志展示）
# 注意：" ppx_packet region is invalid" 已迁移至下方「频率触发型」配置，此处不再列出
CRITICAL_BUG_KEYWORDS: List[str] = [
    "left motor,communication loss!!!",
    "right motor,communication loss!!!",
    "0 motor recv",
    "voice_msg num: 7",
    "reg_addr (00) isunviald",
    "assertion failed",
    "hard fault",
    "param is invalid"
]

# ========== 模糊匹配正则预编译 ==========
def _build_fuzzy_pattern(kw: str) -> re.Pattern:
    """
    将关键字转为宽松正则表达式，用于替代原来的精确 in 匹配：

    规则：
      1. 将关键字按空白拆分为词元（tokens）
      2. 每个词元用 re.escape() 做正则转义，保证特殊字符被字面匹配
      3. 词元之间用 \\s* 连接 —— 允许出现 0 个或任意多个空白字符
         （兼容：多余空格、缺失空格、Tab、换行等边界情况）
      4. re.IGNORECASE —— 大小写不敏感

    示例：
      "[D/motor] left motor,communication loss!!!"
      → \\[D/motor\\]\\s*left\\s*motor\\,communication\\s*loss\\!\\!\\!
      可匹配 "[d/motor]left  motor,communication loss!!!" 等变体
    """
    tokens = kw.split()                           # 按空白切分词元（自动吸收前导/尾随空格）
    escaped = [re.escape(t) for t in tokens]      # 每个词元做正则转义
    return re.compile(r'\s*'.join(escaped), re.IGNORECASE)

# 与 CRITICAL_BUG_KEYWORDS 下标一一对应，脚本启动时一次性预编译
CRITICAL_BUG_PATTERNS: List[re.Pattern] = [
    _build_fuzzy_pattern(kw) for kw in CRITICAL_BUG_KEYWORDS
]

# ========== 频率触发型致命关键字配置 ==========
# 此类关键字偶发出现属正常噪声，需在短时间窗口内连续高频出现才判定为真实故障
#
# 格式：{ 关键字原文: (时间窗口秒数, 最小触发次数) }
# 含义：在 time_window 秒内累计命中次数 >= min_count，才触发停测保现场
#
# 注意：关键字前导空格不影响模糊匹配（_build_fuzzy_pattern 会自动 .split() 吸收空格）
FREQUENCY_BUG_CONFIG: Dict[str, Tuple[float, int]] = {
    " ppx_packet region is invalid": (3.0, 3),  # 3 秒内出现 ≥ 3 次才触发
}

# 预编译频率触发型关键字的模糊匹配正则（与 CRITICAL_BUG_PATTERNS 规则相同）
FREQUENCY_BUG_PATTERNS: Dict[str, re.Pattern] = {
    kw: _build_fuzzy_pattern(kw) for kw in FREQUENCY_BUG_CONFIG
}

# 各频率触发型关键字的历史命中时间戳列表
# —— 仅在 log_listener 线程内读写，无需加锁
# —— 采用滑动窗口：每次检测时剔除超出 time_window 的旧时间戳，再统计剩余数量
frequency_bug_timestamps: Dict[str, List[float]] = {
    kw: [] for kw in FREQUENCY_BUG_CONFIG
}
# =============================================

# 动作时间配置
NFC_LOW_STAY: float  = 2.1   # NFC最低点停留时间（秒）
NFC_HIGH_STAY: float = 3.5   # NFC最高点停留时间（秒）

# ========== 路径与全局状态配置 ==========
DESKTOP_PATH: str    = os.path.join(os.path.expanduser("~"), "Desktop")
CURRENT_RUN_DIR: str = os.path.join(DESKTOP_PATH,
                                    f"NFC_Test_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}")
RAW_LOG_FILE: str    = os.path.join(CURRENT_RUN_DIR, "raw_stream.log")

# 内存缓冲区与并发锁（彻底摒弃时间戳，改用单轮清理机制）
log_text_buffer: List[str] = []
log_listener_running: bool = False
log_serial: Optional[serial.Serial] = None
log_lock: threading.Lock = threading.Lock()

# 致命异常事件与记录（使用滚动字符串防止被 Chunk 截断）
rolling_bug_buffer: str = ""
critical_bug_event: threading.Event = threading.Event()
critical_bug_msg: str = ""

# 统计变量
current_status: str = INIT_DEVICE_STATUS
success_cnt: int = 0
total_cnt: int = 0

# ========== 全局日志记录器 ==========
logger: logging.Logger = logging.getLogger("NFC_Stress_Test")


def setup_logging() -> None:
    """初始化 logging 模块，配置控制台与文件双输出"""
    if not os.path.exists(CURRENT_RUN_DIR):
        os.makedirs(CURRENT_RUN_DIR)

    logger.setLevel(logging.DEBUG)
    formatter = logging.Formatter('[%(asctime)s] [%(levelname)s] %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

    # 控制台处理器
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # 结果文件处理器
    result_file = os.path.join(CURRENT_RUN_DIR, "test_result.log")
    file_handler = logging.FileHandler(result_file, encoding="utf-8")
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)


# ========== 核心工具函数 ==========
def strip_ansi(text: str) -> str:
    """剥离 ANSI 转义码，避免干扰关键字匹配"""
    return ANSI_ESCAPE.sub('', text)


# ========== 日志监听与异常捕获 ==========
def log_listener() -> None:
    """
    日志监听线程
    功能：
    1. 实时抓取串口数据并写入全量日志
    2. 无视换行符，直接存入本轮测试专属 Buffer
    3. 维护滚动窗口，实时监控「立即触发型」致命关键字（模糊正则匹配）
    4. 【新增】频率触发型关键字：在滑动时间窗口内命中次数达到阈值才触发
    """
    global log_serial, log_listener_running, critical_bug_msg, rolling_bug_buffer
    log_listener_running = True

    try:
        raw_file = open(RAW_LOG_FILE, "a", encoding="utf-8", buffering=1)
    except Exception as e:
        logger.error(f"无法创建原始日志文件: {e}")
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
                    raw_data = log_serial.read(log_serial.in_waiting)
                    if not raw_data:
                        continue

                    decoded = raw_data.decode('utf-8', errors='replace')
                    clean_chunk = strip_ansi(decoded)

                    # 实时落盘原始数据
                    raw_file.write(clean_chunk)
                    raw_file.flush()

                    # 1. 存入供主线程状态判断的缓冲区（无视换行，保留原貌）
                    with log_lock:
                        log_text_buffer.append(clean_chunk)

                    # 2. 维护致命异常滚动窗口（防止关键字在两次 read 之间被截断）
                    rolling_bug_buffer += clean_chunk
                    if len(rolling_bug_buffer) > 2048:
                        rolling_bug_buffer = rolling_bug_buffer[-1024:]

                    # 3. 立即触发型致命关键字检测（模糊正则，命中即停）
                    for kw, pat in zip(CRITICAL_BUG_KEYWORDS, CRITICAL_BUG_PATTERNS):
                        if pat.search(rolling_bug_buffer) and not critical_bug_event.is_set():
                            critical_bug_msg = kw
                            critical_bug_event.set()
                            logger.critical(f"实时日志捕获到致命异常: {kw}")

                    # 4. 频率触发型致命关键字检测（滑动时间窗口计数）
                    #
                    # 逻辑：
                    #   a. 用 findall 统计本次 chunk 中关键字的出现次数（不依赖累积的 rolling_bug_buffer，
                    #      避免对同一段文本重复计数）
                    #   b. 将每次命中的时刻（time.time()）追加到对应关键字的时间戳列表
                    #   c. 剔除列表中早于「当前时间 - time_window」的旧时间戳
                    #   d. 若列表长度 >= min_count，则判定为真实故障，触发停测
                    now = time.time()
                    for kw, pat in FREQUENCY_BUG_PATTERNS.items():
                        hit_count = len(pat.findall(clean_chunk))
                        if hit_count == 0:
                            continue

                        time_window, min_count = FREQUENCY_BUG_CONFIG[kw]
                        ts_list = frequency_bug_timestamps[kw]

                        # 追加本次命中的时间戳（每次 findall 命中记录一个）
                        ts_list.extend([now] * hit_count)

                        # 剔除滑动窗口外的旧时间戳
                        ts_list[:] = [t for t in ts_list if now - t <= time_window]

                        # 达到阈值则触发停测
                        if len(ts_list) >= min_count and not critical_bug_event.is_set():
                            critical_bug_msg = (
                                f"{kw.strip()}  "
                                f"（{time_window:.0f} 秒内出现 {len(ts_list)} 次，"
                                f"阈值 {min_count} 次）"
                            )
                            critical_bug_event.set()
                            logger.critical(
                                f"实时日志捕获到致命异常（频率触发）: {critical_bug_msg}"
                            )

                except Exception as read_e:
                    logger.error(f"串口读取或解码异常: {read_e}")

            time.sleep(0.001)

    except serial.SerialException as e:
        logger.error(f"日志监听串口打开失败：{e}")
    finally:
        if log_serial and log_serial.is_open:
            log_serial.close()
        raw_file.close()
        log_listener_running = False


# ========== 舵机控制 ==========
def servo_send(cmd: str) -> bool:
    """发送舵机指令"""
    for _ in range(2):
        try:
            with serial.Serial(SERVO_PORT, SERVO_BAUD, timeout=0.2, write_timeout=1.0) as ser:
                ser.write(cmd.encode('ascii'))
            return True
        except serial.SerialException:
            time.sleep(0.1)
    logger.error("舵机指令发送失败")
    return False


def nfc_move(side: str) -> None:
    """NFC移动逻辑，使用 Event.wait 代替 sleep 使得异常发生时能立即中断"""
    if side == "low":
        servo_send("#000P2500T1000!")
        critical_bug_event.wait(1.0 + NFC_LOW_STAY)
    else:
        servo_send("#000P0500T1000!")
        critical_bug_event.wait(1.0 + NFC_HIGH_STAY)


# ========== 核心分析逻辑 ==========
def get_final_status(log_text: str) -> Tuple[bool, str]:
    """
    状态判定逻辑（使用绝对最后一次出现的关键字，无视日志串行顺序问题）
    """
    target_status = "开机" if current_status == "关机" else "关机"

    pos_on = -1
    pos_off = -1

    # 查找所有开机关键字在全量字符串中最后一次出现的位置
    for kw in STATUS_KEYS["开机"]:
        p = log_text.rfind(kw)
        if p > pos_on:
            pos_on = p

    # 查找所有关机关键字在全量字符串中最后一次出现的位置
    for kw in STATUS_KEYS["关机"]:
        p = log_text.rfind(kw)
        if p > pos_off:
            pos_off = p

    # 若都没有找到，认为维持原状态
    if pos_on == -1 and pos_off == -1:
        return False, current_status

    # 比较开机和关机关键字，谁在字符串更靠后的位置，设备就处于什么最终状态
    if pos_on > pos_off:
        final_status = "开机"
    elif pos_off > pos_on:
        final_status = "关机"
    else:
        final_status = current_status

    is_success = (final_status == target_status)
    return is_success, final_status


# ========== 单次测试 ==========
def run_test(test_num: int) -> None:
    """单次开/关机测试流程"""
    global current_status, success_cnt, total_cnt
    total_cnt += 1

    logger.info("-" * 40)
    logger.info(f"开始第 {test_num} 次测试")

    target_status = "开机" if current_status == "关机" else "关机"
    logger.info(f"当前状态: {current_status} -> 期望状态: {target_status}")

    # 【关键修复】：在每次动作执行前，强制清空本轮测试的日志缓冲区，保证无历史污染
    with log_lock:
        log_text_buffer.clear()

    # 下压扫NFC → 抬起
    nfc_move("low")

    if critical_bug_event.is_set():
        return

    nfc_move("high")

    # 额外等待，让日志吐完。支持被严重错误打断
    critical_bug_event.wait(1.0)
    if critical_bug_event.is_set():
        return

    # 提取本轮抓取到的所有日志拼接为单一字符串
    with log_lock:
        round_log_text = "".join(log_text_buffer)

    is_success, final_status = get_final_status(round_log_text)

    if is_success:
        success_cnt += 1
        logger.info(f"结果: 成功 (设备已{final_status})")
    else:
        logger.warning(f"结果: 失败 (期望{target_status}，实际{final_status})")
        logger.debug("--- 失败时段日志片段 Start ---")
        # 仅取最后 10 行打印，防止刷屏
        for line in round_log_text.splitlines()[-10:]:
            logger.debug(line)
        logger.debug("--- 失败时段日志片段 End ---")

    # 【关键修复】：不论测试成功与否，强行将设备状态同步为日志反馈的最终真实状态
    current_status = final_status

    fail_cnt = total_cnt - success_cnt
    rate = (success_cnt / total_cnt) * 100 if total_cnt > 0 else 0.0
    logger.info(f"统计: 总计 {total_cnt} | 成功 {success_cnt} | 失败 {fail_cnt} | 成功率 {rate:.2f}%")


# ========== 主函数 ==========
def main() -> None:
    setup_logging()

    logger.info(f"本次测试日志存放于: {CURRENT_RUN_DIR}")
    logger.info(f"原始全量日志路径: {RAW_LOG_FILE}")

    # 启动后台日志监听（守护线程）
    listener_thread = threading.Thread(target=log_listener, daemon=True)
    listener_thread.start()

    # 等待串口初始化
    time.sleep(2.0)

    logger.info("初始化舵机位置 (抬起)...")
    nfc_move("high")

    try:
        for i in range(1, TOTAL_TESTS + 1):
            if critical_bug_event.is_set():
                break

            run_test(i)

            # 随机间隔，支持被事件中断
            if critical_bug_event.wait(random.uniform(0.5, 1.5)):
                break

        # 检查循环结束的原因是否为严重BUG触发
        if critical_bug_event.is_set():
            logger.critical("============== 发现致命异常，终止测试 ==============")
            logger.critical(f"触发原因: {critical_bug_msg}")
            logger.info("正在执行电机复位（舵机抬起）以保留BUG现场...")
            servo_send("#000P0500T1000!")  # 强制抬起指令，不依赖 wait
            logger.info("电机已复位，脚本停止运行。")

    except KeyboardInterrupt:
        logger.warning("用户手动停止测试")
    except Exception as e:
        logger.error(f"主程序运行中发生异常: {e}", exc_info=True)
    finally:
        logger.info("测试任务结束。")


if __name__ == "__main__":
    main()
