# -*- coding: utf-8 -*-
"""
皮皮熊 R3_FCB 四通道自动测试程序 (优化版)
===========================================================
功能：
1. 直接键盘输入4个SN，等待4通道开始测试
2. OpenCV + HSV 检测各状态指示灯
3. 状态灯全绿后自动点击“通过”，并等待数据库与最终结果
4. 自动统计测试结果，计算成功率
"""

import os
import json
import time
import logging
from pathlib import Path
from datetime import datetime
from dataclasses import dataclass
from typing import Dict, Tuple, Optional, List, Any

import cv2
import numpy as np
import pyautogui

# ============================================================
# 一、CONFIG 基础配置
# ============================================================
CONFIG: Dict[str, Any] = {
    "LOOP_COUNT": 1,
    "BETWEEN_ROUNDS_DELAY": 3.0,
    "SN_BASE": "2023006001R000GD006800001",
    "SN_SUFFIX_WIDTH": 6,
    "CHANNEL_COUNT": 4,
    "TYPE_INTERVAL": 0.01,
    "SCAN_AFTER_ENTER_DELAY": 0.5,
    "SCAN_START_DELAY": 2.0,
    "PASS_BUTTON_COORDS": {
        1: (1092, 201), 2: (1092, 406), 3: (1092, 611), 4: (1092, 816),
    },
    "HSV_LOWER": (35, 80, 60),
    "HSV_UPPER": (90, 255, 255),
    "STATUS_POINTS": {
        1: {"power": (539, 185), "version": (641, 185), "calibrate": (777, 185), "opamp": (879, 185), "sn": (965, 185)},
        2: {"power": (539, 390), "version": (641, 390), "calibrate": (777, 390), "opamp": (879, 390), "sn": (965, 390)},
        3: {"power": (539, 595), "version": (641, 595), "calibrate": (777, 595), "opamp": (879, 595), "sn": (965, 595)},
        4: {"power": (539, 800), "version": (641, 800), "calibrate": (777, 800), "opamp": (879, 800), "sn": (965, 800)},
    },
    "DATABASE_POINTS": {
        1: (1328, 185), 2: (1328, 390), 3: (1328, 595), 4: (1328, 800),
    },
    "FINAL_PASS_SEARCH_ROIS": {
        1: (10, 100, 140, 320), 2: (10, 305, 140, 525), 3: (10, 510, 140, 730), 4: (10, 715, 140, 950),
    },
    "STATUS_X_RADIUS": 22,
    "STATUS_Y_RADIUS": 45,
    "DATABASE_X_RADIUS": 22,
    "DATABASE_Y_RADIUS": 45,
    "DOT_GREEN_RATIO": 0.20,
    "PASS_GREEN_RATIO": 0.20,
    "GREEN_CONFIRM_COUNT": 3,
    "POLL_INTERVAL": 0.20,
    "STATUS_TIMEOUT": 120,
    "DATABASE_TIMEOUT": 30,
    "FINAL_PASS_TIMEOUT": 30,
    "CHANNEL_TOTAL_TIMEOUT": 180,
    "SAVE_TIMEOUT_SCREENSHOT": True,
    "SAVE_LAYOUT_SCREENSHOT": True,
    "DEBUG_RECOGNITION": True,
    "DEBUG_DIR": "debug_screenshots",
    "LOG_DIR": "logs",
    "STATISTICS_FILE": "test_statistics.json",
    # 提示：如果不需要历史统计，可将其改为 False；否则请在每次新批次测试前删除 test_statistics.json 文件
    "PERSIST_STATISTICS": True,
}

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.05

# ============================================================
# 二、日志与数据结构
# ============================================================
def setup_logging() -> logging.Logger:
    """初始化日志配置，双向输出到终端和文件"""
    log_dir = Path(CONFIG["LOG_DIR"])
    log_dir.mkdir(exist_ok=True)
    filename = log_dir / f"FCB_AutoTest_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    if logger.handlers:
        logger.handlers.clear()
        
    formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s", datefmt="%H:%M:%S")
    
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    
    file_handler = logging.FileHandler(filename, encoding="utf-8")
    file_handler.setFormatter(formatter)
    
    logger.addHandler(console_handler)
    logger.addHandler(file_handler)
    return logger

LOGGER: logging.Logger = setup_logging()

@dataclass
class ChannelState:
    channel: int
    state: str = "WAIT_STATUS"
    start_time: float = 0.0
    phase_start_time: float = 0.0
    green_streak: int = 0
    fail_reason: str = ""
    pass_button_clicked: bool = False

# ============================================================
# 三、核心图像与坐标处理逻辑
# ============================================================
def take_screenshot() -> np.ndarray:
    """获取全屏截图并转换为 numpy 数组"""
    screenshot = pyautogui.screenshot()
    return np.array(screenshot)

def generate_sn_list() -> List[str]:
    """根据配置的后缀宽度，生成 4 个连续的 SN"""
    base_sn = CONFIG["SN_BASE"]
    width = CONFIG["SN_SUFFIX_WIDTH"]
    if len(base_sn) < width:
        raise ValueError("SN_BASE长度小于SN_SUFFIX_WIDTH")
    prefix = base_sn[:-width]
    start_number = int(base_sn[-width:])
    
    sn_list: List[str] = []
    for i in range(CONFIG["CHANNEL_COUNT"]):
        number = start_number + i
        sn_list.append(f"{prefix}{number:0{width}d}")
    return sn_list

def point_to_roi(point: Tuple[int, int], x_radius: int, y_radius: int) -> Tuple[int, int, int, int]:
    """将中心点转换为 ROI 矩形坐标 (x1, y1, x2, y2)"""
    center_x, center_y = point
    return (center_x - x_radius, center_y - y_radius, center_x + x_radius, center_y + y_radius)

def clip_roi(roi: Tuple[int, int, int, int], image_width: int, image_height: int) -> Tuple[int, int, int, int]:
    """限制 ROI 在图像尺寸范围内"""
    x1, y1, x2, y2 = roi
    x1 = max(0, min(x1, image_width - 1))
    y1 = max(0, min(y1, image_height - 1))
    x2 = max(0, min(x2, image_width))
    y2 = max(0, min(y2, image_height))
    return x1, y1, x2, y2

def is_green_dot(image: np.ndarray, region: Tuple[int, int, int, int], green_ratio_threshold: Optional[float] = None) -> Tuple[bool, float, int]:
    """检测指定区域是否存在符合要求的绿色连通域或绿色占比"""
    if green_ratio_threshold is None:
        green_ratio_threshold = CONFIG["DOT_GREEN_RATIO"]
    height, width = image.shape[:2]
    x1, y1, x2, y2 = clip_roi(region, width, height)
    
    if x2 <= x1 or y2 <= y1:
        return False, 0.0, 0
    roi = image[y1:y2, x1:x2]
    if roi.size == 0:
        return False, 0.0, 0
        
    hsv = cv2.cvtColor(roi, cv2.COLOR_RGB2HSV)
    lower = np.array(CONFIG["HSV_LOWER"], dtype=np.uint8)
    upper = np.array(CONFIG["HSV_UPPER"], dtype=np.uint8)
    mask = cv2.inRange(hsv, lower, upper)
    
    green_pixels = cv2.countNonZero(mask)
    total_pixels = roi.shape[0] * roi.shape[1]
    green_ratio = green_pixels / total_pixels if total_pixels > 0 else 0.0
    
    num_labels, _, stats, _ = cv2.connectedComponentsWithStats(mask, connectivity=8)
    max_green_area = int(np.max(stats[1:, cv2.CC_STAT_AREA])) if num_labels > 1 else 0
    
    is_green = bool((green_ratio >= green_ratio_threshold) or (max_green_area >= 180))
    return is_green, green_ratio, max_green_area

def check_channel_status(image: np.ndarray, channel: int) -> Tuple[bool, Dict[str, Any]]:
    """检查通道对应的 5 个指示灯是否全绿"""
    points = CONFIG["STATUS_POINTS"][channel]
    results: Dict[str, Any] = {}
    all_green = True
    
    for name, point in points.items():
        roi = point_to_roi(point, CONFIG["STATUS_X_RADIUS"], CONFIG["STATUS_Y_RADIUS"])
        green, ratio, area = is_green_dot(image, roi, CONFIG["DOT_GREEN_RATIO"])
        results[name] = {"green": green, "ratio": ratio, "area": area, "roi": roi}
        if not green:
            all_green = False
    return all_green, results

def check_database_status(image: np.ndarray, channel: int) -> Tuple[bool, float, int]:
    """检测数据库绿灯"""
    point = CONFIG["DATABASE_POINTS"][channel]
    roi = point_to_roi(point, CONFIG["DATABASE_X_RADIUS"], CONFIG["DATABASE_Y_RADIUS"])
    return is_green_dot(image, roi, CONFIG["DOT_GREEN_RATIO"])

def check_final_pass(image: np.ndarray, channel: int) -> Tuple[bool, float, int]:
    """检测最终的 PASS 矩形区域"""
    roi = CONFIG["FINAL_PASS_SEARCH_ROIS"][channel]
    return is_green_dot(image, roi, CONFIG["PASS_GREEN_RATIO"])

def log_status_results(channel: int, results: Dict[str, Any]) -> None:
    """按指定格式输出通道指示灯状态日志"""
    if not CONFIG["DEBUG_RECOGNITION"]:
        return
    parts: List[str] = []
    for name in ["power", "version", "calibrate", "opamp", "sn"]:
        item = results[name]
        status = "绿" if item["green"] else "灰"
        parts.append(f"{name}={status}(ratio={item['ratio']:.2f},area={item['area']})")
    LOGGER.info(f"通道{channel}状态：" + " | ".join(parts))

# ============================================================
# 四、界面控制与调试输出
# ============================================================
def click_pass_button(channel: int) -> None:
    """通过自动化点击通道对应的通过按钮"""
    x, y = CONFIG["PASS_BUTTON_COORDS"][channel]
    LOGGER.info(f"通道{channel}：准备点击“通过”按钮 坐标=({x},{y})")
    try:
        pyautogui.moveTo(x, y, duration=0.1)
        pyautogui.click()
        LOGGER.info(f"通道{channel}：已点击“通过”按钮")
        time.sleep(0.3)
    except Exception as e:
        LOGGER.error(f"通道{channel} 点击“通过”按钮失败: {e}")
        raise

def scan_barcodes(sn_list: List[str]) -> None:
    """自动完成输入与扫码提交流程"""
    LOGGER.info("=" * 70)
    LOGGER.info("开始4通道扫码")
    LOGGER.info("=" * 70)
    time.sleep(CONFIG["SCAN_START_DELAY"])
    
    for channel, sn in enumerate(sn_list, start=1):
        LOGGER.info(f"通道{channel}：输入SN -> {sn}")
        try:
            pyautogui.write(sn, interval=CONFIG["TYPE_INTERVAL"])
            pyautogui.press("enter")
            LOGGER.info(f"通道{channel}：SN输入完成并回车")
            time.sleep(CONFIG["SCAN_AFTER_ENTER_DELAY"])
        except Exception as e:
            LOGGER.exception(f"通道{channel}：扫码失败：{e}")
            raise
    LOGGER.info("4个SN全部输入完成")

def save_timeout_screenshot(channel: int, reason: str) -> None:
    """遇到超时失败时保存当前屏幕截图，便于追溯问题"""
    if not CONFIG["SAVE_TIMEOUT_SCREENSHOT"]:
        return
    try:
        debug_dir = Path(CONFIG["DEBUG_DIR"])
        debug_dir.mkdir(exist_ok=True)
        filename = debug_dir / f"CH{channel}_{reason}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        pyautogui.screenshot().save(filename)
        LOGGER.info(f"超时截图已保存：{filename}")
    except Exception as e:
        LOGGER.exception(f"保存超时截图失败：{e}")

def save_layout_screenshot() -> None:
    """绘制并保存当前识别区域的边界框图，辅助校准位置"""
    if not CONFIG["SAVE_LAYOUT_SCREENSHOT"]:
        return
    try:
        debug_dir = Path(CONFIG["DEBUG_DIR"])
        debug_dir.mkdir(exist_ok=True)
        image_rgb = take_screenshot()
        image = cv2.cvtColor(image_rgb, cv2.COLOR_RGB2BGR)
        
        for channel in range(1, CONFIG["CHANNEL_COUNT"] + 1):
            # 状态灯边界
            points = CONFIG["STATUS_POINTS"][channel]
            for name, point in points.items():
                x1, y1, x2, y2 = point_to_roi(point, CONFIG["STATUS_X_RADIUS"], CONFIG["STATUS_Y_RADIUS"])
                cv2.rectangle(image, (x1, y1), (x2, y2), (0, 0, 255), 2)
                cv2.putText(image, f"CH{channel}-{name}", (x1, max(10, y1 - 5)), cv2.FONT_HERSHEY_SIMPLEX, 0.45, (0, 0, 255), 1, cv2.LINE_AA)
            
            # 数据库边界
            db_point = CONFIG["DATABASE_POINTS"][channel]
            dx1, dy1, dx2, dy2 = point_to_roi(db_point, CONFIG["DATABASE_X_RADIUS"], CONFIG["DATABASE_Y_RADIUS"])
            cv2.rectangle(image, (dx1, dy1), (dx2, dy2), (255, 0, 0), 2)
            cv2.putText(image, f"CH{channel}-DB", (dx1, max(10, dy1 - 5)), cv2.FONT_HERSHEY_SIMPLEX, 0.45, (255, 0, 0), 1, cv2.LINE_AA)
            
            # PASS边界
            px1, py1, px2, py2 = CONFIG["FINAL_PASS_SEARCH_ROIS"][channel]
            cv2.rectangle(image, (px1, py1), (px2, py2), (0, 255, 255), 2)
            cv2.putText(image, f"CH{channel}-FINAL", (px1, max(10, py1 - 5)), cv2.FONT_HERSHEY_SIMPLEX, 0.45, (0, 255, 255), 1, cv2.LINE_AA)

        filename = debug_dir / f"layout_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        cv2.imwrite(str(filename), image)
        LOGGER.info(f"ROI调试截图已保存：{filename}")
    except Exception as e:
        LOGGER.exception(f"保存ROI调试图失败：{e}")

# ============================================================
# 五、测试主流程与状态机
# ============================================================
def run_parallel_test() -> Dict[int, ChannelState]:
    """并行监控4个测试通道的状态流转"""
    LOGGER.info("=" * 70)
    LOGGER.info("开始4通道并行状态监控")
    LOGGER.info("=" * 70)
    
    now = time.time()
    channel_states: Dict[int, ChannelState] = {
        channel: ChannelState(channel=channel, state="WAIT_STATUS", start_time=now, phase_start_time=now)
        for channel in range(1, CONFIG["CHANNEL_COUNT"] + 1)
    }
    
    finished_count = 0
    last_status_log_time = {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0}

    while finished_count < CONFIG["CHANNEL_COUNT"]:
        image = take_screenshot()
        current_time = time.time()
        
        for channel in range(1, CONFIG["CHANNEL_COUNT"] + 1):
            state = channel_states[channel]
            if state.state in ("PASS", "FAIL"):
                continue
                
            # 全局超时管控
            if (current_time - state.start_time) >= CONFIG["CHANNEL_TOTAL_TIMEOUT"]:
                state.state = "FAIL"
                state.fail_reason = f"通道总测试超时：{CONFIG['CHANNEL_TOTAL_TIMEOUT']}秒"
                LOGGER.error(f"通道{channel}：{state.fail_reason}")
                save_timeout_screenshot(channel, "CHANNEL_TOTAL_TIMEOUT")
                finished_count += 1
                continue

            # 阶段一：等待5个状态灯全部变绿
            if state.state == "WAIT_STATUS":
                if (current_time - state.phase_start_time) >= CONFIG["STATUS_TIMEOUT"]:
                    state.state = "FAIL"
                    state.fail_reason = "5个状态灯全部变绿等待超时"
                    LOGGER.error(f"通道{channel}：{state.fail_reason}")
                    save_timeout_screenshot(channel, "WAIT_STATUS_TIMEOUT")
                    finished_count += 1
                    continue

                all_green, results = check_channel_status(image, channel)
                if (current_time - last_status_log_time[channel]) >= 1.0:
                    log_status_results(channel, results)
                    last_status_log_time[channel] = current_time

                if all_green:
                    state.green_streak += 1
                    LOGGER.info(f"通道{channel}：5个状态灯全部检测为绿色 [{state.green_streak}/{CONFIG['GREEN_CONFIRM_COUNT']}]")
                    if state.green_streak >= CONFIG["GREEN_CONFIRM_COUNT"]:
                        LOGGER.info(f"通道{channel}：所有状态灯PASS，尝试点击通过")
                        try:
                            click_pass_button(channel)
                            state.pass_button_clicked = True
                            state.state = "WAIT_DATABASE"
                            state.phase_start_time = current_time
                            state.green_streak = 0
                        except Exception as e:
                            state.state = "FAIL"
                            state.fail_reason = f"点击通过按钮异常：{e}"
                            LOGGER.exception(f"通道{channel}：{state.fail_reason}")
                            finished_count += 1
                else:
                    state.green_streak = 0
                continue

            # 阶段二：等待数据库绿灯
            if state.state == "WAIT_DATABASE":
                if (current_time - state.phase_start_time) >= CONFIG["DATABASE_TIMEOUT"]:
                    state.state = "FAIL"
                    state.fail_reason = "数据库绿灯等待超时"
                    LOGGER.error(f"通道{channel}：{state.fail_reason}")
                    save_timeout_screenshot(channel, "WAIT_DATABASE_TIMEOUT")
                    finished_count += 1
                    continue

                db_green, ratio, area = check_database_status(image, channel)
                if db_green:
                    state.green_streak += 1
                    LOGGER.info(f"通道{channel}：数据库检测为绿色 [{state.green_streak}/{CONFIG['GREEN_CONFIRM_COUNT']}]")
                    if state.green_streak >= CONFIG["GREEN_CONFIRM_COUNT"]:
                        LOGGER.info(f"通道{channel}：数据库PASS")
                        state.state = "WAIT_FINAL_PASS"
                        state.phase_start_time = current_time
                        state.green_streak = 0
                else:
                    state.green_streak = 0
                continue

            # 阶段三：等待最终验证框变为PASS
            if state.state == "WAIT_FINAL_PASS":
                if (current_time - state.phase_start_time) >= CONFIG["FINAL_PASS_TIMEOUT"]:
                    state.state = "FAIL"
                    state.fail_reason = "左侧最终PASS等待超时"
                    LOGGER.error(f"通道{channel}：{state.fail_reason}")
                    save_timeout_screenshot(channel, "WAIT_FINAL_PASS_TIMEOUT")
                    finished_count += 1
                    continue

                final_pass, ratio, area = check_final_pass(image, channel)
                if final_pass:
                    state.green_streak += 1
                    LOGGER.info(f"通道{channel}：最终PASS检测为绿色 [{state.green_streak}/{CONFIG['GREEN_CONFIRM_COUNT']}]")
                    if state.green_streak >= CONFIG["GREEN_CONFIRM_COUNT"]:
                        state.state = "PASS"
                        LOGGER.info(f"通道{channel}：================ PASS ================")
                        finished_count += 1
                else:
                    state.green_streak = 0
                continue

    return channel_states

# ============================================================
# 六、统计与报表记录模块
# ============================================================
def create_empty_statistics() -> Dict[str, Any]:
    """生成空统计模板"""
    return {
        "total_rounds": 0,
        "machine": {"pass": 0, "fail": 0},
        "channels": {str(c): {"pass": 0, "fail": 0} for c in range(1, 5)}
    }

def load_statistics() -> Dict[str, Any]:
    """加载历史统计数据"""
    if not CONFIG["PERSIST_STATISTICS"]:
        return create_empty_statistics()
    path = Path(CONFIG["STATISTICS_FILE"])
    if not path.exists():
        return create_empty_statistics()
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        LOGGER.info("已读取历史统计数据")
        return data
    except Exception as e:
        LOGGER.exception(f"统计文件读取失败，重置为新表：{e}")
        return create_empty_statistics()

def save_statistics(statistics: Dict[str, Any]) -> None:
    """持久化统计数据到本地 JSON"""
    if not CONFIG["PERSIST_STATISTICS"]:
        return
    try:
        with open(CONFIG["STATISTICS_FILE"], "w", encoding="utf-8") as f:
            json.dump(statistics, f, ensure_ascii=False, indent=4)
    except Exception as e:
        LOGGER.exception(f"统计文件保存失败：{e}")

def update_statistics(statistics: Dict[str, Any], channel_states: Dict[int, ChannelState]) -> None:
    """根据本轮状态更新统计表"""
    statistics["total_rounds"] += 1
    machine_pass = True
    
    for channel in range(1, CONFIG["CHANNEL_COUNT"] + 1):
        state = channel_states[channel]
        key = str(channel)
        if state.state == "PASS":
            statistics["channels"][key]["pass"] += 1
        else:
            statistics["channels"][key]["fail"] += 1
            machine_pass = False

    if machine_pass:
        statistics["machine"]["pass"] += 1
    else:
        statistics["machine"]["fail"] += 1

def calculate_rate(success: int, failure: int) -> Tuple[float, float]:
    """安全计算成功率与失败率"""
    total = success + failure
    if total <= 0:
        return 0.0, 0.0
    return (success / total * 100), (failure / total * 100)

def print_round_result(round_index: int, channel_states: Dict[int, ChannelState]) -> None:
    """输出单轮通道测试结果日志"""
    LOGGER.info("\n" + "=" * 70)
    LOGGER.info(f"第{round_index}轮测试结果")
    LOGGER.info("=" * 70)
    machine_pass = True
    
    for channel in range(1, CONFIG["CHANNEL_COUNT"] + 1):
        state = channel_states[channel]
        if state.state == "PASS":
            LOGGER.info(f"通道{channel}：PASS")
        else:
            machine_pass = False
            LOGGER.error(f"通道{channel}：FAIL, 失败原因：{state.fail_reason}")

    LOGGER.info("-" * 70)
    if machine_pass:
        LOGGER.info("整机结果：PASS")
    else:
        LOGGER.error("整机结果：FAIL")
    LOGGER.info("=" * 70 + "\n")

def print_statistics(statistics: Dict[str, Any]) -> None:
    """打印最终汇总报告"""
    LOGGER.info("\n\n" + "#" * 80)
    LOGGER.info("                  自动测试统计报告")
    LOGGER.info("#" * 80)

    total_rounds = statistics["total_rounds"]
    m_pass = statistics["machine"]["pass"]
    m_fail = statistics["machine"]["fail"]
    m_pass_rate, m_fail_rate = calculate_rate(m_pass, m_fail)

    LOGGER.info(f"总测试轮数：{total_rounds}")
    LOGGER.info("\n-------------------- 整机统计 --------------------")
    LOGGER.info(f"整机成功次数：{m_pass} | 整机失败次数：{m_fail}")
    LOGGER.info(f"整机成功率：{m_pass_rate:.2f}% | 整机失败率：{m_fail_rate:.2f}%")

    LOGGER.info("\n-------------------- 通道统计 --------------------")
    total_channel_pass = 0
    total_channel_fail = 0
    for channel in range(1, CONFIG["CHANNEL_COUNT"] + 1):
        c_pass = statistics["channels"][str(channel)]["pass"]
        c_fail = statistics["channels"][str(channel)]["fail"]
        total_channel_pass += c_pass
        total_channel_fail += c_fail
        c_pass_rate, c_fail_rate = calculate_rate(c_pass, c_fail)
        LOGGER.info(f"通道{channel}：PASS={c_pass} | FAIL={c_fail} | 成功率={c_pass_rate:.2f}% | 失败率={c_fail_rate:.2f}%")

    total_channel_test = total_rounds * CONFIG["CHANNEL_COUNT"]
    tc_pass_rate, tc_fail_rate = calculate_rate(total_channel_pass, total_channel_fail)

    LOGGER.info("\n-------------------- 汇总统计 --------------------")
    LOGGER.info(f"总通道测试次数：{total_channel_test}")
    LOGGER.info(f"总通道PASS次数：{total_channel_pass} | 总通道FAIL次数：{total_channel_fail}")
    LOGGER.info(f"总通道成功率：{tc_pass_rate:.2f}% | 总通道失败率：{tc_fail_rate:.2f}%")
    LOGGER.info("#" * 80 + "\n")

# ============================================================
# 七、主入口执行
# ============================================================
def main() -> None:
    LOGGER.info("\n" + "=" * 80)
    LOGGER.info("          皮皮熊 FCB 四通道自动测试程序")
    LOGGER.info("=" * 80)
    LOGGER.info(f"计划测试轮数：{CONFIG['LOOP_COUNT']} | SN起始值：{CONFIG['SN_BASE']}")

    screen_w, screen_h = pyautogui.size()
    LOGGER.info(f"当前屏幕分辨率：{screen_w} x {screen_h} (注：脚本使用绝对坐标)")
    
    sn_list = generate_sn_list()
    LOGGER.info("\n本轮SN：")
    for idx, sn in enumerate(sn_list, start=1):
        LOGGER.info(f"通道{idx}：{sn}")

    LOGGER.info("\n5秒后开始测试。不会自动点击SN输入框，请确认焦点在第1个框。")
    time.sleep(5)
    save_layout_screenshot()

    statistics = load_statistics()

    for round_index in range(1, CONFIG["LOOP_COUNT"] + 1):
        LOGGER.info(f"\n\n{'=' * 80}\n开始第{round_index}/{CONFIG['LOOP_COUNT']}轮测试\n{'=' * 80}")
        round_start = time.time()
        try:
            scan_barcodes(sn_list)
            channel_states = run_parallel_test()
            print_round_result(round_index, channel_states)
            update_statistics(statistics, channel_states)
            save_statistics(statistics)
            
            elapsed = time.time() - round_start
            LOGGER.info(f"第{round_index}轮测试耗时：{elapsed:.2f}秒")
            
            if round_index < CONFIG["LOOP_COUNT"]:
                LOGGER.info(f"{CONFIG['BETWEEN_ROUNDS_DELAY']}秒后开始下一轮")
                time.sleep(CONFIG["BETWEEN_ROUNDS_DELAY"])

        except KeyboardInterrupt:
            LOGGER.warning("用户主动终止程序")
            break
        except Exception as e:
            LOGGER.exception(f"第{round_index}轮出现异常退出：{e}")
            
            # 【BUG修复】程序异常时，不仅要记录整机FAIL，也要记录对应所有通道的FAIL，以此保障统计基数一致性
            statistics["total_rounds"] += 1
            statistics["machine"]["fail"] += 1
            for channel_idx in range(1, CONFIG["CHANNEL_COUNT"] + 1):
                statistics["channels"][str(channel_idx)]["fail"] += 1
            
            save_statistics(statistics)
            break

    print_statistics(statistics)
    LOGGER.info("自动测试程序结束")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        LOGGER.warning("程序被用户手动终止")
    except Exception as e:
        LOGGER.exception(f"程序发生未处理的全局异常：{e}")
    finally:
        LOGGER.info("程序完全退出")