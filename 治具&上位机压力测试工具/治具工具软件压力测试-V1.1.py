# filename: auto_test_input.py
"""
治具工具自动化测试脚本 V1.3
功能：
- 在窗口中聚焦序列号输入框（Tab 或相对坐标回退）
- 清空并输入序列号（不回车），点击“开始测试”按钮
- 依次点击：显示 / 大灯 / 喇叭 的 PASS（严格使用 adjust=(100, 0)）
- 等待检测测试结果：支持识别 "通过"（result_pass.png）或 "不通过"（result_fail.png）
- 将每次测试结果写入 test_results.log
- 兼容 Windows DPI 缩放
注意：请把 `start_button.png`、`display_pass.png`、`headlight_pass.png`、
      `horn_pass.png`、`result_pass.png`（通过）和/或 `result_fail.png`（不通过）
      放在脚本同一目录下（根据你需要截取的结果图片）。
"""

import time
import os
import ctypes
import datetime
import pyautogui
import pygetwindow as gw
import winsound  # Windows 上播放提示音

# ---------- 配置区 ----------
WINDOW_TITLE_SUBSTR = "CCB 测试 V1.0.00"
SERIAL = "2030003003R538ZJ005600004"

IMG_START_BUTTON = "start_button.png"
IMG_DISPLAY_PASS = "display_pass.png"
IMG_HEADLIGHT_PASS = "headlight_pass.png"
IMG_HORN_PASS = "horn_pass.png"
IMG_RESULT_PASS = "result_pass.png"   # “通过”截图（必须）
IMG_RESULT_FAIL = "result_fail.png"   # 可选：显式“不通过”截图

LOCATE_CONFIDENCE = 0.8
WAIT_WINDOW_TIMEOUT = 30
WAIT_IMAGE_TIMEOUT = 20
WAIT_RESULT_TIMEOUT = 30

# 聚焦输入框的回退配置（如果你知道相对坐标可填入，否则使用 Tab）
SERIAL_BOX_COORD = None    # e.g. (450, 280) 相对于窗口左上角
SERIAL_TAB_STEPS = 3

# 日志文件
LOG_FILE = "test_results.log"

# ---------- 初始化 ----------
# 尝试启用 DPI 感知
try:
    ctypes.windll.user32.SetProcessDPIAware()
    print("[系统] DPI 感知已启用。")
except Exception:
    print("[警告] 无法启用 DPI 感知。")

pyautogui.FAILSAFE = True
try:
    SCALE = ctypes.windll.user32.GetDpiForSystem() / 96.0
except Exception:
    SCALE = 1.0
print(f"[系统] 当前缩放比例：{SCALE*100:.0f}%")

def log_result(serial, result):
    """把测试结果追加到日志文件（时间, 序列号, 结果）"""
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"{ts} - SN: {serial} - {result}\n"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line)
    except Exception as e:
        print("[日志写入失败]", e)

def beep_ok():
    try:
        winsound.Beep(1000, 150)  # 1kHz 150ms
    except Exception:
        pass

def beep_fail():
    try:
        winsound.Beep(400, 300)  # 400Hz 300ms
    except Exception:
        pass

# ---------- 工具函数 ----------
def wait_for_window_and_activate(substring, timeout=WAIT_WINDOW_TIMEOUT):
    print(f"等待窗口包含标题片段：'{substring}'（超时 {timeout}s）...")
    end = time.time() + timeout
    while time.time() < end:
        wins = gw.getWindowsWithTitle(substring)
        if wins:
            w = wins[0]
            try:
                w.activate()
                time.sleep(0.3)
            except Exception:
                pass
            print("✅ 已找到并激活窗口：", w.title)
            return w
        time.sleep(0.4)
    print("⚠️ 未在超时时间内找到目标窗口。")
    return None

def window_relative_to_screen(window, rel_x, rel_y):
    return window.left + rel_x, window.top + rel_y

def screen_click(x, y):
    """单次点击（会处理 SCALE）"""
    pyautogui.moveTo(x * SCALE, y * SCALE)
    time.sleep(0.06)
    pyautogui.click()

def focus_serial_box(window=None):
    """把焦点移动到序列号输入框：优先用坐标回退，否则用 Tab"""
    if SERIAL_BOX_COORD and window:
        abs_x, abs_y = window_relative_to_screen(window, *SERIAL_BOX_COORD)
        screen_click(abs_x, abs_y)
        print("🎯 使用坐标聚焦输入框")
    else:
        for _ in range(SERIAL_TAB_STEPS):
            pyautogui.press('tab')
            time.sleep(0.08)
        print(f"🎯 使用 Tab {SERIAL_TAB_STEPS} 次聚焦输入框")
    time.sleep(0.25)

def clear_and_type_serial(serial):
    """清空输入框并输入序列号（不回车）"""
    # 多重清空策略以提高成功率
    pyautogui.hotkey('ctrl', 'a'); time.sleep(0.05)
    pyautogui.press('delete'); time.sleep(0.05)
    for _ in range(3):
        pyautogui.press('backspace'); time.sleep(0.02)
    pyautogui.hotkey('ctrl', 'a'); time.sleep(0.03)
    pyautogui.press('delete'); time.sleep(0.03)
    pyautogui.typewrite(serial, interval=0.01)
    print("✏️ 已输入新序列号（未回车）")

def safe_locate_on_screen(img, confidence=LOCATE_CONFIDENCE):
    """封装 locateOnScreen，找不到或出错返回 None（不抛异常）"""
    if not os.path.exists(img):
        return None
    try:
        box = pyautogui.locateOnScreen(img, confidence=confidence)
        return box
    except Exception:
        return None

def wait_and_click_image(img, timeout=WAIT_IMAGE_TIMEOUT, confidence=LOCATE_CONFIDENCE, adjust=(0,0)):
    """
    找到图片并点击（返回 True/False）。不会抛出 locate 异常。
    adjust: (dx, dy) — 单位为屏幕像素（正数向右/下）。
    """
    end = time.time() + timeout
    while time.time() < end:
        box = safe_locate_on_screen(img, confidence=confidence)
        if box:
            cx, cy = pyautogui.center(box)
            # apply adjust (注意：cx/cy 已是屏幕坐标)
            tx = cx + adjust[0]
            ty = cy + adjust[1]
            pyautogui.moveTo(tx, ty)
            time.sleep(0.06)
            pyautogui.click(tx, ty)
            return True
        time.sleep(0.25)
    return False

# ---------- 主流程函数 ----------
def type_serial_and_start(serial):
    """激活窗口 -> 聚焦输入框 -> 清空并输入序列号 -> 点击开始测试按钮"""
    window = wait_for_window_and_activate(WINDOW_TITLE_SUBSTR)
    if not window:
        print("❌ 找不到目标窗口，退出。")
        return False
    focus_serial_box(window)
    clear_and_type_serial(serial)
    # 点击“开始测试”
    if os.path.exists(IMG_START_BUTTON):
        ok = wait_and_click_image(IMG_START_BUTTON, timeout=8, adjust=(0, 0))
        if ok:
            print("✅ 已点击“开始测试”按钮。")
            return True
        else:
            print("⚠️ 未检测到 start_button.png（或匹配失败）。请确认图片或手动点击。")
            return False
    else:
        print("⚠️ 未提供 start_button.png，请手动点击开始测试。")
        return False

def click_pass_sequence():
    """严格按照要求使用 adjust=(100, 0) 点击三项 PASS"""
    print("尝试依次点击：显示 -> 大灯 -> 喇叭 的 PASS 按钮（adjust=(100, 0)）")
    time.sleep(2)
    # **严格使用 adjust=(100, 0)**
    if wait_and_click_image(IMG_DISPLAY_PASS, timeout=12, adjust=(100, 0)):
        print("点击：显示 PASS （已用 adjust=(100,0)）")
    else:
        print("提示：未检测到显示 PASS（跳过）")
    time.sleep(0.8)
    if wait_and_click_image(IMG_HEADLIGHT_PASS, timeout=12, adjust=(100, 0)):
        print("点击：大灯 PASS （已用 adjust=(100,0)）")
    else:
        print("提示：未检测到大灯 PASS（跳过）")
    time.sleep(0.8)
    if wait_and_click_image(IMG_HORN_PASS, timeout=12, adjust=(100, 0)):
        print("点击：喇叭 PASS （已用 adjust=(100,0)）")
    else:
        print("提示：未检测到喇叭 PASS（跳过）")
    print("✅ 三项 PASS 点击流程结束。")

def wait_for_result(timeout=WAIT_RESULT_TIMEOUT):
    """
    等待“通过”或“不通过”出现，优先检测通过（result_pass），若 result_fail 存在也检测。
    返回: "PASS", "FAIL", 或 "TIMEOUT"
    """
    print(f"🕐 等待测试结果出现（超时 {timeout}s）...")
    end = time.time() + timeout
    while time.time() < end:
        # 先检测通过
        box_pass = safe_locate_on_screen(IMG_RESULT_PASS, confidence=LOCATE_CONFIDENCE)
        if box_pass:
            print("✅ 检测到: 测试结果 = 通过")
            return "PASS"
        # 再检测不通过（如果有图片）
        box_fail = safe_locate_on_screen(IMG_RESULT_FAIL, confidence=LOCATE_CONFIDENCE) if os.path.exists(IMG_RESULT_FAIL) else None
        if box_fail:
            print("❌ 检测到: 测试结果 = 不通过")
            return "FAIL"
        time.sleep(0.5)
    print("❌ 等待超时：未检测到测试结果（视为不通过）")
    return "TIMEOUT"

# ---------- 主程序 ----------
def main():
    print("=== 自动化测试脚本开始 ===")
    ok = type_serial_and_start(SERIAL)
    if not ok:
        print("⚠️ 未能启动测试（未点击开始测试），脚本结束。")
        log_result(SERIAL, "START_FAILED")
        beep_fail()
        return

    # 等待软件进入测试流程
    time.sleep(6)

    # 点击 PASS 三项（使用你指定的 adjust）
    click_pass_sequence()

    # 等待并判断最终结果
    result = wait_for_result(WAIT_RESULT_TIMEOUT)
    if result == "PASS":
        log_result(SERIAL, "PASS")
        beep_ok()
    elif result == "FAIL":
        log_result(SERIAL, "FAIL")
        beep_fail()
    else:  # TIMEOUT
        log_result(SERIAL, "TIMEOUT_AS_FAIL")
        beep_fail()

    print("=== 所有测试流程结束 ===")

if __name__ == "__main__":
    main()
