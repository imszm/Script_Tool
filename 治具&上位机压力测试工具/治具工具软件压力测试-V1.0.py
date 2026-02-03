# filename: auto_test_input.py
import time
import os
import ctypes
import pyautogui
import pygetwindow as gw
from PIL import ImageGrab, ImageDraw

# ---------- 配置区 ----------
WINDOW_TITLE_SUBSTR = "CCB 测试 V1.0.00"   # 程序窗口标题部分
SERIAL = "2030003003R538ZJ005600004"       # 序列号

IMG_SERIAL_BOX = "serial_box.png"
IMG_START_BUTTON = "start_button.png"      # 开始测试按钮截图
IMG_DISPLAY_PASS = "display_pass.png"
IMG_HEADLIGHT_PASS = "headlight_pass.png"
IMG_HORN_PASS = "horn_pass.png"
IMG_PASS_GENERIC = "pass_generic.png"

LOCATE_CONFIDENCE = 0.8
WAIT_WINDOW_TIMEOUT = 30
WAIT_IMAGE_TIMEOUT = 20
GENERIC_PASS_SCAN_DURATION = 40

# -------- 调试模式开关 --------
DEBUG_MODE = True  # True 开启调试模式，会保存截图并打印坐标
DEBUG_FOLDER = "debug_shots"
# --------------------------------

# 启用 DPI 感知
try:
    ctypes.windll.user32.SetProcessDPIAware()
    print("[系统] DPI 感知模式已启用。")
except Exception:
    print("[警告] 无法启用 DPI 感知，可能存在点击偏移。")

pyautogui.FAILSAFE = True  # 左上角中断

# 自动检测缩放比例
SCALE = ctypes.windll.user32.GetDpiForSystem() / 96.0  # 96 DPI = 100%
print(f"[系统] 当前缩放比例: {SCALE*100:.0f}%")

def debug_log(msg):
    if DEBUG_MODE:
        print("[DEBUG]", msg)

def save_debug_screenshot(region=None, tag=""):
    if not DEBUG_MODE:
        return
    os.makedirs(DEBUG_FOLDER, exist_ok=True)
    img = ImageGrab.grab()
    if region:
        x0, y0 = region.left, region.top
        x1, y1 = region.left + region.width, region.top + region.height
        draw = ImageDraw.Draw(img)
        draw.rectangle([x0, y0, x1, y1], outline="red", width=3)
    filename = os.path.join(DEBUG_FOLDER, f"{int(time.time())}_{tag}.png")
    img.save(filename)
    debug_log(f"已保存调试截图: {filename}")

def wait_for_window_and_activate(substring, timeout=WAIT_WINDOW_TIMEOUT):
    if not substring:
        print("未指定窗口标题，假定你已手动打开并置前程序。")
        return True
    print(f"等待窗口包含标题片段：'{substring}'（超时 {timeout}s）...")
    end = time.time() + timeout
    while time.time() < end:
        wins = gw.getWindowsWithTitle(substring)
        if wins:
            w = wins[0]
            try:
                w.activate()
            except Exception:
                pass
            print("已找到并尝试激活窗口：", w.title)
            return True
        time.sleep(0.5)
    print("⚠️ 未在超时时间内找到目标窗口。")
    return False

def wait_and_click_image(img, timeout=WAIT_IMAGE_TIMEOUT, confidence=LOCATE_CONFIDENCE, adjust=(0, 0)):
    """找到图片并点击，自动缩放修正"""
    if not os.path.exists(img):
        print(f"❌ 找不到参考图：{img}")
        return False
    print(f"查找图片 {img}（超时 {timeout}s）...")
    end = time.time() + timeout
    while time.time() < end:
        box = pyautogui.locateOnScreen(img, confidence=confidence)
        if box:
            center = pyautogui.center(box)
            # 缩放修正 + adjust微调
            x = center.x * SCALE + adjust[0]
            y = center.y * SCALE + adjust[1]
            debug_log(f"{img} box={box}, center={center}, adjusted=({x},{y})")
            save_debug_screenshot(region=box, tag=os.path.basename(img))
            pyautogui.moveTo(x, y)
            time.sleep(0.3)
            pyautogui.click(x, y)
            return True
        time.sleep(0.4)
    print(f"未找到图片：{img}")
    return False

def type_serial_to_box(serial, serial_box_img=IMG_SERIAL_BOX, start_button_img=IMG_START_BUTTON):
    """输入序列号并点击开始测试按钮"""
    ok = wait_and_click_image(serial_box_img, timeout=10)
    time.sleep(0.2)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.1)
    pyautogui.press('backspace')
    time.sleep(0.1)
    print("输入序列号：", serial)
    pyautogui.typewrite(serial, interval=0.01)
    # 点击开始测试
    if wait_and_click_image(start_button_img, timeout=10):
        print("已点击 '开始测试' 按钮。")
    else:
        print("⚠️ 未找到 '开始测试' 按钮，请手动点击。")

def click_pass_sequence():
    """依次点击显示 -> 大灯 -> 喇叭 PASS，支持 adjust 微调"""
    print("尝试依次点击：显示 -> 大灯 -> 喇叭 的 PASS 按钮")
    time.sleep(2)
    if not wait_and_click_image(IMG_DISPLAY_PASS, timeout=15, adjust=(100,0)):
        print("提示：未检测到显示 PASS")
    time.sleep(1)
    if not wait_and_click_image(IMG_HEADLIGHT_PASS, timeout=15, adjust=(100,0)):
        print("提示：未检测到大灯 PASS")
    time.sleep(1)
    if not wait_and_click_image(IMG_HORN_PASS, timeout=15, adjust=(100,1)):
        print("提示：未检测到喇叭 PASS")
    print("三项 PASS 点击流程结束。")

def generic_pass_scanner(duration=GENERIC_PASS_SCAN_DURATION):
    if not os.path.exists(IMG_PASS_GENERIC):
        print("未提供通用 PASS 图片，跳过。")
        return
    print(f"启动通用 PASS 扫描 {duration}s...")
    end = time.time() + duration
    while time.time() < end:
        box = pyautogui.locateOnScreen(IMG_PASS_GENERIC, confidence=LOCATE_CONFIDENCE)
        if box:
            center = pyautogui.center(box)
            x = center.x * SCALE
            y = center.y * SCALE
            print("找到通用 PASS，点击：", (x, y))
            save_debug_screenshot(region=box, tag="generic_pass")
            pyautogui.click(x, y)
            time.sleep(1.2)
        else:
            time.sleep(0.4)
    print("通用扫描结束。")

def main():
    print("=== 自动化脚本开始 ===")
    screen_w, screen_h = pyautogui.size()
    print(f"[系统] 当前屏幕分辨率: {screen_w}x{screen_h}")
    ok = wait_for_window_and_activate(WINDOW_TITLE_SUBSTR)
    if not ok:
        print("警告：未激活窗口，脚本仍会继续。")
        time.sleep(2)

    time.sleep(1)
    type_serial_to_box(SERIAL)
    time.sleep(6)
    click_pass_sequence()
    generic_pass_scanner(GENERIC_PASS_SCAN_DURATION)
    print("=== 脚本执行完毕 ===")

if __name__ == "__main__":
    main()
