# -*- coding: utf-8 -*-
from pywinauto import Application
from pywinauto.keyboard import send_keys
import time
import datetime
import os
import serial
import win32api
import win32con

# ================= 配置区域 =================
PASS_LIGHT_POS = (1165, 224)
PASS_HORN_POS  = (1171, 274)

LOOP_COUNT = 10000

# [优化点1] 启动后的等待时间由 8秒 改为 3秒
# 如果软件反应慢，按钮还没出来就点了，请把这个数字改大一点（例如 5）
WAIT_AFTER_START = 5 

WAIT_FULL_PROCESS = 80
LOOP_INTERVAL = 1
SERIAL_PREFIX = "2010007005R615GD005900"

RELAY_PORT = "COM4"
RELAY_BAUD = 9600

FAIL_SCREENSHOT_DIR = "fail_screenshots"
os.makedirs(FAIL_SCREENSHOT_DIR, exist_ok=True)
# ==================================================

def log(msg):
    print(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}")

# ----------------- 极速点击函数 -----------------
def fast_click(x, y):
    win32api.SetCursorPos((x, y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)
    time.sleep(0.1)

# ----------------- 继电器控制 -----------------
def relay_on(ser):
    ser.write(b"O")
    log("继电器 -> 开启(O)")

def relay_off(ser):
    ser.write(b"P")
    log("继电器 -> 关闭(P)")

# ----------------- 初始化连接 -----------------
try:
    relay_ser = serial.Serial(RELAY_PORT, RELAY_BAUD, timeout=1)
    log("继电器串口连接成功")
except Exception as e:
    log(f"继电器串口连接失败: {e}")
    raise e

log("首次启动初始化：继电器重新关闭/打开...")
relay_off(relay_ser)
time.sleep(3)
relay_on(relay_ser)
time.sleep(3)
log("首次继电器初始化完成")

log("连接测试软件窗口...")
try:
    app = Application(backend="uia").connect(title_re="CCB 测试 V3.1.00.*")
    dlg = app.window(title_re="CCB 测试 V3.1.00.*")
    dlg.set_focus() 
    log("连接成功")
except Exception as e:
    log(f"连接窗口失败: {e}")
    raise e

btn_open = None
for btn in dlg.descendants(control_type="Button"):
    if btn.window_text() in ["打开", "关闭"]:
        btn_open = btn
        break
if btn_open is None:
    raise Exception("串口按钮未找到")

def capture_fail_screen(serial_number):
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = os.path.join(FAIL_SCREENSHOT_DIR, f"FAIL_{serial_number}_{timestamp}.png")
    try:
        dlg.capture_as_image().save(filename)
        log(f"失败截图保存: {filename}")
    except:
        pass

def is_fail_by_big_label():
    try:
        return dlg.child_window(title="不通过", control_type="Text").exists(timeout=2)
    except:
        return False

# =================== 主测试循环 ===================
for idx in range(1, LOOP_COUNT + 1):
    SERIAL_NUMBER = f"{SERIAL_PREFIX}{str(idx).zfill(3)}"
    log(f"=== 开始测试: {SERIAL_NUMBER} (第 {idx} 台) ===")

    # 检查串口状态
    if btn_open.window_text() == "打开":
        log("点击打开串口")
        btn_open.click_input()
        time.sleep(4)

    # 输入SN
    try:
        edit_sn = dlg.child_window(auto_id="Widget.lineEditSerialNumber", control_type="Edit")
        edit_sn.set_focus() 
        edit_sn.set_edit_text(SERIAL_NUMBER)
    except Exception as e:
        log(f"输入SN错误: {e}")
    
    time.sleep(1)

    # 开始测试
    log("按回车键开始")
    send_keys('{ENTER}')
    
    # [优化点2] 删除了这里原本的 time.sleep(10)，直接使用配置的短等待
    time.sleep(WAIT_AFTER_START) # 现在只等 3 秒
    
    log("点击 PASS_LIGHT")
    fast_click(PASS_LIGHT_POS[0], PASS_LIGHT_POS[1])
    
    # [优化点3] 两次点击间隔由 2秒 缩短为 0.5秒
    time.sleep(0.5) 
    
    log("点击 PASS_HORN")
    fast_click(PASS_HORN_POS[0], PASS_HORN_POS[1])
    
    time.sleep(0.5)

    log(f"等待流程结束 ({WAIT_FULL_PROCESS}s)...")
    time.sleep(WAIT_FULL_PROCESS)

    if is_fail_by_big_label():
        log("结果：不通过 -> 截图")
        capture_fail_screen(SERIAL_NUMBER)
    else:
        log("结果：通过")

    log("继电器重启...")
    relay_off(relay_ser)
    time.sleep(10) # 这里如果觉得继电器断电太久，也可以改小，比如 5
    relay_on(relay_ser)
    time.sleep(10)

    log(f"=== 第 {idx} 台结束 ===\n")
    time.sleep(LOOP_INTERVAL)

log("测试完成")
