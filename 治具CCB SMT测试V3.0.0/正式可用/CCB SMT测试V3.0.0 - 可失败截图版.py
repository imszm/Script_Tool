# -*- coding: utf-8 -*-
from pywinauto import Application
from pywinauto.keyboard import send_keys
from pywinauto.mouse import click
import time
import datetime
import os

# ================= 配置区域 =================
PASS_LIGHT_POS = (881, 255)
PASS_HORN_POS  = (884, 312)

LOOP_COUNT = 10000
WAIT_AFTER_START = 6
WAIT_FULL_PROCESS = 19
LOOP_INTERVAL = 1
SERIAL_PREFIX = "2010007005R615GD005900"

FAIL_SCREENSHOT_DIR = "fail_screenshots"
os.makedirs(FAIL_SCREENSHOT_DIR, exist_ok=True)
# ==================================================

def log(msg):
    print(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}")

# ----------------- 连接测试软件窗口 -----------------
log("连接测试软件窗口...")
app = Application(backend="uia").connect(title_re="CCB 测试 V3.0.00.*")
dlg = app.window(title_re="CCB 测试 V3.0.00.*")
log("连接成功！")

# ----------------- 自动识别串口按钮 -----------------
btn_open = None
for btn in dlg.descendants(control_type="Button"):
    if btn.window_text() in ["打开", "关闭"]:
        btn_open = btn
        break

if btn_open is None:
    log("未找到串口按钮！请确认窗口是否正确打开。")
    raise Exception("串口按钮未找到")

# ----------------- 截图函数 -----------------
def capture_fail_screen(serial_number):
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = os.path.join(FAIL_SCREENSHOT_DIR, f"FAIL_{serial_number}_{timestamp}.png")
    dlg.capture_as_image().save(filename)
    log(f"失败截图已保存：{filename}")

# ----------------- 只判断右下角大字“不通过” -----------------
def is_fail_by_big_label():
    try:
        fail_big = dlg.child_window(title="不通过", control_type="Text")
        return fail_big.exists()
    except:
        return False

# =================== 测试循环 ===================
for idx in range(1, LOOP_COUNT + 1):
    SERIAL_NUMBER = f"{SERIAL_PREFIX}{str(idx).zfill(3)}"
    log(f"======== 开始测试序列号 {SERIAL_NUMBER} （第 {idx} 台） ========")

    # 串口打开
    if btn_open.window_text() == "打开":
        log("串口当前为关闭状态，点击打开串口")
        btn_open.click_input()
        time.sleep(1)
    else:
        log("串口已打开，直接回车执行测试")

    # 写 SN
    edit_sn = dlg.child_window(auto_id="Widget.lineEditSerialNumber", control_type="Edit")
    edit_sn.set_edit_text("")
    edit_sn.set_edit_text(SERIAL_NUMBER)
    time.sleep(0.5)

    # 开始测试
    log("按回车键触发开始测试")
    send_keys('{ENTER}')

    # PASS 两项
    time.sleep(WAIT_AFTER_START)
    click(coords=PASS_LIGHT_POS)
    time.sleep(0.5)
    click(coords=PASS_HORN_POS)
    time.sleep(0.5)

    log(f"等待测试流程结束，大约 {WAIT_FULL_PROCESS} 秒...")
    time.sleep(WAIT_FULL_PROCESS)

    # -------------------- 判断右下角“不通过” --------------------
    if is_fail_by_big_label():
        log(f"{SERIAL_NUMBER} 检测到右下角“不通过”，正在截图")
        capture_fail_screen(SERIAL_NUMBER)
    else:
        log(f"{SERIAL_NUMBER} 测试通过")

    log(f"==================== 第 {idx} 台测试结束 ====================\n")
    time.sleep(LOOP_INTERVAL)

log("所有测试完成！")
