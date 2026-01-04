# filename: 治具工具软件压力测试-V2.0.py
"""
治具工具软件自动化压力测试脚本 V2.0
------------------------------------------------
功能说明：
- 自动执行 1000 次完整测试流程
- 每次执行：
    1. 激活窗口 -> 输入序列号 -> 点击“开始测试”
    2. 点击 显示/大灯/喇叭 PASS（adjust=(100,0)）
    3. 检测“通过”或“不通过”结果
- 实时统计成功率与执行次数
- 自动记录到 test_results.log
- 支持中断 (Ctrl+C)
"""

import time
import os
import ctypes
import datetime
import pyautogui
import pygetwindow as gw
import winsound
import random

# ==========================================================
# 一、基础参数配置
# ==========================================================
WINDOW_TITLE_SUBSTR = "CCB 测试 V1.0.00"   # 要操作的程序窗口标题关键字

# 图像模板文件名（截图文件）
IMG_START_BUTTON   = "start_button.png"    # “开始测试”按钮
IMG_DISPLAY_PASS   = "display_pass.png"    # 显示 PASS 按钮
IMG_HEADLIGHT_PASS = "headlight_pass.png"  # 大灯 PASS 按钮
IMG_HORN_PASS      = "horn_pass.png"       # 喇叭 PASS 按钮
IMG_RESULT_PASS    = "result_pass.png"     # “通过”提示
IMG_RESULT_FAIL    = "result_fail.png"     # “不通过”提示（可选）

# 图像识别参数
LOCATE_CONFIDENCE = 0.8        # 识别相似度
WAIT_WINDOW_TIMEOUT = 30       # 等待窗口出现超时时间
WAIT_IMAGE_TIMEOUT  = 20       # 等待图片识别超时
WAIT_RESULT_TIMEOUT = 30       # 等待结果出现超时
TEST_COUNT = 1000              # 压力测试次数

# 序列号前缀（后面会自动加编号）
SERIAL_PREFIX = "2030003003R538ZJ005600"

LOG_FILE = "test_results.log"  # 测试结果保存文件

# ==========================================================
# 二、系统初始化：DPI 处理、PyAutoGUI 设置
# ==========================================================
try:
    # 启用 DPI 感知，防止高分屏坐标偏移
    ctypes.windll.user32.SetProcessDPIAware()
    print("[系统] DPI 感知已启用。")
except Exception:
    pass

pyautogui.FAILSAFE = True  # 鼠标移到左上角可紧急终止程序

# 获取当前系统缩放比例（100%、125%、150% 等）
try:
    SCALE = ctypes.windll.user32.GetDpiForSystem() / 96.0
except Exception:
    SCALE = 1.0
print(f"[系统] 当前缩放比例：{SCALE*100:.0f}%")

# ==========================================================
# 三、通用工具函数
# ==========================================================

# 播放成功提示音
def beep_ok():
    winsound.Beep(1000, 150)

# 播放失败提示音
def beep_fail():
    winsound.Beep(400, 300)

# 记录测试结果到日志文件
def log_result(serial, result):
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"{ts} - SN: {serial} - {result}\n")

# 等待目标窗口出现并激活
def wait_for_window_and_activate(substring, timeout=WAIT_WINDOW_TIMEOUT):
    end = time.time() + timeout
    while time.time() < end:
        wins = gw.getWindowsWithTitle(substring)
        if wins:
            w = wins[0]
            try:
                w.activate()  # 激活窗口到最前
            except:
                pass
            time.sleep(0.3)
            return w
        time.sleep(0.3)
    return None

# 安全查找图片（带异常捕获）
def safe_locate_on_screen(img, confidence=LOCATE_CONFIDENCE):
    if not os.path.exists(img):
        return None
    try:
        return pyautogui.locateOnScreen(img, confidence=confidence)
    except Exception:
        return None

# 等待图片出现并点击（支持微调 adjust）
def wait_and_click_image(img, timeout=WAIT_IMAGE_TIMEOUT, confidence=LOCATE_CONFIDENCE, adjust=(0,0)):
    end = time.time() + timeout
    while time.time() < end:
        box = safe_locate_on_screen(img, confidence=confidence)
        if box:
            # 计算中心坐标 + 微调偏移
            cx, cy = pyautogui.center(box)
            tx, ty = cx + adjust[0], cy + adjust[1]
            pyautogui.moveTo(tx, ty)
            time.sleep(0.05)
            pyautogui.click(tx, ty)
            return True
        time.sleep(0.25)
    return False

# 连续按三次 Tab 聚焦到序列号输入框
def focus_serial_box():
    for _ in range(3):
        pyautogui.press('tab')
        time.sleep(0.05)

# 清空并输入新的序列号
def clear_and_type_serial(serial):
    pyautogui.hotkey('ctrl', 'a')  # 全选
    pyautogui.press('delete')      # 删除
    pyautogui.typewrite(serial, interval=0.01)  # 输入序列号

# 输入序列号并点击“开始测试”
def type_serial_and_start(serial):
    window = wait_for_window_and_activate(WINDOW_TITLE_SUBSTR)
    if not window:
        return False
    focus_serial_box()
    clear_and_type_serial(serial)
    ok = wait_and_click_image(IMG_START_BUTTON, timeout=8, adjust=(0, 0))
    return ok

# 依次点击三个 PASS 按钮
def click_pass_sequence():
    wait_and_click_image(IMG_DISPLAY_PASS, timeout=12, adjust=(100, 0))
    wait_and_click_image(IMG_HEADLIGHT_PASS, timeout=12, adjust=(100, 0))
    wait_and_click_image(IMG_HORN_PASS, timeout=12, adjust=(100, 0))

# 等待测试结果（识别“通过”或“不通过”）
def wait_for_result(timeout=WAIT_RESULT_TIMEOUT):
    end = time.time() + timeout
    while time.time() < end:
        if safe_locate_on_screen(IMG_RESULT_PASS):
            return "PASS"
        if os.path.exists(IMG_RESULT_FAIL) and safe_locate_on_screen(IMG_RESULT_FAIL):
            return "FAIL"
        time.sleep(0.5)
    return "TIMEOUT"  # 超时未识别到结果

# ==========================================================
# 四、主程序：执行自动化压力测试
# ==========================================================
def main():
    success_count = 0
    fail_count = 0

    print(f"=== 自动化压力测试开始，共 {TEST_COUNT} 次 ===")
    start_time = time.time()

    try:
        for i in range(1, TEST_COUNT + 1):
            serial = SERIAL_PREFIX + str(i).zfill(3)  # 生成序列号（带编号）
            print(f"\n—— 第 {i}/{TEST_COUNT} 次测试 | SN={serial} ——")

            # 步骤 1：输入序列号并启动测试
            ok = type_serial_and_start(serial)
            if not ok:
                print("❌ 启动失败（未点击到开始按钮）")
                result = "START_FAIL"
            else:
                # 步骤 2：等待界面加载，然后依次点击三个 PASS 按钮
                time.sleep(6)
                click_pass_sequence()

                # 步骤 3：等待“通过”或“不通过”结果出现
                result = wait_for_result(WAIT_RESULT_TIMEOUT)

            # 步骤 4：统计结果
            if result == "PASS":
                success_count += 1
                beep_ok()
            else:
                fail_count += 1
                beep_fail()

            # 写入日志文件
            log_result(serial, result)

            # 打印当前统计信息
            rate = success_count / (success_count + fail_count) * 100
            print(f"结果：{result} | 当前成功率：{rate:.1f}% | 用时：{int(time.time()-start_time)}s")

            # 可选延时，避免点击过快导致界面卡顿
            time.sleep(1)

    except KeyboardInterrupt:
        print("\n用户中断测试。")

    # 测试结束汇总
    total = success_count + fail_count
    print("\n=== 压力测试完成 ===")
    print(f"总测试次数：{total}")
    print(f"成功次数：{success_count}")
    print(f"失败次数：{fail_count}")
    print(f"成功率：{(success_count/total*100 if total>0 else 0):.2f}%")

# ==========================================================
# 程序入口
# ==========================================================
if __name__ == "__main__":
    main()
