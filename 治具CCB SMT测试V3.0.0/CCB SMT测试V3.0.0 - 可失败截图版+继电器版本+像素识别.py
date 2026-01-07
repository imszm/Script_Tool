# -*- coding: utf-8 -*-
from pywinauto import Application
from pywinauto.keyboard import send_keys
from PIL import ImageGrab
import time
import datetime
import os
import serial
import win32api
import win32con

# ================= 配置区域 =================

# 1. 屏幕坐标配置
PASS_LIGHT_POS = (1165, 224)
PASS_HORN_POS  = (1171, 274)
STATUS_CHECK_POINTS = [(1701, 820), (1846, 812)] # 结果检测点

# 2. 序列号
START_INDEX = 0000

# 计划测试的总数量
LOOP_COUNT = 10000      

# [修改] 序列号前缀
SERIAL_PREFIX = "2010007005R615GD00590"

# 3. 时间与超时配置
WAIT_AFTER_START = 5          # 启动后的初始等待
WAIT_FULL_PROCESS_TIMEOUT = 80 # 结果监听超时时间
LOOP_INTERVAL = 1             # 两次测试间的间隔

# 4. 硬件配置
RELAY_PORT = "COM12"
RELAY_BAUD = 9600

# 5. 文件存储配置
FAIL_SCREENSHOT_DIR = "fail_screenshots"
os.makedirs(FAIL_SCREENSHOT_DIR, exist_ok=True)

# ==================================================

def log(msg):
    """打印带时间戳的日志"""
    print(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}")

def fast_click(x, y):
    """快速模拟鼠标点击"""
    win32api.SetCursorPos((x, y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)
    time.sleep(0.1)

# ----------------- 颜色检测逻辑 -----------------
def check_test_status_fuzzy():
    """
    模糊颜色检测，用于判断 PASS (绿色) 或 FAIL (红色)
    """
    try:
        screen = ImageGrab.grab()
        width, height = screen.size
        SEARCH_RADIUS = 10 
        
        for (center_x, center_y) in STATUS_CHECK_POINTS:
            for x in range(center_x - SEARCH_RADIUS, center_x + SEARCH_RADIUS, 2):
                for y in range(center_y - SEARCH_RADIUS, center_y + SEARCH_RADIUS, 2):
                    if x < 0 or x >= width or y < 0 or y >= height:
                        continue
                        
                    rgb = screen.getpixel((x, y))
                    r, g, b = rgb
                    
                    # --- 红色判断 (FAIL) ---
                    # 逻辑：R值高，且G和B值明显低，排除橙色/黄色干扰
                    if r > 200 and g < 100 and b < 100: 
                        log(f"检测到【红色 FAIL】 坐标:({x},{y}) RGB:{rgb}")
                        return "FAIL"
            
                    # --- 绿色判断 (PASS) ---
                    # 逻辑：G值高，且G必须比R和B都高出一定阈值
                    if g > 140 and g > r + 30 and g > b + 30:
                        log(f"检测到【绿色 PASS】 坐标:({x},{y}) RGB:{rgb}")
                        return "PASS"
        
    except Exception as e:
        log(f"图像识别错误: {e}")
    
    return None

# ----------------- 串口/继电器操作 -----------------
try:
    relay_ser = serial.Serial(RELAY_PORT, RELAY_BAUD, timeout=1)
    log("继电器串口连接成功")
except Exception as e:
    log(f"继电器串口连接失败: {e}")
    raise e

def relay_restart():
    """执行继电器断电重启流程"""
    log("执行继电器重启流程...")
    relay_ser.write(b"P") # 发送关闭指令
    time.sleep(2)
    relay_ser.write(b"O") # 发送开启指令
    time.sleep(2)

# ----------------- 软件连接流程 -----------------
log("初始化继电器...")
relay_restart()

log("连接测试软件...")
try:
    # 使用 uia 后端连接测试软件
    app = Application(backend="uia").connect(title_re="CCB 测试 V3.2.00.*")
    dlg = app.window(title_re="CCB 测试 V3.2.00.*")
    dlg.set_focus() 
    log("连接成功")
except Exception as e:
    log(f"连接失败: {e}")
    raise e

# 寻找打开/关闭串口的按钮
btn_open = None
for btn in dlg.descendants(control_type="Button"):
    if btn.window_text() in ["打开", "关闭"]:
        btn_open = btn
        break
if btn_open is None:
    raise Exception("串口按钮未找到")

def capture_fail_screen(serial_number):
    """保存失败时的截图"""
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = os.path.join(FAIL_SCREENSHOT_DIR, f"FAIL_{serial_number}_{timestamp}.png")
    try:
        dlg.capture_as_image().save(filename)
        log(f"失败截图保存: {filename}")
    except:
        pass

def is_fail_by_big_label():
    """通过界面文字辅助判断失败"""
    try:
        return dlg.child_window(title="不通过", control_type="Text").exists(timeout=2)
    except:
        return False

# =================== 主循环测试 ===================
# 循环范围：从 START_INDEX 开始，往后测 LOOP_COUNT 个
end_index = START_INDEX + LOOP_COUNT

for idx in range(START_INDEX, end_index):
    
    # [修正后的序列号生成逻辑]
    # 1. zfill(4) 保证数字至少4位 (1->0001, 2831->2831)
    # 2. 拼接前缀 ...590 (少一个0)
    # 结果：...590 + 2831 = ...5902831
    SERIAL_NUMBER = f"{SERIAL_PREFIX}{str(idx).zfill(4)}"
    
    log(f"=== 开始测试: {SERIAL_NUMBER} (第 {idx} 台) ===")

    # 检查串口状态，确保是打开的
    if btn_open.window_text() == "打开":
        log("点击打开串口")
        btn_open.click_input()
        time.sleep(4)

    # 输入序列号
    try:
        edit_sn = dlg.child_window(auto_id="Widget.lineEditSerialNumber", control_type="Edit")
        edit_sn.set_focus() 
        edit_sn.set_edit_text(SERIAL_NUMBER)
    except Exception as e:
        log(f"输入SN错误: {e}")
    
    time.sleep(1)

    log("按回车键开始")
    send_keys('{ENTER}')
    
    time.sleep(WAIT_AFTER_START)
    
    # 模拟人工点击操作
    log("点击 PASS_LIGHT")
    fast_click(PASS_LIGHT_POS[0], PASS_LIGHT_POS[1])
    time.sleep(0.5) 
    
    log("点击 PASS_HORN")
    fast_click(PASS_HORN_POS[0], PASS_HORN_POS[1])
    time.sleep(0.5)

    # 强制等待，避开测试初期的黄色/过渡色
    log("等待界面响应 (避开过渡色)...")
    time.sleep(3)

    log(f"开始监听结果 (上限 {WAIT_FULL_PROCESS_TIMEOUT}s)...")
    
    test_start_time = time.time()
    final_result = "UNKNOWN"
    
    # 结果监听循环
    while time.time() - test_start_time < WAIT_FULL_PROCESS_TIMEOUT:
        status = check_test_status_fuzzy()
        
        if status == "FAIL":
            # 二次确认机制：防止闪烁误判
            time.sleep(0.5)
            if check_test_status_fuzzy() == "FAIL":
                final_result = "FAIL"
                log(">>> 确认失败信号！")
                break
            else:
                log(">>> 虚晃一枪，继续监听...")
            
        if status == "PASS":
            final_result = "PASS"
            log(">>> 捕捉到通过信号！")
            break
        
        time.sleep(1)

    # === 结果后处理 ===
    if final_result == "FAIL":
        log("结果：【不通过】 -> 截图并等待")
        capture_fail_screen(SERIAL_NUMBER)
        time.sleep(2) 
        
    elif final_result == "PASS":
        log("结果：【通过】 -> 等待 2 秒...")
        time.sleep(2)
        
    else:
        # 超时处理，再次检查是否有文字提示失败
        if is_fail_by_big_label():
            log("结果：超时判定为不通过")
            capture_fail_screen(SERIAL_NUMBER)
        else:
            log("结果：超时 -> 默认通过")

    # 重启板子，准备下一次测试
    relay_restart()

    log(f"=== 第 {idx} 台结束 ===\n")
    time.sleep(LOOP_INTERVAL)

log("所有测试完成")
