# -*- coding: utf-8 -*-
"""
W3 PCTOOL 压力测试脚本 V1.6
改进说明：
1. 检测到 PCTOOL 闪退后不再无限重连，直接终止脚本；
2. 每个动作延时参数均添加详细注释；
3. 日志记录更加清晰；
4. 每次循环记录成功率、失败点；
"""

import time
import random
from pywinauto import Application
from datetime import datetime
import sys

# ================== 基本配置 ==================
APP_TITLE = "W3 PCTOOL V5.5.00"   # 工具窗口标题
TEST_CYCLES = 100000              # 测试循环次数
SN_PREFIX = "W30300GZP004A00"     # SN 编号前缀
LOG_FILE = "tool_stress_log.txt"  # 日志文件名

# ================== 各动作独立延时（秒） ==================
# 每个时间参数都对应一个具体按钮或动作
DELAY_START = 2          # 点击“开始”按钮后等待时间
DELAY_SN_WRITE = 6      # 点击“SN写入”按钮后等待时间
DELAY_LOG_ENABLE = 5     # 点击“开启日志”后等待时间
DELAY_LOG_DISABLE = 5    # 点击“关闭日志”后等待时间（同一个按钮）
DELAY_IMU = 20           # 点击“IMU校准”后等待时间
DELAY_RESTORE = 20       # 点击“恢复出厂设置”后等待时间
DELAY_PARAM_SET = 8      # 点击“参数设置->确认”后等待时间
DELAY_JOYSTICK = 15      # 点击“摇杆校准”后等待时间

# ================== 工具函数 ==================
def log(msg):
    """写入日志并打印"""
    t = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{t}] {msg}\n")
    print(f"[{t}] {msg}")

def find_and_click_confirm(app, timeout=8):
    """检测‘参数设置’弹窗并点击‘确认’按钮"""
    start_t = time.time()
    while time.time() - start_t < timeout:
        try:
            param_win = app.window(title_re=".*参数设置.*")
            if param_win.exists(timeout=1):
                confirm_btn = param_win.child_window(title="确认", control_type="Button")
                if confirm_btn.exists(timeout=1):
                    confirm_btn.click_input()
                    log("已点击 参数设置弹窗 的 ‘确认’ 按钮")
                    return True
        except Exception:
            pass
        time.sleep(0.5)
    log("未检测到 参数设置弹窗 或 ‘确认’ 按钮")
    return False

def wait_action(label, seconds):
    """通用延时函数"""
    log(f"等待 {seconds} 秒 -> {label}")
    time.sleep(seconds)

def connect_app():
    """连接 W3 PCTOOL 主窗口"""
    try:
        app = Application(backend="uia").connect(title_re=APP_TITLE)
        dlg = app.window(title_re=APP_TITLE)
        dlg.wait("ready", timeout=5)
        return app, dlg
    except Exception as e:
        log(f"连接 PCTOOL 失败: {e}")
        return None, None

# ================== 主程序 ==================
log("=== 启动 PC 工具压力测试 V1.6 ===")
app, dlg = connect_app()
if dlg is None:
    sys.exit("❌ 未检测到 PCTOOL，请先启动软件再运行脚本。")

# ========== 控件绑定 ==========
sn_input = dlg.child_window(auto_id="Widget.lineEditSN", control_type="Edit")       # SN输入框
button_start = dlg.child_window(title="开始", control_type="Button")                # “开始”按钮
button_write_sn = dlg.child_window(auto_id="Widget.buttonWriteSN", control_type="Button")  # “SN写入”
button_restore = dlg.child_window(auto_id="Widget.buttonRestore", control_type="Button")   # “恢复出厂设置”
button_params = dlg.child_window(auto_id="Widget.buttonSetParams", control_type="Button")  # “参数设置”
button_imu = dlg.child_window(auto_id="Widget.buttonIMUCali", control_type="Button")       # “IMU校准”
button_joystick = dlg.child_window(auto_id="Widget.buttonJoystickCali", control_type="Button")  # “摇杆校准”
button_log = dlg.child_window(auto_id="Widget.buttonLogEnable", control_type="Button")     # “开启/关闭日志”

# ========== 统计变量 ==========
total = 0
success = 0
failed = 0
last_action = ""

# ========== 主循环 ==========
for i in range(1, TEST_CYCLES + 1):
    total += 1
    try:
        if not dlg.exists():
            raise Exception("检测到 PCTOOL 窗口丢失，可能已闪退。")

        log(f"===== 循环测试 {i} 开始 =====")

        # 1. 输入 SN
        sn_number = f"{SN_PREFIX}{random.randint(100,999)}"
        last_action = f"输入 SN: {sn_number}"
        sn_input.set_edit_text(sn_number)
        log(last_action)

        # 2. 点击 “开始”
        last_action = "点击 ‘开始’ 按钮"
        button_start.click_input()
        log(last_action)
        wait_action("开始", DELAY_START)

        # 3. 点击 “SN写入”
        last_action = "点击 ‘SN写入’ 按钮"
        button_write_sn.click_input()
        log(last_action)
        wait_action("SN写入", DELAY_SN_WRITE)

        # 4. 点击 “开启日志”
        last_action = "点击 ‘开启日志’"
        button_log.click_input()
        log(last_action)
        wait_action("开启日志", DELAY_LOG_ENABLE)

        # 5. 点击 “关闭日志”
        last_action = "点击 ‘关闭日志’"
        button_log.click_input()
        log(last_action)
        wait_action("关闭日志", DELAY_LOG_DISABLE)

        # 6. 点击 “IMU校准”
        last_action = "点击 ‘IMU校准’"
        button_imu.click_input()
        log(last_action)
        wait_action("IMU校准", DELAY_IMU)

        # 7. 点击 “恢复出厂设置”
        last_action = "点击 ‘恢复出厂设置’"
        button_restore.click_input()
        log(last_action)
        wait_action("恢复出厂设置", DELAY_RESTORE)

        # 8. 点击 “参数设置 -> 确认”
        last_action = "点击 ‘参数设置 -> 确认’"
        button_params.click_input()
        log(last_action)
        find_and_click_confirm(app)
        wait_action("参数设置", DELAY_PARAM_SET)

        # 9. 点击 “摇杆校准”
        last_action = "点击 ‘摇杆校准’"
        button_joystick.click_input()
        log(last_action)
        wait_action("摇杆校准", DELAY_JOYSTICK)

        # 统计成功
        success += 1
        log(f"✅ 第 {i} 次测试完成。当前成功率：{success}/{total} = {success/total:.2%}\n")

    except Exception as e:
        failed += 1
        log(f"❌ 第 {i} 次循环异常: {e}")
        log(f"最后动作: {last_action}")

        # 检测闪退 -> 直接终止脚本
        app, dlg = connect_app()
        if dlg is None:
            log("⚠️ 检测到 PCTOOL 闪退！已记录异常并终止脚本。")
            log(f"最后执行动作: {last_action}")
            log(f"已执行次数: {total}, 成功: {success}, 失败: {failed}")
            sys.exit(1)
        continue

# ========== 总结 ==========
log("=== 压力测试任务结束 ===")
log(f"总次数: {total}, 成功: {success}, 失败: {failed}, 成功率: {success/total:.2%}")
