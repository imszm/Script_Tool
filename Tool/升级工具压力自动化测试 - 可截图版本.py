import time
import csv
import os
from datetime import datetime
from pywinauto.application import Application

# ================================
# 配置
# ================================
APP_TITLE = "L5 PCTOOL V3.9.00"
CYCLE_WAIT = 170
LOG_PATH = "test_log.csv"
SCREENSHOT_DIR = "screenshots"

# ================================
# 初始化
# ================================
if not os.path.exists(SCREENSHOT_DIR):
    os.mkdir(SCREENSHOT_DIR)

if not os.path.exists(LOG_PATH):
    with open(LOG_PATH, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["时间", "结果", "详情"])

print("正在连接程序…")

app = Application(backend="uia").connect(title=APP_TITLE, timeout=10)
win = app.window(title=APP_TITLE)

button_start = win.child_window(auto_id="Widget.buttonUpgrade", control_type="Button")
log_edit = win.child_window(auto_id="Widget.textEditLog", control_type="Edit")

total = 0
fail = 0
success = 0

print("开始压力测试… Ctrl+C 可停止")

# ================================
# 主循环
# ================================
while True:
    total += 1
    start_time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    print(f"\n============== 第 {total} 轮测试开始 ==============")

    # ------- 读取开始前日志长度 -------
    old_log = log_edit.get_value()
    old_len = len(old_log)

    # 点击开始
    button_start.click()

    print(f"等待 {CYCLE_WAIT} 秒…")
    time.sleep(CYCLE_WAIT)

    # ------- 获取本轮新增内容 -------
    new_log_full = log_edit.get_value()
    new_part = new_log_full[old_len:] if len(new_log_full) > old_len else ""

    if not new_part.strip():
        new_part = "(本轮无日志输出)"

    # ------- 判断Pass/Fail -------
    if ("失败" in new_part) or ("超时" in new_part):
        result = "失败"
        fail += 1

        # -------- 截图 --------
        screenshot_path = os.path.join(
            SCREENSHOT_DIR,
            f"fail_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        )
        win.capture_as_image().save(screenshot_path)
        print(f"❌ 检测到失败，本次测试终止，截图已保存：{screenshot_path}")

        # -------- 写入 CSV --------
        with open(LOG_PATH, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow([start_time_str, result, new_part])

        # -------- 打印最终统计并退出 --------
        print("\n========= 终止测试（检测到失败） =========")
        print(f"总轮数：{total}")
        print(f"成功：{success}")
        print(f"失败：{fail}")
        success_rate = (success / total) * 100
        print(f"成功率：{success_rate:.2f}%")
        print("====================================\n")

        break   # <<< 核心：失败后立即终止脚本

    else:
        result = "成功"
        success += 1
        print("✔ 本轮成功")

        # 写入成功记录
        with open(LOG_PATH, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow([start_time_str, result, new_part])

    # ------- 显示统计 -------
    success_rate = (success / total) * 100
    print("\n========= 当前统计 =========")
    print(f"总轮数：{total}")
    print(f"成功：{success}")
    print(f"失败：{fail}")
    print(f"成功率：{success_rate:.2f}%")
    print("===========================\n")
