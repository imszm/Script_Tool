import pyautogui
import pygetwindow as gw
import time
import logging
from datetime import datetime

# 设置点击次数
click_count = 10000
current_click = 0  # 当前点击次数计数
valid_click_count = 0  # 有效点击次数计数

# 使用时间戳创建日志文件名
log_filename = f'click_log_{datetime.now().strftime("%Y%m%d_%H%M%S")}.txt'
logging.basicConfig(filename=log_filename, level=logging.INFO, format='%(asctime)s - %(message)s')

def end_logging(clicks):
    """结束日志并记录点击次数"""
    logging.info(f"脚本结束，总点击次数：{clicks}")
    print(f"脚本结束，总点击次数：{clicks}")

def check_window_exists(title):
    """检查目标窗口是否存在"""
    windows = gw.getWindowsWithTitle(title)
    return len(windows) > 0

try:
    # 查找目标窗口
    window_title = 'W3 升级工具 V4.0.3'
    windows = gw.getWindowsWithTitle(window_title)

    if not windows:
        print("未找到目标窗口，请确保窗口标题正确且窗口已打开。")
        logging.warning("未找到目标窗口。")
        end_logging(current_click)
    else:
        window = windows[0]  # 获取窗口列表中的第一个

        # 在开始之前确保切换到英文输入法
        pyautogui.hotkey('alt', 'shift')  # 切换到英文输入法，使用适合你系统的快捷键

        for i in range(click_count):
            try:
                window_x, window_y = window.topright

                click_actions = [
                    ((1857, 146), 75, "开始升级"),
                    ((1844, 463), 30, "摇杆校准")
                ]

                for pos, wait_time, description in click_actions:
                    if not check_window_exists(window_title):
                        logging.error("检测到目标窗口已关闭，可能是软件闪退。")
                        print("检测到目标窗口已关闭，脚本停止执行。")
                        raise Exception("目标窗口已关闭")

                    x = window_x + (pos[0] - window_x)
                    y = window_y + (pos[1] - window_y)

                    pyautogui.click(x, y)
                    current_click += 1
                    logging.info(f"第 {current_click} 次点击，位置 ({x}, {y})，操作：{description}")
                    print(f"第 {current_click} 次点击，位置 ({x}, {y})，操作：{description}")

                    if pos in [(1840, 215), (1839, 188), (1842, 357), (1836, 430), (847, 630), (1834, 506)]:
                        valid_click_count += 1
                        logging.info(f"有效点击次数更新为：{valid_click_count}")
                        print(f"有效点击次数更新为：{valid_click_count}")

                    # 输入SN号
                    if pos == (1558, 188):
                        time.sleep(1)
                        pyautogui.click(x, y)

                        # 清除文本框内容
                        pyautogui.hotkey('ctrl', 'a')  # 选择全部文本
                        pyautogui.press('backspace')    # 清除文本

                        # 输入大写的SN号
                        text = 'W30300RZPUS4800001'
                        pyautogui.typewrite(text, interval=0.1)

                    time.sleep(wait_time)

            except Exception as e:
                logging.error(f"发生错误: {e}")
                print(f"发生错误: {e}")
                break

except Exception as e:
    logging.error(f"检测到软件崩溃或意外错误: {e}")

end_logging(current_click)
