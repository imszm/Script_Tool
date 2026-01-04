import pyautogui
import pygetwindow as gw
import time
import logging
from datetime import datetime
import psutil
import os

# 设置点击次数
click_count = 30000
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

def restart_software(program_path, window_title):
    """重新启动目标软件"""
    try:
        # 检查是否有已运行的目标软件进程并终止
        logging.info("检查是否有正在运行的目标软件...")
        for proc in psutil.process_iter(['pid', 'name']):
            if 'EW01_PCTOOL' in proc.info['name']:
                logging.info(f"发现进程: {proc.info['name']} (PID: {proc.info['pid']}), 尝试终止...")
                proc.terminate()  # 发送终止信号
                proc.wait(timeout=5)  # 等待进程结束
                logging.info(f"已成功终止进程: {proc.info['name']} (PID: {proc.info['pid']})")
        
        # 再次确认所有相关进程是否已终止
        for proc in psutil.process_iter(['name']):
            if 'EW01_PCTOOL' in proc.info['name']:
                raise Exception(f"进程未能终止: {proc.info['name']}")
        
        # 启动目标软件
        logging.info("尝试启动目标软件...")
        os.startfile(program_path)
        time.sleep(10)  # 等待窗口完全加载
        
        # 检查目标窗口是否存在
        if not check_window_exists(window_title):
            raise Exception("目标软件启动后未检测到窗口。")
        logging.info("目标软件启动成功。")
        print("目标软件启动成功。")
        return True

    except Exception as e:
        logging.error(f"重新启动软件失败: {e}")
        print(f"重新启动软件失败: {e}")
        return False

try:
    # 目标窗口和程序路径
    window_title = 'W3 PCTOOL V5.4.00'
    program_path = r"C:\Users\szm21\Downloads\Test_Tool\EW01_PCTOOL_V5.4\c\EW01_FACTORY_V4.3.0"  # 替换为程序的实际路径

    for i in range(click_count):
        try:
            # 检查目标窗口是否存在
            if not check_window_exists(window_title):
                logging.warning("检测到目标窗口已关闭，可能是软件闪退。")
                print("检测到目标窗口已关闭，尝试重新启动软件...")
                if not restart_software(program_path, window_title):
                    raise Exception("软件重新启动失败，脚本停止执行。")

            # 获取目标窗口
            windows = gw.getWindowsWithTitle(window_title)
            if not windows:
                raise Exception("未找到目标窗口。")
            window = windows[0]

            window_x, window_y = window.topleft

            # 点击流程
            click_actions = [
                ((1685, 324), 100, "开始升级"),
                ((1231, 363), 6, "SN序列号输入框点击"),
                ((1683, 370), 6, "SN写入"),
                ((1687, 410), 3, "开启日志"),
                ((1687, 410), 3, "关闭日志"),
                ((1683, 491), 20, "IMU校准"),
                ((1684, 551), 7, "恢复出厂设置"),
                ((1684, 581), 1, "打开参数设置"),
                ((862, 623), 6, "确认参数设置"),
                ((1686, 674), 30, "摇杆校准")
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

                if pos in [(1695, 324), (1683, 370), (1687, 410), (1683, 491), (1684, 551), (1684, 581), (864, 621), (1686, 674)]:
                    valid_click_count += 1
                    logging.info(f"有效点击次数更新为：{valid_click_count}")
                    print(f"有效点击次数更新为：{valid_click_count}")

                # 输入SN号
                if pos == (1231, 363):
                    time.sleep(1)
                    pyautogui.click(x, y)

                     #清除文本框内容
                    pyautogui.hotkey('ctrl', 'a')  # 选择全部文本
                    pyautogui.press('backspace')  # 清除文本

                    # 输入大写的SN号
                    text = 'W303010ZP004A00001'
                    pyautogui.typewrite(text, interval=0.1)

                time.sleep(wait_time)

        except Exception as e:
            logging.error(f"发生错误: {e}")
            print(f"发生错误: {e}")
            break

except Exception as e:
    logging.error(f"检测到软件崩溃或意外错误: {e}")

end_logging(current_click)
