import time
import logging
import sys
from typing import Tuple, Dict, List
import pyautogui

# ================= 配置区域 =================
# 测试参数配置
SERIAL_NUMBER: str = "2021005002R201GD006400001"

# 操作坐标配置
APP_FOCUS_POS: Tuple[int, int] = (1448, 796)  # 上位机软件窗口坐标，用于获取输入焦点
CONFIRM_BUTTON_POS: Tuple[int, int] = (1733, 949)
OK_BUTTON_POS: Tuple[int, int] = (1570, 323)

# 结果判定坐标配置 (必须全部为绿色才算通过)
CHECK_POSITIONS: List[Tuple[int, int]] = [
    (997, 116),
    (997, 258),
    (997, 322),
    (997, 390),
    (997, 497),
    (997, 579),
    (1846, 948)  # 右下角“通过”文字及边框区域
]

# 时间配置
MAX_WAIT_SECONDS: int = 92
UI_RESET_DELAY: float = 2.0  # 点击确认后，等待上位机清空上次绿色结果的缓冲时间
POST_RESULT_DELAY: float = 2.0  # 判定出结果后等待进入下一次循环的时间
OK_BUTTON_DELAY_SECONDS: int = 10  # 新增：点击OK按钮前的等待时间（秒）

# PyAutoGUI安全设置：将鼠标移动到屏幕四个角之一即可强制中止脚本并释放鼠标
pyautogui.FAILSAFE = True


# ================= 日志配置 =================
def setup_logger() -> logging.Logger:
    """
    初始化日志配置，将日志输出到控制台和本地文件，防止重复打印。
    """
    logger = logging.getLogger("FCT_Stress_Test")
    logger.setLevel(logging.DEBUG)
    logger.propagate = False  # 阻断日志向上传播，防止在某些终端环境下重复打印

    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

    # 文件处理器
    file_handler = logging.FileHandler("fct_stress_test.log", encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)

    # 控制台处理器
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)

    if not logger.handlers:
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)

    return logger


# ================= 核心判断逻辑 =================
def is_color_green(r: int, g: int, b: int) -> bool:
    """
    判断给定RGB值是否符合绿色特征。
    """
    return g > r + 30 and g > b + 30


def is_color_red(r: int, g: int, b: int) -> bool:
    """
    判断给定RGB值是否符合红色特征。
    """
    return r > g + 30 and r > b + 30


def scan_region_colors(pos: Tuple[int, int], drift_range: int, logger: logging.Logger) -> Tuple[bool, bool]:
    """
    在指定坐标的周边区域内搜寻是否存在绿色或红色。

    Args:
        pos: 待检测的中心坐标 (x, y)
        drift_range: 漂移搜索半径（像素值）
        logger: 日志记录器

    Returns:
        Tuple[bool, bool]: (是否存在绿色, 是否存在红色)
    """
    try:
        # 定义截图区域，只截取目标点周围的一个小正方形
        bbox = (pos[0] - drift_range, pos[1] - drift_range, drift_range * 2, drift_range * 2)
        screenshot = pyautogui.screenshot(region=bbox)

        has_green: bool = False
        has_red: bool = False

        # 遍历该区域内的像素点
        for x in range(screenshot.width):
            for y in range(screenshot.height):
                r, g, b = screenshot.getpixel((x, y))
                if is_color_green(r, g, b):
                    has_green = True
                elif is_color_red(r, g, b):
                    has_red = True

                if has_green and has_red:
                    break
            if has_green and has_red:
                break

        return has_green, has_red

    except Exception as e:
        logger.warning("区域像素扫描发生异常，尝试降级为单像素点检测。坐标: %s, 异常信息: %s", pos, str(e))
        try:
            r, g, b = pyautogui.pixel(pos[0], pos[1])
            return is_color_green(r, g, b), is_color_red(r, g, b)
        except Exception as inner_e:
            logger.error("单像素点检测同样失败，坐标: %s, 错误: %s", pos, str(inner_e))
            return False, False


def log_statistics(stats: Dict[str, int], logger: logging.Logger) -> None:
    """
    计算并打印当前压力测试的成功与失败报告。
    """
    total_runs: int = stats["pass"] + stats["fail"]

    if total_runs > 0:
        pass_rate: float = (stats["pass"] / total_runs) * 100
    else:
        pass_rate = 0.0

    logger.info("统计报告 运行总次数: %d 次", total_runs)
    logger.info("统计报告 成功 (PASS): %d 次", stats["pass"])
    logger.info("统计报告 失败 (FAIL): %d 次", stats["fail"])
    logger.info("统计报告 实时成功率: %.2f%%", pass_rate)


# ================= 主控制流程 =================
def run_stress_test(total_cycles: int, logger: logging.Logger) -> None:
    """
    执行压力测试主循环，支持快捷键随时中断并释放鼠标。
    """
    logger.info("开始执行FCT自动化测试，计划最大循环次数: %d 次", total_cycles)

    test_stats: Dict[str, int] = {
        "pass": 0,
        "fail": 0
    }

    for cycle in range(1, total_cycles + 1):
        try:
            logger.info("开始第 %d 次测试循环", cycle)

            # 步骤0：点击上位机激活窗口
            logger.info("步骤0: 点击上位机激活窗口，坐标 %s", APP_FOCUS_POS)
            pyautogui.click(x=APP_FOCUS_POS[0], y=APP_FOCUS_POS[1])
            time.sleep(0.5)

            # 第一步：模拟键盘输入序列号
            logger.info("步骤1: 键盘输入序列号 [%s]", SERIAL_NUMBER)
            pyautogui.typewrite(SERIAL_NUMBER, interval=0.02)

            # 第二步：模拟按下回车键
            logger.info("步骤2: 模拟按下回车键 (Enter)")
            pyautogui.press('enter')
            time.sleep(0.5)

            # 第三步：点击确认按钮
            logger.info("步骤3: 点击确认按钮，坐标 %s", CONFIRM_BUTTON_POS)
            pyautogui.click(x=CONFIRM_BUTTON_POS[0], y=CONFIRM_BUTTON_POS[1])

            logger.info("延迟缓冲: 等待 %.1f 秒，确保上位机已清空上一次的测试结果...", UI_RESET_DELAY)
            time.sleep(UI_RESET_DELAY)

            # 第四步：动态轮询结果
            logger.info("步骤4: 正在动态监测测试状态，最大死等时长 %d 秒...", MAX_WAIT_SECONDS)

            test_result: str = "TIMEOUT"
            failed_positions: List[Tuple[int, int]] = []

            for sec in range(MAX_WAIT_SECONDS):
                # 新增逻辑：延迟设定的时间（默认10秒）后，再执行OK按钮的持续点击
                if sec >= OK_BUTTON_DELAY_SECONDS:
                    try:
                        pyautogui.click(x=OK_BUTTON_POS[0], y=OK_BUTTON_POS[1])
                    except Exception as e:
                        logger.debug("点击OK按钮时发生异常: %s", str(e))

                # 预留半秒给UI重新渲染或响应点击
                time.sleep(0.5)

                all_green: bool = True
                has_red_detected: bool = False
                current_not_green: List[Tuple[int, int]] = []

                # 遍历校验所有判定点位
                for pos in CHECK_POSITIONS:
                    is_green, is_red = scan_region_colors(pos=pos, drift_range=25, logger=logger)

                    if is_green:
                        continue
                    elif is_red:
                        has_red_detected = True
                        all_green = False
                        current_not_green.append(pos)
                    else:
                        all_green = False
                        current_not_green.append(pos)

                # 动态条件触发判定
                if all_green:
                    test_result = "PASS"
                    logger.info("通知: 用时约 %d 秒，检测到所有点位变绿，提前结束本次等待。", sec + 1)
                    break
                elif has_red_detected:
                    test_result = "FAIL"
                    failed_positions = current_not_green
                    logger.warning("通知: 用时约 %d 秒，检测到红色异常点位 %s，提前结束本次等待。", sec + 1,
                                   current_not_green)
                    break

                # 若无明确结果，补足剩下的半秒，进入下一秒的轮询
                time.sleep(0.5)

            # 第五步：收尾当前循环并计入统计
            if test_result == "PASS":
                logger.info("本次测试最终结果: [成功]")
                test_stats["pass"] += 1
            elif test_result == "FAIL":
                logger.warning("本次测试最终结果: [失败] (检测到非绿色坐标: %s)", failed_positions)
                test_stats["fail"] += 1
            elif test_result == "TIMEOUT":
                logger.error("本次测试最终结果: [失败] (达到死等上限 %d 秒，未能全绿，异常坐标: %s)",
                             MAX_WAIT_SECONDS, current_not_green)
                test_stats["fail"] += 1

            # 输出当前统计结果
            log_statistics(stats=test_stats, logger=logger)
            logger.info("第 %d 次测试循环结束\n", cycle)

            time.sleep(POST_RESULT_DELAY)

        except pyautogui.FailSafeException:
            logger.error("触发PyAutoGUI安全机制（鼠标移至屏幕角落），脚本已强制终止并释放鼠标。")
            logger.info("正在输出退出前的最终统计数据：")
            log_statistics(stats=test_stats, logger=logger)
            break

        except KeyboardInterrupt:
            logger.warning("接收到用户通过终端发送的 CTRL+C 中断信号，脚本立即终止并释放鼠标。")
            logger.info("正在输出退出前的最终统计数据：")
            log_statistics(stats=test_stats, logger=logger)
            break

        except Exception as e:
            logger.error("第 %d 次循环中发生未知异常: %s", cycle, e, exc_info=True)
            test_stats["fail"] += 1
            log_statistics(stats=test_stats, logger=logger)
            logger.info("脚本将跳过当前错误，等待5秒后继续尝试下一次执行...")
            time.sleep(5)


if __name__ == "__main__":
    app_logger = setup_logger()
    TARGET_CYCLES: int = 9999

    try:
        run_stress_test(total_cycles=TARGET_CYCLES, logger=app_logger)
    except KeyboardInterrupt:
        app_logger.info("运行前被外部信号中断，脚本未执行。")