import time
import logging
import sys
from typing import Tuple, Dict, List, Optional
import pyautogui
import serial

# ================= 配置区域 =================
# 操作坐标配置
START_BUTTON_POS: Tuple[int, int] = (1802, 822)
OK_BUTTON_POS: Tuple[int, int] = (1518, 559)

# 串口开关操作坐标 (点击两次实现关闭再打开)
SERIAL_TOGGLE_POS: Tuple[int, int] = (1870, 88)

# 结果判定坐标配置 (必须全部为绿色才算通过)
CHECK_POSITIONS: List[Tuple[int, int]] = [
    (997, 116),
    (997, 233),
    (997, 342),
    (998, 407),
    (997, 493)
]

# 时间配置
MAX_WAIT_SECONDS: int = 100
UI_RESET_DELAY: float = 2.0      # 点击开始后，等待上位机清空上次绿色结果的缓冲时间
POST_RESULT_DELAY: float = 1.0   # 判定出结果后等待进入下一次循环的时间
SERIAL_TOGGLE_DELAY: float = 3.0 # 根据需求：关闭和打开串口后各等待 3 秒

# 继电器硬件配置
RELAY_PORT: str = "COM17"         # 继电器使用的串口号
RELAY_BAUDRATE: int = 9600        # 继电器波特率
RELAY_CMD_ON: int = 0x50          # 上电闭合指令
RELAY_CMD_OFF: int = 0x4F         # 断电断开指令
POWER_OFF_DURATION: float = 3.0   # 判定失败时断电维持的时长（秒）
REBOOT_WAIT_DELAY: float = 2.0    # 重新上电后，留给治具系统冷启动的缓冲时间（秒）

# PyAutoGUI安全设置：将鼠标移动到屏幕四个角之一即可强制中止脚本并释放鼠标
pyautogui.FAILSAFE = True

# ================= 日志配置 =================
def setup_logger() -> logging.Logger:
    """
    初始化日志配置，将日志输出到控制台和本地文件，防止重复打印。
    """
    logger = logging.getLogger("FCT_Stress_Test")
    logger.setLevel(logging.DEBUG)
    logger.propagate = False  # 阻断日志向上传播

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

# ================= 继电器控制逻辑 =================
def send_relay_command(port: str, baudrate: int, command: int, logger: logging.Logger) -> bool:
    """
    向指定的串口发送单字节继电器控制命令，包含完善的错误处理。
    """
    ser: Optional[serial.Serial] = None
    try:
        ser = serial.Serial(port, baudrate, timeout=1.0)
        cmd_bytes: bytes = bytes([command])
        ser.write(cmd_bytes)
        ser.flush()
        time.sleep(0.1)  # 给硬件链路预留微小的响应时间
        logger.info("继电器指令下发成功，串口: %s, 命令字节: %s", port, hex(command))
        return True
    except serial.SerialException as e:
        logger.error("继电器串口通信失败 (端口: %s), 错误信息: %s", port, e)
        return False
    except Exception as e:
        logger.error("操控继电器时发生未预料的底层错误: %s", e, exc_info=True)
        return False
    finally:
        if ser and ser.is_open:
            ser.close()

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

# ================= 辅助操作逻辑 =================
def toggle_software_serial(logger: logging.Logger) -> None:
    """
    执行软件层面的串口重启：先关闭串口（等待3秒），再打开串口（等待3秒）。
    """
    try:
        logger.info("开始执行步骤2: 重启上位机串口，目标坐标 %s", SERIAL_TOGGLE_POS)
        
        # 第一次点击：关闭串口
        pyautogui.click(x=SERIAL_TOGGLE_POS[0], y=SERIAL_TOGGLE_POS[1])
        logger.info("已点击关闭串口，开始等待 %d 秒...", int(SERIAL_TOGGLE_DELAY))
        time.sleep(SERIAL_TOGGLE_DELAY)
        
        # 第二次点击：打开串口
        pyautogui.click(x=SERIAL_TOGGLE_POS[0], y=SERIAL_TOGGLE_POS[1])
        logger.info("已点击打开串口，开始等待 %d 秒...", int(SERIAL_TOGGLE_DELAY))
        time.sleep(SERIAL_TOGGLE_DELAY)
        
    except pyautogui.FailSafeException:
        raise 
    except Exception as e:
        logger.error("重启上位机串口时发生错误: %s", e, exc_info=True)

# ================= 主控制流程 =================
def run_stress_test(total_cycles: int, logger: logging.Logger) -> None:
    """
    执行压力测试主循环。
    判定失败后执行：1.下电上电 -> 2.关闭/打开串口(各等3秒) -> 3.下一轮循环自动点击开始测试。
    """
    logger.info("开始执行FCT自动化测试，计划最大循环次数: %d 次", total_cycles)
    
    # 脚本初始化，确保继电器处于开启状态
    logger.info("初始化阶段: 正在使能并开启继电器以保持治具供电状态...")
    send_relay_command(RELAY_PORT, RELAY_BAUDRATE, RELAY_CMD_ON, logger)
    
    test_stats: Dict[str, int] = {
        "pass": 0,
        "fail": 0
    }
    
    for cycle in range(1, total_cycles + 1):
        try:
            logger.info("开始第 %d 次测试循环", cycle)
            
            # 步骤3（针对上一轮失败而言）/ 正常流程起点：点击开始测试
            logger.info("步骤3 / 起点: 点击 开始测试 按钮，坐标 %s", START_BUTTON_POS)
            pyautogui.click(x=START_BUTTON_POS[0], y=START_BUTTON_POS[1])
            
            # 引入缓冲：等待上位机响应点击并重置UI颜色
            logger.info("延迟缓冲: 等待 %.1f 秒，确保上位机已清空上一次的测试结果...", UI_RESET_DELAY)
            time.sleep(UI_RESET_DELAY)
            
            # 动态轮询结果
            logger.info("检测阶段: 正在动态监测测试状态，最大死等时长 %d 秒...", MAX_WAIT_SECONDS)
            
            test_result: str = "TIMEOUT"
            failed_positions: List[Tuple[int, int]] = []
            
            for sec in range(MAX_WAIT_SECONDS):
                # 轮询中持续点击OK按钮，确保常规提示弹窗被及时关闭
                pyautogui.click(x=OK_BUTTON_POS[0], y=OK_BUTTON_POS[1])
                
                time.sleep(0.5)
                
                all_green: bool = True
                has_red: bool = False
                current_not_green: List[Tuple[int, int]] = []
                
                for pos in CHECK_POSITIONS:
                    r, g, b = pyautogui.pixel(x=pos[0], y=pos[1])
                    
                    if is_color_green(r, g, b):
                        continue
                    elif is_color_red(r, g, b):
                        has_red = True
                        all_green = False
                        current_not_green.append(pos)
                    else:
                        all_green = False
                        current_not_green.append(pos)
                
                if all_green:
                    test_result = "PASS"
                    logger.info("通知: 用时约 %d 秒，检测到所有点位变绿，提前结束本次等待并锁定结果。", sec + 1)
                    break
                elif has_red:
                    test_result = "FAIL"
                    failed_positions = current_not_green
                    logger.warning("通知: 用时约 %d 秒，检测到红色异常点位 %s，提前结束本次等待并锁定结果。", sec + 1, current_not_green)
                    break
                
                time.sleep(0.5)
            
            # 收尾当前循环并计入统计
            need_recovery: bool = False
            
            if test_result == "PASS":
                logger.info("本次测试最终结果: [成功]")
                test_stats["pass"] += 1
                logger.info("需求触发：已判定为成功(PASS)，现在执行固定死等100秒...")
                time.sleep(100.0)
                logger.info("固定死等100秒结束，准备恢复后续常规流程。")
                
            elif test_result == "FAIL":
                logger.warning("本次测试最终结果: [失败] (检测到非绿色坐标: %s)", failed_positions)
                test_stats["fail"] += 1
                logger.warning("需求触发：已判定为失败(FAIL)，现在执行固定死等100秒...")
                time.sleep(100.0)
                logger.warning("固定死等100秒结束，准备进入失败恢复流程。")
                need_recovery = True
                
            elif test_result == "TIMEOUT":
                logger.error("本次测试最终结果: [失败] (达到死等上限 %d 秒，未能全绿)", MAX_WAIT_SECONDS)
                test_stats["fail"] += 1
                need_recovery = True
            
            log_statistics(stats=test_stats, logger=logger)
            logger.info("第 %d 次测试循环结束\n", cycle)
            
            # 判定失败后的处理步骤
            if need_recovery:
                logger.warning("执行判定失败后的异常恢复步骤...")
                
                # 1. 下电，上电
                logger.info("开始执行步骤1: 继电器断电，保持 %s 秒...", POWER_OFF_DURATION)
                send_relay_command(RELAY_PORT, RELAY_BAUDRATE, RELAY_CMD_OFF, logger)
                time.sleep(POWER_OFF_DURATION)
                
                logger.info("开始执行步骤1: 继电器重新通电...")
                send_relay_command(RELAY_PORT, RELAY_BAUDRATE, RELAY_CMD_ON, logger)
                logger.info("等待治具系统启动就绪，缓冲 %s 秒...", REBOOT_WAIT_DELAY)
                time.sleep(REBOOT_WAIT_DELAY)
                
                # 2. 关闭串口（等待3秒），打开串口（等待3秒）
                toggle_software_serial(logger=logger)
                
                # 3. 点击开始测试 (当下一次循环开始时，会自动执行循环开头的点击操作)
                logger.info("失败恢复步骤执行完毕，即将进入下一轮循环触发测试。")
            else:
                # 判定成功，按照设定的常规延迟进入下一轮循环
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
            
            logger.warning("未知异常触发治具安全重启恢复流程...")
            # 保持恢复步骤一致性：1.下电上电 -> 2.重启软件串口
            send_relay_command(RELAY_PORT, RELAY_BAUDRATE, RELAY_CMD_OFF, logger)
            time.sleep(POWER_OFF_DURATION)
            send_relay_command(RELAY_PORT, RELAY_BAUDRATE, RELAY_CMD_ON, logger)
            time.sleep(REBOOT_WAIT_DELAY)
            
            toggle_software_serial(logger=logger)
            
            logger.info("脚本将跳过当前错误，等待5秒后继续尝试下一次执行...")
            time.sleep(5)

if __name__ == "__main__":
    app_logger = setup_logger()
    TARGET_CYCLES: int = 9999 
    
    try:
        run_stress_test(total_cycles=TARGET_CYCLES, logger=app_logger)
    except KeyboardInterrupt:
        app_logger.info("运行前被外部信号中断，脚本未执行。")