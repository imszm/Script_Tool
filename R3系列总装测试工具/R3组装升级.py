import pyautogui
import time
import logging
import sys
from typing import Tuple, Optional

# 配置logging日志格式
# 使用stdout输出，包含时间、日志级别和具体信息
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler(sys.stdout)]
)

def wait_for_ui_element(
    image_path: str,
    timeout: float = 182.0,
    confidence: float = 0.8
) -> bool:
    """
    等待直到屏幕上出现指定的图像。
    
    :param image_path: 目标图像的路径（建议截取图二中红色的“请扫描左电机条形码”文字作为特征图）
    :param timeout: 最大等待时间（秒）
    :param confidence: 匹配置信度
    :return: 如果在超时前找到图像返回 True，超时未找到返回 False
    """
    logging.info(f"开始等待弹窗图像: {image_path}，最大等待时间: {timeout} 秒")
    start_time: float = time.time()
    
    while (time.time() - start_time) < timeout:
        try:
            # 只需判断是否在屏幕上存在，无需获取中心坐标
            location = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if location is not None:
                logging.info(f"成功检测到目标弹窗，耗时: {time.time() - start_time:.2f} 秒。")
                return True
        except pyautogui.ImageNotFoundException:
            # 找不到图像属于正常轮询现象，直接跳过并等待下一次检测
            pass
        except Exception as e:
            # 记录其他非预期异常
            logging.debug(f"等待图像匹配时发生非致命错误: {e}")
            
        time.sleep(0.5)
        
    logging.error(f"等待超时：在 {timeout} 秒内未能检测到弹窗图像 '{image_path}'。")
    return False

def click_ui_element_by_image(
    image_path: str, 
    timeout: float = 10.0, 
    confidence: float = 0.8
) -> bool:
    """
    通过图像识别动态定位并点击屏幕上的静态UI元素（如固定的按钮）。
    
    :param image_path: 目标UI元素的截图路径
    :param timeout: 寻找图像的超时时间（秒）
    :param confidence: 图像匹配的置信度（需安装 opencv-python 库）
    :return: 成功找到并点击返回 True，超时未找到返回 False
    """
    logging.info(f"正在屏幕中动态寻找目标图像: {image_path}")
    start_time: float = time.time()
    
    while (time.time() - start_time) < timeout:
        try:
            coords = pyautogui.locateCenterOnScreen(image_path, confidence=confidence)
            if coords is not None:
                logging.info(f"成功定位目标图像，动态坐标为: (X:{coords.x}, Y:{coords.y})")
                pyautogui.click(coords.x, coords.y)
                return True
        except pyautogui.ImageNotFoundException:
            pass
        except Exception as e:
            logging.debug(f"图像匹配时发生非致命错误: {e}")
            
        time.sleep(0.5)
        
    logging.error(f"定位失败：在 {timeout} 秒内未能找到特征图 '{image_path}'。")
    return False

def click_relative_to_anchor(
    anchor_image_path: str, 
    x_offset: int, 
    y_offset: int, 
    timeout: float = 5.0, 
    confidence: float = 0.8
) -> bool:
    """
    锚点偏移点击法：寻找屏幕上的静态锚点（如固定的文本标签），
    并在其坐标基础上增加偏移量后进行点击。适用于内容动态变化的输入框。
    
    :param anchor_image_path: 静态锚点（如"VIN S/N"标签）的截图路径
    :param x_offset: 目标点击位置相对于锚点中心的X轴偏移量（正数向右，负数向左）
    :param y_offset: 目标点击位置相对于锚点中心的Y轴偏移量（正数向下，负数向上）
    :param timeout: 寻找图像的超时时间
    :param confidence: 图像匹配的置信度
    :return: 成功找到锚点并点击偏移位置返回 True，否则返回 False
    """
    logging.info(f"正在寻找静态锚点图像: {anchor_image_path}")
    start_time: float = time.time()
    
    while (time.time() - start_time) < timeout:
        try:
            # 找到静态标签的中心点
            coords = pyautogui.locateCenterOnScreen(anchor_image_path, confidence=confidence)
            if coords is not None:
                # 计算实际需要点击的输入框坐标
                target_x: int = int(coords.x + x_offset)
                target_y: int = int(coords.y + y_offset)
                logging.info(f"找到锚点 (X:{coords.x}, Y:{coords.y})，计算偏移后点击目标坐标: (X:{target_x}, Y:{target_y})")
                pyautogui.click(target_x, target_y)
                return True
        except pyautogui.ImageNotFoundException:
            pass
        except Exception as e:
            logging.debug(f"锚点匹配时发生非致命错误: {e}")
        time.sleep(0.5)
        
    logging.error(f"锚点定位失败：在 {timeout} 秒内未能找到特征图 '{anchor_image_path}'。")
    return False

def run_embedded_assembly_test(
    base_vin: str = "R30100GZK0066",
    start_suffix: int = 1,
    max_suffix: int = 99999,
    motor_left_sn: str = "PX01-8120Y063000002",
    motor_right_sn: str = "PX01-8120Y063000002",
    battery_sn: str = "AV06002LH201F3140001",
    vin_anchor_image: str = "vin_label.png",
    popup_window_image: str = "popup_window.png",  # 新增：用于识别扫码弹窗的截图
    confirm_btn_image: str = "confirm_btn.png"
) -> None:
    """
    执行整机自动化组装升级测试的循环操作。
    全程使用图像识别与锚点偏移，摆脱对绝对坐标的依赖。
    """
    logging.info("自动化测试脚本已启动。安全提示：若需紧急停止，请将鼠标移动到屏幕角落。")
    
    execution_count: int = 0

    try:
        for current_suffix in range(start_suffix, max_suffix + 1):
            current_vin: str = f"{base_vin}{current_suffix:05d}"
            logging.info(f"========== 开始执行测试循环，当前 VIN SN: {current_vin} ==========")

            # 1. 锚点偏移点击 VIN SN 输入框
            logging.info("寻找VIN标签并点击右侧输入框获取焦点...")
            is_vin_clicked: bool = click_relative_to_anchor(
                anchor_image_path=vin_anchor_image,
                x_offset=200, 
                y_offset=0
            )
            
            if not is_vin_clicked:
                logging.error("无法点击VIN输入框，终止当前循环。")
                continue
                
            time.sleep(1.0)

            # 2. 清除输入框中原有的序列号
            logging.info("执行全选并删除，确保输入框清空...")
            pyautogui.hotkey('ctrl', 'a') 
            time.sleep(0.3)
            pyautogui.press('delete')     
            time.sleep(0.3)               

            # 3. 输入新的 VIN 序列号并回车
            logging.info(f"写入新的 VIN SN: {current_vin}")
            pyautogui.typewrite(current_vin, interval=0.02)
            pyautogui.press('enter')

            # 4. 动态等待扫码弹窗出现（替代原有的死等11秒）
            logging.info("开始检测扫码弹窗状态...")
            is_popup_ready: bool = wait_for_ui_element(
                image_path=popup_window_image,
                timeout=181.0,  # 迟迟不出现弹窗，则死等181秒
                confidence=0.8
            )
            
            if is_popup_ready:
                logging.info("弹窗已就位，1秒后开始执行后续脚本动作...")
                time.sleep(1.0)
            else:
                logging.error("等待 181 秒后未能检测到弹窗，停止当前 VIN 流程，进入下一次循环。")
                continue

            # 5. 输入左电机SN并回车
            logging.info(f"输入左电机序列号: {motor_left_sn}")
            pyautogui.typewrite(motor_left_sn, interval=0.02)
            pyautogui.press('enter')
            time.sleep(0.5)

            # 6. 输入右电机SN并回车
            logging.info(f"输入右电机序列号: {motor_right_sn}")
            pyautogui.typewrite(motor_right_sn, interval=0.02)
            pyautogui.press('enter')
            time.sleep(0.5)

            # 7. 输入电池SN并回车
            logging.info(f"输入电池序列号: {battery_sn}")
            pyautogui.typewrite(battery_sn, interval=0.02)
            pyautogui.press('enter')
            time.sleep(0.5)

            # 8. 直接图像识别点击确认按钮
            logging.info("准备点击确认按钮...")
            is_confirm_clicked: bool = click_ui_element_by_image(
                image_path=confirm_btn_image, 
                timeout=5.0, 
                confidence=0.8
            )
            
            if not is_confirm_clicked:
                logging.error("未能点击到确认按钮，跳过本次循环下的等待，直接进入下一次迭代。")
                continue

            # 9. 等待20秒上位机组装升级完成
            logging.info("等待 241 秒，等待上位机组装及固件升级完成...")
            time.sleep(241.0)

            execution_count += 1
            logging.info(f"========== VIN SN: {current_vin} 测试循环执行完毕。当前已完成总次数: {execution_count} ==========\n")

    except KeyboardInterrupt:
        logging.info("检测到手动中断，脚本已安全停止。")
    except pyautogui.FailSafeException:
        logging.error("触发防故障机制：鼠标移动至屏幕边缘。脚本已强制退出。")
    except Exception as e:
        logging.error(f"脚本运行过程中发生意外错误: {e}", exc_info=True)
    finally:
        logging.info(f"自动化测试任务结束，总计成功执行次数: {execution_count}")

if __name__ == "__main__":
    logging.info("脚本将在 1 秒后开始执行，请尽快将焦点切换至目标上位机软件窗口...")
    time.sleep(1.0)
    run_embedded_assembly_test()