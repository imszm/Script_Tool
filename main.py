# main.py
import time
import logging
from src import config
from src.logger import setup_logging
from src.driver import RelayController

def main():
    # 1. 启动日志系统
    logger = setup_logging()
    logger.info(">>> 自动化测试任务启动 <<<")

    # 2. 实例化驱动类
    relay = RelayController(config.SERIAL_PORT, config.BAUDRATE)
    
    # 统计变量
    success_count = 0
    fail_count = 0

    # 3. 连接设备
    if not relay.connect():
        logger.critical("程序终止：无法连接硬件")
        return

    try:
        # 初始化复位
        logger.info("初始化继电器状态...")
        relay.send_cmd(config.RELAY_COMMANDS["K2_OFF"], "关闭K2")
        relay.send_cmd(config.RELAY_COMMANDS["K3_OFF"], "关闭K3")
        time.sleep(0.5)

        current_side = "LEFT" # 初始方向

        # 4. 开始循环
        for i in range(1, config.TEST_COUNT + 1):
            logger.info(f"--- 开始执行第 {i}/{config.TEST_COUNT} 次测试 (当前侧: {current_side}) ---")

            # 步骤 A: 模拟按压
            logger.info("动作: 按压按钮")
            time.sleep(config.PRESS_TIME)

            # 步骤 B: 根据方向亮灯
            if current_side == "LEFT":
                relay.send_cmd(config.RELAY_COMMANDS["K2_ON"], "左灯开")
                relay.send_cmd(config.RELAY_COMMANDS["K3_OFF"], "右灯关")
            else:
                relay.send_cmd(config.RELAY_COMMANDS["K3_ON"], "右灯开")
                relay.send_cmd(config.RELAY_COMMANDS["K2_OFF"], "左灯关")
            
            # 保持亮灯
            time.sleep(config.LIGHT_ON_TIME)

            # 步骤 C: 松开回弹（全灭）
            logger.info("动作: 松开按钮")
            time.sleep(config.RELEASE_TIME)
            
            relay.send_cmd(config.RELAY_COMMANDS["K2_OFF"], "左灯灭")
            relay.send_cmd(config.RELAY_COMMANDS["K3_OFF"], "右灯灭")
            
            # 步骤 D: 等待切换
            time.sleep(config.RELEASE_TIME + config.INTERVAL_BETWEEN_SWITCH)

            # 逻辑切换
            current_side = "RIGHT" if current_side == "LEFT" else "LEFT"
            success_count += 1
    
    except KeyboardInterrupt:
        logger.warning("用户强制停止测试")
    except Exception as e:
        logger.error(f"测试过程发生未预期的错误: {e}")
        fail_count += 1
    finally:
        # 5. 收尾工作
        logger.info("正在清理资源...")
        relay.send_cmd(config.RELAY_COMMANDS["K2_OFF"], "最终复位K2")
        relay.send_cmd(config.RELAY_COMMANDS["K3_OFF"], "最终复位K3")
        relay.close()

        # 6. 输出报告
        total = success_count + fail_count
        rate = (success_count / total * 100) if total > 0 else 0
        logger.info("=" * 30)
        logger.info(f"测试结束报告")
        logger.info(f"总次数: {total}")
        logger.info(f"成功: {success_count}")
        logger.info(f"失败: {fail_count}")
        logger.info(f"成功率: {rate:.2f}%")
        logger.info("=" * 30)

if __name__ == "__main__":
    main()