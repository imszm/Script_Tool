# src/logger.py
import logging
import os
from datetime import datetime

def setup_logging():
    """
    配置日志系统：同时输出到控制台和文件
    """
    # 1. 创建 logs 目录（如果不存在）
    if not os.path.exists('logs'):
        os.makedirs('logs')

    # 2. 定义日志格式
    # 格式：时间 - 模块名 - 等级 - 消息
    log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    
    # 3. 获取根 Logger
    logger = logging.getLogger("AutoTest")
    logger.setLevel(logging.INFO)  # 设置最低记录级别为 INFO

    # 防止重复添加 Handler（避免日志重复打印）
    if not logger.handlers:
        # File Handler: 写入文件
        file_name = f"logs/test_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        file_handler = logging.FileHandler(file_name, encoding='utf-8')
        file_handler.setFormatter(logging.Formatter(log_format))
        logger.addHandler(file_handler)

        # Stream Handler: 输出到屏幕
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(logging.Formatter(log_format))
        logger.addHandler(console_handler)

    return logger