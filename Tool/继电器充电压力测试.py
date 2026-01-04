# -*- coding: utf-8 -*-
import serial
import serial.tools.list_ports
import time
import datetime
import random
import sys
import logging
import threading
import os

# 尝试导入 win32api 用于弹窗提醒
try:
    import win32api
    import win32con

    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

# ================= 动态日志配置 =================
START_TIME_STR = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

LOG_DIR = "logs"
if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR)

# 1. 全量格式化日志（带分析结果）
LOG_FILE_PATH = os.path.join(LOG_DIR, f"test_{START_TIME_STR}_full.log")
# 2. 错误日志（仅记录严重报错）
ERR_FILE_PATH = os.path.join(LOG_DIR, f"test_{START_TIME_STR}_error.log")
# 3. 原始数据日志（原汁原味，给开发Debug用）
RAW_FILE_PATH = os.path.join(LOG_DIR, f"test_{START_TIME_STR}_raw.log")

CONFIG = {
    # 串口设置
    'RELAY_BAUDRATE': 9600,
    'DEVICE_BAUDRATE': 115200,
    'SERIAL_TIMEOUT': 1.0,

    # 端口识别关键字 (请根据实际情况调整)
    'RELAY_PORT_KEYWORD': "4",
    'DEVICE_PORT_KEYWORD': "cp210x",

    # 测试循环设置
    'TEST_CYCLES': 500000,
    'POWER_ON_MIN': 3.0,
    'POWER_ON_MAX': 5.0,
    'POWER_OFF_TIME': 5.0,
    'DELAY_AFTER_OFF': 20,  # 关机后等待日志的时间

    # 路径引用
    'LOG_FILENAME': LOG_FILE_PATH,
    'ERROR_LOG_FILENAME': ERR_FILE_PATH,
    'RAW_LOG_FILENAME': RAW_FILE_PATH
}

# ================= 关键字定义 =================
KEYWORDS = {
    'SUCCESS': ["voice_msgnum:9", "voice_msgnum:10"],
    'EXCEPTION': ["assertionfailedatfunction"],
    'INFO': ["voice_msgnum"]
}


# ================= 日志系统配置 =================
class LoggerSetup:
    @staticmethod
    def setup():
        logger = logging.getLogger("RelayTester")
        logger.setLevel(logging.INFO)
        logger.handlers = []

        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

        # 控制台输出
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)

        # 全量文件日志
        file_handler = logging.FileHandler(CONFIG['LOG_FILENAME'], encoding='utf-8')
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

        return logger

    @staticmethod
    def log_exception_to_file(msg):
        """记录严重错误到 error.log"""
        timestamp = datetime.datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
        with open(CONFIG['ERROR_LOG_FILENAME'], "a", encoding="utf-8") as f:
            f.write(f"{timestamp} {msg}\n")

    @staticmethod
    def log_raw_data(text_data):
        """记录原始数据到 raw.log"""
        timestamp = datetime.datetime.now().strftime("[%H:%M:%S.%f] ")
        try:
            with open(CONFIG['RAW_LOG_FILENAME'], "a", encoding="utf-8") as f:
                # 记录时间戳和原始数据
                f.write(f"{timestamp}--->\n{text_data}\n")
        except Exception as e:
            print(f"写入原始日志失败: {e}")


logger = LoggerSetup.setup()


# ================= 测试核心类 =================

class RelayTester:
    def __init__(self):
        self.relay_ser = None
        self.device_ser = None
        self.stats = {
            'success': 0,  # 检测到 voice_msg
            'exceptions': 0,  # 检测到代码断言失败
            'failures': 0,  # 超时/未检测到关键字
            'cycles': 0
        }

    def show_alert(self, msg):
        """显示弹窗提示"""
        logger.info(f"系统提示: {msg}")
        if HAS_WIN32:
            threading.Thread(target=lambda: win32api.MessageBox(
                0, msg, f"提示 {datetime.datetime.now().strftime('%H:%M:%S')}",
                win32con.MB_ICONINFORMATION | win32con.MB_SYSTEMMODAL
            )).start()

    def detect_ports(self):
        """自动扫描串口"""
        ports = list(serial.tools.list_ports.comports())
        relay_port = None
        device_port = None

        logger.info("正在扫描串口...")
        for p in ports:
            desc = p.description.lower()
            if CONFIG['RELAY_PORT_KEYWORD'].lower() in desc:
                relay_port = p.device
            elif CONFIG['DEVICE_PORT_KEYWORD'].lower() in desc:
                device_port = p.device
        return device_port, relay_port

    def open_serials(self):
        """打开串口连接"""
        dev, relay = self.detect_ports()
        if not dev or not relay:
            logger.error(f"串口识别失败! Device: {dev}, Relay: {relay}")
            return False
        try:
            self.relay_ser = serial.Serial(relay, CONFIG['RELAY_BAUDRATE'], timeout=CONFIG['SERIAL_TIMEOUT'])
            self.device_ser = serial.Serial(dev, CONFIG['DEVICE_BAUDRATE'], timeout=CONFIG['SERIAL_TIMEOUT'])
            self.relay_ser.reset_input_buffer()
            self.device_ser.reset_input_buffer()
            logger.info(f"串口连接成功: Device={dev}, Relay={relay}")
            return True
        except Exception as e:
            logger.error(f"串口打开异常: {e}")
            return False

    def close_serials(self):
        """关闭串口连接"""
        if self.relay_ser and self.relay_ser.is_open:
            try:
                # 退出时尝试断电
                self.relay_ser.write(bytes([0x50]))
            except:
                pass
            self.relay_ser.close()
        if self.device_ser and self.device_ser.is_open:
            self.device_ser.close()
        logger.info("串口已关闭")

    def init_relay_hardware(self):
        """
        初始化继电器逻辑：
        1. 发送 0x50 复位
        2. 发送 0x51 使能/握手，识别继电器类型
        3. 识别完成后，发送 0x50 关闭继电器，保持初始化状态
        """
        if not self.relay_ser or not self.relay_ser.is_open:
            logger.error("初始化失败：继电器串口未打开")
            return

        logger.info(">>> 开始执行继电器硬件初始化...")
        try:
            # 1. 发送 0x50 (复位信号)
            logger.info("STEP 1: 发送复位指令 (0x50)...")
            self.relay_ser.write(bytes([0x50]))
            time.sleep(1)
            # 读取缓存防止干扰
            if self.relay_ser.in_waiting:
                self.relay_ser.read(self.relay_ser.in_waiting)

            # 2. 发送 0x51 (使能/查询)
            logger.info("STEP 2: 发送使能/查询指令 (0x51)...")
            self.relay_ser.write(bytes([0x51]))
            time.sleep(1)

            # 3. 读取响应并判断类型
            if self.relay_ser.in_waiting:
                resp = self.relay_ser.read(self.relay_ser.in_waiting)
                resp_hex = resp.hex().lower()
                logger.info(f"继电器握手响应(Hex): {resp_hex}")

                if "ac" in resp_hex:
                    logger.info("=== 检测到硬件：8路继电器 ===")
                elif "ab" in resp_hex:
                    logger.info("=== 检测到硬件：4路继电器 ===")
                elif "ad" in resp_hex:
                    logger.info("=== 检测到硬件：2路继电器 ===")
                else:
                    logger.warning(f"=== 未知继电器类型，响应码：{resp_hex} ===")
            else:
                logger.warning("=== 警告：继电器未返回握手数据 ===")

            # 4. 【关键步骤】初始化完成后，立即关闭继电器
            logger.info("STEP 3: 初始化完成，强制关闭继电器以保持初始状态 (0x50)...")
            self.relay_ser.write(bytes([0x50]))
            time.sleep(2)  # 给硬件一点反应时间
            logger.info(">>> 继电器已就绪 (当前状态: OFF)")

        except Exception as e:
            logger.error(f"继电器初始化异常: {e}")

    def relay_control(self, state):
        """控制继电器开关"""
        if not self.relay_ser or not self.relay_ser.is_open: return
        try:
            # 0x4F: 开, 0x50: 关
            cmd = bytes([0x4F]) if state else bytes([0x50])
            self.relay_ser.write(cmd)
        except Exception as e:
            logger.error(f"继电器控制失败: {e}")

    def read_device_buffer(self):
        """读取数据，同时写入 Raw 日志"""
        if not self.device_ser or not self.device_ser.is_open: return []
        logs = []
        try:
            if self.device_ser.in_waiting > 0:
                raw = self.device_ser.read(self.device_ser.in_waiting)
            else:
                raw = self.device_ser.read_all()

            if raw:
                # 1. 尝试解码
                try:
                    text_decoded = raw.decode("utf-8", errors="ignore")
                except:
                    text_decoded = raw.decode("latin1", errors="ignore")

                # 2.  写入原始日志 (给开发看)
                LoggerSetup.log_raw_data(text_decoded)

                # 3. 处理成列表供脚本分析
                for line in text_decoded.split('\n'):
                    if line.strip():
                        logs.append(line.strip())
        except Exception as e:
            logger.error(f"读取设备日志出错: {e}")
            self.device_ser = None
        return logs

    def analyze_logs(self, log_lines):
        """分析日志关键字"""
        found_success = False
        found_exception = False

        for line in log_lines:
            # 简单预处理用于匹配
            processed_line = line.replace(" ", "").lower()

            # 检查异常 (Assertion Failed)
            for kw in KEYWORDS['EXCEPTION']:
                if kw in processed_line:
                    found_exception = True
                    msg = f"检测到异常报错: {line}"
                    logger.error(msg)
                    LoggerSetup.log_exception_to_file(msg)

            # 检查成功 (Voice Msg)
            for kw in KEYWORDS['SUCCESS']:
                if kw in processed_line:
                    found_success = True
                    logger.info(f"检测到成功关键字: {line}")

        return found_success, found_exception

    def run_cycle(self, cycle_num):
        """执行单次测试循环"""
        self.stats['cycles'] = cycle_num
        logger.info(f"{'=' * 20} 第 {cycle_num} 轮开始 {'=' * 20}")

        # 1. 开启充电
        logger.info("动作: 开启继电器 (ON)")
        self.relay_control(True)
        time.sleep(random.uniform(CONFIG['POWER_ON_MIN'], CONFIG['POWER_ON_MAX']))
        logs_stage_1 = self.read_device_buffer()

        # 2. 关闭充电
        logger.info("动作: 关闭继电器 (OFF)")
        self.relay_control(False)
        time.sleep(CONFIG['POWER_OFF_TIME'])

        # 3. 关机等待
        logger.info(f"等待 {CONFIG['DELAY_AFTER_OFF']} 秒 (捕获关机/休眠日志)...")
        time.sleep(CONFIG['DELAY_AFTER_OFF'])
        logs_stage_2 = self.read_device_buffer()

        # 4. 分析结果
        is_success, is_exception = self.analyze_logs(logs_stage_1 + logs_stage_2)

        # 5. 统计逻辑
        if is_exception:
            self.stats['exceptions'] += 1
            logger.error(f"第 {cycle_num} 轮结果: 🔴 严重异常 (代码报错)")
        elif is_success:
            self.stats['success'] += 1
            logger.info(f"第 {cycle_num} 轮结果: 🟢 成功")
        else:
            self.stats['failures'] += 1
            logger.warning(f"第 {cycle_num} 轮结果: 🟡 失败 (未检测到关键字)")

        logger.info(
            f"当前统计 -> 成功: {self.stats['success']} | 失败: {self.stats['failures']} | 异常: {self.stats['exceptions']}")

    def run(self):
        """主运行函数"""
        if not self.open_serials():
            self.show_alert("串口打开失败")
            return

        logger.info(f"日志目录: {os.path.abspath(LOG_DIR)}")

        # ==========================================
        #  执行初始化 (使能 -> 识别 -> 关断)
        # ==========================================
        self.init_relay_hardware()
        # ==========================================

        try:
            for i in range(1, CONFIG['TEST_CYCLES'] + 1):
                self.run_cycle(i)
        except KeyboardInterrupt:
            logger.warning("\n用户强制停止测试")
        except Exception as e:
            logger.critical(f"发生错误: {e}", exc_info=True)
        finally:
            self.close_serials()
            msg = (f"测试结束\n"
                   f"成功: {self.stats['success']}\n"
                   f"失败: {self.stats['failures']}\n"
                   f"异常: {self.stats['exceptions']}")
            logger.info(msg)
            self.show_alert(msg)


if __name__ == "__main__":
    RelayTester().run()