import time
import os
from pywinauto.application import Application
from src import config
from src.tests.base_test import BaseTest


class PcUpgradeTest(BaseTest):
    """
    对应原脚本: 升级工具压力自动化测试.py
    """

    def run(self, loops):
        cfg = config.PC_TOOL_CONFIG
        title = cfg["UPGRADE_APP_TITLE"]

        try:
            self.logger.info(f"连接应用: {title}")
            app = Application(backend="uia").connect(title=title, timeout=10)
            win = app.window(title=title)
        except Exception as e:
            self.logger.error(f"连接失败，请先启动软件: {e}")
            return

        btn = win.child_window(auto_id=cfg["UPGRADE_BTN_ID"], control_type="Button")
        log_box = win.child_window(auto_id=cfg["UPGRADE_LOG_ID"], control_type="Edit")

        if not os.path.exists("logs/screenshots"):
            os.makedirs("logs/screenshots")

        for i in range(1, loops + 1):
            if self.stop_flag: break
            self.logger.info(f"--- 升级 Cycle {i} ---")

            # 记录旧日志长度
            old_len = len(log_box.get_value())

            # 点击升级
            btn.click()
            self.logger.info(f"等待 {cfg['UPGRADE_WAIT_TIME']} 秒...")
            time.sleep(cfg['UPGRADE_WAIT_TIME'])

            # 检查新日志
            full_log = log_box.get_value()
            new_log = full_log[old_len:]

            if "失败" in new_log or "超时" in new_log:
                self.logger.error("升级失败！截图保存中...")
                win.capture_as_image().save(f"logs/screenshots/fail_{i}.png")
                break
            else:
                self.logger.info("本轮成功")