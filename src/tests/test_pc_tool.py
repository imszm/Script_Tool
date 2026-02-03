# src/tests/test_pc_tool.py
import time
import os
from datetime import datetime
from pywinauto.application import Application
from src.tests.base_test import BaseTest


class UpgradeToolTest(BaseTest):
    """
    PC升级工具 UI 自动化测试
    """

    def setup(self):
        self.app_path = r"C:\Path\To\Your\Tool.exe"  # 建议放到config里
        self.app_title = "L5 PCTOOL V3.9.00"

        # 尝试连接或启动应用
        try:
            self.logger.info("正在连接升级工具...")
            self.app = Application(backend="uia").connect(title=self.app_title, timeout=10)
            self.win = self.app.window(title=self.app_title)
        except Exception as e:
            self.logger.error(f"无法连接程序，请先手动启动: {e}")
            raise e

    def run(self, loops):
        # 获取控件
        btn_start = self.win.child_window(auto_id="Widget.buttonUpgrade", control_type="Button")
        edit_log = self.win.child_window(auto_id="Widget.textEditLog", control_type="Edit")

        for i in range(1, loops + 1):
            self.logger.info(f"--- 升级测试第 {i} 轮 ---")

            # 记录旧日志长度
            old_log_content = edit_log.get_value()
            old_len = len(old_log_content)

            # 点击开始
            btn_start.click()
            self.logger.info("点击升级按钮，等待 170秒...")

            # 这里的等待可以优化，比如每秒检查一次日志状态
            time.sleep(170)

            # 获取新日志
            full_log = edit_log.get_value()
            new_log = full_log[old_len:]

            if "失败" in new_log:
                self.logger.error("检测到升级失败！")
                # 截图功能
                screenshot_name = f"logs/fail_{datetime.now().strftime('%H%M%S')}.png"
                self.win.capture_as_image().save(screenshot_name)
                self.logger.info(f"截图已保存: {screenshot_name}")
                break  # 失败停止
            else:
                self.logger.info("本轮升级成功")