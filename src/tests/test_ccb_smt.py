import time
import win32api, win32con
from PIL import ImageGrab
from pywinauto.application import Application
from src import config
from src.tests.base_test import BaseTest


class CcbSmtTest(BaseTest):
    """
    对应原脚本: CCB SMT测试V3.0.0...
    包含：继电器控制(ASCII) + 鼠标点击 + 屏幕像素识别
    """

    def fast_click(self, x, y):
        win32api.SetCursorPos((x, y))
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
        time.sleep(0.05)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)

    def check_pixel_status(self):
        """像素颜色检测逻辑"""
        try:
            screen = ImageGrab.grab()
            for (cx, cy) in config.PC_TOOL_CONFIG["CCB_CHECK_POINTS"]:
                rgb = screen.getpixel((cx, cy))
                r, g, b = rgb
                # 红色判定(FAIL)
                if r > 200 and g < 100 and b < 100: return "FAIL"
                # 绿色判定(PASS)
                if g > 140 and g > r + 30 and g > b + 30: return "PASS"
        except:
            pass
        return None

    def run(self, loops):
        relay = self.drivers.get('relay')  # 注意：这里实际上需要连接 COM12
        cfg = config.PC_TOOL_CONFIG

        # 连接软件
        try:
            app = Application(backend="uia").connect(title_re=cfg["CCB_TITLE_REGEX"])
            win = app.window(title_re=cfg["CCB_TITLE_REGEX"])
            win.set_focus()
        except:
            self.logger.error("无法连接CCB测试软件")
            return

        start_sn = 0

        for i in range(loops):
            if self.stop_flag: break

            # 生成SN
            sn = f"{cfg['CCB_SERIAL_PREFIX']}{str(start_sn + i).zfill(4)}"
            self.logger.info(f"测试 SN: {sn}")

            # 1. 继电器重启 (ASCII指令 P/O)
            if relay:
                relay.send_bytes(config.COMMANDS['ASCII_OFF'], "断电")
                time.sleep(2)
                relay.send_bytes(config.COMMANDS['ASCII_ON'], "上电")
                time.sleep(2)

            # 2. 输入SN
            try:
                edit = win.child_window(auto_id="Widget.lineEditSerialNumber", control_type="Edit")
                edit.set_edit_text(sn)
                win.type_keys('{ENTER}')
            except:
                self.logger.error("输入SN失败")
                continue

            time.sleep(5)

            # 3. 模拟点击
            c1 = cfg["CCB_COORDS"]["PASS_LIGHT"]
            c2 = cfg["CCB_COORDS"]["PASS_HORN"]
            self.fast_click(c1[0], c1[1])
            time.sleep(0.5)
            self.fast_click(c2[0], c2[1])

            # 4. 颜色检测 (超时80秒)
            self.logger.info("等待颜色结果...")
            res = "TIMEOUT"
            for _ in range(80):
                status = self.check_pixel_status()
                if status:
                    res = status
                    break
                time.sleep(1)

            self.logger.info(f"结果: {res}")
            if res == "FAIL":
                win.capture_as_image().save(f"logs/fail_{sn}.png")