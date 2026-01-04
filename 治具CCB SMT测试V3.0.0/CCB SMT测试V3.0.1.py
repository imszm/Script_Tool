# -*- coding: utf-8 -*-
import time
import os
from datetime import datetime
from pywinauto.application import Application
from pywinauto.keyboard import send_keys

APP_TITLE = "CCB 测试 V3.0.00"

log_file = f"log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
success_cnt = 0
fail_cnt = 0


def log(msg):
    print(msg)
    with open(log_file, "a", encoding="utf-8") as f:
        f.write(msg + "\n")


def is_fail(win):
    """通过控件直接判断结果"""
    try:
        label = win.child_window(auto_id="Widget.labelResult", control_type="Text")
        txt = label.window_text().strip()
        return ("不通过" in txt) or ("fail" in txt.lower())
    except:
        return False


def test_main():
    global success_cnt, fail_cnt

    log("脚本启动，按 CTRL+Q 可退出\n")

    # 连接窗口
    app = Application(backend="uia").connect(title_re=APP_TITLE)
    win = app.window(title_re=APP_TITLE)

    # 获取控件
    btn_start = win.child_window(auto_id="Widget.pushButtonStart", control_type="Button")
    btn_restart = win.child_window(auto_id="Widget.pushButtonRestart", control_type="Button")

    log("控件绑定成功！开始循环...\n")

    while True:

        # 退出检测（用户按 Ctrl+Q 会触发 pywinauto 的异常）
        try:
            send_keys("", pause=0.01)
        except:
            log("检测到 CTRL+Q，脚本退出")
            break

        # 开始测试
        try:
            btn_start.click_input()
        except:
            log("点击开始失败")
            fail_cnt += 1
            continue

        time.sleep(1.0)

        # 重新测试
        try:
            btn_restart.click_input()
        except:
            log("点击重新测试失败")
            fail_cnt += 1
            continue

        time.sleep(1.0)

        # 判断结果（控件判断）
        if is_fail(win):
            fail_cnt += 1
            filename = f"FAIL_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
            win.capture_as_image().save(filename)
            log(f"[FAIL] 保存失败截图：{filename}\n")
        else:
            success_cnt += 1
            log("[PASS] 测试通过\n")

        # 输出成功率
        total = success_cnt + fail_cnt
        percent = success_cnt / total * 100
        log(f"当前成功率: {success_cnt}/{total} = {percent:.2f}%\n")

        time.sleep(0.5)

    # 总结
    total = success_cnt + fail_cnt
    percent = success_cnt / total * 100 if total else 0

    log("\n===== 测试结束 =====")
    log(f"PASS：{success_cnt}")
    log(f"FAIL：{fail_cnt}")
    log(f"成功率：{percent:.2f}%")
    log("日志：" + log_file)


if __name__ == "__main__":
    test_main()
