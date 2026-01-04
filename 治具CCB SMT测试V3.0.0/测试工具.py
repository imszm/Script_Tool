from pywinauto.application import Application

app = Application(backend="uia").connect(title="L5 PCTOOL V3.8.00")
win = app.window(title="CCB 测试 L5 PCTOOL V3.8.00")

for c in win.descendants():
    print(c.window_text(), "|", c.element_info.control_type, "|", c.element_info.automation_id)
