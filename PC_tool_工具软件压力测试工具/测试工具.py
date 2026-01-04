from pywinauto import Application

app = Application(backend="uia").connect(title_re="W3 PCTOOL.*")
dlg = app.window(title_re="W3 PCTOOL.*")
dlg.print_control_identifiers()
