# Script 工程化入口

你现在可以通过一个统一的 `main.py` 来运行：

- `src/` 下已经工程化的用例（更易维护、可复用）
- `Tool/` 下历史零散脚本（先用“兼容运行”方式接入，逐步再模块化）

## 使用方法

进入目录 `C:\Users\szm21\Downloads\Script` 后运行：

```bash
python main.py --help
```

示例：

```bash
# src 版本（工程化）
python main.py w3-power -n 1000
python main.py charging -n 1000
python main.py turn-signal -n 1000 --side left
python main.py horn -n 1000

# Tool 版本（兼容运行历史脚本）
python main.py relay-power-cycle
python main.py ota-gui
python main.py time-diff
```

## 依赖说明

不同子命令依赖不同库：

- 串口类脚本：`pyserial`
- Windows 弹窗/鼠标：`pywin32`
- GUI 自动化：`pywinauto`（部分脚本还用到 `pyautogui`、`pygetwindow`、`psutil`、`pillow`、`pytesseract` 等）
- Excel 用例生成：`pandas`、`openpyxl`

建议按你实际要跑的命令按需安装依赖。

