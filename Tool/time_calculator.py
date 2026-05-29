import tkinter as tk
from tkinter import messagebox
from datetime import datetime
import logging

# 1. 配置日志记录模块，替代所有 print 语句
# 设置日志级别为 INFO，并规范日志输出格式
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def calculate_time_diff_hours(start_str: str, end_str: str, fmt: str = "%Y/%m/%d %H:%M:%S") -> float:
    """
    计算两个时间字符串之间的时间差（单位：小时）。
    
    参数:
        start_str (str): 起始时间字符串
        end_str (str): 结束时间字符串
        fmt (str): 时间格式，默认为 "%Y/%m/%d %H:%M:%S"
        
    返回:
        float: 相差的小时数
    """
    try:
        # 将传入的字符串按照指定格式转换为 datetime 对象
        start_dt: datetime = datetime.strptime(start_str, fmt)
        end_dt: datetime = datetime.strptime(end_str, fmt)
        
        # 计算时间差，并提取为总秒数
        diff_seconds: float = (end_dt - start_dt).total_seconds()
        
        # 将秒数转换为小时数
        hours: float = diff_seconds / 3600.0
        
        # 记录成功的计算操作
        logging.info(f"成功计算时间差: {start_str} 至 {end_str}，差值: {hours:.4f} 小时")
        return hours
        
    except ValueError as e:
        # 捕获时间格式不匹配的错误
        logging.error(f"时间解析错误: {e}")
        raise ValueError("输入的时间格式不正确，请确保格式类似于：2026/4/8 16:45:28")
    except Exception as e:
        # 捕获并记录其他潜在的未知异常
        logging.error(f"发生未知错误: {e}")
        raise Exception("计算过程中发生未知错误，请检查日志。")

class TimeCalculatorApp:
    """
    基于 Tkinter 的时间差计算器桌面应用程序类。
    """
    def __init__(self, root: tk.Tk) -> None:
        self.root: tk.Tk = root
        self.root.title("时间差计算工具")
        
        # 设置窗口大小并禁止调整大小以保持布局美观
        self.root.geometry("450x300")
        self.root.resizable(False, False)
        
        # 初始化界面组件
        self._setup_ui()
        logging.info("应用程序界面初始化完成。")

    def _setup_ui(self) -> None:
        """
        初始化并布局所有的图形化组件。
        """
        # 起始时间输入区域
        tk.Label(self.root, text="起始时间 (例: 2026/4/8 16:45:28):", font=("Arial", 10)).pack(pady=(20, 5))
        self.start_entry: tk.Entry = tk.Entry(self.root, width=40, font=("Arial", 10))
        self.start_entry.pack()

        # 结束时间输入区域
        tk.Label(self.root, text="结束时间 (例: 2026/4/9 2:49:39):", font=("Arial", 10)).pack(pady=(15, 5))
        self.end_entry: tk.Entry = tk.Entry(self.root, width=40, font=("Arial", 10))
        self.end_entry.pack()

        # 计算触发按钮
        self.calc_button: tk.Button = tk.Button(
            self.root, 
            text="计算时间差 (小时)", 
            command=self._on_calculate,
            width=20,
            bg="lightgray"
        )
        self.calc_button.pack(pady=25)

        # 结果显示标签，初始状态显示提示信息
        self.result_label: tk.Label = tk.Label(self.root, text="等待输入计算...", font=("Arial", 12, "bold"), fg="blue")
        self.result_label.pack(pady=10)

    def _on_calculate(self) -> None:
        """
        处理点击“计算”按钮的事件，提取输入框文本并调用后台计算逻辑。
        """
        start_text: str = self.start_entry.get().strip()
        end_text: str = self.end_entry.get().strip()

        # 数据校验：检查输入是否为空
        if not start_text or not end_text:
            logging.warning("用户尝试在输入为空的情况下进行计算。")
            messagebox.showwarning("输入警告", "起始时间和结束时间不能为空，请填写完整。")
            return

        try:
            # 调用独立的时间计算函数
            hours: float = calculate_time_diff_hours(start_text, end_text)
            
            # 格式化并展示结果（保留四位小数）
            result_text: str = f"时间差: {hours:.4f} 小时"
            self.result_label.config(text=result_text, fg="green")
            
        except ValueError as ve:
            # 弹出警告窗口提示用户格式错误
            messagebox.showerror("格式错误", str(ve))
            self.result_label.config(text="计算失败：格式错误", fg="red")
        except Exception as e:
            # 弹出警告窗口提示系统或未知错误
            messagebox.showerror("系统错误", str(e))
            self.result_label.config(text="计算失败：系统错误", fg="red")

def main() -> None:
    """
    主程序入口点。负责创建根窗口并启动事件循环。
    """
    logging.info("正在启动时间差计算器应用程序...")
    try:
        root: tk.Tk = tk.Tk()
        # 实例化应用程序类
        app: TimeCalculatorApp = TimeCalculatorApp(root)
        # 运行主循环，等待用户交互
        root.mainloop()
    except Exception as e:
        logging.critical(f"应用程序启动失败，发生致命错误: {e}")

if __name__ == "__main__":
    main()