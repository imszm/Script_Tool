import tkinter as tk
from tkinter import ttk
import math
import logging
from typing import Any

# 1. 配置日志记录模块
# 记录 GUI 状态、计算过程以及潜在的异常数据
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class SpeedCalculatorApp:
    def __init__(self, root: tk.Tk) -> None:
        """
        初始化 GUI 应用程序并构建界面布局。
        增加变量追踪 (Trace) 以实现实时自动计算。
        """
        self.root = root
        self.root.title("固件车速工程计算器")
        # 稍微增加一点高度以容纳两行结果文本
        self.root.geometry("400x280")
        self.root.resizable(False, False)
        
        # 配置主框架
        self.main_frame = ttk.Frame(self.root, padding="20 20 20 20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # 2. 构建输入数据行及变量监听
        self.rpm_var: tk.StringVar = tk.StringVar(value="3728")
        self._create_input_row("电机转速 (RPM):", self.rpm_var, 0)
        # 绑定 write 事件，当输入框内容发生任何改变时，触发自动计算
        self.rpm_var.trace_add("write", self._perform_calculation)

        self.ratio_var: tk.StringVar = tk.StringVar(value="20.7")
        self._create_input_row("减速箱减速比:", self.ratio_var, 1)
        self.ratio_var.trace_add("write", self._perform_calculation)

        self.diameter_var: tk.StringVar = tk.StringVar(value="10")
        self._create_input_row("轮胎外径 (英寸):", self.diameter_var, 2)
        self.diameter_var.trace_add("write", self._perform_calculation)

        # 3. 构建结果显示标签 (已移除手动计算按钮)
        self.result_label = ttk.Label(
            self.main_frame, 
            text="计算结果: \n-- km/h\n-- mp/h", 
            font=("Arial", 14, "bold"),
            justify=tk.CENTER
        )
        self.result_label.grid(row=3, column=0, columnspan=2, pady=30)

        logging.info("GUI 界面加载完毕，已开启自动计算监听。")
        
        # 初始化启动时先执行一次计算，展示默认参数的结果
        self._perform_calculation()

    def _create_input_row(self, label_text: str, text_var: tk.StringVar, row_index: int) -> None:
        """
        辅助方法：在网格布局中创建带标签的输入行。
        """
        label = ttk.Label(self.main_frame, text=label_text)
        label.grid(row=row_index, column=0, sticky=tk.W, pady=10)
        
        entry = ttk.Entry(self.main_frame, textvariable=text_var, width=18)
        entry.grid(row=row_index, column=1, sticky=tk.E, pady=10)

    def _perform_calculation(self, *args: Any) -> None:
        """
        核心方法：获取输入值，处理异常，执行计算逻辑并更新 UI。
        注意：因用于 trace_add 回调，需接收 *args 参数。
        处理异常时采用更新 Label 而非弹窗，避免打断用户实时输入。
        """
        try:
            # 数据获取与去除前后空格
            rpm_str: str = self.rpm_var.get().strip()
            ratio_str: str = self.ratio_var.get().strip()
            diameter_str: str = self.diameter_var.get().strip()

            # 若用户正在删除输入框内容（即变为空），则静默重置结果，不记作错误
            if not rpm_str or not ratio_str or not diameter_str:
                self.result_label.config(text="等待输入完整参数...")
                return

            # 类型转换
            motor_rpm: float = float(rpm_str)
            gear_ratio: float = float(ratio_str)
            diameter_inch: float = float(diameter_str)

            # 物理与业务逻辑合法性校验
            if gear_ratio <= 0 or diameter_inch <= 0:
                self.result_label.config(text="参数有误: 减速比和外径需大于 0")
                return

            # 核心公式计算逻辑
            # 步骤 a: 计算实际轮毂转速
            wheel_rpm: float = motor_rpm / gear_ratio
            # 步骤 b: 计算轮胎周长 (英寸换算为米)
            circumference_m: float = math.pi * diameter_inch * 0.0254
            
            # 步骤 c: 计算出公制速度 (km/h)
            speed_kmh: float = (wheel_rpm * 60 * circumference_m) / 1000.0
            
            # 步骤 d: 将公制速度换算为英制速度 (mp/h)，1 英里 ≈ 1.609344 公里
            speed_mph: float = speed_kmh / 1.609344

            # 更新 GUI 显示结果 (保留两位小数)
            result_text: str = f"计算结果: \n{speed_kmh:.2f} km/h\n{speed_mph:.2f} mp/h"
            self.result_label.config(text=result_text)
            
            # 记录成功日志
            logging.info(f"自动计算成功 | 输入: RPM={motor_rpm}, Ratio={gear_ratio}, Dia={diameter_inch} -> 输出: {speed_kmh:.2f} km/h, {speed_mph:.2f} mp/h")

        except ValueError:
            # 捕获输入格式错误 (如用户输入字母或多个小数点)
            # 实时输入过程中这是常见情况，更新界面提示即可，避免干扰用户
            self.result_label.config(text="输入格式不合法...")
            
        except Exception as e:
            # 捕获其他不可预知的底层异常并记录日志
            logging.error(f"发生未知的核心计算异常: {e}")
            self.result_label.config(text="系统计算异常，请查看日志")

def main() -> None:
    """
    主程序入口：初始化 Tkinter 实例并启动主事件循环。
    """
    try:
        root = tk.Tk()
        app = SpeedCalculatorApp(root)
        root.mainloop()
    except Exception as e:
        logging.critical(f"GUI 进程启动失败，程序终止: {e}")

if __name__ == "__main__":
    main()
    