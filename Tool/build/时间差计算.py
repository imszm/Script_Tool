import tkinter as tk
from tkinter import messagebox
from datetime import datetime

class TimeCalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("时间差计算工具 (支持复制版)")
        self.root.geometry("480x450")
        
        padding = {'padx': 15, 'pady': 5}
        
        # --- UI 布局 ---
        tk.Label(root, text="开始时间 (YYYY-MM-DD HH:MM:SS):").pack(**padding)
        self.start_entry = tk.Entry(root, width=35)
        self.start_entry.insert(0, "2025-12-24 20:33:50")
        self.start_entry.pack()
        
        tk.Label(root, text="结束时间 (YYYY-MM-DD HH:MM:SS):").pack(**padding)
        self.end_entry = tk.Entry(root, width=35)
        self.end_entry.insert(0, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        self.end_entry.pack()
        
        # 按钮容器
        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=15)

        self.calc_btn = tk.Button(btn_frame, text="开始计算", command=self.calculate, bg="#4CAF50", fg="white", width=12)
        self.calc_btn.pack(side=tk.LEFT, padx=5)

        self.copy_btn = tk.Button(btn_frame, text="复制结果", command=self.copy_to_clipboard, bg="#2196F3", fg="white", width=12)
        self.copy_btn.pack(side=tk.LEFT, padx=5)
        
        # 结果显示区域（改用 Text 组件，支持选中和复制）
        tk.Label(root, text="计算结果:").pack()
        self.result_text = tk.Text(root, height=10, width=55, font=("Consolas", 10), bg="#f4f4f4")
        self.result_text.pack(padx=15, pady=5)
        
        # 初始化状态
        self.result_text.insert(tk.END, "等待计算...")
        self.result_text.config(state=tk.DISABLED) # 初始状态设为只读

    def calculate(self):
        start_str = self.start_entry.get().strip()
        end_str = self.end_entry.get().strip()
        
        try:
            start_dt = datetime.strptime(start_str, "%Y-%m-%d %H:%M:%S")
            end_dt = datetime.strptime(end_str, "%Y-%m-%d %H:%M:%S")
            duration = end_dt - start_dt
            
            total_seconds = duration.total_seconds()
            hours, remainder = divmod(int(total_seconds), 3600)
            minutes, seconds = divmod(remainder, 60)
            
            res_text = (
                f"【时间差详情】\n"
                f"--------------------------\n"
                f"标准时长: {hours}小时 {minutes}分钟 {seconds}秒\n"
                f"Delta格式: {duration}\n"
                f"总计秒数: {total_seconds:,} s\n"
                f"总计分钟: {total_seconds / 60:.2f} min\n"
                f"总计小时: {total_seconds / 3600:.4f} h"
            )
            
            # 更新文本框内容
            self.result_text.config(state=tk.NORMAL) # 切换为可写状态
            self.result_text.delete(1.0, tk.END)     # 清空内容
            self.result_text.insert(tk.END, res_text) # 写入结果
            self.result_text.config(state=tk.NORMAL)  # 保持开启，以便手动选择复制
            
        except ValueError:
            messagebox.showerror("格式错误", "请输入正确的格式：\nYYYY-MM-DD HH:MM:SS")

    def copy_to_clipboard(self):
        """一键复制功能"""
        content = self.result_text.get(1.0, tk.END).strip()
        if content and content != "等待计算...":
            self.root.clipboard_clear()
            self.root.clipboard_append(content)
            messagebox.showinfo("成功", "结果已成功复制到剪贴板！")
        else:
            messagebox.showwarning("提示", "没有可复制的内容，请先点击计算。")

if __name__ == "__main__":
    root = tk.Tk()
    app = TimeCalculator(root)
    root.mainloop()
