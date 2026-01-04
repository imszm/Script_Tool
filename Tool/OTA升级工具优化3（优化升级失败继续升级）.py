import time
import subprocess
import serial
import serial.tools.list_ports
import psutil
import os
import threading
import logging
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
from datetime import datetime
import pyautogui
import pygetwindow as gw

class OTAUpgradeGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("OTA循环升级自动化工具")
        self.root.geometry("900x800")
        
        # 初始化自动化类
        self.automation = OTAUpgradeAutomation(self)
        
        # 创建界面
        self.create_widgets()
        
        # 运行状态
        self.is_running = False
        self.current_cycle = 0
        
    def create_widgets(self):
        """创建GUI组件"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="OTA循环升级自动化工具", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # 配置区域
        config_frame = ttk.LabelFrame(main_frame, text="配置参数", padding="10")
        config_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        config_frame.columnconfigure(1, weight=1)
        
        # 升级工具路径
        ttk.Label(config_frame, text="升级工具路径:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.upgrade_tool_entry = ttk.Entry(config_frame, width=50)
        self.upgrade_tool_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))
        self.upgrade_tool_entry.insert(0, "升级工具.exe")
        ttk.Button(config_frame, text="浏览", command=self.browse_upgrade_tool).grid(row=0, column=2, padx=(5, 0))
        
        # 串口设置
        ttk.Label(config_frame, text="串口:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.serial_port_combo = ttk.Combobox(config_frame, width=20)
        self.serial_port_combo.grid(row=1, column=1, sticky=tk.W, pady=2, padx=(5, 0))
        
        ttk.Label(config_frame, text="波特率:").grid(row=1, column=1, sticky=tk.E, pady=2)
        self.baud_rate_combo = ttk.Combobox(config_frame, width=10, values=["9600", "115200", "57600", "38400"])
        self.baud_rate_combo.grid(row=1, column=2, sticky=tk.W, pady=2, padx=(5, 0))
        self.baud_rate_combo.set("115200")
        
        # 刷新串口按钮
        ttk.Button(config_frame, text="刷新串口", command=self.refresh_serial_ports).grid(row=1, column=2, sticky=tk.E, padx=(5, 0))
        
        # 密码和命令
        ttk.Label(config_frame, text="验证密码:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.password_entry = ttk.Entry(config_frame, width=20, show="*")
        self.password_entry.grid(row=2, column=1, sticky=tk.W, pady=2, padx=(5, 0))
        self.password_entry.insert(0, "ppx1220")
        
        ttk.Label(config_frame, text="OTA命令:").grid(row=2, column=1, sticky=tk.E, pady=2)
        self.ota_command_entry = ttk.Entry(config_frame, width=20)
        self.ota_command_entry.grid(row=2, column=2, sticky=tk.W, pady=2, padx=(5, 0))
        self.ota_command_entry.insert(0, "ota_begin 0 23")
        
        # 升级工具成功关键字设置
        ttk.Label(config_frame, text="中控升级成功关键字:").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.central_keyword_entry = ttk.Entry(config_frame, width=20)
        self.central_keyword_entry.grid(row=3, column=1, sticky=tk.W, pady=2, padx=(5, 0))
        self.central_keyword_entry.insert(0, "中控升级成功")
        
        ttk.Label(config_frame, text="电控升级成功关键字:").grid(row=4, column=0, sticky=tk.W, pady=2)
        self.electric_keyword_entry = ttk.Entry(config_frame, width=20)
        self.electric_keyword_entry.grid(row=4, column=1, sticky=tk.W, pady=2, padx=(5, 0))
        self.electric_keyword_entry.insert(0, "电控升级成功")
        
        ttk.Label(config_frame, text="BLE升级成功关键字:").grid(row=5, column=0, sticky=tk.W, pady=2)
        self.ble_keyword_entry = ttk.Entry(config_frame, width=20)
        self.ble_keyword_entry.grid(row=5, column=1, sticky=tk.W, pady=2, padx=(5, 0))
        self.ble_keyword_entry.insert(0, "BLE升级成功")
        
        # OTA串口升级成功关键字设置
        ttk.Label(config_frame, text="串口中控成功关键字:").grid(row=3, column=2, sticky=tk.W, pady=2)
        self.serial_central_entry = ttk.Entry(config_frame, width=15)
        self.serial_central_entry.grid(row=3, column=2, sticky=tk.E, pady=2, padx=(5, 0))
        self.serial_central_entry.insert(0, "new version: V0452R307C01L0")
        
        ttk.Label(config_frame, text="串口电控成功关键字:").grid(row=4, column=2, sticky=tk.W, pady=2)
        self.serial_electric_entry = ttk.Entry(config_frame, width=15)
        self.serial_electric_entry.grid(row=4, column=2, sticky=tk.E, pady=2, padx=(5, 0))
        self.serial_electric_entry.insert(0, "read version: V1330R617C01L0")
        
        ttk.Label(config_frame, text="串口BLE成功关键字:").grid(row=5, column=2, sticky=tk.W, pady=2)
        self.serial_ble_entry = ttk.Entry(config_frame, width=15)
        self.serial_ble_entry.grid(row=5, column=2, sticky=tk.E, pady=2, padx=(5, 0))
        self.serial_ble_entry.insert(0, "read version: V3632R206C01")
        
        # 循环设置
        ttk.Label(config_frame, text="升级循环次数:").grid(row=6, column=0, sticky=tk.W, pady=2)
        self.cycles_entry = ttk.Entry(config_frame, width=10)
        self.cycles_entry.grid(row=6, column=1, sticky=tk.W, pady=2, padx=(5, 0))
        self.cycles_entry.insert(0, "1")
        
        ttk.Label(config_frame, text="OTA超时时间(秒):").grid(row=6, column=1, sticky=tk.E, pady=2)
        self.timeout_entry = ttk.Entry(config_frame, width=10)
        self.timeout_entry.grid(row=6, column=2, sticky=tk.W, pady=2, padx=(5, 0))
        self.timeout_entry.insert(0, "1800")
        
        # 日志设置
        ttk.Label(config_frame, text="日志保存路径:").grid(row=7, column=0, sticky=tk.W, pady=2)
        self.log_path_entry = ttk.Entry(config_frame, width=50)
        self.log_path_entry.grid(row=7, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))
        self.log_path_entry.insert(0, "ota_upgrade_logs")
        ttk.Button(config_frame, text="浏览", command=self.browse_log_path).grid(row=7, column=2, padx=(5, 0))
        
        # 控制按钮区域
        control_frame = ttk.Frame(main_frame)
        control_frame.grid(row=2, column=0, columnspan=3, pady=10)
        
        self.start_button = ttk.Button(control_frame, text="开始升级", command=self.start_upgrade)
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        self.stop_button = ttk.Button(control_frame, text="停止升级", command=self.stop_upgrade, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=5)
        
        self.single_cycle_button = ttk.Button(control_frame, text="单次循环", command=self.single_cycle)
        self.single_cycle_button.pack(side=tk.LEFT, padx=5)
        
        # 状态显示区域
        status_frame = ttk.LabelFrame(main_frame, text="状态信息", padding="10")
        status_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        main_frame.rowconfigure(3, weight=1)
        status_frame.columnconfigure(0, weight=1)
        status_frame.rowconfigure(0, weight=1)
        
        self.status_text = scrolledtext.ScrolledText(status_frame, width=80, height=20, state=tk.DISABLED)
        self.status_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 进度条
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        # 当前状态显示
        self.current_status = ttk.Label(main_frame, text="就绪", relief=tk.SUNKEN, anchor=tk.W)
        self.current_status.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        # 初始化时刷新串口
        self.refresh_serial_ports()
        
    def browse_upgrade_tool(self):
        """浏览选择升级工具"""
        from tkinter import filedialog
        filename = filedialog.askopenfilename(title="选择升级工具", filetypes=[("Executable files", "*.exe")])
        if filename:
            self.upgrade_tool_entry.delete(0, tk.END)
            self.upgrade_tool_entry.insert(0, filename)
    
    def browse_log_path(self):
        """浏览选择日志保存路径"""
        from tkinter import filedialog
        folder = filedialog.askdirectory(title="选择日志保存目录")
        if folder:
            self.log_path_entry.delete(0, tk.END)
            self.log_path_entry.insert(0, folder)
    
    def refresh_serial_ports(self):
        """刷新串口列表"""
        ports = serial.tools.list_ports.comports()
        port_list = [port.device for port in ports]
        self.serial_port_combo['values'] = port_list
        if port_list:
            self.serial_port_combo.set(port_list[0])
    
    def log_message(self, message, level="INFO"):
        """在日志区域显示消息"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        formatted_message = f"[{timestamp}] [{level}] {message}\n"
        
        self.status_text.config(state=tk.NORMAL)
        self.status_text.insert(tk.END, formatted_message)
        self.status_text.see(tk.END)
        self.status_text.config(state=tk.DISABLED)
        
        # 更新状态栏
        self.current_status.config(text=message)
        
        # 强制更新界面
        self.root.update()
    
    def update_progress(self, start=True):
        """更新进度条"""
        if start:
            self.progress.start()
        else:
            self.progress.stop()
    
    def start_upgrade(self):
        """开始连续升级"""
        if not self.validate_inputs():
            return
        
        self.is_running = True
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.single_cycle_button.config(state=tk.DISABLED)
        
        # 更新配置
        self.update_config()
        
        # 在新线程中运行升级
        thread = threading.Thread(target=self.run_continuous_upgrade)
        thread.daemon = True
        thread.start()
    
    def stop_upgrade(self):
        """停止升级"""
        self.is_running = False
        self.automation.stop_requested = True
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.single_cycle_button.config(state=tk.NORMAL)
        self.update_progress(False)
        self.log_message("升级已停止")
    
    def single_cycle(self):
        """执行单次升级循环"""
        if not self.validate_inputs():
            return
        
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.DISABLED)
        self.single_cycle_button.config(state=tk.DISABLED)
        
        # 更新配置
        self.update_config()
        
        # 在新线程中运行单次升级
        thread = threading.Thread(target=self.run_single_upgrade)
        thread.daemon = True
        thread.start()
    
    def validate_inputs(self):
        """验证输入参数"""
        if not self.upgrade_tool_entry.get().strip():
            messagebox.showerror("错误", "请选择升级工具路径")
            return False
        
        if not self.serial_port_combo.get().strip():
            messagebox.showerror("错误", "请选择串口")
            return False
        
        # 验证升级工具成功关键字
        if not self.central_keyword_entry.get().strip():
            messagebox.showerror("错误", "请输入中控升级成功关键字")
            return False
        
        if not self.electric_keyword_entry.get().strip():
            messagebox.showerror("错误", "请输入电控升级成功关键字")
            return False
        
        if not self.ble_keyword_entry.get().strip():
            messagebox.showerror("错误", "请输入BLE升级成功关键字")
            return False
        
        # 验证串口升级成功关键字
        if not self.serial_central_entry.get().strip():
            messagebox.showerror("错误", "请输入串口中控成功关键字")
            return False
        
        if not self.serial_electric_entry.get().strip():
            messagebox.showerror("错误", "请输入串口电控成功关键字")
            return False
        
        if not self.serial_ble_entry.get().strip():
            messagebox.showerror("错误", "请输入串口BLE成功关键字")
            return False
        
        try:
            int(self.cycles_entry.get())
        except ValueError:
            messagebox.showerror("错误", "循环次数必须是数字")
            return False
        
        try:
            int(self.timeout_entry.get())
        except ValueError:
            messagebox.showerror("错误", "超时时间必须是数字")
            return False
        
        return True
    
    def update_config(self):
        """更新配置参数"""
        self.automation.upgrade_tool_path = self.upgrade_tool_entry.get().strip()
        self.automation.serial_port = self.serial_port_combo.get().strip()
        self.automation.baud_rate = int(self.baud_rate_combo.get())
        self.automation.upgrade_password = self.password_entry.get().strip()
        self.automation.ota_command = self.ota_command_entry.get().strip()
        self.automation.upgrade_timeout = int(self.timeout_entry.get())
        self.automation.log_directory = self.log_path_entry.get().strip()
        
        # 更新升级工具成功关键字
        self.automation.central_success_keyword = self.central_keyword_entry.get().strip()
        self.automation.electric_success_keyword = self.electric_keyword_entry.get().strip()
        self.automation.ble_success_keyword = self.ble_keyword_entry.get().strip()
        
        # 更新串口升级成功关键字
        self.automation.serial_central_success = self.serial_central_entry.get().strip()
        self.automation.serial_electric_success = self.serial_electric_entry.get().strip()
        self.automation.serial_ble_success = self.serial_ble_entry.get().strip()
    
    def run_continuous_upgrade(self):
        """运行连续升级"""
        self.update_progress(True)
        max_cycles = int(self.cycles_entry.get())
        self.automation.run_continuous_cycles(max_cycles)
        self.upgrade_completed()
    
    def run_single_upgrade(self):
        """运行单次升级"""
        self.update_progress(True)
        success = self.automation.run_single_cycle(1)
        if success:
            self.log_message("单次升级完成")
        else:
            self.log_message("单次升级失败", "ERROR")
        self.upgrade_completed()
    
    def upgrade_completed(self):
        """升级完成后的清理工作"""
        self.update_progress(False)
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.single_cycle_button.config(state=tk.NORMAL)
        self.is_running = False

class OTAUpgradeAutomation:
    def __init__(self, gui):
        self.gui = gui
        self.upgrade_tool_path = "升级工具.exe"
        self.serial_port = None
        self.baud_rate = 115200
        self.upgrade_password = "ppx1220"
        self.ota_command = "ota_begin 0 23"
        self.upgrade_timeout = 1800
        self.stop_requested = False
        self.log_directory = "ota_upgrade_logs"
        
        # 升级工具成功关键字
        self.central_success_keyword = "中控升级成功"
        self.electric_success_keyword = "电控升级成功"
        self.ble_success_keyword = "BLE升级成功"
        
        # 串口升级成功关键字
        self.serial_central_success = "new version: V0452R307C01L0"
        self.serial_electric_success = "read version: V1330R617C01L0"
        self.serial_ble_success = "read version: V3632R207C01"
        
        # 串口提示符检测关键字
        self.msh_prompt = "msh >"
        self.password_prompt = "password"
        
        # 固定坐标
        self.start_button_x = 1276
        self.start_button_y = 317
        
        # 串口对象
        self.serial_conn = None
        
        # 配置pyautogui
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 1.0
        
        # 创建日志目录
        os.makedirs(self.log_directory, exist_ok=True)
    
    def log(self, message, level="INFO"):
        """通过GUI记录日志"""
        self.gui.log_message(message, level)
    
    def save_serial_log(self, data, direction="RX"):
        """保存串口通信日志到文件"""
        try:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
            log_entry = f"[{timestamp}] [{direction}] {data}\n"
            
            log_filename = f"serial_log_{datetime.now().strftime('%Y%m%d')}.txt"
            log_path = os.path.join(self.log_directory, log_filename)
            
            with open(log_path, 'a', encoding='utf-8') as f:
                f.write(log_entry)
                
        except Exception as e:
            self.log(f"保存串口日志失败: {e}", "ERROR")
    
    def kill_process_by_name(self, process_name):
        """根据进程名结束进程"""
        try:
            killed_count = 0
            for proc in psutil.process_iter(['pid', 'name']):
                if process_name.lower() in proc.info['name'].lower():
                    try:
                        proc.kill()
                        self.log(f"已结束进程: {proc.info['name']} (PID: {proc.info['pid']})")
                        killed_count += 1
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        pass
            
            if killed_count > 0:
                time.sleep(3)  # 等待进程完全结束
                self.log(f"成功结束 {killed_count} 个进程")
            else:
                self.log(f"未找到进程: {process_name}")
                
        except Exception as e:
            self.log(f"结束进程时出错: {e}", "ERROR")
    
    def find_window_by_title(self, title_keyword):
        """根据标题关键词查找窗口"""
        try:
            windows = gw.getWindowsWithTitle(title_keyword)
            if windows:
                return windows[0]
            return None
        except Exception as e:
            self.log(f"查找窗口时出错: {e}", "ERROR")
            return None
    
    def click_start_upgrade(self):
        """点击开始升级按钮"""
        try:
            # 查找升级工具窗口 - 使用新的窗口标题
            upgrade_window = self.find_window_by_title("L5 升级工具")
            if not upgrade_window:
                # 如果找不到完整标题，尝试部分匹配
                upgrade_window = self.find_window_by_title("升级工具")
            
            if upgrade_window:
                upgrade_window.activate()
                time.sleep(2)
                
                # 使用固定坐标点击开始升级按钮
                self.log(f"点击固定坐标: ({self.start_button_x}, {self.start_button_y})")
                pyautogui.click(self.start_button_x, self.start_button_y)
                self.log("已点击开始升级按钮")
                return True
            else:
                self.log("未找到升级工具窗口", "ERROR")
                return False
            
        except Exception as e:
            self.log(f"点击开始升级按钮时出错: {e}", "ERROR")
            return False
    
    def check_upgrade_status(self):
        """检查升级状态 - 分别检测三个模块的升级状态"""
        try:
            self.log("检查升级状态...")
            
            # 需要检测的3个关键词
            required_keywords = [
                self.central_success_keyword,
                self.electric_success_keyword, 
                self.ble_success_keyword
            ]
            
            found_keywords = []
            
            # 激活升级工具窗口 - 使用新的窗口标题
            upgrade_window = self.find_window_by_title("L5 升级工具")
            if not upgrade_window:
                upgrade_window = self.find_window_by_title("升级工具")
                
            if upgrade_window:
                upgrade_window.activate()
                time.sleep(2)
                
                # 截取屏幕区域进行文本识别
                window_rect = upgrade_window.box
                if window_rect:
                    # 截取窗口区域
                    screenshot = pyautogui.screenshot(region=(
                        window_rect.left, 
                        window_rect.top, 
                        window_rect.width, 
                        window_rect.height
                    ))
                    
                    # 保存截图用于调试
                    debug_image_path = os.path.join(self.log_directory, f"debug_upgrade_status_{int(time.time())}.png")
                    screenshot.save(debug_image_path)
                    self.log(f"升级状态截图已保存: {debug_image_path}")
                    
                    # 使用OCR检查
                    try:
                        import pytesseract
                        # 设置tesseract路径（如果需要）
                        # pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
                        
                        # 使用中文识别
                        text = pytesseract.image_to_string(screenshot, lang='chi_sim')
                        self.log(f"OCR识别结果: {text}")
                        
                        # 分别检查每个关键词
                        for keyword in required_keywords:
                            if keyword in text:
                                found_keywords.append(keyword)
                                self.log(f"✓ 找到关键词: {keyword}")
                            else:
                                self.log(f"✗ 未找到关键词: {keyword}")
                        
                    except ImportError:
                        self.log("未安装pytesseract，无法进行OCR检测", "WARNING")
                        # 如果没有OCR，模拟找到所有关键词（仅用于测试）
                        found_keywords = required_keywords.copy()
                        self.log("模拟找到所有关键词（测试模式）")
                    except Exception as e:
                        self.log(f"OCR检测失败: {e}", "ERROR")
            
            # 检查是否找到所有关键词
            if set(found_keywords) == set(required_keywords):
                self.log("所有模块升级状态检测成功")
                return True
            else:
                missing_keywords = set(required_keywords) - set(found_keywords)
                self.log(f"缺少关键词: {missing_keywords}", "WARNING")
                return False
            
        except Exception as e:
            self.log(f"检查升级状态时出错: {e}", "ERROR")
            return False
    
    def initialize_serial(self):
        """初始化串口连接"""
        try:
            self.log(f"初始化串口: {self.serial_port}, 波特率: {self.baud_rate}")
            self.serial_conn = serial.Serial(
                port=self.serial_port,
                baudrate=self.baud_rate,
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                timeout=1,  # 读取超时1秒
                write_timeout=1  # 写入超时1秒
            )
            
            if self.serial_conn.is_open:
                self.log("串口连接成功")
                # 清空输入缓冲区
                self.serial_conn.reset_input_buffer()
                return True
            else:
                self.log("串口连接失败", "ERROR")
                return False
                
        except Exception as e:
            self.log(f"串口初始化失败: {e}", "ERROR")
            return False
    
    def close_serial(self):
        """关闭串口连接"""
        try:
            if self.serial_conn and self.serial_conn.is_open:
                self.serial_conn.close()
                self.log("串口连接已关闭")
        except Exception as e:
            self.log(f"关闭串口时出错: {e}", "ERROR")
    
    def send_serial_command(self, command):
        """发送串口命令（自动添加回车）"""
        try:
            if not self.serial_conn or not self.serial_conn.is_open:
                self.log("串口未连接", "ERROR")
                return False
            
            # 添加回车符
            if not command.endswith('\r\n'):
                command += '\r\n'
            
            self.log(f"发送命令: {command.strip()}")
            self.serial_conn.write(command.encode('utf-8'))
            self.save_serial_log(command.strip(), "TX")
            return True
            
        except Exception as e:
            self.log(f"发送串口命令失败: {e}", "ERROR")
            return False
    
    def read_serial_response_with_timeout(self, timeout=10, check_prompts=False, check_success=False):
        """读取串口响应，支持检测特定提示符和成功关键字"""
        try:
            if not self.serial_conn or not self.serial_conn.is_open:
                return None, None
            
            start_time = time.time()
            response = ""
            
            while time.time() - start_time < timeout and not self.stop_requested:
                if self.serial_conn.in_waiting > 0:
                    data = self.serial_conn.readline().decode('utf-8', errors='ignore').strip()
                    if data:
                        response += data + "\n"
                        self.save_serial_log(data, "RX")
                        self.log(f"串口响应: {data}")
                        
                        # 检查是否包含成功关键词
                        if check_success:
                            if self.serial_central_success in data:
                                self.log(f"检测到中控升级成功: {self.serial_central_success}")
                                return response, "central_success"
                            elif self.serial_electric_success in data:
                                self.log(f"检测到电控升级成功: {self.serial_electric_success}")
                                return response, "electric_success"
                            elif self.serial_ble_success in data:
                                self.log(f"检测到BLE升级成功: {self.serial_ble_success}")
                                return response, "ble_success"
                        
                        # 检查是否包含成功关键词（旧版兼容）
                        if self.serial_central_success in data:
                            self.log(f"检测到升级成功关键词: {self.serial_central_success}")
                            return response, "success"
                        
                        # 如果启用了提示符检测
                        if check_prompts:
                            if self.msh_prompt in data.lower():
                                self.log(f"检测到MSH提示符: {self.msh_prompt}")
                                return response, "msh"
                            elif self.password_prompt in data.lower():
                                self.log(f"检测到密码提示符: {self.password_prompt}")
                                return response, "password"
                
                time.sleep(0.1)
            
            return response if response else None, None
            
        except Exception as e:
            self.log(f"读取串口响应失败: {e}", "ERROR")
            return None, None
    
    def wait_for_prompt(self, timeout=60):
        """等待特定的串口提示符"""
        self.log(f"等待串口提示符，超时时间: {timeout}秒")
        
        start_time = time.time()
        while time.time() - start_time < timeout and not self.stop_requested:
            response, prompt_type = self.read_serial_response_with_timeout(5, check_prompts=True)
            
            if prompt_type == "msh":
                self.log("检测到MSH提示符，无需输入密码")
                return "msh"
            elif prompt_type == "password":
                self.log("检测到密码提示符，需要输入密码")
                return "password"
            elif response:
                # 有响应但没有检测到特定提示符，继续等待
                pass
            else:
                # 没有响应，继续等待
                pass
            
            elapsed_time = int(time.time() - start_time)
            if elapsed_time % 10 == 0:
                self.log(f"等待提示符中... 已等待 {elapsed_time} 秒")
        
        self.log("等待提示符超时", "WARNING")
        return None
    
    def check_serial_upgrade_success(self, timeout=300):
        """检查串口OTA升级是否成功 - 分别检测三个模块"""
        self.log("开始检查串口OTA升级状态...")
        
        start_time = time.time()
        success_modules = {
            "central": False,
            "electric": False, 
            "ble": False
        }
        
        required_modules = list(success_modules.keys())
        
        while time.time() - start_time < timeout and not self.stop_requested:
            # 读取串口响应并检查成功关键字
            response, success_type = self.read_serial_response_with_timeout(10, check_success=True)
            
            if success_type:
                if success_type == "central_success" or success_type == "success":
                    success_modules["central"] = True
                    self.log("✓ 中控升级成功")
                elif success_type == "electric_success":
                    success_modules["electric"] = True
                    self.log("✓ 电控升级成功")
                elif success_type == "ble_success":
                    success_modules["ble"] = True
                    self.log("✓ BLE升级成功")
            
            # 检查是否所有模块都升级成功
            all_success = all(success_modules.values())
            if all_success:
                self.log("🎉 所有模块串口OTA升级成功完成！")
                return True
            
            # 显示当前进度
            elapsed_time = int(time.time() - start_time)
            completed_count = sum(success_modules.values())
            self.log(f"串口升级进度: {completed_count}/3 已完成 - 中控: {'✓' if success_modules['central'] else '✗'}, "
                    f"电控: {'✓' if success_modules['electric'] else '✗'}, BLE: {'✓' if success_modules['ble'] else '✗'} "
                    f"- 已等待 {elapsed_time} 秒")
            
            time.sleep(5)
        
        # 超时处理
        failed_modules = [module for module, success in success_modules.items() if not success]
        self.log(f"串口OTA升级超时，未完成的模块: {failed_modules}", "ERROR")
        return False
    
    def execute_ota_via_serial(self):
        """通过串口执行OTA升级流程"""
        try:
            self.log("开始串口OTA升级流程")
            
            # 初始化串口
            if not self.initialize_serial():
                return False
            
            time.sleep(2)  # 等待串口稳定
            
            # 步骤1: 发送5次回车，每次间隔2秒
            self.log("发送5次回车激活串口...")
            for i in range(5):
                self.log(f"发送第 {i+1} 次回车")
                if not self.send_serial_command(""):
                    self.log(f"第 {i+1} 次回车发送失败", "WARNING")
                time.sleep(2)  # 每次间隔2秒
            
            # 步骤2: 等待并检测提示符
            self.log("等待30秒并检测串口提示符...")
            prompt_detected = False
            prompt_type = None
            
            for i in range(30):
                if self.stop_requested:
                    self.log("用户停止请求，中断等待")
                    return False
                
                # 每5秒检查一次串口响应
                if i % 5 == 0:
                    response, detected_type = self.read_serial_response_with_timeout(1, check_prompts=True)
                    if detected_type == "msh":
                        self.log("检测到MSH提示符，跳过密码输入")
                        prompt_detected = True
                        prompt_type = "msh"
                        break
                    elif detected_type == "password":
                        self.log("检测到密码提示符，等待30秒后输入密码")
                        prompt_detected = True
                        prompt_type = "password"
                        # 继续等待剩余的30秒
                        remaining_time = 30 - i
                        if remaining_time > 0:
                            self.log(f"继续等待 {remaining_time} 秒后输入密码")
                            for j in range(remaining_time):
                                if self.stop_requested:
                                    return False
                                time.sleep(1)
                        break
                
                time.sleep(1)
                if (i + 1) % 10 == 0:
                    self.log(f"已等待 {i+1} 秒")
            
            # 步骤3: 根据检测到的提示符决定是否输入密码
            if not prompt_detected:
                self.log("未检测到特定提示符，尝试检测当前状态")
                prompt_result = self.wait_for_prompt(10)
                if prompt_result == "msh":
                    self.log("检测到MSH提示符，跳过密码输入")
                    prompt_type = "msh"
                elif prompt_result == "password":
                    self.log("检测到密码提示符，输入密码")
                    prompt_type = "password"
                else:
                    self.log("未检测到明确提示符，默认需要输入密码", "WARNING")
                    prompt_type = "password"
            
            # 如果检测到密码提示符，则输入密码
            if prompt_type == "password":
                self.log("发送验证密码...")
                if not self.send_serial_command(self.upgrade_password):
                    return False
                
                time.sleep(2)
                
                # 读取密码验证响应
                response, _ = self.read_serial_response_with_timeout(5)
                if response:
                    self.log(f"密码验证响应: {response}")
            
            # 步骤4: 发送OTA命令（带回车）
            self.log("发送OTA命令...")
            if not self.send_serial_command(self.ota_command):
                return False
            
            # 步骤5: 等待OTA升级完成并分别检测三个模块
            ota_success = self.check_serial_upgrade_success(self.upgrade_timeout)
            
            # 关闭串口
            self.close_serial()
            
            return ota_success
            
        except Exception as e:
            self.log(f"串口OTA升级失败: {e}", "ERROR")
            self.close_serial()
            return False
    
    def start_upgrade_tool(self):
        """启动升级工具并执行升级流程"""
        upgrade_process = None
        try:
            self.log("启动升级工具...")
            # 先结束可能存在的旧进程 - 使用新的进程名
            self.kill_process_by_name("PC_TOOLS.exe")
            self.kill_process_by_name("L01_PCTOOL.exe")  # 同时结束旧进程名
            
            # 获取升级工具所在目录
            exe_path = self.upgrade_tool_path
            working_dir = os.path.dirname(exe_path)
            
            # 确保工作目录存在
            if not working_dir:
                working_dir = os.getcwd()
            
            self.log(f"升级工具路径: {exe_path}")
            self.log(f"工作目录: {working_dir}")
            
            # 使用正确的工作目录启动程序
            upgrade_process = subprocess.Popen([exe_path], cwd=working_dir, shell=True)
            self.log("升级工具已启动")
            
            # 等待程序完全启动
            time.sleep(10)  # 增加等待时间确保程序完全启动
            
            # 点击开始升级按钮
            self.log("点击开始升级按钮...")
            if not self.click_start_upgrade():
                self.log("点击开始升级按钮失败", "ERROR")
                if upgrade_process:
                    upgrade_process.terminate()
                return False
            
            # 等待升级完成并检查状态
            upgrade_success = self.wait_for_upgrade_completion()
            
            if upgrade_success:
                self.log("升级工具流程完成，所有模块升级成功")
                return True
            else:
                self.log("升级工具流程失败", "ERROR")
                return False
                
        except Exception as e:
            self.log(f"启动升级工具时出错: {e}", "ERROR")
            if upgrade_process:
                upgrade_process.terminate()
            return False
        finally:
            # 无论成功与否，都关闭升级工具
            if upgrade_process:
                self.log("关闭升级工具...")
                try:
                    upgrade_process.terminate()
                    time.sleep(2)
                    # 强制结束进程 - 使用新的进程名
                    self.kill_process_by_name("PC_TOOLS.exe")
                    self.kill_process_by_name("L01_PCTOOL.exe")  # 同时结束旧进程名
                    self.log("升级工具已关闭")
                except Exception as e:
                    self.log(f"关闭升级工具时出错: {e}", "ERROR")
    
    def wait_for_upgrade_completion(self):
        """等待升级完成并检查状态"""
        max_wait_time = 180  # 3分钟超时
        start_time = time.time()
        status_checked = False
        
        while time.time() - start_time < max_wait_time and not self.stop_requested:
            elapsed_time = int(time.time() - start_time)
            self.log(f"等待升级完成... 已等待 {elapsed_time} 秒")
            
            # 每隔20秒检查一次升级状态
            if elapsed_time > 160 :  # 160秒后开始检查
                if self.check_upgrade_status():
                    self.log("升级状态检查通过")
                    status_checked = True
                    break
            
            time.sleep(5)
        
        if status_checked:
            return True
        else:
            self.log("升级超时或状态检查失败", "ERROR")
            return False
    
    def run_single_cycle(self, cycle_num):
        """执行单个升级循环"""
        self.log(f"开始第 {cycle_num} 次升级循环")
        
        # 步骤1: 升级工具流程
        self.log("=== 步骤1: 升级工具流程 ===")
        if not self.start_upgrade_tool():
            self.log("升级工具流程失败", "ERROR")
            return False
        
        if self.stop_requested:
            return False
        
        # 步骤2: 串口OTA流程
        self.log("=== 步骤2: 串口OTA流程 ===")
        if not self.execute_ota_via_serial():
            self.log("串口OTA流程失败", "ERROR")
            return False
        
        self.log(f"第 {cycle_num} 次升级循环完成")
        return True
    
    def run_continuous_cycles(self, max_cycles=None):
        """连续运行升级循环"""
        cycle_count = 0
        
        try:
            while (max_cycles is None or cycle_count < max_cycles) and not self.stop_requested:
                cycle_count += 1
                self.log(f"\n{'='*50}")
                self.log(f"开始执行第 {cycle_count} 个升级循环")
                self.log(f"{'='*50}")
                
                success = self.run_single_cycle(cycle_count)
                
                if not success:
                    self.log(f"第 {cycle_count} 个升级循环失败", "ERROR")
                    break
                
                if self.stop_requested:
                    break
                
                self.log("等待下一次循环...")
                time.sleep(10)
                
        except KeyboardInterrupt:
            self.log("用户中断了升级流程")
        except Exception as e:
            self.log(f"升级过程中出错: {e}", "ERROR")
        finally:
            self.stop_requested = False
            self.close_serial()
            self.log("升级自动化程序已结束")

def main():
    """主函数"""
    root = tk.Tk()
    app = OTAUpgradeGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
