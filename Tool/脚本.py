import tkinter as tk
from tkinter import ttk, messagebox
import logging
import json
import re
from typing import Dict, List, Any, Optional

# 使用 logging 模块记录日志，替代简单的 print
logging.basicConfig(
    level=logging.INFO, 
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class SpeedConfigAnalyzer:
    """
    速度配置解析器，负责将原始 JSON 数据转化为显示速度到实际速度的映射关系字典。
    """
    def __init__(self, config_data: Dict[str, Any]) -> None:
        self.config_data = config_data
        # mapping_data 结构：{显示速度: [{"gear": "D1", "region": "GE", "actual_speed": 50}]}
        self.mapping_data: Dict[int, List[Dict[str, Any]]] = self._build_mapping()

    def _build_mapping(self) -> Dict[int, List[Dict[str, Any]]]:
        """
        构建排列组合的映射关系
        """
        mapping: Dict[int, List[Dict[str, Any]]] = {}
        
        # 错误处理结构：防止配置文件节点缺失导致崩溃
        try:
            sn_config = self.config_data.get("sn", {})
            args_config = self.config_data.get("args", {})

            gears: List[str] = ["d1", "d2", "d3"]
            
            # 1. 提取 args 中的默认配置
            for gear in gears:
                show_key = f"speed_show_{gear}"
                gear_key = f"speed_gear_{gear}"

                def_show: Optional[int] = args_config.get(show_key)
                def_gear: Optional[int] = args_config.get(gear_key)

                if def_show is not None and def_gear is not None:
                    if def_show not in mapping:
                        mapping[def_show] = []
                    mapping[def_show].append({
                        "gear": gear.upper(),
                        "region": "默认 (Default)",
                        "actual_speed": def_gear
                    })

            # 2. 提取 sn 中的特殊地区匹配配置
            for gear in gears:
                show_match_list: List[Dict[str, int]] = sn_config.get(f"speed_show_{gear}_match", [])
                gear_match_list: List[Dict[str, int]] = sn_config.get(f"speed_gear_{gear}_match", [])

                # 遍历显示速度的区域映射，如 {"(JP|GE)": 20}
                for show_dict in show_match_list:
                    for region_pattern, show_speed in show_dict.items():
                        # 使用正则提取出具体的国家代码，如 JP, GE
                        regions = re.findall(r'[A-Z]{2}', region_pattern)

                        for target_region in regions:
                            actual_speed: Optional[int] = None
                            
                            # 在实际速度匹配列表中寻找对应国家的设定
                            for gear_dict in gear_match_list:
                                for g_pattern, g_speed in gear_dict.items():
                                    if target_region in g_pattern:
                                        actual_speed = g_speed
                                        break
                                if actual_speed is not None:
                                    break

                            # 如果找到了该区域对应的实际速度，则加入映射字典
                            if actual_speed is not None:
                                if show_speed not in mapping:
                                    mapping[show_speed] = []
                                mapping[show_speed].append({
                                    "gear": gear.upper(),
                                    "region": target_region,
                                    "actual_speed": actual_speed
                                })
                                
            logging.info("配置解析与映射关系构建成功。")
            
        except Exception as e:
            logging.error(f"构建速度映射关系时发生异常: {e}")
            
        return mapping

    def query_actual_speed(self, show_speed: int) -> List[Dict[str, Any]]:
        """
        根据给定的显示速度，返回所有可能的实际速度组合。
        """
        return self.mapping_data.get(show_speed, [])


class SpeedGUI:
    """
    使用 Tkinter 构建的图形用户界面类。
    """
    def __init__(self, root: tk.Tk, analyzer: SpeedConfigAnalyzer) -> None:
        self.root = root
        self.analyzer = analyzer
        self.root.title("挡位速度查询组合工具")
        self.root.geometry("500x350")

        self._create_widgets()
        logging.info("GUI 界面初始化完成。")

    def _create_widgets(self) -> None:
        """
        初始化并布局 UI 组件。
        """
        # 顶部输入容器
        input_frame = tk.Frame(self.root)
        input_frame.pack(pady=20)

        tk.Label(input_frame, text="请输入显示速度配置值 (如 30, 40, 60):").grid(row=0, column=0, padx=5)
        self.speed_entry = tk.Entry(input_frame, width=10)
        self.speed_entry.grid(row=0, column=1, padx=5)

        query_btn = tk.Button(input_frame, text="查询组合", command=self._on_query)
        query_btn.grid(row=0, column=2, padx=5)

        # 底部结果展示容器
        result_frame = tk.Frame(self.root)
        result_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # 使用 Scrollbar 配合 Text 以防数据过多
        scrollbar = tk.Scrollbar(result_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.result_text = tk.Text(result_frame, state=tk.DISABLED, wrap=tk.WORD, yscrollcommand=scrollbar.set)
        self.result_text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.result_text.yview)

    def _on_query(self) -> None:
        """
        查询按钮的事件处理函数。
        """
        input_val = self.speed_entry.get().strip()
        
        # 错误处理：校验输入合法性
        try:
            show_speed = int(input_val)
            logging.info(f"用户发起查询，输入的显示速度为: {show_speed}")
        except ValueError:
            messagebox.showwarning("输入错误", "请输入有效的整数配置值！")
            logging.warning(f"用户输入了无效的数值: '{input_val}'")
            return

        results = self.analyzer.query_actual_speed(show_speed)

        # 开放 Text 控件编辑权限以写入结果
        self.result_text.config(state=tk.NORMAL)
        self.result_text.delete("1.0", tk.END)

        if not results:
            self.result_text.insert(tk.END, f"未找到显示速度为 【{show_speed}】 的相关配置文件组合。\n")
        else:
            self.result_text.insert(tk.END, f"输入显示速度 【{show_speed}】 产生以下实际速度组合：\n\n")
            for item in results:
                # 按照文件注释，实际物理速度为配置值除以 10
                physical_speed = item['actual_speed'] / 10
                line = (f"• 挡位: {item['gear']:<4} | "
                        f"地区: {item['region']:<12} | "
                        f"配置实际速度: {item['actual_speed']:<4} "
                        f"(约 {physical_speed:.1f} km/h)\n")
                self.result_text.insert(tk.END, line)

        # 锁定 Text 控件防止用户误修改
        self.result_text.config(state=tk.DISABLED)


def main() -> None:
    # 错误处理结构：包裹主入口以捕获致命错误
    try:
        # 提取用户提供的文件内容核心部分作为默认数据源（便于单文件直接运行）
        json_str = """
        {
            "sn": {
                "speed_gear_d1_match": [{"(JP|GE)":20}],
                "speed_gear_d2_match": [{"(GE)":50},{"(JP)":40}],
                "speed_gear_d3_match": [{"(GE)":80},{"(JP)":60}],
                "speed_show_d1_match": [{"(JP|GE)":20}],
                "speed_show_d2_match": [{"(JP|GE)":40}],
                "speed_show_d3_match": [{"(JP|GE)":60}]
            },
            "args": {
                "speed_gear_d1": 30,
                "speed_gear_d2": 60,
                "speed_gear_d3": 100,
                "speed_show_d1": 30,
                "speed_show_d2": 60,
                "speed_show_d3": 100
            }
        }
        """
        config_data: Dict[str, Any] = json.loads(json_str)
        
        analyzer = SpeedConfigAnalyzer(config_data)
        
        root = tk.Tk()
        app = SpeedGUI(root, analyzer)
        root.mainloop()
        
    except Exception as e:
        logging.critical(f"程序启动发生致命错误: {e}")

if __name__ == "__main__":
    main()