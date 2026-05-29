import csv
import logging
import itertools
from typing import List, Dict, Any

# 1. 配置日志记录 (替代简单的print)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def generate_combinatorial_test_cases(
    parameters: Dict[str, List[Any]], 
    output_filename: str
) -> None:
    """
    根据输入的参数字典，生成所有可能的笛卡尔积组合测试用例，并导出为CSV文件。
    
    Args:
        parameters: 包含测试变量及其对应取值列表的字典
        output_filename: 输出的CSV文件路径
    """
    logging.info(f"开始生成测试用例，目标文件: {output_filename}")
    
    try:
        # 提取参数名称和对应的取值列表
        keys = list(parameters.keys())
        values_lists = list(parameters.values())
        
        # 使用 itertools.product 生成所有可能的组合
        # 例如: (温度高, 电压稳, 风扇开), (温度高, 电压稳, 风扇关)...
        combinations = list(itertools.product(*values_lists))
        
        total_cases = len(combinations)
        logging.info(f"共计算出 {total_cases} 条组合用例。")
        
        # 将生成的组合写入CSV文件
        with open(output_filename, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            
            # 写入表头：用例编号 + 各个参数名 + 预期结果（留空供人工或AI后续补充）
            header = ["用例编号"] + keys + ["预期结果"]
            writer.writerow(header)
            
            # 写入具体数据
            for index, combination in enumerate(combinations, start=1):
                case_id = f"TC_{index:04d}"
                row_data = [case_id] + list(combination) + ["待补充"]
                writer.writerow(row_data)
                
        logging.info(f"测试用例已成功导出至 {output_filename}")

    except IOError as e:
        # 处理文件读写异常
        logging.error(f"文件操作失败: {e}")
    except Exception as e:
        # 捕获其他未知异常，防止程序直接崩溃
        logging.error(f"生成测试用例时发生未知错误: {e}", exc_info=True)

# ---------------- 测试执行入口 ----------------
if __name__ == "__main__":
    # 定义测试因子（你可以让AI帮你梳理出这些因子，然后填入这里）
    test_factors: Dict[str, List[Any]] = {
        "传感器温度": ["-40度(下限)", "25度(常温)", "85度(上限)", "100度(越界)", "断路异常"],
        "输入电压": ["9V(欠压)", "12V(额定)", "16V(过压)"],
        "散热风扇状态": ["正常运转", "堵转告警", "未连接"],
        "运行模式": ["待机", "全负载", "低功耗"]
    }
    
    csv_file_path = "embedded_temp_control_cases.csv"
    
    # 5 * 3 * 3 * 3 = 135条基础组合，如果增加因子，轻松生成上千条
    generate_combinatorial_test_cases(test_factors, csv_file_path)