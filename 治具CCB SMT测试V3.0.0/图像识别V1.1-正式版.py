import os
import cv2
import logging
from collections import Counter

# --- 环境配置 ---
# 抑制 PaddlePaddle 的底层 C++ 日志
os.environ['FLAGS_minloglevel'] = '2'
# 抑制 Python 层的 logging 日志
logging.getLogger("ppocr").setLevel(logging.ERROR)

def extract_text(result):
    """
    递归提取识别结果中的所有文本信息
    兼容 PaddleOCR 不同版本的返回格式 (list/dict/object)
    """
    texts = []
    if isinstance(result, (list, tuple)):
        for item in result:
            texts.extend(extract_text(item))
    elif isinstance(result, dict):
        # 提取字典中可能的文本字段
        for key in ['rec_text', 'text', 'label']:
            if key in result:
                texts.append(result[key])
        for val in result.values():
            texts.extend(extract_text(val))
    elif isinstance(result, str):
        # 过滤掉文件后缀等非文字信息
        if len(result) > 0 and not result.endswith(('.jpg', '.png')):
            texts.append(result)
    elif hasattr(result, 'text'):
        texts.append(result.text)
    elif hasattr(result, 'rec_text'):
        texts.append(result.rec_text)
    return texts

def main():
    print("=" * 50)
    print("正在启动分析脚本 (优化版 V1.3 - 修正兼容性)")
    print("策略: 顶部强制切割 + 关键词深度过滤 + 智能干扰剔除")
    print("=" * 50 + "\n")

    # 1. 初始化 OCR 引擎
    try:
        from paddleocr import PaddleOCR
        # [修正] 移除不支持的 show_log 参数，使用默认初始化
        ocr = PaddleOCR(lang='ch')
    except Exception as e:
        print(f"[错误] OCR 初始化失败: {e}")
        print("请检查是否安装了 paddlepaddle 和 paddleocr")
        return

    # --- 请在此处确认或修改图片文件夹路径 ---
    images_folder = r"C:\Users\szm21\Desktop\fail_screenshots-L5"
    # -------------------------------------

    if not os.path.exists(images_folder):
        print(f"[错误] 文件夹路径不存在: {images_folder}")
        return

    # 获取所有图片文件
    files = [f for f in os.listdir(images_folder) if f.lower().endswith(('.png', '.jpg', '.bmp', '.jpeg'))]
    total_files = len(files)
    print(f"准备分析 {total_files} 张图片...\n")

    # 统计计数器
    target_count = 0
    other_errors = Counter()
    
    # 用于记录非功耗异常的详细识别结果 (方便后续排查)
    unknown_details = []

    # 定义无意义的干扰词 (黑名单)
    # [优化] 增加了大量中文界面常用词，防止它们被错误拼接到异常描述中
    noise_words = {
        'general', 'min', 'test', 'time', 'sec', 'pass', 'fail', 'prd', 'batt', 'lang', 'seconds',
        '测试用时', '参数配置', '软件版本', '硬件版本', '序列号', '串口', '关闭', '开启', '数据库', 
        '开始测试', '重新测试', '不通过', 'adc校准', '显示', '电池电压', '按键', '触摸'
    }

    # 定义明确的异常关键词 (白名单)
    # [优化] 增加了 "检测", "正在", "禁止" 等词，确保能捕获非标准Error的异常标题
    error_keywords = [
        "异常", "失败", "Error", "Fail", "超时", "崩溃", "无响应", 
        "检测", "正在", "禁止", "Warning", "Fatal", "断开"
    ]

    for idx, file in enumerate(files):
        path = os.path.join(images_folder, file)
        
        try:
            # 2. 读取并切割图片
            img = cv2.imread(path)
            if img is None:
                continue
                
            h, w = img.shape[:2]
            
            # 核心策略: 只切取顶部 25% 的区域进行识别
            limit_h = int(h * 0.25)
            img_header = img[0:limit_h, 0:w]

            # 3. 执行识别
            # 继续使用 predict 接口以保持与您环境的兼容性
            result = ocr.predict(img_header)
            txts = extract_text(result)
            
            # 生成去空格的全文，用于匹配
            full_text_nospace = "".join(txts).replace(" ", "")

            # --- 判定逻辑 ---
            
            # 情况 A: 包含目标关键词 "功耗异常"
            if "功耗异常" in full_text_nospace:
                target_count += 1
            
            # 情况 B: 其他异常分类
            else:
                best_label = "无法识别文字"
                found_keyword = False
                
                # 策略 1: 优先查找明确的报错关键词
                for t in txts:
                    # 只要包含关键词库中的任意一个词，直接作为分类依据，不再拼接后续内容
                    if any(k in t for k in error_keywords):
                        best_label = t
                        found_keyword = True
                        break
                
                # 策略 2: 如果没有关键词，过滤掉 noise_words，尝试拼凑描述
                if not found_keyword:
                    # 过滤黑名单，且保留长度大于1的词 (排除单字干扰)
                    meaningful_words = [t for t in txts if t.lower() not in noise_words and len(t) > 1]
                    
                    if meaningful_words:
                        # [优化] 如果第一个词已经比较长(>4个字)，可能本身就是标题，就不拼接后面的了
                        if len(meaningful_words[0]) > 4:
                            best_label = meaningful_words[0]
                        else:
                            # 否则取前2个词拼接
                            best_label = "".join(meaningful_words[:2])
                    elif txts:
                        # 实在没有筛选出结果，取原始识别的第一行
                        best_label = txts[0]

                other_errors[best_label] += 1
                
                # 记录该文件的详细识别结果
                unknown_details.append(f"文件名: {file} -> 最终标签: [{best_label}] | 原始识别: {txts}")

        except Exception as e:
            print(f"[错误] 处理图片 {file} 时出错: {e}")

        # 进度打印
        if (idx + 1) % 10 == 0:
            print(f"已处理 {idx + 1}/{total_files} ...")

    # --- 输出最终报告 ---
    print("\n" + "=" * 50)
    print("分析报告")
    print("=" * 50)
    print(f"分析文件总数: {total_files}")
    print(f"功耗异常 数量: {target_count}")
    print("-" * 50)
    print("其他异常情况分布 (已过滤干扰词):")
    
    if not other_errors:
        print("   [无其他异常发现]")
    else:
        # 按出现次数从多到少排序
        for error_text, count in other_errors.most_common():
            print(f"   [{count} 张] -> {error_text}")

    print("\n" + "=" * 50)
    print("非功耗异常的详细记录 (Top 5 样本):")
    for i, detail in enumerate(unknown_details[:5]):
        print(f"   {i+1}. {detail}")
    
    print(f"   (共记录 {len(unknown_details)} 条)")
    print("=" * 50)

    input("按回车键退出程序...")

if __name__ == "__main__":
    main()
