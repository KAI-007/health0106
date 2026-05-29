import os
from datetime import datetime, timedelta

import pandas as pd
import config


def jhsx():  # 将函数名从hbh改为jhsx，以匹配文件名
    # 读取健康度判别表
    input_file = config.output_file
    output_file = config.output_file

    # 验证文件路径
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"文件未找到: {input_file}")

    # 读取Excel文件中的"Sheet1"工作表
    try:
        df = pd.read_excel(input_file, sheet_name='Sheet1')
    except ValueError as e:
        # 如果找不到名为"Sheet1"的工作表，给出提示
        print(f"读取工作表时出错: {e}")
        return

    # 确保'实际上线日期'列存在且为日期格式
    if '实际上线日期' not in df.columns:
        print("错误：找不到'实际上线日期'列")
        return

    # 将'实际上线日期'列转换为日期格式
    df['实际上线日期'] = pd.to_datetime(df['实际上线日期'], errors='coerce')

    # 过滤掉无效的日期
    valid_dates_df = df.dropna(subset=['实际上线日期'])
    
    # 计算当前周的过去两周时间范围
    today = datetime.today()
    # 计算两周前的日期（14天前）
    start_date = today - timedelta(days=14)
    
    # 筛选实际上线日期在指定时间范围内的条目
    filtered_df = valid_dates_df[(valid_dates_df['实际上线日期'] >= start_date) & (valid_dates_df['实际上线日期'] <= today)]

    # 读取现有的Excel文件并保存到'计划上线表格'工作表
    with pd.ExcelWriter(output_file, mode='a', if_sheet_exists='replace', engine='openpyxl') as writer:
        filtered_df.to_excel(writer, sheet_name='计划上线', index=False)


if __name__ == "__main__":
    jhsx()