import os

import pandas as pd
import config


def zj():
    # 读取健康度判别表
    input_file = config.output_file
    output_file = config.output_file

    # 验证文件路径
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"文件未找到: {input_file}")

    # 读取Excel文件
    df = pd.read_excel(input_file)

    # 筛选需求状态为"开发中", "开发完成", "已提交测试", "已通过测试", "已通过审核", "已计划上线"的需求条目
    filtered_df = df[
        df['需求状态'].isin(["开发中", "开发完成", "已提交测试", "已通过测试", "已通过审核", "已计划上线","首次上线"])]


    # 读取现有的Excel文件
    with pd.ExcelWriter(output_file, mode='a', if_sheet_exists='new', engine='openpyxl') as writer:
        filtered_df.to_excel(writer, sheet_name='zj', index=False)
