import os

import pandas as pd
import config


def hld():
    # 读取健康度判别表
    input_file = config.output_file
    output_file = config.output_file

    # 验证文件路径
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"文件未找到: {input_file}")

    # 读取Excel文件中的"在建"sheet
    try:
        df = pd.read_excel(input_file, sheet_name='zj')
        k1=1
    except ValueError as e:
        # 如果找不到名为"在建"的工作表，给出提示
        print(f"读取工作表时出错: {e}")
        print("请确保已运行zj.py生成'在建'工作表")
        return

    # 筛选审批级别校准为"行办会"的条目
    filtered_df = df[df['审批级别校准'] == '行领导']

    # 读取现有的Excel文件
    with pd.ExcelWriter(output_file, mode='a', if_sheet_exists='replace', engine='openpyxl') as writer:
        filtered_df.to_excel(writer, sheet_name='hld', index=False)