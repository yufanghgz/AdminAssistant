#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
根据入职人员信息表，通过列名匹配在最终版考勤表中增加人员信息
"""

import pandas as pd
import numpy as np
from datetime import datetime
import os

def add_employee_info_by_columns():
    """根据入职人员信息表，通过列名匹配在考勤表中增加人员信息"""
    
    # 数据目录
    data_dir = "/Users/heguangzhong/Documents/8月考勤"
    
    # 读取入职人员表
    new_employee_file = os.path.join(data_dir, "8月入职人员.xlsx")
    new_employee_df = pd.read_excel(new_employee_file)
    
    # 读取最终版考勤记录表
    final_attendance_file = os.path.join(data_dir, "考勤记录_考勤表_2025-8_完整版_20250905_092527.xlsx")
    attendance_df = pd.read_excel(final_attendance_file)
    
    print("开始处理数据...")
    print(f"入职人员表: {len(new_employee_df)} 条记录")
    print(f"最终版考勤记录表: {len(attendance_df)} 条记录")
    
    # 显示两个表的列名
    print("\n入职人员表列名:")
    print(list(new_employee_df.columns))
    print("\n考勤记录表列名:")
    print(list(attendance_df.columns))
    
    # 分析可匹配的列
    new_emp_cols = set(new_employee_df.columns)
    attendance_cols = set(attendance_df.columns)
    
    # 找到入职人员表中有但考勤记录表中没有的列
    new_cols = new_emp_cols - attendance_cols
    print(f"\n入职人员表中有但考勤记录表中没有的列: {list(new_cols)}")
    
    # 找到两个表都有的列
    common_cols = new_emp_cols & attendance_cols
    print(f"\n两个表都有的列: {list(common_cols)}")
    
    # 创建入职人员信息字典，以姓名为键
    new_employee_info = {}
    for _, row in new_employee_df.iterrows():
        name = row['姓名']
        # 只保存考勤记录表中没有的列的信息
        info = {}
        for col in new_cols:
            if col in row and pd.notna(row[col]):
                info[col] = row[col]
        new_employee_info[name] = info
    
    print(f"\n入职人员新增信息字典:")
    for name, info in new_employee_info.items():
        print(f"  {name}: {info}")
    
    # 在考勤记录表中添加新列
    for col in new_cols:
        attendance_df[col] = ''
    
    # 按姓名匹配，填充入职人员信息
    matched_count = 0
    matched_people = []
    for idx, row in attendance_df.iterrows():
        name = row['人员']
        if name in new_employee_info:
            for col, value in new_employee_info[name].items():
                attendance_df.at[idx, col] = value
            matched_count += 1
            matched_people.append(name)
            print(f"匹配成功: {name}")
    
    print(f"\n匹配结果:")
    print(f"成功匹配: {matched_count} 人")
    print(f"匹配的人员: {matched_people}")
    print(f"未匹配的入职人员: {len(new_employee_df) - matched_count} 人")
    
    # 检查未匹配的入职人员
    attendance_names = set(attendance_df['人员'].tolist())
    new_employee_names = set(new_employee_df['姓名'].tolist())
    unmatched = new_employee_names - attendance_names
    if unmatched:
        print(f"未在考勤记录中找到的入职人员: {list(unmatched)}")
    
    # 保存更新后的考勤记录表
    output_file = os.path.join(data_dir, f"考勤记录_考勤表_2025-8_完整版_按列匹配_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    attendance_df.to_excel(output_file, index=False)
    
    print(f"\n完整版考勤记录表已保存到: {output_file}")
    print(f"新增列: {list(new_cols)}")
    
    # 显示有入职信息的记录
    new_employee_records = attendance_df[attendance_df['入职日期'] != '']
    if not new_employee_records.empty:
        print(f"\n包含入职信息的记录 ({len(new_employee_records)} 条):")
        # 显示关键列
        key_cols = ['人员', '入职日期', '转正日期', '转正状态', '最高学历', '政治面貌']
        available_cols = [col for col in key_cols if col in new_employee_records.columns]
        print(new_employee_records[available_cols].to_string(index=False))
    
    # 显示最终统计
    print(f"\n最终统计:")
    print(f"总记录数: {len(attendance_df)}")
    print(f"包含入职信息: {len(new_employee_records)} 条")
    print(f"包含离职信息: {len(attendance_df[attendance_df['离职日期'] != ''])} 条")
    print(f"普通员工: {len(attendance_df) - len(new_employee_records) - len(attendance_df[attendance_df['离职日期'] != ''])} 条")
    
    # 显示新增的列信息
    print(f"\n新增列详情:")
    for col in new_cols:
        non_empty_count = (attendance_df[col] != '').sum()
        print(f"  {col}: {non_empty_count} 条记录有数据")
    
    return output_file

if __name__ == "__main__":
    add_employee_info_by_columns()
