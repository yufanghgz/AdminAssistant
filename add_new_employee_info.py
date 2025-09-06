#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
根据入职人员信息表，在最终版考勤表中增加人员信息
"""

import pandas as pd
import numpy as np
from datetime import datetime
import os

def add_new_employee_info():
    """根据入职人员信息表在考勤表中增加人员信息"""
    
    # 数据目录
    data_dir = "/Users/heguangzhong/Documents/8月考勤"
    
    # 读取入职人员表
    new_employee_file = os.path.join(data_dir, "8月入职人员.xlsx")
    new_employee_df = pd.read_excel(new_employee_file)
    
    # 读取最终版考勤记录表
    final_attendance_file = os.path.join(data_dir, "考勤记录_考勤表_2025-8_最终版_20250905_091717.xlsx")
    attendance_df = pd.read_excel(final_attendance_file)
    
    print("开始处理数据...")
    print(f"入职人员表: {len(new_employee_df)} 条记录")
    print(f"最终版考勤记录表: {len(attendance_df)} 条记录")
    
    # 显示入职人员表结构
    print("\n入职人员表结构:")
    print(f"列名: {list(new_employee_df.columns)}")
    print("\n前5行数据:")
    print(new_employee_df.head())
    
    # 显示考勤记录表结构
    print("\n考勤记录表结构:")
    print(f"列名: {list(attendance_df.columns)}")
    print("\n前5行数据:")
    print(attendance_df.head())
    
    # 创建入职人员信息字典，以姓名为键
    new_employee_info = {}
    for _, row in new_employee_df.iterrows():
        name = row['姓名']
        new_employee_info[name] = {
            '工号': row['工号'],
            '入职日期': row['入职日期'],
            '转正日期': row['转正日期'],
            '人员类型': row['人员类型'],
            '转正状态': row['转正状态'],
            '最高学历': row['最高学历'],
            '最高学历学校': row['最高学历学校'],
            '政治面貌': row['政治面貌']
        }
    
    print(f"\n入职人员信息字典:")
    for name, info in new_employee_info.items():
        print(f"  {name}: {info}")
    
    # 在考勤记录表中添加新列
    attendance_df['入职日期'] = ''
    attendance_df['转正日期'] = ''
    attendance_df['转正状态'] = ''
    attendance_df['最高学历'] = ''
    attendance_df['最高学历学校'] = ''
    attendance_df['政治面貌'] = ''
    
    # 按姓名匹配，填充入职人员信息
    matched_count = 0
    matched_people = []
    for idx, row in attendance_df.iterrows():
        name = row['人员']
        if name in new_employee_info:
            attendance_df.at[idx, '入职日期'] = new_employee_info[name]['入职日期']
            attendance_df.at[idx, '转正日期'] = new_employee_info[name]['转正日期']
            attendance_df.at[idx, '转正状态'] = new_employee_info[name]['转正状态']
            attendance_df.at[idx, '最高学历'] = new_employee_info[name]['最高学历']
            attendance_df.at[idx, '最高学历学校'] = new_employee_info[name]['最高学历学校']
            attendance_df.at[idx, '政治面貌'] = new_employee_info[name]['政治面貌']
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
    output_file = os.path.join(data_dir, f"考勤记录_考勤表_2025-8_完整版_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    attendance_df.to_excel(output_file, index=False)
    
    print(f"\n完整版考勤记录表已保存到: {output_file}")
    print(f"新增列: 入职日期, 转正日期, 转正状态, 最高学历, 最高学历学校, 政治面貌")
    
    # 显示有入职信息的记录
    new_employee_records = attendance_df[attendance_df['入职日期'] != '']
    if not new_employee_records.empty:
        print(f"\n包含入职信息的记录 ({len(new_employee_records)} 条):")
        print(new_employee_records[['人员', '入职日期', '转正日期', '转正状态', '最高学历', '政治面貌']].to_string(index=False))
    
    # 显示最终统计
    print(f"\n最终统计:")
    print(f"总记录数: {len(attendance_df)}")
    print(f"包含入职信息: {len(new_employee_records)} 条")
    print(f"包含离职信息: {len(attendance_df[attendance_df['离职日期'] != ''])} 条")
    print(f"普通员工: {len(attendance_df) - len(new_employee_records) - len(attendance_df[attendance_df['离职日期'] != ''])} 条")
    
    return output_file

if __name__ == "__main__":
    add_new_employee_info()






