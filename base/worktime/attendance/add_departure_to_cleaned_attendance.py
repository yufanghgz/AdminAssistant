#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
在已删除离职备注的考勤表中，按姓名匹配添加离职信息四列
"""

import pandas as pd
import numpy as np
from datetime import datetime
import os

def add_departure_to_cleaned_attendance():
    """在已删除离职备注的考勤表中添加离职信息"""
    
    # 数据目录
    data_dir = "/Users/heguangzhong/Documents/8月考勤"
    
    # 读取离职人员表
    departure_file = os.path.join(data_dir, "8月离职人员.xlsx")
    departure_df = pd.read_excel(departure_file)
    
    # 读取已删除离职备注的考勤记录表
    cleaned_attendance_file = os.path.join(data_dir, "考勤记录_考勤表_2025-8_已删除离职备注_20250905_091150.xlsx")
    attendance_df = pd.read_excel(cleaned_attendance_file)
    
    print("开始处理数据...")
    print(f"离职人员表: {len(departure_df)} 条记录")
    print(f"已清理的考勤记录表: {len(attendance_df)} 条记录")
    
    # 创建离职人员信息字典，以姓名为键
    departure_info = {}
    for _, row in departure_df.iterrows():
        name = row['姓名']
        departure_info[name] = {
            '离职日期': row['离职日期'],
            '离职类型': row['离职类型'],
            '离职原因': row['离职原因'],
            '离职备注': row['离职备注']
        }
    
    print(f"\n离职人员信息字典:")
    for name, info in departure_info.items():
        print(f"  {name}: {info}")
    
    # 在考勤记录表中添加四列
    attendance_df['离职日期'] = ''
    attendance_df['离职类型'] = ''
    attendance_df['离职原因'] = ''
    attendance_df['离职备注'] = ''
    
    # 按姓名匹配，填充离职信息
    matched_count = 0
    matched_people = []
    for idx, row in attendance_df.iterrows():
        name = row['人员']
        if name in departure_info:
            attendance_df.at[idx, '离职日期'] = departure_info[name]['离职日期']
            attendance_df.at[idx, '离职类型'] = departure_info[name]['离职类型']
            attendance_df.at[idx, '离职原因'] = departure_info[name]['离职原因']
            attendance_df.at[idx, '离职备注'] = departure_info[name]['离职备注']
            matched_count += 1
            matched_people.append(name)
            print(f"匹配成功: {name}")
    
    print(f"\n匹配结果:")
    print(f"成功匹配: {matched_count} 人")
    print(f"匹配的人员: {matched_people}")
    print(f"未匹配的离职人员: {len(departure_df) - matched_count} 人")
    
    # 检查未匹配的离职人员
    attendance_names = set(attendance_df['人员'].tolist())
    departure_names = set(departure_df['姓名'].tolist())
    unmatched = departure_names - attendance_names
    if unmatched:
        print(f"未在考勤记录中找到的离职人员: {list(unmatched)}")
    
    # 保存更新后的考勤记录表
    output_file = os.path.join(data_dir, f"考勤记录_考勤表_2025-8_最终版_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    attendance_df.to_excel(output_file, index=False)
    
    print(f"\n最终版考勤记录表已保存到: {output_file}")
    print(f"新增列: 离职日期, 离职类型, 离职原因, 离职备注")
    
    # 显示有离职信息的记录
    departure_records = attendance_df[attendance_df['离职日期'] != '']
    if not departure_records.empty:
        print(f"\n包含离职信息的记录 ({len(departure_records)} 条):")
        print(departure_records[['人员', '离职日期', '离职类型', '离职原因', '离职备注']].to_string(index=False))
    
    # 显示最终统计
    print(f"\n最终统计:")
    print(f"总记录数: {len(attendance_df)}")
    print(f"包含离职信息: {len(departure_records)} 条")
    print(f"不含离职信息: {len(attendance_df) - len(departure_records)} 条")
    
    return output_file

if __name__ == "__main__":
    add_departure_to_cleaned_attendance()











