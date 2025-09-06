#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
删除原考勤表中备注列包含"离职"信息的人员
"""

import pandas as pd
import numpy as np
from datetime import datetime
import os

def remove_departure_from_remarks():
    """删除原考勤表中备注列包含"离职"信息的人员"""
    
    # 数据目录
    data_dir = "/Users/heguangzhong/Documents/8月考勤"
    
    # 读取原考勤记录表
    attendance_file = os.path.join(data_dir, "考勤记录_考勤表_2025-8.xlsx")
    attendance_df = pd.read_excel(attendance_file)
    
    print("开始处理数据...")
    print(f"原考勤记录表: {len(attendance_df)} 条记录")
    
    # 检查备注列的内容
    print("\n备注列内容分析:")
    print(f"备注列非空记录数: {attendance_df['备注'].notna().sum()}")
    
    # 显示所有非空的备注内容
    non_empty_remarks = attendance_df[attendance_df['备注'].notna()]
    if not non_empty_remarks.empty:
        print("\n所有非空备注内容:")
        for idx, row in non_empty_remarks.iterrows():
            print(f"  {row['人员']}: {row['备注']}")
    
    # 查找包含"离职"的备注
    departure_in_remarks = attendance_df[attendance_df['备注'].str.contains('离职', na=False)]
    print(f"\n包含'离职'的备注记录数: {len(departure_in_remarks)}")
    
    if not departure_in_remarks.empty:
        print("包含'离职'的备注记录:")
        for idx, row in departure_in_remarks.iterrows():
            print(f"  {row['人员']}: {row['备注']}")
        
        # 删除包含"离职"的备注记录
        attendance_df_cleaned = attendance_df[~attendance_df['备注'].str.contains('离职', na=False)]
        
        print(f"\n删除后记录数: {len(attendance_df_cleaned)}")
        print(f"删除了 {len(attendance_df) - len(attendance_df_cleaned)} 条记录")
        
        # 保存清理后的考勤记录表
        output_file = os.path.join(data_dir, f"考勤记录_考勤表_2025-8_已删除离职备注_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        attendance_df_cleaned.to_excel(output_file, index=False)
        
        print(f"\n清理后的考勤记录表已保存到: {output_file}")
        
        # 显示被删除的人员名单
        deleted_people = departure_in_remarks['人员'].tolist()
        print(f"\n被删除的人员: {deleted_people}")
        
        return output_file, deleted_people
    else:
        print("没有找到包含'离职'的备注记录")
        return None, []

if __name__ == "__main__":
    remove_departure_from_remarks()






