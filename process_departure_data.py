#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
处理离职人员数据，将离职信息添加到考勤记录表中
"""

import pandas as pd
import numpy as np
from datetime import datetime
import os

def process_departure_data():
    """处理离职人员数据，按姓名匹配添加到考勤记录表"""
    
    # 数据目录
    data_dir = "/Users/heguangzhong/Documents/8月考勤"
    
    # 读取离职人员表
    departure_file = os.path.join(data_dir, "8月离职人员.xlsx")
    departure_df = pd.read_excel(departure_file)
    
    print("离职人员表结构:")
    print(f"列名: {list(departure_df.columns)}")
    print(f"数据形状: {departure_df.shape}")
    print("\n前5行数据:")
    print(departure_df.head())
    
    # 读取考勤记录表
    attendance_file = os.path.join(data_dir, "考勤记录_考勤表_2025-8.xlsx")
    attendance_df = pd.read_excel(attendance_file)
    
    print("\n考勤记录表结构:")
    print(f"列名: {list(attendance_df.columns)}")
    print(f"数据形状: {attendance_df.shape}")
    print("\n前5行数据:")
    print(attendance_df.head())
    
    # 检查姓名列
    print("\n离职人员表中的姓名列:")
    if '姓名' in departure_df.columns:
        print(departure_df['姓名'].tolist())
    else:
        print("未找到'姓名'列，可用列名:", list(departure_df.columns))
    
    print("\n考勤记录表中的姓名列:")
    if '人员' in attendance_df.columns:
        print(attendance_df['人员'].unique()[:10])  # 显示前10个不重复的姓名
    else:
        print("未找到'人员'列，可用列名:", list(attendance_df.columns))

if __name__ == "__main__":
    process_departure_data()






