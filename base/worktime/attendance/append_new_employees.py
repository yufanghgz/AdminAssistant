#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
根据入职人员信息表，提取姓名、工号、人员类型、银行卡号四列信息，
追加到考勤表中，如果工号相同则不追加
"""

import pandas as pd
import numpy as np
from datetime import datetime
import os

def append_new_employees():
    """从入职人员信息表提取信息追加到考勤表中"""
    
    # 数据目录
    data_dir = "/Users/heguangzhong/Documents/8月考勤"
    
    # 读取入职人员表
    new_employee_file = os.path.join(data_dir, "8月入职人员.xlsx")
    new_employee_df = pd.read_excel(new_employee_file)
    
    # 读取目标考勤表
    target_attendance_file = os.path.join(data_dir, "02-考勤记录_考勤表_2025-8_增加8月离职人员相关列_20250905_091717.xlsx")
    attendance_df = pd.read_excel(target_attendance_file)
    
    print("开始处理数据...")
    print(f"入职人员表: {len(new_employee_df)} 条记录")
    print(f"目标考勤表: {len(attendance_df)} 条记录")
    
    # 显示入职人员表结构
    print("\n入职人员表列名:")
    print(list(new_employee_df.columns))
    
    # 显示目标考勤表结构
    print("\n目标考勤表列名:")
    print(list(attendance_df.columns))
    
    # 从入职人员表中提取需要的四列信息
    required_cols = ['姓名', '工号', '人员类型', '银行卡号']
    new_employee_data = new_employee_df[required_cols].copy()
    
    print(f"\n提取的入职人员数据:")
    print(new_employee_data)
    
    # 检查目标考勤表中现有的工号
    existing_employee_ids = set(attendance_df['工号'].tolist())
    print(f"\n目标考勤表中现有的工号: {sorted(existing_employee_ids)}")
    
    # 检查入职人员表中的工号
    new_employee_ids = set(new_employee_data['工号'].tolist())
    print(f"入职人员表中的工号: {sorted(new_employee_ids)}")
    
    # 找出需要追加的记录（工号不重复的）
    ids_to_append = new_employee_ids - existing_employee_ids
    print(f"需要追加的工号: {sorted(ids_to_append)}")
    
    # 过滤出需要追加的记录
    records_to_append = new_employee_data[new_employee_data['工号'].isin(ids_to_append)]
    print(f"\n需要追加的记录数: {len(records_to_append)}")
    
    if not records_to_append.empty:
        print("需要追加的记录:")
        print(records_to_append)
        
        # 为追加的记录创建完整的行数据
        # 首先获取目标表的列结构
        target_columns = list(attendance_df.columns)
        
        # 创建新的DataFrame，包含所有目标表的列
        new_rows = pd.DataFrame(columns=target_columns)
        
        for idx, row in records_to_append.iterrows():
            new_row = pd.Series(index=target_columns, dtype=object)
            
            # 填充基本信息
            new_row['工号'] = row['工号']
            new_row['人员'] = row['姓名']
            new_row['员工类型'] = row['人员类型']
            new_row['银行卡号'] = row['银行卡号']
            
            # 设置默认值
            new_row['归属'] = '待分配'  # 默认归属
            new_row['月份'] = '2025-8'  # 8月
            new_row['年假'] = 0
            new_row['事假'] = 0
            new_row['病假'] = 0
            new_row['产检假'] = 0
            new_row['婚假'] = 0
            new_row['育儿假'] = 0
            new_row['丧假'] = 0
            new_row['陪产假'] = 0
            new_row['请假天数'] = 0
            new_row['应出勤天数'] = 0
            new_row['实际出勤天数'] = 0
            new_row['多休年假天数'] = 0
            new_row['实际发工资天数'] = 0
            new_row['备注'] = '8月新入职'
            new_row['图片'] = ''
            new_row['请确认'] = 0
            new_row['父记录'] = ''
            new_row['离职日期'] = ''
            new_row['离职类型'] = ''
            new_row['离职原因'] = ''
            new_row['离职备注'] = ''
            
            new_rows = pd.concat([new_rows, new_row.to_frame().T], ignore_index=True)
        
        # 追加到目标表
        attendance_df_updated = pd.concat([attendance_df, new_rows], ignore_index=True)
        
        print(f"\n追加后的记录数: {len(attendance_df_updated)}")
        print(f"新增记录数: {len(attendance_df_updated) - len(attendance_df)}")
        
        # 保存更新后的考勤表
        output_file = os.path.join(data_dir, f"考勤记录_考勤表_2025-8_追加新员工_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        attendance_df_updated.to_excel(output_file, index=False)
        
        print(f"\n更新后的考勤表已保存到: {output_file}")
        
        # 显示追加的记录
        print(f"\n追加的新员工记录:")
        appended_records = attendance_df_updated[attendance_df_updated['工号'].isin(ids_to_append)]
        print(appended_records[['工号', '人员', '员工类型', '归属', '银行卡号', '备注']].to_string(index=False))
        
        return output_file, len(records_to_append)
    else:
        print("没有需要追加的记录（所有工号都已存在）")
        return None, 0

if __name__ == "__main__":
    append_new_employees()











