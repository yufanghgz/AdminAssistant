#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
考勤汇总和统计工具
用于处理8月考勤数据，包括考勤记录、入职离职人员、请假数据等
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os
import warnings
warnings.filterwarnings('ignore')

class AttendanceProcessor:
    def __init__(self, data_dir):
        self.data_dir = data_dir
        self.attendance_data = None
        self.new_employees = None
        self.departed_employees = None
        self.leave_data = None
        
    def load_data(self):
        """加载所有考勤相关数据"""
        print("正在加载考勤数据...")
        
        # 加载考勤记录
        attendance_file = os.path.join(self.data_dir, "考勤记录_考勤表_2025-8.xlsx")
        if os.path.exists(attendance_file):
            self.attendance_data = pd.read_excel(attendance_file)
            print(f"✓ 考勤记录已加载: {len(self.attendance_data)} 条记录")
        else:
            print("✗ 考勤记录文件不存在")
            
        # 加载入职人员数据
        new_emp_file = os.path.join(self.data_dir, "8月入职人员.xlsx")
        if os.path.exists(new_emp_file):
            self.new_employees = pd.read_excel(new_emp_file)
            print(f"✓ 入职人员数据已加载: {len(self.new_employees)} 人")
        else:
            print("✗ 入职人员文件不存在")
            
        # 加载离职人员数据
        departed_emp_file = os.path.join(self.data_dir, "8月离职人员.xlsx")
        if os.path.exists(departed_emp_file):
            self.departed_employees = pd.read_excel(departed_emp_file)
            print(f"✓ 离职人员数据已加载: {len(self.departed_employees)} 人")
        else:
            print("✗ 离职人员文件不存在")
            
        # 加载请假数据
        leave_file = os.path.join(self.data_dir, "请假数据.xlsx")
        if os.path.exists(leave_file):
            self.leave_data = pd.read_excel(leave_file)
            print(f"✓ 请假数据已加载: {len(self.leave_data)} 条记录")
        else:
            print("✗ 请假数据文件不存在")
    
    def analyze_attendance_structure(self):
        """分析考勤数据结构"""
        if self.attendance_data is None:
            print("考勤数据未加载")
            return
            
        print("\n=== 考勤记录数据结构分析 ===")
        print(f"数据形状: {self.attendance_data.shape}")
        print(f"列名: {list(self.attendance_data.columns)}")
        print("\n前5行数据:")
        print(self.attendance_data.head())
        print("\n数据类型:")
        print(self.attendance_data.dtypes)
        print("\n缺失值统计:")
        print(self.attendance_data.isnull().sum())
    
    def analyze_employee_changes(self):
        """分析人员变动情况"""
        print("\n=== 人员变动分析 ===")
        
        if self.new_employees is not None:
            print(f"8月新入职人员: {len(self.new_employees)} 人")
            if not self.new_employees.empty:
                print("新入职人员列表:")
                print(self.new_employees)
        
        if self.departed_employees is not None:
            print(f"8月离职人员: {len(self.departed_employees)} 人")
            if not self.departed_employees.empty:
                print("离职人员列表:")
                print(self.departed_employees)
    
    def analyze_leave_data(self):
        """分析请假数据"""
        if self.leave_data is None:
            print("请假数据未加载")
            return
            
        print("\n=== 请假数据分析 ===")
        print(f"请假记录数: {len(self.leave_data)} 条")
        print(f"列名: {list(self.leave_data.columns)}")
        print("\n前5行数据:")
        print(self.leave_data.head())
    
    def generate_summary_report(self):
        """生成汇总报告"""
        print("\n=== 8月考勤汇总报告 ===")
        print(f"报告生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # 人员变动汇总
        if self.new_employees is not None and self.departed_employees is not None:
            net_change = len(self.new_employees) - len(self.departed_employees)
            print(f"\n人员变动情况:")
            print(f"- 新入职: {len(self.new_employees)} 人")
            print(f"- 离职: {len(self.departed_employees)} 人")
            print(f"- 净增加: {net_change} 人")
        
        # 考勤数据汇总
        if self.attendance_data is not None:
            print(f"\n考勤记录统计:")
            print(f"- 总记录数: {len(self.attendance_data)} 条")
            print(f"- 涉及人员数: {self.attendance_data.iloc[:, 0].nunique() if not self.attendance_data.empty else 0} 人")
        
        # 请假数据汇总
        if self.leave_data is not None:
            print(f"\n请假记录统计:")
            print(f"- 总请假记录: {len(self.leave_data)} 条")
    
    def export_summary(self, output_file=None):
        """导出汇总数据到Excel文件"""
        if output_file is None:
            output_file = os.path.join(self.data_dir, f"8月考勤汇总_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 人员变动汇总
            if self.new_employees is not None:
                self.new_employees.to_excel(writer, sheet_name='新入职人员', index=False)
            
            if self.departed_employees is not None:
                self.departed_employees.to_excel(writer, sheet_name='离职人员', index=False)
            
            # 考勤记录
            if self.attendance_data is not None:
                self.attendance_data.to_excel(writer, sheet_name='考勤记录', index=False)
            
            # 请假记录
            if self.leave_data is not None:
                self.leave_data.to_excel(writer, sheet_name='请假记录', index=False)
        
        print(f"\n汇总数据已导出到: {output_file}")
        return output_file

def main():
    # 数据目录
    data_dir = "/Users/heguangzhong/Documents/8月考勤"
    
    # 创建处理器实例
    processor = AttendanceProcessor(data_dir)
    
    # 加载数据
    processor.load_data()
    
    # 分析数据结构
    processor.analyze_attendance_structure()
    processor.analyze_employee_changes()
    processor.analyze_leave_data()
    
    # 生成汇总报告
    processor.generate_summary_report()
    
    # 导出汇总数据
    output_file = processor.export_summary()
    
    print(f"\n考勤汇总和统计完成！")
    print(f"输出文件: {output_file}")

if __name__ == "__main__":
    main()











