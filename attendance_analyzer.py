#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
考勤详细分析工具
专门处理考勤数据的深度分析和统计
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os
import warnings
warnings.filterwarnings('ignore')

class AttendanceAnalyzer:
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
            
        # 加载请假数据 - 尝试不同的读取方式
        leave_file = os.path.join(self.data_dir, "请假数据.xlsx")
        if os.path.exists(leave_file):
            try:
                # 尝试读取所有工作表
                excel_file = pd.ExcelFile(leave_file)
                print(f"请假数据文件包含工作表: {excel_file.sheet_names}")
                
                # 尝试读取第一个有效的工作表
                for sheet_name in excel_file.sheet_names:
                    try:
                        df = pd.read_excel(leave_file, sheet_name=sheet_name, header=1)  # 跳过第一行标题
                        if len(df) > 0 and not df.empty:
                            self.leave_data = df
                            print(f"✓ 请假数据已加载 (工作表: {sheet_name}): {len(self.leave_data)} 条记录")
                            break
                    except Exception as e:
                        print(f"读取工作表 {sheet_name} 失败: {e}")
                        continue
                        
                if self.leave_data is None:
                    print("✗ 无法读取请假数据")
            except Exception as e:
                print(f"✗ 请假数据文件读取失败: {e}")
        else:
            print("✗ 请假数据文件不存在")
    
    def analyze_attendance_details(self):
        """详细分析考勤数据"""
        if self.attendance_data is None:
            print("考勤数据未加载")
            return
            
        print("\n=== 考勤详细分析 ===")
        
        # 基本统计
        total_employees = len(self.attendance_data)
        print(f"总员工数: {total_employees}")
        
        # 按部门统计
        if '归属' in self.attendance_data.columns:
            dept_stats = self.attendance_data['归属'].value_counts()
            print(f"\n按部门统计:")
            for dept, count in dept_stats.items():
                print(f"  {dept}: {count} 人")
        
        # 出勤情况统计
        if '实际出勤天数' in self.attendance_data.columns:
            attendance_stats = self.attendance_data['实际出勤天数'].describe()
            print(f"\n实际出勤天数统计:")
            print(f"  平均出勤天数: {attendance_stats['mean']:.1f} 天")
            print(f"  最少出勤天数: {attendance_stats['min']:.1f} 天")
            print(f"  最多出勤天数: {attendance_stats['max']:.1f} 天")
            
            # 出勤率分析
            if '应出勤天数' in self.attendance_data.columns:
                self.attendance_data['出勤率'] = (self.attendance_data['实际出勤天数'] / 
                                            self.attendance_data['应出勤天数'] * 100).round(2)
                low_attendance = self.attendance_data[self.attendance_data['出勤率'] < 80]
                print(f"  出勤率低于80%的员工: {len(low_attendance)} 人")
        
        # 请假情况统计
        leave_columns = ['年假', '事假', '病假', '产检假', '婚假', '育儿假', '丧假', '陪产假']
        print(f"\n请假情况统计:")
        for col in leave_columns:
            if col in self.attendance_data.columns:
                total_leave = self.attendance_data[col].sum()
                employees_with_leave = (self.attendance_data[col] > 0).sum()
                print(f"  {col}: {total_leave:.1f} 天 (涉及 {employees_with_leave} 人)")
    
    def analyze_employee_changes_detailed(self):
        """详细分析人员变动"""
        print("\n=== 人员变动详细分析 ===")
        
        if self.new_employees is not None and not self.new_employees.empty:
            print(f"\n新入职人员详情:")
            print(f"总人数: {len(self.new_employees)} 人")
            
            # 按人员类型统计
            if '人员类型' in self.new_employees.columns:
                type_stats = self.new_employees['人员类型'].value_counts()
                print("按人员类型:")
                for emp_type, count in type_stats.items():
                    print(f"  {emp_type}: {count} 人")
            
            # 按学历统计
            if '最高学历' in self.new_employees.columns:
                edu_stats = self.new_employees['最高学历'].value_counts()
                print("按学历:")
                for edu, count in edu_stats.items():
                    print(f"  {edu}: {count} 人")
        
        if self.departed_employees is not None and not self.departed_employees.empty:
            print(f"\n离职人员详情:")
            print(f"总人数: {len(self.departed_employees)} 人")
            
            # 按人员类型统计
            if '人员类型' in self.departed_employees.columns:
                type_stats = self.departed_employees['人员类型'].value_counts()
                print("按人员类型:")
                for emp_type, count in type_stats.items():
                    print(f"  {emp_type}: {count} 人")
            
            # 按部门统计
            if '部门' in self.departed_employees.columns:
                dept_stats = self.departed_employees['部门'].value_counts()
                print("按部门:")
                for dept, count in dept_stats.items():
                    print(f"  {dept}: {count} 人")
    
    def analyze_leave_data_detailed(self):
        """详细分析请假数据"""
        if self.leave_data is None:
            print("请假数据未加载")
            return
            
        print("\n=== 请假数据详细分析 ===")
        print(f"请假记录数: {len(self.leave_data)} 条")
        
        # 显示列名
        print(f"数据列: {list(self.leave_data.columns)}")
        
        # 尝试找到关键列
        key_columns = ['姓名', '人员', '员工', '申请人', '请假类型', '类型', '开始时间', '结束时间', '天数', '状态']
        found_columns = []
        for col in self.leave_data.columns:
            if any(key in str(col) for key in key_columns):
                found_columns.append(col)
        
        if found_columns:
            print(f"关键列: {found_columns}")
            print("\n前5行关键数据:")
            print(self.leave_data[found_columns].head())
        else:
            print("前5行原始数据:")
            print(self.leave_data.head())
    
    def generate_comprehensive_report(self):
        """生成综合报告"""
        print("\n" + "="*50)
        print("8月考勤综合统计报告")
        print("="*50)
        print(f"报告生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # 人员变动汇总
        if self.new_employees is not None and self.departed_employees is not None:
            net_change = len(self.new_employees) - len(self.departed_employees)
            print(f"\n📊 人员变动情况:")
            print(f"  • 新入职: {len(self.new_employees)} 人")
            print(f"  • 离职: {len(self.departed_employees)} 人")
            print(f"  • 净增加: {net_change} 人")
            print(f"  • 人员变动率: {((len(self.new_employees) + len(self.departed_employees)) / 72 * 100):.1f}%")
        
        # 考勤数据汇总
        if self.attendance_data is not None:
            print(f"\n📈 考勤数据汇总:")
            print(f"  • 总员工数: {len(self.attendance_data)} 人")
            
            if '实际出勤天数' in self.attendance_data.columns and '应出勤天数' in self.attendance_data.columns:
                total_actual = self.attendance_data['实际出勤天数'].sum()
                total_expected = self.attendance_data['应出勤天数'].sum()
                overall_attendance_rate = (total_actual / total_expected * 100) if total_expected > 0 else 0
                print(f"  • 整体出勤率: {overall_attendance_rate:.1f}%")
                print(f"  • 总应出勤天数: {total_expected:.0f} 天")
                print(f"  • 总实际出勤天数: {total_actual:.0f} 天")
        
        # 请假数据汇总
        if self.leave_data is not None:
            print(f"\n📋 请假数据汇总:")
            print(f"  • 总请假记录: {len(self.leave_data)} 条")
        
        print(f"\n✅ 报告生成完成！")
    
    def export_detailed_report(self, output_file=None):
        """导出详细报告到Excel文件"""
        if output_file is None:
            output_file = os.path.join(self.data_dir, f"8月考勤详细分析_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 人员变动汇总
            if self.new_employees is not None:
                self.new_employees.to_excel(writer, sheet_name='新入职人员', index=False)
            
            if self.departed_employees is not None:
                self.departed_employees.to_excel(writer, sheet_name='离职人员', index=False)
            
            # 考勤记录
            if self.attendance_data is not None:
                self.attendance_data.to_excel(writer, sheet_name='考勤记录', index=False)
                
                # 按部门统计
                if '归属' in self.attendance_data.columns:
                    dept_stats = self.attendance_data['归属'].value_counts().reset_index()
                    dept_stats.columns = ['部门', '人数']
                    dept_stats.to_excel(writer, sheet_name='部门统计', index=False)
                
                # 出勤率统计
                if '实际出勤天数' in self.attendance_data.columns and '应出勤天数' in self.attendance_data.columns:
                    attendance_analysis = self.attendance_data[['工号', '人员', '归属', '应出勤天数', '实际出勤天数']].copy()
                    attendance_analysis['出勤率'] = (attendance_analysis['实际出勤天数'] / 
                                                attendance_analysis['应出勤天数'] * 100).round(2)
                    attendance_analysis.to_excel(writer, sheet_name='出勤率分析', index=False)
            
            # 请假记录
            if self.leave_data is not None:
                self.leave_data.to_excel(writer, sheet_name='请假记录', index=False)
        
        print(f"\n详细报告已导出到: {output_file}")
        return output_file

def main():
    # 数据目录
    data_dir = "/Users/heguangzhong/Documents/8月考勤"
    
    # 创建分析器实例
    analyzer = AttendanceAnalyzer(data_dir)
    
    # 加载数据
    analyzer.load_data()
    
    # 详细分析
    analyzer.analyze_attendance_details()
    analyzer.analyze_employee_changes_detailed()
    analyzer.analyze_leave_data_detailed()
    
    # 生成综合报告
    analyzer.generate_comprehensive_report()
    
    # 导出详细报告
    output_file = analyzer.export_detailed_report()
    
    print(f"\n🎉 考勤详细分析完成！")
    print(f"📁 输出文件: {output_file}")

if __name__ == "__main__":
    main()






