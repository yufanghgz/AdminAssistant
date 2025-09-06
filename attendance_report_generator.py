#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
考勤汇总报告生成器
生成完整的考勤统计报告，包含数据分析和图表
"""

import pandas as pd
import numpy as np
from datetime import datetime
import os
import warnings
warnings.filterwarnings('ignore')

class AttendanceReportGenerator:
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
        
        # 加载入职人员数据
        new_emp_file = os.path.join(self.data_dir, "8月入职人员.xlsx")
        if os.path.exists(new_emp_file):
            self.new_employees = pd.read_excel(new_emp_file)
            print(f"✓ 入职人员数据已加载: {len(self.new_employees)} 人")
        
        # 加载离职人员数据
        departed_emp_file = os.path.join(self.data_dir, "8月离职人员.xlsx")
        if os.path.exists(departed_emp_file):
            self.departed_employees = pd.read_excel(departed_emp_file)
            print(f"✓ 离职人员数据已加载: {len(self.departed_employees)} 人")
        
        # 加载请假数据
        leave_file = os.path.join(self.data_dir, "请假数据.xlsx")
        if os.path.exists(leave_file):
            try:
                excel_file = pd.ExcelFile(leave_file)
                for sheet_name in excel_file.sheet_names:
                    try:
                        df = pd.read_excel(leave_file, sheet_name=sheet_name, header=1)
                        if len(df) > 0 and not df.empty:
                            self.leave_data = df
                            print(f"✓ 请假数据已加载: {len(self.leave_data)} 条记录")
                            break
                    except:
                        continue
            except:
                print("✗ 请假数据文件读取失败")
    
    def generate_text_report(self):
        """生成文本格式的详细报告"""
        report = []
        report.append("=" * 80)
        report.append("8月考勤数据汇总统计报告")
        report.append("=" * 80)
        report.append(f"报告生成时间: {datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}")
        report.append("")
        
        # 1. 人员变动情况
        report.append("一、人员变动情况")
        report.append("-" * 40)
        if self.new_employees is not None and self.departed_employees is not None:
            new_count = len(self.new_employees)
            departed_count = len(self.departed_employees)
            net_change = new_count - departed_count
            change_rate = ((new_count + departed_count) / 72 * 100) if 72 > 0 else 0
            
            report.append(f"新入职人员: {new_count} 人")
            report.append(f"离职人员: {departed_count} 人")
            report.append(f"净增加: {net_change} 人")
            report.append(f"人员变动率: {change_rate:.1f}%")
            
            # 新入职人员详情
            if not self.new_employees.empty:
                report.append("")
                report.append("新入职人员详情:")
                if '人员类型' in self.new_employees.columns:
                    type_stats = self.new_employees['人员类型'].value_counts()
                    for emp_type, count in type_stats.items():
                        report.append(f"  {emp_type}: {count} 人")
                
                if '最高学历' in self.new_employees.columns:
                    edu_stats = self.new_employees['最高学历'].value_counts()
                    report.append("按学历分布:")
                    for edu, count in edu_stats.items():
                        report.append(f"  {edu}: {count} 人")
            
            # 离职人员详情
            if not self.departed_employees.empty:
                report.append("")
                report.append("离职人员详情:")
                if '人员类型' in self.departed_employees.columns:
                    type_stats = self.departed_employees['人员类型'].value_counts()
                    for emp_type, count in type_stats.items():
                        report.append(f"  {emp_type}: {count} 人")
                
                if '部门' in self.departed_employees.columns:
                    dept_stats = self.departed_employees['部门'].value_counts()
                    report.append("按部门分布:")
                    for dept, count in dept_stats.items():
                        report.append(f"  {dept}: {count} 人")
        
        report.append("")
        
        # 2. 考勤数据统计
        report.append("二、考勤数据统计")
        report.append("-" * 40)
        if self.attendance_data is not None:
            total_employees = len(self.attendance_data)
            report.append(f"总员工数: {total_employees} 人")
            
            # 按部门统计
            if '归属' in self.attendance_data.columns:
                dept_stats = self.attendance_data['归属'].value_counts()
                report.append("按部门分布:")
                for dept, count in dept_stats.items():
                    percentage = (count / total_employees * 100) if total_employees > 0 else 0
                    report.append(f"  {dept}: {count} 人 ({percentage:.1f}%)")
            
            # 出勤情况统计
            if '实际出勤天数' in self.attendance_data.columns and '应出勤天数' in self.attendance_data.columns:
                total_actual = self.attendance_data['实际出勤天数'].sum()
                total_expected = self.attendance_data['应出勤天数'].sum()
                overall_attendance_rate = (total_actual / total_expected * 100) if total_expected > 0 else 0
                
                report.append("")
                report.append("出勤情况:")
                report.append(f"  总应出勤天数: {total_expected:.0f} 天")
                report.append(f"  总实际出勤天数: {total_actual:.0f} 天")
                report.append(f"  整体出勤率: {overall_attendance_rate:.1f}%")
                
                # 出勤率分布
                attendance_rate = (self.attendance_data['实际出勤天数'] / 
                                 self.attendance_data['应出勤天数'] * 100).round(2)
                low_attendance = (attendance_rate < 80).sum()
                high_attendance = (attendance_rate >= 95).sum()
                
                report.append(f"  出勤率低于80%: {low_attendance} 人")
                report.append(f"  出勤率95%以上: {high_attendance} 人")
                
                # 各部门出勤率
                if '归属' in self.attendance_data.columns:
                    report.append("")
                    report.append("各部门出勤率:")
                    dept_attendance = self.attendance_data.groupby('归属').apply(
                        lambda x: (x['实际出勤天数'].sum() / x['应出勤天数'].sum() * 100).round(2)
                    )
                    for dept, rate in dept_attendance.items():
                        report.append(f"  {dept}: {rate:.1f}%")
        
        report.append("")
        
        # 3. 请假情况统计
        report.append("三、请假情况统计")
        report.append("-" * 40)
        if self.attendance_data is not None:
            leave_columns = ['年假', '事假', '病假', '产检假', '婚假', '育儿假', '丧假', '陪产假']
            total_leave_days = 0
            total_leave_people = 0
            
            report.append("各类请假统计:")
            for col in leave_columns:
                if col in self.attendance_data.columns:
                    total_days = self.attendance_data[col].sum()
                    people_count = (self.attendance_data[col] > 0).sum()
                    if total_days > 0:
                        report.append(f"  {col}: {total_days:.1f} 天 (涉及 {people_count} 人)")
                        total_leave_days += total_days
                        total_leave_people += people_count
            
            report.append("")
            report.append(f"请假总计: {total_leave_days:.1f} 天")
            report.append(f"涉及人员: {total_leave_people} 人")
        
        # 请假申请数据
        if self.leave_data is not None:
            report.append("")
            report.append("请假申请记录:")
            report.append(f"  总申请记录: {len(self.leave_data)} 条")
            
            if '申请状态' in self.leave_data.columns:
                status_stats = self.leave_data['申请状态'].value_counts()
                report.append("按状态分布:")
                for status, count in status_stats.items():
                    report.append(f"    {status}: {count} 条")
            
            if '假期类型' in self.leave_data.columns:
                leave_type_stats = self.leave_data['假期类型'].value_counts()
                report.append("按假期类型分布:")
                for leave_type, count in leave_type_stats.items():
                    report.append(f"    {leave_type}: {count} 条")
        
        report.append("")
        
        # 4. 数据质量分析
        report.append("四、数据质量分析")
        report.append("-" * 40)
        if self.attendance_data is not None:
            missing_data = self.attendance_data.isnull().sum()
            report.append("缺失数据统计:")
            for col, missing_count in missing_data.items():
                if missing_count > 0:
                    percentage = (missing_count / len(self.attendance_data) * 100)
                    report.append(f"  {col}: {missing_count} 条 ({percentage:.1f}%)")
        
        report.append("")
        
        # 5. 总结和建议
        report.append("五、总结和建议")
        report.append("-" * 40)
        
        if self.attendance_data is not None and '实际出勤天数' in self.attendance_data.columns:
            attendance_rate = (self.attendance_data['实际出勤天数'] / 
                             self.attendance_data['应出勤天数'] * 100).round(2)
            avg_attendance = attendance_rate.mean()
            
            if avg_attendance >= 95:
                report.append("✓ 整体出勤情况良好，出勤率较高")
            elif avg_attendance >= 90:
                report.append("⚠ 整体出勤情况一般，建议关注出勤率较低的人员")
            else:
                report.append("⚠ 整体出勤情况需要改善，建议加强考勤管理")
        
        if self.new_employees is not None and self.departed_employees is not None:
            change_rate = ((len(self.new_employees) + len(self.departed_employees)) / 72 * 100) if 72 > 0 else 0
            if change_rate > 20:
                report.append("⚠ 人员变动较为频繁，建议关注人员稳定性")
            else:
                report.append("✓ 人员变动在正常范围内")
        
        report.append("")
        report.append("建议:")
        report.append("1. 定期监控出勤率，及时处理异常情况")
        report.append("2. 关注人员变动趋势，做好人员规划")
        report.append("3. 完善请假管理制度，提高数据质量")
        report.append("4. 定期生成考勤报告，为管理决策提供数据支持")
        
        report.append("")
        report.append("=" * 80)
        report.append("报告结束")
        report.append("=" * 80)
        
        return "\n".join(report)
    
    def export_comprehensive_report(self, output_file=None):
        """导出综合报告到Excel文件"""
        if output_file is None:
            output_file = os.path.join(self.data_dir, f"8月考勤综合报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 1. 报告摘要
            summary_data = {
                '统计项目': [
                    '总员工数',
                    '新入职人数',
                    '离职人数',
                    '净增加人数',
                    '整体出勤率',
                    '请假总天数',
                    '请假申请记录数'
                ],
                '数值': [
                    len(self.attendance_data) if self.attendance_data is not None else 0,
                    len(self.new_employees) if self.new_employees is not None else 0,
                    len(self.departed_employees) if self.departed_employees is not None else 0,
                    (len(self.new_employees) - len(self.departed_employees)) if (self.new_employees is not None and self.departed_employees is not None) else 0,
                    f"{((self.attendance_data['实际出勤天数'].sum() / self.attendance_data['应出勤天数'].sum() * 100) if (self.attendance_data is not None and '实际出勤天数' in self.attendance_data.columns and '应出勤天数' in self.attendance_data.columns) else 0):.1f}%",
                    f"{sum([self.attendance_data[col].sum() for col in ['年假', '事假', '病假', '产检假', '婚假', '育儿假', '丧假', '陪产假'] if col in self.attendance_data.columns]) if self.attendance_data is not None else 0:.1f} 天",
                    len(self.leave_data) if self.leave_data is not None else 0
                ]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='报告摘要', index=False)
            
            # 2. 人员变动详情
            if self.new_employees is not None:
                self.new_employees.to_excel(writer, sheet_name='新入职人员', index=False)
            
            if self.departed_employees is not None:
                self.departed_employees.to_excel(writer, sheet_name='离职人员', index=False)
            
            # 3. 考勤记录
            if self.attendance_data is not None:
                self.attendance_data.to_excel(writer, sheet_name='考勤记录', index=False)
                
                # 部门统计
                if '归属' in self.attendance_data.columns:
                    dept_stats = self.attendance_data['归属'].value_counts().reset_index()
                    dept_stats.columns = ['部门', '人数']
                    dept_stats['占比'] = (dept_stats['人数'] / dept_stats['人数'].sum() * 100).round(2)
                    dept_stats.to_excel(writer, sheet_name='部门统计', index=False)
                
                # 出勤率分析
                if '实际出勤天数' in self.attendance_data.columns and '应出勤天数' in self.attendance_data.columns:
                    attendance_analysis = self.attendance_data[['工号', '人员', '归属', '应出勤天数', '实际出勤天数']].copy()
                    attendance_analysis['出勤率'] = (attendance_analysis['实际出勤天数'] / 
                                                attendance_analysis['应出勤天数'] * 100).round(2)
                    attendance_analysis = attendance_analysis.sort_values('出勤率')
                    attendance_analysis.to_excel(writer, sheet_name='出勤率分析', index=False)
                
                # 请假统计
                leave_columns = ['年假', '事假', '病假', '产检假', '婚假', '育儿假', '丧假', '陪产假']
                leave_stats = []
                for col in leave_columns:
                    if col in self.attendance_data.columns:
                        total_days = self.attendance_data[col].sum()
                        people_count = (self.attendance_data[col] > 0).sum()
                        leave_stats.append({
                            '请假类型': col,
                            '总天数': total_days,
                            '涉及人数': people_count,
                            '平均天数': (total_days / people_count) if people_count > 0 else 0
                        })
                
                if leave_stats:
                    leave_stats_df = pd.DataFrame(leave_stats)
                    leave_stats_df.to_excel(writer, sheet_name='请假统计', index=False)
            
            # 4. 请假申请记录
            if self.leave_data is not None:
                self.leave_data.to_excel(writer, sheet_name='请假申请记录', index=False)
        
        print(f"\n综合报告已导出到: {output_file}")
        return output_file
    
    def save_text_report(self, output_file=None):
        """保存文本报告"""
        if output_file is None:
            output_file = os.path.join(self.data_dir, f"8月考勤报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
        
        report_text = self.generate_text_report()
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(report_text)
        
        print(f"文本报告已保存到: {output_file}")
        return output_file

def main():
    # 数据目录
    data_dir = "/Users/heguangzhong/Documents/8月考勤"
    
    # 创建报告生成器实例
    generator = AttendanceReportGenerator(data_dir)
    
    # 加载数据
    generator.load_data()
    
    # 生成并保存文本报告
    text_report_file = generator.save_text_report()
    
    # 生成并保存Excel综合报告
    excel_report_file = generator.export_comprehensive_report()
    
    # 显示报告摘要
    print("\n" + "="*60)
    print("8月考勤汇总统计报告摘要")
    print("="*60)
    
    if generator.attendance_data is not None:
        total_employees = len(generator.attendance_data)
        print(f"总员工数: {total_employees} 人")
        
        if '实际出勤天数' in generator.attendance_data.columns and '应出勤天数' in generator.attendance_data.columns:
            total_actual = generator.attendance_data['实际出勤天数'].sum()
            total_expected = generator.attendance_data['应出勤天数'].sum()
            overall_attendance_rate = (total_actual / total_expected * 100) if total_expected > 0 else 0
            print(f"整体出勤率: {overall_attendance_rate:.1f}%")
    
    if generator.new_employees is not None and generator.departed_employees is not None:
        new_count = len(generator.new_employees)
        departed_count = len(generator.departed_employees)
        net_change = new_count - departed_count
        print(f"人员变动: 新入职 {new_count} 人，离职 {departed_count} 人，净增加 {net_change} 人")
    
    if generator.leave_data is not None:
        print(f"请假申请记录: {len(generator.leave_data)} 条")
    
    print(f"\n📄 文本报告: {os.path.basename(text_report_file)}")
    print(f"📊 Excel报告: {os.path.basename(excel_report_file)}")
    print(f"\n🎉 考勤汇总和统计工作完成！")

if __name__ == "__main__":
    main()






