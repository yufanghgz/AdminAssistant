#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
考勤数据可视化工具
生成考勤统计图表和可视化报告
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime
import os
import warnings
warnings.filterwarnings('ignore')

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei', 'Arial Unicode MS', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

class AttendanceVisualizer:
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
    
    def create_department_chart(self):
        """创建部门人员分布图"""
        if self.attendance_data is None or '归属' not in self.attendance_data.columns:
            return None
            
        plt.figure(figsize=(10, 6))
        dept_counts = self.attendance_data['归属'].value_counts()
        
        # 创建饼图
        plt.subplot(1, 2, 1)
        colors = ['#ff9999', '#66b3ff', '#99ff99']
        wedges, texts, autotexts = plt.pie(dept_counts.values, labels=dept_counts.index, 
                                          autopct='%1.1f%%', colors=colors, startangle=90)
        plt.title('部门人员分布', fontsize=14, fontweight='bold')
        
        # 创建柱状图
        plt.subplot(1, 2, 2)
        bars = plt.bar(dept_counts.index, dept_counts.values, color=colors)
        plt.title('部门人员数量', fontsize=14, fontweight='bold')
        plt.xlabel('部门')
        plt.ylabel('人数')
        
        # 在柱状图上添加数值标签
        for bar in bars:
            height = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                    f'{int(height)}', ha='center', va='bottom')
        
        plt.tight_layout()
        return plt.gcf()
    
    def create_attendance_chart(self):
        """创建出勤情况图表"""
        if self.attendance_data is None or '实际出勤天数' not in self.attendance_data.columns:
            return None
            
        plt.figure(figsize=(15, 5))
        
        # 出勤天数分布直方图
        plt.subplot(1, 3, 1)
        plt.hist(self.attendance_data['实际出勤天数'], bins=20, alpha=0.7, color='skyblue', edgecolor='black')
        plt.title('出勤天数分布', fontsize=12, fontweight='bold')
        plt.xlabel('出勤天数')
        plt.ylabel('人数')
        plt.grid(True, alpha=0.3)
        
        # 出勤率分布
        if '应出勤天数' in self.attendance_data.columns:
            attendance_rate = (self.attendance_data['实际出勤天数'] / 
                             self.attendance_data['应出勤天数'] * 100).round(2)
            
            plt.subplot(1, 3, 2)
            plt.hist(attendance_rate, bins=20, alpha=0.7, color='lightgreen', edgecolor='black')
            plt.title('出勤率分布', fontsize=12, fontweight='bold')
            plt.xlabel('出勤率 (%)')
            plt.ylabel('人数')
            plt.axvline(x=80, color='red', linestyle='--', label='80%基准线')
            plt.legend()
            plt.grid(True, alpha=0.3)
        
        # 部门出勤率对比
        if '归属' in self.attendance_data.columns and '应出勤天数' in self.attendance_data.columns:
            plt.subplot(1, 3, 3)
            dept_attendance = self.attendance_data.groupby('归属').apply(
                lambda x: (x['实际出勤天数'].sum() / x['应出勤天数'].sum() * 100).round(2)
            )
            
            bars = plt.bar(dept_attendance.index, dept_attendance.values, 
                          color=['#ff9999', '#66b3ff', '#99ff99'])
            plt.title('各部门出勤率对比', fontsize=12, fontweight='bold')
            plt.xlabel('部门')
            plt.ylabel('出勤率 (%)')
            plt.ylim(0, 100)
            
            # 添加数值标签
            for bar in bars:
                height = bar.get_height()
                plt.text(bar.get_x() + bar.get_width()/2., height + 1,
                        f'{height:.1f}%', ha='center', va='bottom')
        
        plt.tight_layout()
        return plt.gcf()
    
    def create_leave_chart(self):
        """创建请假情况图表"""
        if self.attendance_data is None:
            return None
            
        leave_columns = ['年假', '事假', '病假', '产检假', '婚假', '育儿假', '丧假', '陪产假']
        leave_data = {}
        
        for col in leave_columns:
            if col in self.attendance_data.columns:
                total_days = self.attendance_data[col].sum()
                if total_days > 0:
                    leave_data[col] = total_days
        
        if not leave_data:
            return None
            
        plt.figure(figsize=(12, 5))
        
        # 请假天数统计
        plt.subplot(1, 2, 1)
        leave_types = list(leave_data.keys())
        leave_days = list(leave_data.values())
        
        bars = plt.bar(leave_types, leave_days, color='lightcoral', alpha=0.8)
        plt.title('各类请假天数统计', fontsize=12, fontweight='bold')
        plt.xlabel('请假类型')
        plt.ylabel('天数')
        plt.xticks(rotation=45)
        
        # 添加数值标签
        for bar in bars:
            height = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                    f'{height:.1f}', ha='center', va='bottom')
        
        # 请假人数统计
        plt.subplot(1, 2, 2)
        leave_people = {}
        for col in leave_columns:
            if col in self.attendance_data.columns:
                people_count = (self.attendance_data[col] > 0).sum()
                if people_count > 0:
                    leave_people[col] = people_count
        
        if leave_people:
            people_types = list(leave_people.keys())
            people_counts = list(leave_people.values())
            
            bars = plt.bar(people_types, people_counts, color='lightblue', alpha=0.8)
            plt.title('各类请假人数统计', fontsize=12, fontweight='bold')
            plt.xlabel('请假类型')
            plt.ylabel('人数')
            plt.xticks(rotation=45)
            
            # 添加数值标签
            for bar in bars:
                height = bar.get_height()
                plt.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                        f'{int(height)}', ha='center', va='bottom')
        
        plt.tight_layout()
        return plt.gcf()
    
    def create_employee_changes_chart(self):
        """创建人员变动图表"""
        if self.new_employees is None and self.departed_employees is None:
            return None
            
        plt.figure(figsize=(12, 5))
        
        # 人员变动对比
        plt.subplot(1, 2, 1)
        categories = ['新入职', '离职']
        counts = [
            len(self.new_employees) if self.new_employees is not None else 0,
            len(self.departed_employees) if self.departed_employees is not None else 0
        ]
        colors = ['lightgreen', 'lightcoral']
        
        bars = plt.bar(categories, counts, color=colors, alpha=0.8)
        plt.title('8月人员变动情况', fontsize=12, fontweight='bold')
        plt.ylabel('人数')
        
        # 添加数值标签
        for bar in bars:
            height = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                    f'{int(height)}', ha='center', va='bottom')
        
        # 新入职人员类型分布
        if self.new_employees is not None and '人员类型' in self.new_employees.columns:
            plt.subplot(1, 2, 2)
            type_counts = self.new_employees['人员类型'].value_counts()
            
            bars = plt.bar(type_counts.index, type_counts.values, 
                          color=['lightblue', 'lightyellow'], alpha=0.8)
            plt.title('新入职人员类型分布', fontsize=12, fontweight='bold')
            plt.xlabel('人员类型')
            plt.ylabel('人数')
            
            # 添加数值标签
            for bar in bars:
                height = bar.get_height()
                plt.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                        f'{int(height)}', ha='center', va='bottom')
        
        plt.tight_layout()
        return plt.gcf()
    
    def create_comprehensive_dashboard(self):
        """创建综合仪表板"""
        fig = plt.figure(figsize=(20, 12))
        
        # 设置整体标题
        fig.suptitle('8月考勤数据综合仪表板', fontsize=20, fontweight='bold', y=0.95)
        
        # 1. 部门分布 (左上)
        if self.attendance_data is not None and '归属' in self.attendance_data.columns:
            plt.subplot(2, 3, 1)
            dept_counts = self.attendance_data['归属'].value_counts()
            colors = ['#ff9999', '#66b3ff', '#99ff99']
            plt.pie(dept_counts.values, labels=dept_counts.index, autopct='%1.1f%%', 
                   colors=colors, startangle=90)
            plt.title('部门人员分布', fontsize=14, fontweight='bold')
        
        # 2. 出勤率分布 (右上)
        if (self.attendance_data is not None and '实际出勤天数' in self.attendance_data.columns 
            and '应出勤天数' in self.attendance_data.columns):
            plt.subplot(2, 3, 2)
            attendance_rate = (self.attendance_data['实际出勤天数'] / 
                             self.attendance_data['应出勤天数'] * 100).round(2)
            plt.hist(attendance_rate, bins=15, alpha=0.7, color='lightgreen', edgecolor='black')
            plt.title('出勤率分布', fontsize=14, fontweight='bold')
            plt.xlabel('出勤率 (%)')
            plt.ylabel('人数')
            plt.axvline(x=80, color='red', linestyle='--', alpha=0.7)
        
        # 3. 请假类型统计 (中左)
        if self.attendance_data is not None:
            plt.subplot(2, 3, 3)
            leave_columns = ['年假', '事假', '病假', '产检假', '婚假', '育儿假', '丧假', '陪产假']
            leave_data = {}
            for col in leave_columns:
                if col in self.attendance_data.columns:
                    total_days = self.attendance_data[col].sum()
                    if total_days > 0:
                        leave_data[col] = total_days
            
            if leave_data:
                leave_types = list(leave_data.keys())
                leave_days = list(leave_data.values())
                bars = plt.bar(leave_types, leave_days, color='lightcoral', alpha=0.8)
                plt.title('请假天数统计', fontsize=14, fontweight='bold')
                plt.xlabel('请假类型')
                plt.ylabel('天数')
                plt.xticks(rotation=45)
        
        # 4. 人员变动 (中右)
        plt.subplot(2, 3, 4)
        categories = ['新入职', '离职']
        counts = [
            len(self.new_employees) if self.new_employees is not None else 0,
            len(self.departed_employees) if self.departed_employees is not None else 0
        ]
        colors = ['lightgreen', 'lightcoral']
        bars = plt.bar(categories, counts, color=colors, alpha=0.8)
        plt.title('人员变动情况', fontsize=14, fontweight='bold')
        plt.ylabel('人数')
        
        # 5. 部门出勤率对比 (下左)
        if (self.attendance_data is not None and '归属' in self.attendance_data.columns 
            and '实际出勤天数' in self.attendance_data.columns and '应出勤天数' in self.attendance_data.columns):
            plt.subplot(2, 3, 5)
            dept_attendance = self.attendance_data.groupby('归属').apply(
                lambda x: (x['实际出勤天数'].sum() / x['应出勤天数'].sum() * 100).round(2)
            )
            bars = plt.bar(dept_attendance.index, dept_attendance.values, 
                          color=['#ff9999', '#66b3ff', '#99ff99'])
            plt.title('各部门出勤率对比', fontsize=14, fontweight='bold')
            plt.xlabel('部门')
            plt.ylabel('出勤率 (%)')
            plt.ylim(0, 100)
        
        # 6. 考勤统计摘要 (下右)
        plt.subplot(2, 3, 6)
        plt.axis('off')
        
        # 计算统计数据
        total_employees = len(self.attendance_data) if self.attendance_data is not None else 0
        new_employees = len(self.new_employees) if self.new_employees is not None else 0
        departed_employees = len(self.departed_employees) if self.departed_employees is not None else 0
        net_change = new_employees - departed_employees
        
        if (self.attendance_data is not None and '实际出勤天数' in self.attendance_data.columns 
            and '应出勤天数' in self.attendance_data.columns):
            total_actual = self.attendance_data['实际出勤天数'].sum()
            total_expected = self.attendance_data['应出勤天数'].sum()
            overall_attendance_rate = (total_actual / total_expected * 100) if total_expected > 0 else 0
        else:
            overall_attendance_rate = 0
        
        # 显示统计摘要
        summary_text = f"""
        考勤统计摘要
        
        总员工数: {total_employees} 人
        新入职: {new_employees} 人
        离职: {departed_employees} 人
        净增加: {net_change} 人
        
        整体出勤率: {overall_attendance_rate:.1f}%
        
        报告生成时间:
        {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
        """
        
        plt.text(0.1, 0.9, summary_text, transform=plt.gca().transAxes, 
                fontsize=12, verticalalignment='top', fontfamily='monospace',
                bbox=dict(boxstyle="round,pad=0.3", facecolor="lightgray", alpha=0.5))
        
        plt.tight_layout()
        return fig
    
    def save_all_charts(self, output_dir=None):
        """保存所有图表"""
        if output_dir is None:
            output_dir = self.data_dir
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        # 保存各个图表
        charts = [
            ('department_distribution', self.create_department_chart),
            ('attendance_analysis', self.create_attendance_chart),
            ('leave_analysis', self.create_leave_chart),
            ('employee_changes', self.create_employee_changes_chart)
        ]
        
        saved_files = []
        for chart_name, chart_func in charts:
            try:
                fig = chart_func()
                if fig is not None:
                    filename = f"8月考勤_{chart_name}_{timestamp}.png"
                    filepath = os.path.join(output_dir, filename)
                    fig.savefig(filepath, dpi=300, bbox_inches='tight')
                    plt.close(fig)
                    saved_files.append(filepath)
                    print(f"✓ 图表已保存: {filename}")
            except Exception as e:
                print(f"✗ 保存图表 {chart_name} 失败: {e}")
        
        # 保存综合仪表板
        try:
            dashboard = self.create_comprehensive_dashboard()
            dashboard_filename = f"8月考勤综合仪表板_{timestamp}.png"
            dashboard_filepath = os.path.join(output_dir, dashboard_filename)
            dashboard.savefig(dashboard_filepath, dpi=300, bbox_inches='tight')
            plt.close(dashboard)
            saved_files.append(dashboard_filepath)
            print(f"✓ 综合仪表板已保存: {dashboard_filename}")
        except Exception as e:
            print(f"✗ 保存综合仪表板失败: {e}")
        
        return saved_files

def main():
    # 数据目录
    data_dir = "/Users/heguangzhong/Documents/8月考勤"
    
    # 创建可视化器实例
    visualizer = AttendanceVisualizer(data_dir)
    
    # 加载数据
    visualizer.load_data()
    
    # 生成并保存所有图表
    saved_files = visualizer.save_all_charts()
    
    print(f"\n🎨 考勤数据可视化完成！")
    print(f"📊 共生成 {len(saved_files)} 个图表文件")
    for file in saved_files:
        print(f"  📁 {os.path.basename(file)}")

if __name__ == "__main__":
    main()
