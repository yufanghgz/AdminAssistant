#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
考勤表处理程序
根据考勤处理需求自动处理各种考勤相关文件
"""

import os
import shutil
import pandas as pd
import datetime
import logging
from typing import Dict, List, Any
import warnings
warnings.filterwarnings('ignore')

class AttendanceProcessor:
    def __init__(self, work_dir: str):
        """
        初始化考勤处理器
        
        Args:
            work_dir: 工作目录路径，包含所有考勤相关文件
        """
        self.work_dir = work_dir
        self.current_month = datetime.datetime.now().strftime('%Y-%m')
        
        # 文件路径配置
        self.file_paths = {
            'attendance_template': os.path.join(work_dir, '上月考勤表.xlsx'),
            'current_attendance': os.path.join(work_dir, '当月考勤表.xlsx'),
            'leave_records': os.path.join(work_dir, '当月请假记录.xlsx'),
            'new_employees': os.path.join(work_dir, '当月入职人员信息表.xlsx'),
            'departure_employees': os.path.join(work_dir, '当月离职人员信息表.xlsx')
        }
        
        # 初始化日志
        self._setup_logging()
    
    def _setup_logging(self):
        """
        设置日志功能，日志文件保存在用户主目录下
        """
        # 获取用户主目录
        home_dir = os.path.expanduser("~")
        
        # 创建日志文件名（使用程序名）
        log_filename = f"attendance_processor_{datetime.datetime.now().strftime('%Y%m%d')}.log"
        log_filepath = os.path.join(home_dir, log_filename)
        
        # 配置日志格式
        log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        date_format = '%Y-%m-%d %H:%M:%S'
        
        # 配置日志记录器
        logging.basicConfig(
            level=logging.INFO,
            format=log_format,
            datefmt=date_format,
            handlers=[
                logging.FileHandler(log_filepath, encoding='utf-8'),
                logging.StreamHandler()  # 同时输出到控制台
            ]
        )
        
        # 创建日志记录器
        self.logger = logging.getLogger('AttendanceProcessor')
        
        # 记录程序启动信息
        self.logger.info(f"考勤处理器启动 - 工作目录: {self.work_dir}")
        self.logger.info(f"日志文件: {log_filepath}")
        
    def copy_attendance_file(self):
        """
        复制指定目录下的『上月考勤表.xlsx』为『当月考勤表.xlsx』
        """
        self.logger.info("开始复制上月考勤表为当月考勤表")
        print("1. 复制上月考勤表为当月考勤表...")
        
        # 检查源文件是否存在
        if not os.path.exists(self.file_paths['attendance_template']):
            error_msg = f"上月考勤表文件不存在: {self.file_paths['attendance_template']}"
            self.logger.error(error_msg)
            raise FileNotFoundError(error_msg)
        
        try:
            # 复制文件
            shutil.copy2(self.file_paths['attendance_template'], self.file_paths['current_attendance'])
            self.logger.info(f"文件复制成功: {self.file_paths['current_attendance']}")
            print(f"✓ 已复制为: {self.file_paths['current_attendance']}")
            
            # 验证复制结果
            if os.path.exists(self.file_paths['current_attendance']):
                file_size = os.path.getsize(self.file_paths['current_attendance'])
                self.logger.info(f"文件复制验证成功，文件大小: {file_size} 字节")
                print("✓ 文件复制成功")
            else:
                error_msg = "文件复制失败，目标文件不存在"
                self.logger.error(error_msg)
                raise FileNotFoundError(error_msg)
                
        except Exception as e:
            error_msg = f"复制文件时发生错误: {str(e)}"
            self.logger.error(error_msg)
            raise
    
    def update_attendance_days(self):
        """
        修改『当月考勤表』表中应出勤天数字段为22天
        """
        self.logger.info("开始修改应出勤天数字段为22天")
        print("2. 修改应出勤天数字段为22天...")
        
        try:
            # 读取当月考勤表
            df = pd.read_excel(self.file_paths['current_attendance'])
            self.logger.info(f"读取考勤表成功，共 {len(df)} 条记录")
            
            # 查找应出勤天数字段
            attendance_columns = [col for col in df.columns if '应出勤' in str(col)]
            
            if attendance_columns:
                for col in attendance_columns:
                    old_values = df[col].unique()
                    df[col] = 22
                    self.logger.info(f"已更新字段 '{col}' 为22天，原值: {old_values}")
                
                # 保存修改后的文件
                df.to_excel(self.file_paths['current_attendance'], index=False)
                self.logger.info("应出勤天数字段修改完成并已保存")
                print("✓ 应出勤天数字段已修改为22天")
            else:
                self.logger.warning("未找到应出勤天数字段")
                print("⚠ 未找到应出勤天数字段")
                
        except Exception as e:
            error_msg = f"修改应出勤天数时发生错误: {str(e)}"
            self.logger.error(error_msg)
            raise
    
    def remove_at_symbols(self):
        """
        去掉『当月考勤表』表中人员字段前的@符号
        """
        self.logger.info("开始去掉人员字段前的@符号")
        print("3. 去掉人员字段前的@符号...")
        
        try:
            # 读取当月考勤表
            df = pd.read_excel(self.file_paths['current_attendance'])
            
            # 查找人员字段
            name_columns = [col for col in df.columns if '人员' in str(col) or '姓名' in str(col)]
            
            if name_columns:
                for col in name_columns:
                    # 统计包含@符号的记录数
                    at_count = df[col].astype(str).str.contains('@', na=False).sum()
                    self.logger.info(f"字段 '{col}' 中包含@符号的记录数: {at_count}")
                    
                    # 去掉@符号
                    df[col] = df[col].astype(str).str.replace('@', '', regex=False)
                    
                    # 验证修改结果
                    remaining_at = df[col].astype(str).str.contains('@', na=False).sum()
                    self.logger.info(f"修改后剩余@符号数量: {remaining_at}")
                
                # 保存修改后的文件
                df.to_excel(self.file_paths['current_attendance'], index=False)
                self.logger.info("人员字段@符号去除完成并已保存")
                print("✓ 人员字段前的@符号已去除")
            else:
                self.logger.warning("未找到人员字段")
                print("⚠ 未找到人员字段")
                
        except Exception as e:
            error_msg = f"去除@符号时发生错误: {str(e)}"
            self.logger.error(error_msg)
            raise
    
    def reset_leave_fields(self):
        """
        修改『当月考勤表』表中年假、事假、病假、产检假、婚假、育儿假、丧假、陪产假、请假天数字段为0
        """
        self.logger.info("开始重置各种假期字段为0")
        print("4. 重置各种假期字段为0...")
        
        try:
            # 读取当月考勤表
            df = pd.read_excel(self.file_paths['current_attendance'])
            
            # 需要重置的字段列表
            leave_fields = ['年假', '事假', '病假', '产检假', '婚假', '育儿假', '丧假', '陪产假', '请假天数']
            
            found_fields = []
            for field in leave_fields:
                # 查找匹配的列
                matching_cols = [col for col in df.columns if field in str(col)]
                for col in matching_cols:
                    # 记录修改前的统计信息
                    old_stats = {
                        '非零值数量': (df[col] != 0).sum(),
                        '唯一值': df[col].unique().tolist()
                    }
                    
                    # 重置为0
                    df[col] = 0
                    
                    self.logger.info(f"字段 '{col}' 重置为0，修改前统计: {old_stats}")
                    found_fields.append(col)
            
            if found_fields:
                # 保存修改后的文件
                df.to_excel(self.file_paths['current_attendance'], index=False)
                self.logger.info(f"假期字段重置完成并已保存，共处理 {len(found_fields)} 个字段")
                print(f"✓ 已重置 {len(found_fields)} 个假期字段为0")
            else:
                self.logger.warning("未找到需要重置的假期字段")
                print("⚠ 未找到需要重置的假期字段")
                
        except Exception as e:
            error_msg = f"重置假期字段时发生错误: {str(e)}"
            self.logger.error(error_msg)
            raise
    
    def remove_departure_records(self):
        """
        删除备注字段里包含"离职"的人员记录
        """
        self.logger.info("开始删除备注字段里包含'离职'的人员记录")
        print("5. 删除备注字段里包含'离职'的人员记录...")
        
        try:
            # 读取当月考勤表
            df = pd.read_excel(self.file_paths['current_attendance'])
            original_count = len(df)
            self.logger.info(f"读取考勤表成功，共 {original_count} 条记录")
            
            # 查找备注字段
            remarks_columns = [col for col in df.columns if '备注' in str(col) or 'remark' in str(col).lower()]
            
            if remarks_columns:
                remarks_col = remarks_columns[0]
                
                # 统计包含"离职"的记录数
                departure_mask = df[remarks_col].astype(str).str.contains('离职', na=False)
                departure_count = departure_mask.sum()
                self.logger.info(f"找到 {departure_count} 条包含'离职'的记录")
                
                if departure_count > 0:
                    # 显示即将删除的记录
                    departure_records = df[departure_mask]
                    self.logger.info("即将删除的离职人员记录:")
                    for idx, row in departure_records.iterrows():
                        employee_id = row.get('工号', '未知')
                        employee_name = row.get('人员', '未知')
                        remarks = row[remarks_col]
                        self.logger.info(f"  {employee_id} - {employee_name}: {remarks}")
                    
                    # 删除包含"离职"的记录
                    df = df[~departure_mask]
                    new_count = len(df)
                    removed_count = original_count - new_count
                    
                    # 保存修改后的文件
                    df.to_excel(self.file_paths['current_attendance'], index=False)
                    self.logger.info(f"已删除 {removed_count} 条离职人员记录，剩余 {new_count} 条记录")
                    print(f"✓ 已删除 {removed_count} 条离职人员记录")
                else:
                    self.logger.info("没有找到包含'离职'的记录")
                    print("✓ 没有找到包含'离职'的记录")
            else:
                self.logger.warning("未找到备注字段")
                print("⚠ 未找到备注字段")
                
        except Exception as e:
            error_msg = f"删除离职人员记录时发生错误: {str(e)}"
            self.logger.error(error_msg)
            raise
    
    def process_leave_records(self):
        """
        根据『当月请假记录』表，通过工号匹配处理请假信息
        """
        self.logger.info("开始处理当月请假记录")
        print("6. 处理当月请假记录...")
        
        try:
            # 检查请假记录文件是否存在
            if not os.path.exists(self.file_paths['leave_records']):
                self.logger.warning("当月请假记录表不存在，跳过此步骤")
                print("⚠ 当月请假记录表不存在，跳过此步骤")
                return
            
            # 读取考勤表和请假记录表
            df_attendance = pd.read_excel(self.file_paths['current_attendance'])
            df_leave = pd.read_excel(self.file_paths['leave_records'])
            
            self.logger.info(f"读取考勤表成功，共 {len(df_attendance)} 条记录")
            self.logger.info(f"读取请假记录表成功，共 {len(df_leave)} 条记录")
            
            # 查找工号字段
            attendance_id_cols = [col for col in df_attendance.columns if '工号' in str(col)]
            leave_id_cols = [col for col in df_leave.columns if '工号' in str(col) or '发起人' in str(col)]
            
            if not attendance_id_cols or not leave_id_cols:
                self.logger.warning("未找到工号字段，跳过请假记录处理")
                print("⚠ 未找到工号字段，跳过请假记录处理")
                return
            
            attendance_id_col = attendance_id_cols[0]
            leave_id_col = leave_id_cols[0]
            
            # 查找请假类型和时长字段
            leave_type_cols = [col for col in df_leave.columns if '类型' in str(col) or '假期类型' in str(col)]
            leave_duration_cols = [col for col in df_leave.columns if '时长' in str(col) or '天数' in str(col) or '小时' in str(col)]
            
            if not leave_type_cols or not leave_duration_cols:
                self.logger.warning("未找到请假类型或时长字段，跳过请假记录处理")
                print("⚠ 未找到请假类型或时长字段，跳过请假记录处理")
                return
            
            leave_type_col = leave_type_cols[0]
            leave_duration_col = leave_duration_cols[0]
            
            # 按工号分组汇总请假信息
            leave_summary = df_leave.groupby(leave_id_col).apply(
                lambda x: self._summarize_leave_records(x, leave_type_col, leave_duration_col)
            ).to_dict()
            
            self.logger.info(f"请假记录汇总完成，涉及 {len(leave_summary)} 个员工")
            self.logger.info(f"请假记录中的工号: {list(leave_summary.keys())[:10]}")
            
            # 更新考勤表中的请假信息
            updated_count = 0
            for idx, row in df_attendance.iterrows():
                employee_id = row[attendance_id_col]
                employee_name = row.get('人员', '未知')
                
                # 尝试整数和字符串两种匹配方式
                if employee_id in leave_summary or str(employee_id) in leave_summary:
                    # 获取请假信息
                    leave_info = leave_summary.get(employee_id, leave_summary.get(str(employee_id), {}))
                    self.logger.info(f"员工 {employee_id} - {employee_name} 请假信息: {leave_info}")
                    print(f"  {employee_id} - {employee_name}: {leave_info}")
                    
                    # 更新各种假期字段
                    for leave_type, days in leave_info.items():
                        # 精确匹配字段名，避免误匹配
                        if leave_type == '年假':
                            matching_cols = [col for col in df_attendance.columns if col == '年假']
                        elif leave_type == '事假':
                            matching_cols = [col for col in df_attendance.columns if col == '事假']
                        elif leave_type == '病假':
                            matching_cols = [col for col in df_attendance.columns if col == '病假']
                        elif leave_type == '产检假':
                            matching_cols = [col for col in df_attendance.columns if col == '产检假']
                        elif leave_type == '婚假':
                            matching_cols = [col for col in df_attendance.columns if col == '婚假']
                        elif leave_type == '育儿假':
                            matching_cols = [col for col in df_attendance.columns if col == '育儿假']
                        elif leave_type == '丧假':
                            matching_cols = [col for col in df_attendance.columns if col == '丧假']
                        elif leave_type == '陪产假':
                            matching_cols = [col for col in df_attendance.columns if col == '陪产假']
                        else:
                            # 对于其他类型，使用模糊匹配
                            matching_cols = [col for col in df_attendance.columns if leave_type in str(col)]
                        
                        for col in matching_cols:
                            old_value = df_attendance.at[idx, col]
                            df_attendance.at[idx, col] = days
                            self.logger.info(f"    更新字段 '{col}': {old_value} -> {days}")
                    
                    # 计算总请假天数
                    total_leave = sum(leave_info.values())
                    total_leave_cols = [col for col in df_attendance.columns if '请假天数' in str(col)]
                    for col in total_leave_cols:
                        old_value = df_attendance.at[idx, col]
                        df_attendance.at[idx, col] = total_leave
                        self.logger.info(f"    更新总请假天数 '{col}': {old_value} -> {total_leave}")
                    
                    updated_count += 1
            
            # 保存修改后的考勤表
            df_attendance.to_excel(self.file_paths['current_attendance'], index=False)
            self.logger.info(f"请假记录处理完成，共更新 {updated_count} 名员工")
            print(f"✓ 已更新 {updated_count} 名员工的请假信息")
            
        except Exception as e:
            error_msg = f"处理请假记录时发生错误: {str(e)}"
            self.logger.error(error_msg)
            raise
    
    def _summarize_leave_records(self, group_df: pd.DataFrame, type_col: str, duration_col: str) -> Dict[str, float]:
        """汇总单个员工的请假记录"""
        summary = {}
        for _, record in group_df.iterrows():
            leave_type = str(record[type_col])
            duration = float(record[duration_col]) if pd.notna(record[duration_col]) else 0
            
            if leave_type in summary:
                summary[leave_type] += duration
            else:
                summary[leave_type] = duration
                
        return summary
    
    def update_actual_attendance(self):
        """
        更新『当月考勤表』中的实际出勤天数字段
        实际出勤天数 = 应出勤天数 - 请假天数
        """
        self.logger.info("开始更新实际出勤天数字段")
        print("7. 更新实际出勤天数字段...")
        
        try:
            # 读取当月考勤表
            df = pd.read_excel(self.file_paths['current_attendance'])
            self.logger.info(f"读取考勤表成功，共 {len(df)} 条记录")
            
            # 查找相关字段
            should_attend_cols = [col for col in df.columns if '应出勤' in str(col)]
            leave_days_cols = [col for col in df.columns if '请假天数' in str(col)]
            actual_attend_cols = [col for col in df.columns if '实际出勤' in str(col)]
            
            if should_attend_cols and leave_days_cols and actual_attend_cols:
                should_attend_col = should_attend_cols[0]
                leave_days_col = leave_days_cols[0]
                actual_attend_col = actual_attend_cols[0]
                
                # 记录更新前的统计信息
                old_stats = {
                    '应出勤天数范围': f"{df[should_attend_col].min()} - {df[should_attend_col].max()}",
                    '请假天数范围': f"{df[leave_days_col].min()} - {df[leave_days_col].max()}",
                    '实际出勤天数范围': f"{df[actual_attend_col].min()} - {df[actual_attend_col].max()}"
                }
                self.logger.info(f"更新前统计: {old_stats}")
                
                # 计算实际出勤天数 = 应出勤天数 - 请假天数
                df[actual_attend_col] = df[should_attend_col] - df[leave_days_col]
                
                # 记录更新后的统计信息
                new_stats = {
                    '实际出勤天数范围': f"{df[actual_attend_col].min()} - {df[actual_attend_col].max()}",
                    '实际出勤天数平均值': f"{df[actual_attend_col].mean():.2f}"
                }
                self.logger.info(f"更新后统计: {new_stats}")
                
                # 保存修改后的文件
                df.to_excel(self.file_paths['current_attendance'], index=False)
                self.logger.info("实际出勤天数字段更新完成并已保存")
                print("✓ 实际出勤天数字段已更新")
            else:
                self.logger.warning("未找到相关字段，跳过此步骤")
                print("⚠ 未找到相关字段，跳过此步骤")
                
        except Exception as e:
            error_msg = f"更新实际出勤天数时发生错误: {str(e)}"
            self.logger.error(error_msg)
            raise
    
    def update_actual_salary_days(self):
        """
        更新『当月考勤表』中的实际发工资天数字段
        应发工资天数 = 应出勤天数 - 事假天数 - 病假天数*0.5 - 多休年假天数
        """
        self.logger.info("开始更新实际发工资天数字段")
        print("8. 更新实际发工资天数字段...")
        
        try:
            # 读取当月考勤表
            df = pd.read_excel(self.file_paths['current_attendance'])
            self.logger.info(f"读取考勤表成功，共 {len(df)} 条记录")
            
            # 查找相关字段
            should_attend_cols = [col for col in df.columns if '应出勤' in str(col)]
            personal_leave_cols = [col for col in df.columns if '事假' in str(col)]
            sick_leave_cols = [col for col in df.columns if '病假' in str(col)]
            extra_annual_leave_cols = [col for col in df.columns if '多休年假' in str(col)]
            salary_days_cols = [col for col in df.columns if '发工资' in str(col) or '应发工资' in str(col)]
            
            if (should_attend_cols and personal_leave_cols and sick_leave_cols and 
                extra_annual_leave_cols and salary_days_cols):
                
                should_attend_col = should_attend_cols[0]
                personal_leave_col = personal_leave_cols[0]
                sick_leave_col = sick_leave_cols[0]
                extra_annual_leave_col = extra_annual_leave_cols[0]
                salary_days_col = salary_days_cols[0]
                
                # 记录更新前的统计信息
                old_stats = {
                    '应出勤天数范围': f"{df[should_attend_col].min()} - {df[should_attend_col].max()}",
                    '事假天数范围': f"{df[personal_leave_col].min()} - {df[personal_leave_col].max()}",
                    '病假天数范围': f"{df[sick_leave_col].min()} - {df[sick_leave_col].max()}",
                    '多休年假天数范围': f"{df[extra_annual_leave_col].min()} - {df[extra_annual_leave_col].max()}",
                    '应发工资天数范围': f"{df[salary_days_col].min()} - {df[salary_days_col].max()}"
                }
                self.logger.info(f"更新前统计: {old_stats}")
                
                # 计算应发工资天数 = 应出勤天数 - 事假天数 - 病假天数*0.5 - 多休年假天数
                df[salary_days_col] = (df[should_attend_col] - 
                                     df[personal_leave_col] - 
                                     df[sick_leave_col] * 0.5 - 
                                     df[extra_annual_leave_col])
                
                # 记录更新后的统计信息
                new_stats = {
                    '应发工资天数范围': f"{df[salary_days_col].min()} - {df[salary_days_col].max()}",
                    '应发工资天数平均值': f"{df[salary_days_col].mean():.2f}"
                }
                self.logger.info(f"更新后统计: {new_stats}")
                
                # 保存修改后的文件
                df.to_excel(self.file_paths['current_attendance'], index=False)
                self.logger.info("实际发工资天数字段更新完成并已保存")
                print("✓ 实际发工资天数字段已更新")
            else:
                missing_fields = []
                if not should_attend_cols: missing_fields.append("应出勤天数")
                if not personal_leave_cols: missing_fields.append("事假")
                if not sick_leave_cols: missing_fields.append("病假")
                if not extra_annual_leave_cols: missing_fields.append("多休年假天数")
                if not salary_days_cols: missing_fields.append("应发工资天数")
                
                self.logger.warning(f"未找到相关字段: {', '.join(missing_fields)}，跳过此步骤")
                print(f"⚠ 未找到相关字段: {', '.join(missing_fields)}，跳过此步骤")
                
        except Exception as e:
            error_msg = f"更新实际发工资天数时发生错误: {str(e)}"
            self.logger.error(error_msg)
            raise
    
    def add_sick_leave_verification(self):
        """
        逐一处理每个人的病假信息，如有病假，请在备注栏增加『待核查病假条』内容
        """
        self.logger.info("开始处理病假信息，添加待核查病假条备注")
        print("10. 处理病假信息，添加待核查病假条备注...")
        
        try:
            # 读取当月考勤表
            df = pd.read_excel(self.file_paths['current_attendance'])
            self.logger.info(f"读取考勤表成功，共 {len(df)} 条记录")
            
            # 查找病假和备注字段
            sick_leave_cols = [col for col in df.columns if '病假' in str(col)]
            remarks_cols = [col for col in df.columns if '备注' in str(col)]
            
            if sick_leave_cols and remarks_cols:
                sick_leave_col = sick_leave_cols[0]
                remarks_col = remarks_cols[0]
                
                # 查找有病假的员工
                sick_leave_employees = df[df[sick_leave_col] > 0]
                self.logger.info(f"找到 {len(sick_leave_employees)} 名员工有病假记录")
                
                if len(sick_leave_employees) > 0:
                    # 为有病假的员工添加备注
                    for idx, row in sick_leave_employees.iterrows():
                        employee_id = row.get('工号', '未知')
                        employee_name = row.get('人员', '未知')
                        sick_days = row[sick_leave_col]
                        current_remarks = str(row[remarks_col]) if pd.notna(row[remarks_col]) else ''
                        
                        # 添加待核查病假条备注
                        if '待核查病假条' not in current_remarks:
                            if current_remarks and current_remarks != 'nan':
                                new_remarks = f"{current_remarks} 待核查病假条"
                            else:
                                new_remarks = "待核查病假条"
                            
                            df.at[idx, remarks_col] = new_remarks
                            self.logger.info(f"员工 {employee_id} - {employee_name} 病假{sick_days}天，已添加待核查病假条备注")
                            print(f"  {employee_id} - {employee_name}: 病假{sick_days}天，已添加待核查病假条备注")
                    
                    # 保存修改后的文件
                    df.to_excel(self.file_paths['current_attendance'], index=False)
                    self.logger.info(f"病假信息处理完成，共处理 {len(sick_leave_employees)} 名员工")
                    print(f"✓ 已为 {len(sick_leave_employees)} 名有病假的员工添加待核查病假条备注")
                else:
                    self.logger.info("没有员工有病假记录，跳过此步骤")
                    print("✓ 没有员工有病假记录，跳过此步骤")
            else:
                missing_fields = []
                if not sick_leave_cols: missing_fields.append("病假")
                if not remarks_cols: missing_fields.append("备注")
                
                self.logger.warning(f"未找到相关字段: {', '.join(missing_fields)}，跳过此步骤")
                print(f"⚠ 未找到相关字段: {', '.join(missing_fields)}，跳过此步骤")
                
        except Exception as e:
            error_msg = f"处理病假信息时发生错误: {str(e)}"
            self.logger.error(error_msg)
            raise
    
    def add_new_employees(self):
        """
        根据『当月入职人员信息表』添加入职人员信息到『当月考勤表』
        将银行卡号设置为文本格式，去掉小数位
        将『当月入职人员信息表』中的入职日期填写到『当月考勤表』中的备注字段
        """
        self.logger.info("开始添加入职人员信息")
        print("12. 添加入职人员信息...")
        
        try:
            # 检查入职人员信息表是否存在
            if not os.path.exists(self.file_paths['new_employees']):
                self.logger.warning("当月入职人员信息表不存在，跳过此步骤")
                print("⚠ 当月入职人员信息表不存在，跳过此步骤")
                return
            
            # 读取当月考勤表和入职人员信息表
            df_attendance = pd.read_excel(self.file_paths['current_attendance'])
            df_new_employees = pd.read_excel(self.file_paths['new_employees'])
            
            self.logger.info(f"读取考勤表成功，共 {len(df_attendance)} 条记录")
            self.logger.info(f"读取入职人员信息表成功，共 {len(df_new_employees)} 条记录")
            
            # 查找工号字段
            attendance_id_cols = [col for col in df_attendance.columns if '工号' in str(col)]
            new_employee_id_cols = [col for col in df_new_employees.columns if '工号' in str(col)]
            
            if not attendance_id_cols or not new_employee_id_cols:
                self.logger.warning("未找到工号字段，跳过入职人员信息添加")
                print("⚠ 未找到工号字段，跳过入职人员信息添加")
                return
            
            attendance_id_col = attendance_id_cols[0]
            new_employee_id_col = new_employee_id_cols[0]
            
            # 查找其他必要字段
            name_cols = [col for col in df_new_employees.columns if '姓名' in str(col) or '人员' in str(col)]
            bank_card_cols = [col for col in df_new_employees.columns if '银行卡' in str(col) or '银行' in str(col)]
            entry_date_cols = [col for col in df_new_employees.columns if '入职' in str(col) and '日期' in str(col)]
            department_cols = [col for col in df_new_employees.columns if '归属' in str(col) or '部门' in str(col) or '科室' in str(col)]
            # 查找员工类型字段
            employee_type_cols = [col for col in df_new_employees.columns if '类型' in str(col) or '职位' in str(col) or '岗位' in str(col)]
            
            if not name_cols:
                self.logger.warning("未找到姓名字段，跳过入职人员信息添加")
                print("⚠ 未找到姓名字段，跳过入职人员信息添加")
                return
            
            name_col = name_cols[0]
            bank_card_col = bank_card_cols[0] if bank_card_cols else None
            entry_date_col = entry_date_cols[0] if entry_date_cols else None
            department_col = department_cols[0] if department_cols else None
            employee_type_col = employee_type_cols[0] if employee_type_cols else None
            
            # 获取现有员工的工号列表
            existing_employee_ids = set(df_attendance[attendance_id_col].astype(str))
            
            # 处理新入职员工
            added_count = 0
            for idx, row in df_new_employees.iterrows():
                employee_id = str(row[new_employee_id_col])
                employee_name = row[name_col]
                
                # 检查是否已存在
                if employee_id in existing_employee_ids:
                    self.logger.info(f"员工 {employee_id} - {employee_name} 已存在，跳过")
                    continue
                
                # 创建新员工记录
                new_employee = {}
                
                # 复制考勤表的结构，设置默认值
                for col in df_attendance.columns:
                    if col == attendance_id_col:
                        new_employee[col] = employee_id
                    elif col == '人员':
                        new_employee[col] = employee_name
                    elif col == '归属' and department_col and pd.notna(row[department_col]):
                        # 添加归属字段信息
                        new_employee[col] = row[department_col]
                    elif any(keyword in col for keyword in ['类型', '职位', '岗位']) and employee_type_col and pd.notna(row[employee_type_col]):
                        # 添加员工类型/职位/岗位信息
                        new_employee[col] = row[employee_type_col]
                    elif col == '应出勤天数':
                        new_employee[col] = 22
                    elif col == '实际出勤天数':
                        new_employee[col] = 22
                    elif col == '实际发工资天数':
                        new_employee[col] = 22
                    elif col in ['年假', '事假', '病假', '产检假', '婚假', '育儿假', '丧假', '陪产假', '请假天数', '多休年假天数']:
                        new_employee[col] = 0
                    elif col == '银行卡号' and bank_card_col and pd.notna(row[bank_card_col]):
                        # 将银行卡号设置为文本格式，去掉小数位
                        bank_card = str(row[bank_card_col])
                        if '.' in bank_card:
                            bank_card = bank_card.split('.')[0]
                        new_employee[col] = bank_card
                    elif col == '备注' and entry_date_col and pd.notna(row[entry_date_col]):
                        # 将入职日期填写到备注字段
                        entry_date = str(row[entry_date_col])
                        new_employee[col] = f"入职日期: {entry_date}"
                    else:
                        new_employee[col] = ''
                
                # 添加新员工记录到考勤表
                df_attendance = pd.concat([df_attendance, pd.DataFrame([new_employee])], ignore_index=True)
                
                self.logger.info(f"添加新员工: {employee_id} - {employee_name}")
                print(f"  {employee_id} - {employee_name}: 已添加")
                added_count += 1
            
            # 确保银行卡号字段为文本格式
            if '银行卡号' in df_attendance.columns:
                df_attendance['银行卡号'] = df_attendance['银行卡号'].astype(str)
                # 去掉小数位
                df_attendance['银行卡号'] = df_attendance['银行卡号'].str.replace('.0', '', regex=False)
            
            # 保存修改后的考勤表
            df_attendance.to_excel(self.file_paths['current_attendance'], index=False)
            self.logger.info(f"入职人员信息添加完成，共添加 {added_count} 名新员工")
            print(f"✓ 已添加 {added_count} 名新员工")
            
        except Exception as e:
            error_msg = f"添加入职人员信息时发生错误: {str(e)}"
            self.logger.error(error_msg)
            raise
    
    def process_departure_info(self):
        """
        根据『当月离职人员信息表』按工号匹配，在『当月考勤表』中的备注里添加离职日期+"离职"
        """
        self.logger.info("开始处理离职人员信息")
        print("11. 处理离职人员信息...")
        
        try:
            # 检查离职人员信息表是否存在
            if not os.path.exists(self.file_paths['departure_employees']):
                self.logger.warning("当月离职人员信息表不存在，跳过此步骤")
                print("⚠ 当月离职人员信息表不存在，跳过此步骤")
                return
            
            # 读取考勤表和离职人员信息表
            df_attendance = pd.read_excel(self.file_paths['current_attendance'])
            df_departure = pd.read_excel(self.file_paths['departure_employees'])
            
            self.logger.info(f"读取考勤表成功，共 {len(df_attendance)} 条记录")
            self.logger.info(f"读取离职人员信息表成功，共 {len(df_departure)} 条记录")
            
            # 查找工号字段
            attendance_id_cols = [col for col in df_attendance.columns if '工号' in str(col)]
            departure_id_cols = [col for col in df_departure.columns if '工号' in str(col)]
            
            if not attendance_id_cols or not departure_id_cols:
                self.logger.warning("未找到工号字段，跳过离职人员信息处理")
                print("⚠ 未找到工号字段，跳过离职人员信息处理")
                return
            
            attendance_id_col = attendance_id_cols[0]
            departure_id_col = departure_id_cols[0]
            
            # 查找离职日期字段
            departure_date_cols = [col for col in df_departure.columns if '离职日期' in str(col)]
            if not departure_date_cols:
                self.logger.warning("未找到离职日期字段，跳过离职人员信息处理")
                print("⚠ 未找到离职日期字段，跳过离职人员信息处理")
                return
            
            departure_date_col = departure_date_cols[0]
            
            # 查找备注字段
            remarks_cols = [col for col in df_attendance.columns if '备注' in str(col)]
            if not remarks_cols:
                self.logger.warning("未找到备注字段，跳过离职人员信息处理")
                print("⚠ 未找到备注字段，跳过离职人员信息处理")
                return
            
            remarks_col = remarks_cols[0]
            
            # 创建离职人员信息字典
            departure_info = {}
            for _, row in df_departure.iterrows():
                employee_id = row[departure_id_col]
                departure_date = row[departure_date_col]
                if pd.notna(employee_id) and pd.notna(departure_date):
                    departure_info[employee_id] = departure_date
            
            self.logger.info(f"离职人员信息汇总完成，涉及 {len(departure_info)} 个员工")
            
            # 更新考勤表中的离职信息
            updated_count = 0
            for idx, row in df_attendance.iterrows():
                employee_id = row[attendance_id_col]
                employee_name = row.get('人员', '未知')
                
                # 尝试整数和字符串两种匹配方式
                if employee_id in departure_info or str(employee_id) in departure_info:
                    departure_date = departure_info.get(employee_id, departure_info.get(str(employee_id)))
                    current_remarks = str(row[remarks_col]) if pd.notna(row[remarks_col]) else ''
                    
                    # 添加离职信息到备注
                    if current_remarks and current_remarks != 'nan':
                        new_remarks = f"{current_remarks} {departure_date}离职"
                    else:
                        new_remarks = f"{departure_date}离职"
                    
                    df_attendance.at[idx, remarks_col] = new_remarks
                    self.logger.info(f"员工 {employee_id} - {employee_name} 离职日期: {departure_date}，已添加离职备注")
                    print(f"  {employee_id} - {employee_name}: {departure_date}离职")
                    updated_count += 1
            
            # 保存修改后的考勤表
            df_attendance.to_excel(self.file_paths['current_attendance'], index=False)
            self.logger.info(f"离职人员信息处理完成，共更新 {updated_count} 名员工")
            print(f"✓ 已为 {updated_count} 名离职员工添加离职日期备注")
            
        except Exception as e:
            error_msg = f"处理离职人员信息时发生错误: {str(e)}"
            self.logger.error(error_msg)
            raise
    
    def update_month_field(self):
        """
        修改『当月考勤表』中的月份字段为当月字段，格式为 yyyy-mm
        """
        self.logger.info("开始修改月份字段为当月字段")
        print("13. 修改月份字段为当月字段...")
        
        try:
            # 读取当月考勤表，指定银行卡号列为文本类型
            dtype_dict = {}
            if '银行卡号' in pd.read_excel(self.file_paths['current_attendance'], nrows=0).columns:
                dtype_dict['银行卡号'] = str
            df = pd.read_excel(self.file_paths['current_attendance'], dtype=dtype_dict)
            self.logger.info(f"读取考勤表成功，共 {len(df)} 条记录")
            
            # 查找月份字段
            month_cols = [col for col in df.columns if '月份' in str(col)]
            
            if month_cols:
                month_col = month_cols[0]
                old_values = df[month_col].unique()
                self.logger.info(f"字段 '{month_col}' 修改前值: {old_values}")
                
                # 更新为当前月份
                df[month_col] = self.current_month
                
                self.logger.info(f"字段 '{month_col}' 已更新为: {self.current_month}")
                
                # 确保银行卡号字段保持文本格式
                if '银行卡号' in df.columns:
                    # 先转换为字符串，然后去掉小数位
                    df['银行卡号'] = df['银行卡号'].astype(str)
                    df['银行卡号'] = df['银行卡号'].str.replace('.0', '', regex=False)
                    # 将空字符串和'nan'替换为空值
                    df['银行卡号'] = df['银行卡号'].replace(['nan', ''], '')
                    # 确保银行卡号列的数据类型为object（文本）
                    df['银行卡号'] = df['银行卡号'].astype('object')
                
                # 保存修改后的文件，确保银行卡号以文本格式保存
                with pd.ExcelWriter(self.file_paths['current_attendance'], engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Sheet1')
                    
                    # 获取工作表对象
                    worksheet = writer.sheets['Sheet1']
                    
                    # 找到银行卡号列的索引
                    if '银行卡号' in df.columns:
                        bank_card_col_idx = df.columns.get_loc('银行卡号') + 1  # Excel列索引从1开始
                        
                        # 设置银行卡号列为文本格式
                        for row in range(2, len(df) + 2):  # 从第2行开始（跳过标题行）
                            cell = worksheet.cell(row=row, column=bank_card_col_idx)
                            cell.number_format = '@'  # 设置为文本格式
                
                self.logger.info("月份字段修改完成并已保存")
                print(f"✓ 月份字段已修改为: {self.current_month}")
            else:
                self.logger.warning("未找到月份字段")
                print("⚠ 未找到月份字段")
                
        except Exception as e:
            error_msg = f"修改月份字段时发生错误: {str(e)}"
            self.logger.error(error_msg)
            raise

    def clear_remarks_and_bank(self):
        """
        清除『当月考勤表』表中备注、银行账号字段中的信息
        """
        self.logger.info("开始清除备注、银行账号字段中的信息")
        print("9. 清除备注、银行账号字段中的信息...")
        
        try:
            # 读取当月考勤表
            df = pd.read_excel(self.file_paths['current_attendance'])
            
            # 查找备注和银行账号字段
            remarks_columns = [col for col in df.columns if '备注' in str(col) or 'remark' in str(col).lower()]
            bank_columns = [col for col in df.columns if '银行' in str(col) or '账号' in str(col) or 'account' in str(col).lower()]
            
            cleared_fields = []
            
            # 清除备注字段
            for col in remarks_columns:
                # 记录清除前的统计信息
                non_empty_count = df[col].notna().sum()
                self.logger.info(f"字段 '{col}' 清除前非空记录数: {non_empty_count}")
                
                # 清除字段内容，设置为空字符串
                df[col] = ''
                cleared_fields.append(col)
            
            # 清除银行账号字段
            for col in bank_columns:
                # 记录清除前的统计信息
                non_empty_count = df[col].notna().sum()
                self.logger.info(f"字段 '{col}' 清除前非空记录数: {non_empty_count}")
                
                # 清除字段内容，设置为空字符串
                df[col] = ''
                cleared_fields.append(col)
            
            # 确保空字符串在保存后仍然为空字符串而不是NaN
            for col in cleared_fields:
                df[col] = df[col].fillna('')
            
            if cleared_fields:
                # 保存修改后的文件
                df.to_excel(self.file_paths['current_attendance'], index=False)
                self.logger.info(f"备注和银行账号字段清除完成并已保存，共处理 {len(cleared_fields)} 个字段")
                print(f"✓ 已清除 {len(cleared_fields)} 个字段的信息")
            else:
                self.logger.warning("未找到备注或银行账号字段")
                print("⚠ 未找到备注或银行账号字段")
                
        except Exception as e:
            error_msg = f"清除备注和银行账号字段时发生错误: {str(e)}"
            self.logger.error(error_msg)
            raise

def main():
    """主函数"""
    import sys
    
    if len(sys.argv) != 2:
        print("使用方法: python attendance_processor.py <工作目录>")
        print("示例: python attendance_processor.py /path/to/attendance/files")
        sys.exit(1)
    
    work_dir = sys.argv[1]
    
    if not os.path.exists(work_dir):
        print(f"错误: 工作目录不存在: {work_dir}")
        sys.exit(1)
    
    try:
        # 创建处理器并执行处理流程
        processor = AttendanceProcessor(work_dir)
        
        # 1. 复制文件
        processor.copy_attendance_file()
        
        # 2. 修改应出勤天数字段为22天
        processor.update_attendance_days()
        
        # 3. 去掉人员字段前的@符号
        processor.remove_at_symbols()
        
        # 4. 重置各种假期字段为0
        processor.reset_leave_fields()
        
        # 5. 删除备注字段里包含"离职"的人员记录
        processor.remove_departure_records()
        
        # 6. 处理当月请假记录
        processor.process_leave_records()
        
        # 7. 更新实际出勤天数字段
        processor.update_actual_attendance()
        
        # 8. 更新实际发工资天数字段
        processor.update_actual_salary_days()
        
        # 9. 清除备注、银行账号字段中的信息
        processor.clear_remarks_and_bank()
        
        # 10. 处理病假信息，添加待核查病假条备注
        processor.add_sick_leave_verification()
        
        # 11. 处理离职人员信息
        processor.process_departure_info()
        
        # 12. 添加入职人员信息
        processor.add_new_employees()
        
        # 13. 修改月份字段为当月字段
        processor.update_month_field()
        
        processor.logger.info("所有数据处理步骤完成")
        print("\n✓ 所有数据处理步骤已完成")
        
    except Exception as e:
        print(f"程序执行失败: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()