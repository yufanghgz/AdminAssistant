#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
工时处理工具 - 用于生成工时展示图片
"""
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import os
from datetime import datetime
import logging

# 设置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 设置matplotlib中文字体
plt.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False


class WorktimeProcessor:
    """工时处理器类"""
    
    def __init__(self):
        self.df = None
        self.all_staff = []
        self.date_range = []
        self.assignees = []
        self.status_data = None
        self.project_data = None
    
    def load_excel_file(self, file_path, sheet_name='下周工作计划'):
        """
        加载Excel文件
        
        参数:
            file_path: Excel文件路径
            sheet_name: 工作表名称
        """
        try:
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"文件不存在: {file_path}")
            
            self.df = pd.read_excel(file_path, sheet_name=sheet_name)
            logger.info(f"成功加载Excel文件: {file_path}")
            logger.info(f"数据行数: {len(self.df)}")
            logger.info(f"列名: {list(self.df.columns)}")
            return True
        except Exception as e:
            logger.error(f"加载Excel文件失败: {str(e)}")
            return False
    
    def load_staff_list(self, staff_file_path=None):
        """
        加载人员名单
        
        参数:
            staff_file_path: 人员名单文件路径（可选）
        """
        try:
            if staff_file_path and os.path.exists(staff_file_path):
                staff_df = pd.read_excel(staff_file_path)
                self.all_staff = staff_df.iloc[:, 0].unique().tolist()
                logger.info(f"从文件加载人员名单，共 {len(self.all_staff)} 人")
            else:
                # 使用任务表中的人员名单
                if self.df is not None:
                    # 过滤掉空值和无效数据
                    staff_list = self.df['指派人名称'].dropna().astype(str).str.strip()
                    staff_list = staff_list[staff_list != '']
                    self.all_staff = staff_list.unique().tolist()
                    logger.info(f"从任务表加载人员名单，共 {len(self.all_staff)} 人")
                    logger.info(f"人员列表: {self.all_staff[:10]}{'...' if len(self.all_staff) > 10 else ''}")
                else:
                    raise ValueError("没有可用的数据源来获取人员名单")
        except Exception as e:
            logger.error(f"加载人员名单失败: {str(e)}")
            return False
        return True
    
    def process_dates(self):
        """处理日期数据"""
        try:
            if self.df is None:
                raise ValueError("数据未加载")
            
            # 日期处理 - 处理Excel序列号格式
            # 先检查是否为数字格式（Excel序列号）
            if pd.api.types.is_numeric_dtype(self.df['计划开始日期']):
                # Excel序列号转日期：Excel从1900年1月1日开始，但错误地认为1900是闰年
                self.df['计划开始日期'] = pd.to_datetime(self.df['计划开始日期'], origin='1899-12-30', unit='D')
            else:
                self.df['计划开始日期'] = pd.to_datetime(self.df['计划开始日期'], errors='coerce')
                
            if pd.api.types.is_numeric_dtype(self.df['计划结束日期']):
                self.df['计划结束日期'] = pd.to_datetime(self.df['计划结束日期'], origin='1899-12-30', unit='D')
            else:
                self.df['计划结束日期'] = pd.to_datetime(self.df['计划结束日期'], errors='coerce')
            
            # 过滤掉日期为空的行
            original_count = len(self.df)
            self.df = self.df.dropna(subset=['计划开始日期', '计划结束日期'])
            filtered_count = len(self.df)
            if original_count != filtered_count:
                logger.info(f"过滤掉 {original_count - filtered_count} 行日期为空的数据")
            
            # 转换为日期格式
            self.df['计划开始日期'] = self.df['计划开始日期'].dt.date
            self.df['计划结束日期'] = self.df['计划结束日期'].dt.date
            
            # 生成日期范围
            valid_dates_start = self.df['计划开始日期'].dropna()
            valid_dates_end = self.df['计划结束日期'].dropna()
            
            if valid_dates_start.empty or valid_dates_end.empty:
                logger.warning("没有找到有效的日期数据，使用默认日期范围")
                today = datetime.now().date()
                start_date = (datetime.now() - pd.Timedelta(days=7)).date()
                end_date = (datetime.now() + pd.Timedelta(days=7)).date()
            else:
                # 限制日期范围在合理范围内（当前日期前后3个月）
                today = datetime.now().date()
                min_date = today - pd.Timedelta(days=90)  # 3个月前
                max_date = today + pd.Timedelta(days=90)  # 3个月后
                
                # 过滤异常日期
                valid_dates_start = valid_dates_start[(valid_dates_start >= min_date) & (valid_dates_start <= max_date)]
                valid_dates_end = valid_dates_end[(valid_dates_end >= min_date) & (valid_dates_end <= max_date)]
                
                if valid_dates_start.empty or valid_dates_end.empty:
                    logger.warning("过滤后没有有效的日期数据，使用默认日期范围")
                    start_date = today - pd.Timedelta(days=7)
                    end_date = today + pd.Timedelta(days=7)
                else:
                    start_date = valid_dates_start.min()
                    end_date = valid_dates_end.max()
                    
                # 进一步限制日期范围不超过8周
                if (end_date - start_date).days > 56:  # 8周 = 56天
                    logger.warning(f"日期范围过大 ({end_date - start_date} 天)，限制为8周")
                    start_date = max(start_date, today - pd.Timedelta(days=28))  # 最多4周前
                    end_date = min(end_date, today + pd.Timedelta(days=28))      # 最多4周后
            
            self.date_range = pd.date_range(start=start_date, end=end_date).date.tolist()
            logger.info(f"日期范围: {start_date} 到 {end_date}")
            return True
        except Exception as e:
            logger.error(f"处理日期数据失败: {str(e)}")
            return False
    
    def organize_staff_by_project(self):
        """按项目组织人员"""
        try:
            if self.df is None or not self.all_staff:
                raise ValueError("数据或人员名单未加载")
            
            # 创建项目到人员的映射
            project_to_staff = {}
            no_project_staff = []
            
            for assignee in self.all_staff:
                # 确保assignee是字符串类型
                if pd.isna(assignee):
                    continue
                assignee = str(assignee).strip()
                if not assignee:
                    continue
                    
                assignee_tasks = self.df[self.df['指派人名称'] == assignee]
                if not assignee_tasks.empty:
                    projects = assignee_tasks['项目名称'].unique().tolist()
                    # 过滤掉NaN值
                    projects = [str(p).strip() for p in projects if not pd.isna(p) and str(p).strip()]
                    if projects:
                        main_project = projects[0]
                        if main_project not in project_to_staff:
                            project_to_staff[main_project] = []
                        if assignee not in project_to_staff[main_project]:
                            project_to_staff[main_project].append(assignee)
                else:
                    no_project_staff.append(assignee)
            
            # 按项目名称排序
            sorted_projects = sorted(project_to_staff.keys())
            
            # 构建按项目分组的人员列表
            self.assignees = []
            for project in sorted_projects:
                self.assignees.extend(project_to_staff[project])
            self.assignees.extend(no_project_staff)
            
            logger.info("按项目分组的人员列表:")
            for project in sorted_projects:
                logger.info(f"项目 '{project}': {', '.join(project_to_staff[project])}")
            if no_project_staff:
                logger.info(f"无项目人员: {', '.join(no_project_staff)}")
            
            return True
        except Exception as e:
            logger.error(f"组织人员失败: {str(e)}")
            return False
    
    def generate_status_data(self):
        """生成状态数据"""
        try:
            if not self.assignees or not self.date_range:
                raise ValueError("人员列表或日期范围未设置")
            
            # 初始化状态表和项目信息表
            self.status_data = np.zeros((len(self.assignees), len(self.date_range)))
            self.project_data = [[[] for _ in range(len(self.date_range))] for _ in range(len(self.assignees))]
            
            # 填充状态数据和项目信息
            for idx, assignee in enumerate(self.assignees):
                assignee_tasks = self.df[self.df['指派人名称'] == assignee]
                if not assignee_tasks.empty:
                    for _, task in assignee_tasks.iterrows():
                        project_name = task.get('项目名称', '无项目信息')
                        task_dates = pd.date_range(task['计划开始日期'], task['计划结束日期']).date
                        for date in task_dates:
                            if date in self.date_range:
                                col = self.date_range.index(date)
                                self.status_data[idx, col] = 1
                                if project_name not in self.project_data[idx][col]:
                                    self.project_data[idx][col].append(project_name)
            
            logger.info(f"状态数据生成完成: {np.sum(self.status_data == 1)} 个已安排, {np.sum(self.status_data == 0)} 个未安排")
            return True
        except Exception as e:
            logger.error(f"生成状态数据失败: {str(e)}")
            return False
    
    def generate_worktime_chart(self, output_path, title=None):
        """
        生成工时展示图片
        
        参数:
            output_path: 输出图片路径
            title: 图表标题
        """
        try:
            if self.status_data is None or self.project_data is None:
                raise ValueError("状态数据未生成")
            
            # 设置标题
            if title is None:
                title = f'部门工作状态与项目分配可视化 (共{len(self.assignees)}人)'
            else:
                # 如果用户提供了标题，确保包含人数信息
                if f'共{len(self.assignees)}人' not in title:
                    title = f'{title} (共{len(self.assignees)}人)'
            
            # 创建图表
            fig, ax = plt.subplots(figsize=(20, len(self.assignees)*0.8))
            
            # 创建背景
            background = np.ones((len(self.assignees), len(self.date_range), 3)) * np.array([0.56, 0.93, 0.56])
            ax.imshow(background, aspect='auto')
            
            # 在单元格中添加项目信息
            for i in range(len(self.assignees)):
                for j in range(len(self.date_range)):
                    if self.status_data[i, j] == 1 and self.project_data[i][j]:
                        projects = '\n'.join(self.project_data[i][j])
                        ax.text(j, i, projects, ha='center', va='center', fontsize=13,
                                color='black', weight='bold',
                                bbox=dict(facecolor='none', alpha=0, edgecolor='black', boxstyle='round,pad=0.3'))
                    elif self.status_data[i, j] == 0:
                        rect = plt.Rectangle((j-0.5, i-0.5), 1, 1, facecolor='white', edgecolor='none')
                        ax.add_patch(rect)
                        ax.text(j, i, '未安排', ha='center', va='center', fontsize=12,
                                color='gray', weight='bold',
                                bbox=dict(facecolor='none', alpha=0, edgecolor='none', boxstyle='round,pad=0.2'))
            
            # 设置坐标轴
            ax.set_xticks(np.arange(len(self.date_range)))
            ax.set_xticklabels([d.strftime('%m-%d') for d in self.date_range], rotation=45, fontsize=14, weight='bold')
            ax.set_yticks(np.arange(len(self.assignees)))
            ax.set_yticklabels(self.assignees, fontsize=14, weight='bold')
            
            # 添加网格线
            ax.set_xticks(np.arange(len(self.date_range))-0.5, minor=True)
            ax.set_yticks(np.arange(len(self.assignees))-0.5, minor=True)
            ax.grid(which="minor", color="gray", linestyle='-', linewidth=0.5)
            ax.tick_params(which="minor", size=0)
            
            plt.title(title, pad=20, fontsize=19, weight='bold')
            
            # 添加颜色图例
            green_patch = mpatches.Patch(color='#90EE90', label='已安排工作')
            white_patch = mpatches.Patch(color='white', edgecolor='gray', label='未安排工作')
            plt.legend(handles=[green_patch, white_patch], loc='lower center', 
                      bbox_to_anchor=(0.5, -0.05), ncol=2, fontsize=14, prop={'weight': 'bold'})
            
            plt.tight_layout(rect=[0, 0.05, 1, 0.95])
            
            # 保存图片
            plt.savefig(output_path, dpi=300, bbox_inches='tight')
            plt.close()
            
            logger.info(f"工时展示图片已保存到: {output_path}")
            return True
        except Exception as e:
            logger.error(f"生成工时展示图片失败: {str(e)}")
            return False
    
    def process_worktime(self, excel_file_path, output_path, staff_file_path=None, 
                        sheet_name='下周工作计划', title=None):
        """
        完整的工时处理流程
        
        参数:
            excel_file_path: Excel文件路径
            output_path: 输出图片路径
            staff_file_path: 人员名单文件路径（可选）
            sheet_name: 工作表名称
            title: 图表标题
        """
        try:
            # 1. 加载Excel文件
            if not self.load_excel_file(excel_file_path, sheet_name):
                return False
            
            # 2. 加载人员名单
            if not self.load_staff_list(staff_file_path):
                return False
            
            # 3. 处理日期
            if not self.process_dates():
                return False
            
            # 4. 按项目组织人员
            if not self.organize_staff_by_project():
                return False
            
            # 5. 生成状态数据
            if not self.generate_status_data():
                return False
            
            # 6. 生成工时展示图片
            if not self.generate_worktime_chart(output_path, title):
                return False
            
            logger.info("工时处理完成")
            return True
        except Exception as e:
            logger.error(f"工时处理失败: {str(e)}")
            return False


def process_worktime_file(excel_file_path, output_path, staff_file_path=None, 
                         sheet_name='下周工作计划', title=None):
    """
    工时处理函数 - 对外接口
    
    参数:
        excel_file_path: Excel文件路径
        output_path: 输出图片路径
        staff_file_path: 人员名单文件路径（可选）
        sheet_name: 工作表名称
        title: 图表标题
    
    返回:
        bool: 处理是否成功
    """
    processor = WorktimeProcessor()
    return processor.process_worktime(excel_file_path, output_path, staff_file_path, 
                                    sheet_name, title)


if __name__ == "__main__":
    # 测试代码
    test_excel = "/Users/heguangzhong/Work_Doc/11.项目管理/2025/工时计划/9月1周/合并后的9月第1周工作计划.xlsx"
    test_output = "/Users/heguangzhong/Work_Doc/11.项目管理/2025/工时计划/9月1周/工时展示图.png"
    
    success = process_worktime_file(test_excel, test_output, title="2025年9月第1周部门工作状态")
    if success:
        print("工时处理成功")
    else:
        print("工时处理失败")
