#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
待办任务管理核心模块
提供任务的创建、更新、查询和导出功能
"""
from typing import List, Dict, Any, Optional, Tuple
from datetime import datetime, timedelta
from .task_model import Task, TaskStatus, TaskQuery
from .task_storage import TaskStorage


class TaskManager:
    """待办任务管理器，提供任务的核心管理功能"""
    
    def __init__(self, storage_file: str = None):
        """
        初始化任务管理器
        
        Args:
            storage_file: 存储文件路径
        """
        self.storage = TaskStorage(storage_file)
    
    def create_task(self, 
                   title: str,
                   description: Optional[str] = None,
                   status: str = TaskStatus.PENDING.value,
                   start_time: Optional[str] = None,
                   end_time: Optional[str] = None,
                   is_recurring: bool = False,
                   recurrence_rule: Optional[str] = None,
                   tags: Optional[List[str]] = None) -> Dict[str, Any]:
        """创建新任务
        
        Args:
            title: 任务标题
            description: 任务描述
            status: 任务状态
            start_time: 开始时间
            end_time: 结束时间
            is_recurring: 是否为定期任务
            recurrence_rule: 定期任务规则
            tags: 任务标签列表
            
        Returns:
            创建的任务信息字典
        """
        # 验证状态值
        if status not in [s.value for s in TaskStatus]:
            raise ValueError(f"无效的任务状态: {status}")
        
        # 验证定期任务规则
        if is_recurring and not recurrence_rule:
            raise ValueError("定期任务必须指定重复规则")
        
        # 创建任务对象
        task = Task(
            title=title,
            description=description,
            status=status,
            start_time=start_time,
            end_time=end_time,
            is_recurring=is_recurring,
            recurrence_rule=recurrence_rule,
            tags=tags
        )
        
        # 保存任务
        saved_task = self.storage.save_task(task)
        
        return {
            "success": True,
            "task": saved_task.to_dict(),
            "message": "任务创建成功"
        }
    
    def update_task(self, 
                   update_fields: Dict[str, Any],
                   task_id: Optional[str] = None,
                   filters: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """更新任务信息（支持语义可寻址）
        
        Args:
            task_id: 任务ID（可选，如果提供则直接更新该任务）
            filters: 任务过滤条件（用于语义寻址）
            update_fields: 要更新的字段和值
            
        Returns:
            更新结果字典
        """
        # 确定要更新的任务
        if task_id:
            # 通过ID更新单个任务
            task = self.storage.get_task_by_id(task_id)
            if not task:
                return {
                    "success": False,
                    "message": f"未找到ID为{task_id}的任务"
                }
            
            tasks_to_update = [task]
        elif filters:
            # 通过语义条件更新任务
            query = self._build_query_from_filters(filters)
            tasks = self.storage.query_tasks(query)
            
            if not tasks:
                return {
                    "success": False,
                    "message": "未找到匹配的任务"
                }
            
            # 如果匹配多个任务，需要确认
            if len(tasks) > 1:
                return {
                    "success": False,
                    "multiple_matches": True,
                    "matched_tasks": [task.to_dict() for task in tasks],
                    "message": f"找到{len(tasks)}个匹配的任务，请提供更精确的条件或选择特定任务"
                }
            
            tasks_to_update = tasks
        else:
            return {
                "success": False,
                "message": "必须提供任务ID或过滤条件"
            }
        
        # 更新任务
        updated_tasks = []
        for task in tasks_to_update:
            # 验证要更新的字段
            self._validate_update_fields(update_fields)
            
            # 更新任务字段
            task.update(**update_fields)
            
            # 保存更新后的任务
            updated_task = self.storage.save_task(task)
            updated_tasks.append(updated_task.to_dict())
        
        return {
            "success": True,
            "updated_tasks": updated_tasks,
            "message": f"成功更新{len(updated_tasks)}个任务"
        }
    
    def query_tasks(self, filters: Dict[str, Any]) -> Dict[str, Any]:
        """查询任务列表
        
        Args:
            filters: 查询过滤条件
            
        Returns:
            任务列表和查询结果信息
        """
        # 构建查询对象
        query = self._build_query_from_filters(filters)
        
        # 执行查询
        tasks = self.storage.query_tasks(query)
        
        return {
            "success": True,
            "tasks": [task.to_dict() for task in tasks],
            "total": len(tasks),
            "message": f"找到{len(tasks)}个匹配的任务"
        }
    
    def export_tasks(self, 
                    filters: Dict[str, Any],
                    format_type: str = "json",
                    filename: str = None) -> Dict[str, Any]:
        """导出任务数据
        
        Args:
            filters: 查询过滤条件
            format_type: 导出格式 (json/csv/markdown)
            filename: 输出文件名
            
        Returns:
            导出结果信息
        """
        # 验证导出格式
        if format_type not in ["json", "csv", "markdown"]:
            return {
                "success": False,
                "message": f"不支持的导出格式: {format_type}"
            }
        
        # 构建查询对象
        query = self._build_query_from_filters(filters)
        
        try:
            # 执行导出
            export_path = self.storage.export_tasks(query, format_type, filename)
            
            return {
                "success": True,
                "export_path": export_path,
                "message": f"任务已成功导出到: {export_path}"
            }
        except Exception as e:
            return {
                "success": False,
                "message": f"导出任务失败: {str(e)}"
            }
    
    def _build_query_from_filters(self, filters: Dict[str, Any]) -> TaskQuery:
        """从过滤条件构建查询对象
        
        Args:
            filters: 过滤条件字典
            
        Returns:
            TaskQuery对象
        """
        # 处理相对日期
        filters = self._process_relative_dates(filters)
        
        return TaskQuery(
            logic=filters.get("logic", "and"),
            status=filters.get("status"),
            date_field=filters.get("date_field", "start_time"),
            date_range=filters.get("date_range"),
            keywords=filters.get("keywords"),
            tags=filters.get("tags"),
            conditions=filters.get("conditions"),
            limit=filters.get("limit", 100)
        )
    
    def _process_relative_dates(self, filters: Dict[str, Any]) -> Dict[str, Any]:
        """处理相对日期（如今天、昨天、本周等）
        
        Args:
            filters: 过滤条件字典
            
        Returns:
            处理后的过滤条件字典
        """
        # 复制filters以避免修改原始数据
        processed_filters = filters.copy()
        
        # 检查是否有相对日期关键词
        if "relative_date" in filters:
            relative_date = filters["relative_date"]
            date_range = self._get_date_range_for_relative_term(relative_date)
            
            if date_range:
                processed_filters["date_range"] = date_range
                # 移除relative_date字段
                del processed_filters["relative_date"]
        
        return processed_filters
    
    def _get_date_range_for_relative_term(self, term: str) -> Optional[Dict[str, str]]:
        """获取相对日期术语对应的日期范围
        
        Args:
            term: 相对日期术语（如今天、昨天、本周等）
            
        Returns:
            日期范围字典 {"start": "YYYY-MM-DD", "end": "YYYY-MM-DD"}
        """
        today = datetime.now().date()
        
        term = term.lower()
        
        if term == "今天" or term == "today":
            return {
                "start": today.isoformat(),
                "end": today.isoformat()
            }
        elif term == "昨天" or term == "yesterday":
            yesterday = today - timedelta(days=1)
            return {
                "start": yesterday.isoformat(),
                "end": yesterday.isoformat()
            }
        elif term == "本周" or term == "this week":
            # 本周的周一到周日
            weekday = today.weekday()  # 0=周一, 6=周日
            monday = today - timedelta(days=weekday)
            sunday = monday + timedelta(days=6)
            return {
                "start": monday.isoformat(),
                "end": sunday.isoformat()
            }
        elif term == "上周" or term == "last week":
            # 上周的周一到周日
            weekday = today.weekday()
            last_monday = today - timedelta(days=weekday + 7)
            last_sunday = last_monday + timedelta(days=6)
            return {
                "start": last_monday.isoformat(),
                "end": last_sunday.isoformat()
            }
        elif term == "本月" or term == "this month":
            # 本月的第一天到最后一天
            first_day = today.replace(day=1)
            # 获取下个月的第一天，然后减一天得到本月最后一天
            if today.month == 12:
                next_month = today.replace(year=today.year + 1, month=1, day=1)
            else:
                next_month = today.replace(month=today.month + 1, day=1)
            last_day = next_month - timedelta(days=1)
            return {
                "start": first_day.isoformat(),
                "end": last_day.isoformat()
            }
        elif term == "上月" or term == "last month":
            # 上月的第一天到最后一天
            if today.month == 1:
                last_month_first = today.replace(year=today.year - 1, month=12, day=1)
            else:
                last_month_first = today.replace(month=today.month - 1, day=1)
            
            # 计算上月最后一天
            if last_month_first.month == 12:
                next_month = last_month_first.replace(year=last_month_first.year + 1, month=1, day=1)
            else:
                next_month = last_month_first.replace(month=last_month_first.month + 1, day=1)
            last_month_last = next_month - timedelta(days=1)
            
            return {
                "start": last_month_first.isoformat(),
                "end": last_month_last.isoformat()
            }
        
        return None
    
    def _validate_update_fields(self, update_fields: Dict[str, Any]) -> None:
        """验证要更新的字段是否有效
        
        Args:
            update_fields: 要更新的字段和值
            
        Raises:
            ValueError: 如果字段无效
        """
        valid_fields = [
            "title", "description", "status", "start_time", "end_time",
            "is_recurring", "recurrence_rule", "tags"
        ]
        
        # 检查是否包含无效字段
        for field in update_fields:
            if field not in valid_fields:
                raise ValueError(f"无效的任务字段: {field}")
        
        # 验证状态值
        if "status" in update_fields:
            if update_fields["status"] not in [s.value for s in TaskStatus]:
                raise ValueError(f"无效的任务状态: {update_fields['status']}")
        
        # 验证定期任务相关字段
        if "is_recurring" in update_fields:
            if update_fields["is_recurring"] and "recurrence_rule" not in update_fields:
                raise ValueError("启用定期任务时必须指定重复规则")