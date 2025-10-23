#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
待办任务存储模块
负责任务数据的持久化存储和读取
"""
import json
import os
from typing import List, Dict, Any, Optional
from datetime import datetime
from .task_model import Task, TaskQuery


class TaskStorage:
    """任务存储类，提供任务数据的持久化存储功能"""
    
    def __init__(self, storage_file: str = None):
        """
        初始化任务存储
        
        Args:
            storage_file: 存储文件路径，如果不提供则使用默认路径
        """
        # 默认存储文件路径
        if storage_file is None:
            base_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
            storage_dir = os.path.join(base_dir, "data")
            # 确保存储目录存在
            os.makedirs(storage_dir, exist_ok=True)
            storage_file = os.path.join(storage_dir, "tasks.json")
        
        self.storage_file = storage_file
        # 确保存储文件存在
        if not os.path.exists(storage_file):
            with open(storage_file, 'w', encoding='utf-8') as f:
                json.dump([], f, ensure_ascii=False, indent=2)
    
    def _read_all_tasks(self) -> List[Dict[str, Any]]:
        """从存储文件中读取所有任务数据"""
        try:
            with open(self.storage_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, FileNotFoundError):
            # 文件不存在或格式错误，返回空列表
            return []
    
    def _write_all_tasks(self, tasks: List[Dict[str, Any]]) -> None:
        """将所有任务数据写入存储文件"""
        try:
            with open(self.storage_file, 'w', encoding='utf-8') as f:
                json.dump(tasks, f, ensure_ascii=False, indent=2)
        except Exception as e:
            raise IOError(f"写入任务数据失败: {str(e)}")
    
    def save_task(self, task: Task) -> Task:
        """保存任务到存储
        
        Args:
            task: 任务对象
            
        Returns:
            保存后的任务对象
        """
        tasks = self._read_all_tasks()
        task_dict = task.to_dict()
        
        # 检查是否是更新现有任务
        for i, existing_task in enumerate(tasks):
            if existing_task["id"] == task.id:
                tasks[i] = task_dict
                break
        else:
            # 不是更新，添加新任务
            tasks.append(task_dict)
        
        self._write_all_tasks(tasks)
        return task
    
    def get_task_by_id(self, task_id: str) -> Optional[Task]:
        """通过ID获取任务
        
        Args:
            task_id: 任务ID
            
        Returns:
            任务对象，如果不存在则返回None
        """
        tasks = self._read_all_tasks()
        for task_dict in tasks:
            if task_dict["id"] == task_id:
                return Task.from_dict(task_dict)
        return None
    
    def query_tasks(self, query: TaskQuery) -> List[Task]:
        """根据查询条件获取任务列表
        
        Args:
            query: 查询条件对象
            
        Returns:
            符合条件的任务列表
        """
        all_tasks = self._read_all_tasks()
        result_tasks = []
        
        for task_dict in all_tasks:
            if self._matches_query(task_dict, query):
                result_tasks.append(Task.from_dict(task_dict))
        
        # 按创建时间倒序排序（最新的在前）
        result_tasks.sort(key=lambda t: t.created_at, reverse=True)
        
        # 限制返回数量
        return result_tasks[:query.limit]
    
    def delete_task(self, task_id: str) -> bool:
        """删除指定ID的任务
        
        Args:
            task_id: 任务ID
            
        Returns:
            删除是否成功
        """
        tasks = self._read_all_tasks()
        initial_count = len(tasks)
        
        # 过滤掉要删除的任务
        tasks = [task for task in tasks if task["id"] != task_id]
        
        if len(tasks) < initial_count:
            self._write_all_tasks(tasks)
            return True
        
        return False
    
    def _matches_query(self, task_dict: Dict[str, Any], query: TaskQuery) -> bool:
        """检查任务是否符合查询条件
        
        Args:
            task_dict: 任务字典数据
            query: 查询条件
            
        Returns:
            是否符合条件
        """
        conditions_met = []
        
        # 检查状态条件
        if query.status:
            status_match = task_dict.get("status") in query.status
            conditions_met.append(status_match)
        
        # 检查日期范围条件
        if query.date_range:
            date_value = task_dict.get(query.date_field)
            if date_value:
                task_date = datetime.fromisoformat(date_value)
                start_date = datetime.fromisoformat(query.date_range["start"])
                end_date = datetime.fromisoformat(query.date_range["end"])
                date_match = start_date <= task_date <= end_date
                conditions_met.append(date_match)
            else:
                # 没有日期字段，不匹配
                conditions_met.append(False)
        
        # 检查关键词条件
        if query.keywords:
            title = task_dict.get("title", "")
            description = task_dict.get("description", "")
            content = title + " " + description
            # 检查是否包含任一关键词
            keyword_match = any(keyword.lower() in content.lower() for keyword in query.keywords)
            conditions_met.append(keyword_match)
        
        # 检查标签条件
        if query.tags:
            task_tags = task_dict.get("tags", [])
            # 检查是否包含任一标签
            tag_match = any(tag in task_tags for tag in query.tags)
            conditions_met.append(tag_match)
        
        # 检查自定义条件
        for condition in query.conditions:
            field = condition.get("field")
            operator = condition.get("operator", "eq")
            value = condition.get("value")
            
            if field and field in task_dict:
                task_value = task_dict[field]
                match = False
                
                if operator == "eq":
                    match = task_value == value
                elif operator == "ne":
                    match = task_value != value
                elif operator == "contains" and isinstance(task_value, str):
                    match = value in task_value
                elif operator == "in" and isinstance(task_value, list):
                    match = value in task_value
                
                conditions_met.append(match)
            else:
                conditions_met.append(False)
        
        # 应用逻辑操作符
        if not conditions_met:
            # 没有条件时返回True
            return True
        
        if query.logic == "and":
            return all(conditions_met)
        else:
            # 默认使用OR逻辑
            return any(conditions_met)
    
    def export_tasks(self, query: TaskQuery, format_type: str = "json", filename: str = None) -> str:
        """导出任务数据到指定格式的文件
        
        Args:
            query: 查询条件
            format_type: 导出格式 (json/csv/markdown)
            filename: 输出文件名，如果不提供则生成默认文件名
            
        Returns:
            导出文件的路径
        """
        tasks = self.query_tasks(query)
        
        # 生成默认文件名
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
            export_dir = os.path.join(base_dir, "exports")
            os.makedirs(export_dir, exist_ok=True)
            filename = os.path.join(export_dir, f"tasks_export_{timestamp}.{format_type}")
        
        # 根据格式导出数据
        if format_type == "json":
            self._export_json(tasks, filename)
        elif format_type == "csv":
            self._export_csv(tasks, filename)
        elif format_type == "markdown":
            self._export_markdown(tasks, filename)
        else:
            raise ValueError(f"不支持的导出格式: {format_type}")
        
        return filename
    
    def _export_json(self, tasks: List[Task], filename: str) -> None:
        """导出为JSON格式"""
        tasks_data = [task.to_dict() for task in tasks]
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(tasks_data, f, ensure_ascii=False, indent=2)
    
    def _export_csv(self, tasks: List[Task], filename: str) -> None:
        """导出为CSV格式"""
        import csv
        
        if not tasks:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write("")
            return
        
        # 确定CSV字段
        fields = ["id", "title", "description", "status", "start_time", "end_time", 
                 "is_recurring", "recurrence_rule", "tags", "created_at", "updated_at"]
        
        with open(filename, 'w', encoding='utf-8', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=fields)
            writer.writeheader()
            
            for task in tasks:
                task_dict = task.to_dict()
                # 处理列表类型的字段
                if "tags" in task_dict and isinstance(task_dict["tags"], list):
                    task_dict["tags"] = ",".join(task_dict["tags"])
                # 写入行数据
                writer.writerow({k: v for k, v in task_dict.items() if k in fields})
    
    def _export_markdown(self, tasks: List[Task], filename: str) -> None:
        """导出为Markdown格式"""
        with open(filename, 'w', encoding='utf-8') as f:
            # 写入标题
            f.write("# 任务导出清单\n\n")
            f.write(f"导出时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"任务总数: {len(tasks)}\n\n")
            
            # 写入任务列表
            for i, task in enumerate(tasks, 1):
                f.write(f"## {i}. {task.title}\n")
                f.write(f"- **ID**: {task.id}\n")
                if task.description:
                    f.write(f"- **描述**: {task.description}\n")
                f.write(f"- **状态**: {task.status}\n")
                if task.start_time:
                    f.write(f"- **开始时间**: {task.start_time}\n")
                if task.end_time:
                    f.write(f"- **结束时间**: {task.end_time}\n")
                if task.tags:
                    f.write(f"- **标签**: {', '.join(task.tags)}\n")
                if task.is_recurring:
                    f.write(f"- **定期任务**: 是 (规则: {task.recurrence_rule})\n")
                f.write(f"- **创建时间**: {task.created_at}\n")
                f.write(f"- **更新时间**: {task.updated_at}\n\n")