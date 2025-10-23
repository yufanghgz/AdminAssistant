#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
待办任务数据模型定义
"""
import uuid
from datetime import datetime
from typing import Optional, List, Dict, Any
from enum import Enum


class TaskStatus(Enum):
    """任务状态枚举"""
    PENDING = "pending"
    IN_PROGRESS = "in_progress"
    DONE = "done"
    CANCELLED = "cancelled"


class RecurrenceRule(Enum):
    """定期任务规则枚举"""
    DAILY = "daily"
    WEEKLY = "weekly"
    MONTHLY = "monthly"
    CUSTOM = "custom"


class Task:
    """待办任务数据模型"""
    
    def __init__(self,
                 title: str,
                 description: Optional[str] = None,
                 status: str = TaskStatus.PENDING.value,
                 start_time: Optional[str] = None,
                 end_time: Optional[str] = None,
                 is_recurring: bool = False,
                 recurrence_rule: Optional[str] = None,
                 tags: Optional[List[str]] = None,
                 task_id: Optional[str] = None):
        """
        初始化任务对象
        
        Args:
            title: 任务标题
            description: 任务描述
            status: 任务状态
            start_time: 开始时间，格式为ISO 8601字符串
            end_time: 结束时间，格式为ISO 8601字符串
            is_recurring: 是否为定期任务
            recurrence_rule: 定期任务规则
            tags: 任务标签列表
            task_id: 任务ID（如果不提供则自动生成）
        """
        self.id = task_id or str(uuid.uuid4())
        self.title = title
        self.description = description
        self.status = status
        self.start_time = start_time
        self.end_time = end_time
        self.is_recurring = is_recurring
        self.recurrence_rule = recurrence_rule
        self.tags = tags or []
        self.created_at = datetime.now().isoformat()
        self.updated_at = datetime.now().isoformat()
    
    def to_dict(self) -> Dict[str, Any]:
        """将任务对象转换为字典"""
        return {
            "id": self.id,
            "title": self.title,
            "description": self.description,
            "status": self.status,
            "start_time": self.start_time,
            "end_time": self.end_time,
            "is_recurring": self.is_recurring,
            "recurrence_rule": self.recurrence_rule,
            "tags": self.tags,
            "created_at": self.created_at,
            "updated_at": self.updated_at
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'Task':
        """从字典创建任务对象"""
        task = cls(
            title=data["title"],
            description=data.get("description"),
            status=data.get("status", TaskStatus.PENDING.value),
            start_time=data.get("start_time"),
            end_time=data.get("end_time"),
            is_recurring=data.get("is_recurring", False),
            recurrence_rule=data.get("recurrence_rule"),
            tags=data.get("tags"),
            task_id=data["id"]
        )
        # 设置创建和更新时间
        task.created_at = data.get("created_at", task.created_at)
        task.updated_at = data.get("updated_at", task.updated_at)
        return task
    
    def update(self, **kwargs) -> None:
        """更新任务属性"""
        for key, value in kwargs.items():
            if hasattr(self, key):
                setattr(self, key, value)
        # 更新时间戳
        self.updated_at = datetime.now().isoformat()


class TaskQuery:
    """任务查询条件模型"""
    
    def __init__(self,
                 logic: str = "and",
                 status: Optional[List[str]] = None,
                 date_field: str = "start_time",
                 date_range: Optional[Dict[str, str]] = None,
                 keywords: Optional[List[str]] = None,
                 tags: Optional[List[str]] = None,
                 conditions: Optional[List[Dict[str, Any]]] = None,
                 limit: int = 100):
        """
        初始化查询条件
        
        Args:
            logic: 逻辑操作符 (and/or)
            status: 任务状态列表
            date_field: 日期字段 (start_time/end_time)
            date_range: 日期范围 {"start": "2025-10-01", "end": "2025-10-13"}
            keywords: 关键词列表（在标题或描述中）
            tags: 标签列表
            conditions: 自定义条件列表
            limit: 返回结果的最大数量
        """
        self.logic = logic
        self.status = status
        self.date_field = date_field
        self.date_range = date_range
        self.keywords = keywords
        self.tags = tags
        self.conditions = conditions or []
        self.limit = limit
    
    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        return {
            "logic": self.logic,
            "status": self.status,
            "date_field": self.date_field,
            "date_range": self.date_range,
            "keywords": self.keywords,
            "tags": self.tags,
            "conditions": self.conditions,
            "limit": self.limit
        }