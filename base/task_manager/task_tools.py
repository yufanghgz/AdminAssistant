#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
待办任务管理MCP工具接口
提供任务管理功能的MCP工具实现
"""
from typing import Dict, Any, List, Optional
from .task_manager import TaskManager
from mcp.types import TextContent


# 全局任务管理器实例
task_manager = TaskManager()


def create_task(arguments: Dict[str, Any]) -> List[TextContent]:
    """
    创建新任务
    
    Args:
        arguments: 任务参数
        
    Returns:
        包含结果信息的TextContent列表
    """
    try:
        # 提取参数
        title = arguments.get("title")
        if not title:
            return [TextContent(
                type="text",
                text="创建任务失败：请提供任务标题"
            )]
        
        description = arguments.get("description")
        status = arguments.get("status", "pending")
        start_time = arguments.get("start_time")
        end_time = arguments.get("end_time")
        is_recurring = arguments.get("is_recurring", False)
        recurrence_rule = arguments.get("recurrence_rule")
        tags = arguments.get("tags")
        
        # 创建任务
        result = task_manager.create_task(
            title=title,
            description=description,
            status=status,
            start_time=start_time,
            end_time=end_time,
            is_recurring=is_recurring,
            recurrence_rule=recurrence_rule,
            tags=tags
        )
        
        if result["success"]:
            task = result["task"]
            return [TextContent(
                type="text",
                text=f"任务创建成功！\n任务ID: {task['id']}\n任务标题: {task['title']}"
            )]
        else:
            return [TextContent(
                type="text",
                text=f"任务创建失败: {result.get('message', '未知错误')}"
            )]
    except Exception as e:
        return [TextContent(
            type="text",
            text=f"创建任务时发生错误: {str(e)}"
        )]


def update_task(arguments: Dict[str, Any]) -> List[TextContent]:
    """
    更新任务信息（支持语义可寻址）
    
    Args:
        arguments: 更新参数
        
    Returns:
        包含结果信息的TextContent列表
    """
    try:
        # 提取参数
        task_id = arguments.get("id")
        target = arguments.get("target", {})
        filters = target.get("filters")
        update_fields = arguments.get("update_fields", {})
        
        # 兼容旧版本参数格式
        if not filters and 'filters' in arguments:
            filters = arguments['filters']
        
        # 检查必要参数
        if not update_fields:
            return [TextContent(
                type="text",
                text="更新任务失败：请提供要更新的字段"
            )]
        
        # 执行更新
        result = task_manager.update_task(
            task_id=task_id,
            filters=filters,
            update_fields=update_fields
        )
        
        if result["success"]:
            updated_count = len(result["updated_tasks"])
            return [TextContent(
                type="text",
                text=f"成功更新了{updated_count}个任务！"
            )]
        else:
            # 处理多匹配情况
            if result.get("multiple_matches"):
                message = result["message"] + "\n匹配的任务：\n"
                for i, task in enumerate(result["matched_tasks"][:5]):  # 最多显示5个
                    message += f"{i+1}. [{task['status']}] {task['title']} (ID: {task['id']})\n"
                if len(result["matched_tasks"]) > 5:
                    message += f"... 还有{len(result['matched_tasks']) - 5}个任务...\n"
                message += "请使用任务ID或更精确的条件进行更新。"
                return [TextContent(type="text", text=message)]
            else:
                return [TextContent(
                    type="text",
                    text=f"更新任务失败: {result.get('message', '未知错误')}"
                )]
    except Exception as e:
        return [TextContent(
            type="text",
            text=f"更新任务时发生错误: {str(e)}"
        )]


def query_tasks(arguments: Dict[str, Any]) -> List[TextContent]:
    """
    查询任务列表
    
    Args:
        arguments: 查询参数
        
    Returns:
        包含结果信息的TextContent列表
    """
    try:
        # 提取参数
        filters = arguments.get("filters", {})
        
        # 执行查询
        result = task_manager.query_tasks(filters)
        
        if result["success"]:
            tasks = result["tasks"]
            total = result["total"]
            
            if total == 0:
                return [TextContent(
                    type="text",
                    text="未找到匹配的任务"
                )]
            
            # 构建响应消息
            message = f"找到{total}个匹配的任务：\n\n"
            
            # 最多显示20个任务
            display_tasks = tasks[:20]
            for i, task in enumerate(display_tasks, 1):
                # 格式化任务信息
                status_map = {
                    "pending": "待办",
                    "in_progress": "进行中",
                    "done": "已完成",
                    "cancelled": "已取消"
                }
                status_text = status_map.get(task["status"], task["status"])
                
                message += f"{i}. [{status_text}] {task['title']}\n"
                if task.get("description"):
                    # 限制描述长度
                    description = task["description"]
                    if len(description) > 100:
                        description = description[:100] + "..."
                    message += f"   描述: {description}\n"
                if task.get("start_time"):
                    message += f"   开始时间: {task['start_time']}\n"
                if task.get("end_time"):
                    message += f"   结束时间: {task['end_time']}\n"
                if task.get("tags"):
                    message += f"   标签: {', '.join(task['tags'])}\n"
                message += "\n"
            
            if total > 20:
                message += f"... 还有{total - 20}个任务未显示...\n"
                message += "请使用更精确的过滤条件缩小范围。"
            
            return [TextContent(type="text", text=message)]
        else:
            return [TextContent(
                type="text",
                text=f"查询任务失败: {result.get('message', '未知错误')}"
            )]
    except Exception as e:
        return [TextContent(
            type="text",
            text=f"查询任务时发生错误: {str(e)}"
        )]


def export_tasks(arguments: Dict[str, Any]) -> List[TextContent]:
    """
    导出任务数据
    
    Args:
        arguments: 导出参数
        
    Returns:
        包含结果信息的TextContent列表
    """
    try:
        # 提取参数
        filters = arguments.get("filters", {})
        format_type = arguments.get("format", "json")
        filename = arguments.get("filename")
        
        # 执行导出
        result = task_manager.export_tasks(
            filters=filters,
            format_type=format_type,
            filename=filename
        )
        
        if result["success"]:
            return [TextContent(
                type="text",
                text=f"任务导出成功！\n文件保存路径: {result['export_path']}\n导出格式: {format_type}\n请前往该路径查看导出的任务数据。"
            )]
        else:
            return [TextContent(
                type="text",
                text=f"导出任务失败: {result.get('message', '未知错误')}"
            )]
    except Exception as e:
        return [TextContent(
            type="text",
            text=f"导出任务时发生错误: {str(e)}"
        )]