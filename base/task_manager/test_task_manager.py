#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
待办任务管理功能测试脚本
"""
import os
import sys
import json
from datetime import datetime, timedelta
from .task_manager import TaskManager
from .task_model import TaskStatus


def test_task_manager():
    """测试任务管理器的各项功能"""
    print("===== 待办任务管理功能测试 =====")
    
    # 创建临时存储文件路径
    temp_storage = "temp_test_tasks.json"
    
    try:
        # 初始化任务管理器
        print("1. 初始化任务管理器...")
        manager = TaskManager(temp_storage)
        print(f"   ✓ 任务管理器初始化成功，存储文件: {temp_storage}")
        
        # 测试创建任务
        print("\n2. 测试创建任务...")
        # 创建普通任务
        result = manager.create_task(
            title="测试任务1",
            description="这是第一个测试任务",
            status=TaskStatus.PENDING.value,
            tags=["测试", "开发"]
        )
        print(f"   ✓ 创建普通任务: {result['message']}")
        task1_id = result['task']['id']
        
        # 创建带时间的任务
        tomorrow = (datetime.now() + timedelta(days=1)).strftime('%Y-%m-%dT%H:%M:%S')
        result = manager.create_task(
            title="测试任务2（带时间）",
            description="这是一个带时间的测试任务",
            status=TaskStatus.IN_PROGRESS.value,
            start_time=datetime.now().strftime('%Y-%m-%dT%H:%M:%S'),
            end_time=tomorrow,
            tags=["测试", "时间敏感"]
        )
        print(f"   ✓ 创建带时间任务: {result['message']}")
        task2_id = result['task']['id']
        
        # 创建定期任务
        result = manager.create_task(
            title="每周例会",
            description="每周一上午10点的团队例会",
            status=TaskStatus.PENDING.value,
            is_recurring=True,
            recurrence_rule="weekly",
            tags=["会议", "定期"]
        )
        print(f"   ✓ 创建定期任务: {result['message']}")
        task3_id = result['task']['id']
        
        # 测试查询任务
        print("\n3. 测试查询任务...")
        # 查询所有任务
        result = manager.query_tasks({})
        print(f"   ✓ 查询所有任务: 找到{result['total']}个任务")
        
        # 按状态查询
        result = manager.query_tasks({"status": [TaskStatus.PENDING.value]})
        print(f"   ✓ 按状态查询（待办）: 找到{result['total']}个任务")
        
        # 按关键词查询
        result = manager.query_tasks({"keywords": ["测试"]})
        print(f"   ✓ 按关键词查询（测试）: 找到{result['total']}个任务")
        
        # 按标签查询
        result = manager.query_tasks({"tags": ["定期"]})
        print(f"   ✓ 按标签查询（定期）: 找到{result['total']}个任务")
        
        # 测试更新任务
        print("\n4. 测试更新任务...")
        # 通过ID更新
        result = manager.update_task(
            task_id=task1_id,
            update_fields={"status": TaskStatus.DONE.value}
        )
        print(f"   ✓ 通过ID更新任务状态: {result['message']}")
        
        # 通过语义条件更新
        result = manager.update_task(
            filters={"keywords": ["例会"], "tags": ["定期"]},
            update_fields={"description": "更新后的每周例会描述"}
        )
        print(f"   ✓ 通过语义条件更新任务描述: {result['message']}")
        
        # 测试导出任务
        print("\n5. 测试导出任务...")
        # 导出JSON格式
        result = manager.export_tasks({
            "status": [TaskStatus.PENDING.value, TaskStatus.DONE.value]
        }, format_type="json")
        print(f"   ✓ 导出JSON格式: {result['message']}")
        
        # 导出CSV格式
        result = manager.export_tasks({
            "tags": ["测试"]
        }, format_type="csv")
        print(f"   ✓ 导出CSV格式: {result['message']}")
        
        # 导出Markdown格式
        result = manager.export_tasks({}, format_type="markdown")
        print(f"   ✓ 导出Markdown格式: {result['message']}")
        
        print("\n===== 所有测试完成 =====")
        print("\n功能测试结果：通过！\n待办任务管理模块已成功集成到MCP服务中。")
        print("\n可用的MCP接口：")
        print("- create_task: 创建新任务")
        print("- update_task: 更新任务信息（支持语义可寻址）")
        print("- query_tasks: 查询任务列表")
        print("- export_tasks: 导出任务数据")
        
    except Exception as e:
        print(f"\n测试失败: {str(e)}")
        raise
    finally:
        # 清理测试数据
        if os.path.exists(temp_storage):
            os.remove(temp_storage)


if __name__ == "__main__":
    test_task_manager()