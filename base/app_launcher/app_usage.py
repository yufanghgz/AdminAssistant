#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
本地应用智能打开模块 - 使用记录器
记录应用使用频率与时间，用于智能排序推荐
"""
import os
import json
import time
from typing import Dict, Any, List, Tuple
from datetime import datetime
from .config import CONFIG
from ..mcp_logger import MCPToolLogger

logger = MCPToolLogger("app_usage")

class AppUsageTracker:
    """应用使用记录器类，负责记录和管理应用使用情况"""
    
    def __init__(self):
        """初始化应用使用记录器"""
        self.logger = logger
        self.usage_file = CONFIG["USAGE_FILE_PATH"]
        self.usage_data = {}
        self.load_usage_data()
    
    def load_usage_data(self) -> bool:
        """
        从文件加载使用记录数据
        
        Returns:
            bool: 加载是否成功
        """
        self.logger.info(f"尝试加载使用记录文件: {self.usage_file}")
        
        if not os.path.exists(self.usage_file):
            self.logger.info("使用记录文件不存在，将创建新文件")
            self.usage_data = {}
            return True
        
        try:
            with open(self.usage_file, 'r', encoding='utf-8') as f:
                self.usage_data = json.load(f)
            
            # 验证数据格式
            if not isinstance(self.usage_data, dict):
                self.logger.warning("使用记录文件格式无效，重置为新数据")
                self.usage_data = {}
                return False
            
            self.logger.info(f"成功加载使用记录，包含 {len(self.usage_data)} 个应用的使用记录")
            return True
        except json.JSONDecodeError as e:
            self.logger.error(f"使用记录文件解析错误: {str(e)}")
            self.usage_data = {}
            return False
        except Exception as e:
            self.logger.error(f"加载使用记录失败: {str(e)}")
            self.usage_data = {}
            return False
    
    def save_usage_data(self) -> bool:
        """
        保存使用记录数据到文件
        
        Returns:
            bool: 保存是否成功
        """
        self.logger.info(f"尝试保存使用记录到文件: {self.usage_file}")
        
        try:
            # 确保目录存在
            usage_dir = os.path.dirname(self.usage_file)
            if not os.path.exists(usage_dir):
                os.makedirs(usage_dir, exist_ok=True)
            
            # 保存数据
            with open(self.usage_file, 'w', encoding='utf-8') as f:
                json.dump(self.usage_data, f, ensure_ascii=False, indent=2)
            
            self.logger.info("使用记录保存成功")
            return True
        except Exception as e:
            self.logger.error(f"保存使用记录失败: {str(e)}")
            return False
    
    def record_usage(self, app_name: str) -> bool:
        """
        记录应用使用情况
        
        Args:
            app_name: 应用名称
            
        Returns:
            bool: 记录是否成功
        """
        if not app_name or not CONFIG["ENABLE_USAGE_TRACKING"]:
            return False
        
        try:
            # 获取当前时间
            current_time = datetime.now().isoformat()
            current_timestamp = int(time.time())
            
            # 更新使用记录
            if app_name not in self.usage_data:
                self.usage_data[app_name] = {
                    "count": 1,
                    "last_used": current_time,
                    "last_used_timestamp": current_timestamp,
                    "first_used": current_time,
                    "usage_history": [current_timestamp]
                }
            else:
                self.usage_data[app_name]["count"] += 1
                self.usage_data[app_name]["last_used"] = current_time
                self.usage_data[app_name]["last_used_timestamp"] = current_timestamp
                
                # 限制历史记录长度（保留最近100次使用记录）
                if len(self.usage_data[app_name]["usage_history"]) >= 100:
                    self.usage_data[app_name]["usage_history"].pop(0)
                self.usage_data[app_name]["usage_history"].append(current_timestamp)
            
            self.logger.info(f"记录应用使用: {app_name}, 已使用 {self.usage_data[app_name]['count']} 次")
            
            # 异步保存（实际使用时可能需要考虑性能优化）
            self.save_usage_data()
            
            return True
        except Exception as e:
            self.logger.error(f"记录应用使用失败: {str(e)}")
            return False
    
    def get_usage_data(self, app_name: str = None) -> Any:
        """
        获取应用使用数据
        
        Args:
            app_name: 应用名称，如果为None则返回所有应用的使用数据
            
        Returns:
            dict: 应用使用数据或所有应用的使用数据
        """
        if app_name:
            return self.usage_data.get(app_name, {})
        else:
            return self.usage_data.copy()
    
    def get_most_used_apps(self, limit: int = 10) -> List[Tuple[str, Dict[str, Any]]]:
        """
        获取使用频率最高的应用列表
        
        Args:
            limit: 返回的最大应用数量
            
        Returns:
            list: 使用频率最高的应用列表，按使用次数降序排列
        """
        # 按使用次数排序
        sorted_apps = sorted(
            self.usage_data.items(),
            key=lambda x: x[1].get("count", 0),
            reverse=True
        )
        
        # 限制返回数量
        return sorted_apps[:limit]
    
    def get_recently_used_apps(self, limit: int = 10) -> List[Tuple[str, Dict[str, Any]]]:
        """
        获取最近使用的应用列表
        
        Args:
            limit: 返回的最大应用数量
            
        Returns:
            list: 最近使用的应用列表，按使用时间降序排列
        """
        # 按最后使用时间排序
        sorted_apps = sorted(
            self.usage_data.items(),
            key=lambda x: x[1].get("last_used_timestamp", 0),
            reverse=True
        )
        
        # 限制返回数量
        return sorted_apps[:limit]
    
    def clear_usage_data(self, app_name: str = None) -> bool:
        """
        清除使用记录数据
        
        Args:
            app_name: 应用名称，如果为None则清除所有应用的使用记录
            
        Returns:
            bool: 清除是否成功
        """
        try:
            if app_name:
                # 清除指定应用的使用记录
                if app_name in self.usage_data:
                    del self.usage_data[app_name]
                    self.logger.info(f"清除应用 '{app_name}' 的使用记录")
                else:
                    self.logger.warning(f"应用 '{app_name}' 的使用记录不存在")
            else:
                # 清除所有应用的使用记录
                self.usage_data = {}
                self.logger.info("清除所有应用的使用记录")
            
            # 保存更改
            return self.save_usage_data()
        except Exception as e:
            self.logger.error(f"清除使用记录失败: {str(e)}")
            return False
    
    def get_usage_score(self, app_name: str) -> float:
        """
        计算应用的使用得分（用于智能排序）
        
        Args:
            app_name: 应用名称
            
        Returns:
            float: 使用得分，越高表示越常使用或最近使用过
        """
        if not CONFIG["ENABLE_USAGE_PRIORITY"] or app_name not in self.usage_data:
            return 0.0
        
        app_usage = self.usage_data[app_name]
        
        # 获取使用次数和最后使用时间
        count = app_usage.get("count", 0)
        last_used_timestamp = app_usage.get("last_used_timestamp", 0)
        
        # 计算时间衰减因子（最近使用的应用得分更高）
        current_time = int(time.time())
        time_diff = current_time - last_used_timestamp
        
        # 时间衰减因子：使用次数乘以时间衰减权重
        # 时间衰减权重：随时间增加而减少，1小时内为1.0，1天后为0.5，7天后为0.1，30天后为0.05
        if time_diff < 3600:  # 1小时内
            time_weight = 1.0
        elif time_diff < 86400:  # 1天内
            time_weight = 0.7
        elif time_diff < 604800:  # 7天内
            time_weight = 0.3
        elif time_diff < 2592000:  # 30天内
            time_weight = 0.1
        else:
            time_weight = 0.05
        
        # 计算最终得分
        score = count * time_weight
        
        return score
    
    def sort_apps_by_usage(self, apps_list: List[Tuple[str, Any]]) -> List[Tuple[str, Any]]:
        """
        根据应用使用情况对应用列表进行排序
        
        Args:
            apps_list: 应用列表
            
        Returns:
            list: 排序后的应用列表
        """
        if not CONFIG["ENABLE_USAGE_PRIORITY"]:
            return apps_list
        
        # 计算每个应用的使用得分并排序
        sorted_list = sorted(
            apps_list,
            key=lambda x: self.get_usage_score(x[0]),
            reverse=True
        )
        
        return sorted_list

# 初始化应用使用记录器实例
global_usage_tracker = AppUsageTracker()

# 提供模块级函数接口
def record_usage(app_name: str) -> bool:
    """记录应用使用情况"""
    return global_usage_tracker.record_usage(app_name)

def get_usage_data(app_name: str = None) -> Any:
    """获取应用使用数据"""
    return global_usage_tracker.get_usage_data(app_name)

def get_most_used_apps(limit: int = 10) -> List[Tuple[str, Dict[str, Any]]]:
    """获取使用频率最高的应用列表"""
    return global_usage_tracker.get_most_used_apps(limit)

def get_recently_used_apps(limit: int = 10) -> List[Tuple[str, Dict[str, Any]]]:
    """获取最近使用的应用列表"""
    return global_usage_tracker.get_recently_used_apps(limit)

def clear_usage_data(app_name: str = None) -> bool:
    """清除使用记录数据"""
    return global_usage_tracker.clear_usage_data(app_name)

def sort_apps_by_usage(apps_list: List[Tuple[str, Any]]) -> List[Tuple[str, Any]]:
    """根据应用使用情况对应用列表进行排序"""
    return global_usage_tracker.sort_apps_by_usage(apps_list)

# 测试代码
if __name__ == "__main__":
    print("应用使用记录器测试")
    
    # 模拟使用几个应用
    test_apps = ["Google Chrome", "Microsoft Word", "Visual Studio Code", "Google Chrome", "Microsoft Word", "Google Chrome"]
    
    for app in test_apps:
        record_usage(app)
    
    # 获取使用情况
    print("\n使用频率最高的应用:")
    most_used = get_most_used_apps()
    for app_name, app_usage in most_used:
        print(f"{app_name}: 使用 {app_usage['count']} 次，最后使用于 {app_usage['last_used']}")
    
    # 模拟时间流逝（10分钟后）
    print("\n模拟10分钟后...")
    record_usage("Visual Studio Code")
    
    # 再次获取使用情况
    print("\n最近使用的应用:")
    recent_used = get_recently_used_apps()
    for app_name, app_usage in recent_used:
        print(f"{app_name}: 使用 {app_usage['count']} 次，最后使用于 {app_usage['last_used']}")
    
    # 测试清除记录
    print("\n清除Microsoft Word的使用记录")
    clear_usage_data("Microsoft Word")
    
    print("\n清除后的使用频率最高的应用:")
    most_used = get_most_used_apps()
    for app_name, app_usage in most_used:
        print(f"{app_name}: 使用 {app_usage['count']} 次")