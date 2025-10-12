#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
本地应用智能打开模块 - 缓存管理
管理缓存文件（加载、验证、更新）
"""
import os
import json
import time
from typing import Dict, Any, Optional
from .config import CONFIG
from .app_scanner import scan_installed_apps
from ..mcp_logger import MCPToolLogger

logger = MCPToolLogger("app_cache")

class AppCacheManager:
    """应用缓存管理器类，负责管理应用索引的缓存"""
    
    def __init__(self):
        """初始化缓存管理器"""
        self.logger = logger
        self.cache_file = CONFIG["CACHE_FILE_PATH"]
        self.cache_valid_hours = CONFIG["CACHE_VALID_HOURS"]
        self.apps_index = {}
        self.cache_timestamp = None
    
    def load_cache(self) -> bool:
        """
        加载缓存文件
        
        Returns:
            bool: 缓存是否加载成功且有效
        """
        self.logger.info(f"尝试加载缓存文件: {self.cache_file}")
        
        if not os.path.exists(self.cache_file):
            self.logger.info("缓存文件不存在")
            return False
        
        try:
            with open(self.cache_file, 'r', encoding='utf-8') as f:
                cache_data = json.load(f)
            
            # 验证缓存数据格式
            if not self._validate_cache_format(cache_data):
                self.logger.warning("缓存文件格式无效")
                return False
            
            # 验证缓存是否过期
            if self._is_cache_expired(cache_data.get("timestamp", 0)):
                self.logger.info("缓存已过期")
                return False
            
            # 加载有效缓存
            self.apps_index = cache_data.get("apps", {})
            self.cache_timestamp = cache_data.get("timestamp", 0)
            self.logger.info(f"缓存加载成功，包含 {len(self.apps_index)} 个应用")
            return True
        except json.JSONDecodeError as e:
            self.logger.error(f"缓存文件解析错误: {str(e)}")
            return False
        except Exception as e:
            self.logger.error(f"加载缓存失败: {str(e)}")
            return False
    
    def save_cache(self, apps_index: Dict[str, Dict[str, Any]]) -> bool:
        """
        保存应用索引到缓存文件
        
        Args:
            apps_index: 应用索引表
            
        Returns:
            bool: 保存是否成功
        """
        self.logger.info(f"尝试保存缓存到文件: {self.cache_file}")
        
        try:
            # 确保缓存目录存在
            cache_dir = os.path.dirname(self.cache_file)
            if not os.path.exists(cache_dir):
                os.makedirs(cache_dir, exist_ok=True)
            
            # 构建缓存数据
            cache_data = {
                "timestamp": int(time.time()),
                "version": "1.0",
                "apps": apps_index
            }
            
            # 保存到文件
            with open(self.cache_file, 'w', encoding='utf-8') as f:
                json.dump(cache_data, f, ensure_ascii=False, indent=2)
            
            self.apps_index = apps_index
            self.cache_timestamp = cache_data["timestamp"]
            self.logger.info(f"缓存保存成功，共保存 {len(apps_index)} 个应用")
            return True
        except Exception as e:
            self.logger.error(f"保存缓存失败: {str(e)}")
            return False
    
    def update_cache(self) -> bool:
        """
        更新缓存（扫描应用并保存）
        
        Returns:
            bool: 更新是否成功
        """
        self.logger.info("开始更新缓存")
        
        try:
            # 扫描应用
            new_apps_index = scan_installed_apps()
            
            # 增量更新（如果启用）
            if CONFIG["ENABLE_INCREMENTAL_SCAN"] and self.apps_index:
                # 合并新旧缓存
                updated_apps_index = self._merge_cache(self.apps_index, new_apps_index)
                return self.save_cache(updated_apps_index)
            else:
                # 完全替换
                return self.save_cache(new_apps_index)
        except Exception as e:
            self.logger.error(f"更新缓存失败: {str(e)}")
            return False
    
    def get_apps_index(self) -> Dict[str, Dict[str, Any]]:
        """获取当前加载的应用索引表"""
        return self.apps_index
    
    def get_cache_timestamp(self) -> Optional[float]:
        """获取缓存的时间戳"""
        return self.cache_timestamp
    
    def clear_cache(self) -> bool:
        """
        清除缓存文件
        
        Returns:
            bool: 清除是否成功
        """
        self.logger.info(f"尝试清除缓存文件: {self.cache_file}")
        
        try:
            if os.path.exists(self.cache_file):
                os.remove(self.cache_file)
                self.logger.info("缓存文件已清除")
            else:
                self.logger.info("缓存文件不存在，无需清除")
            
            # 重置内部状态
            self.apps_index = {}
            self.cache_timestamp = None
            return True
        except Exception as e:
            self.logger.error(f"清除缓存失败: {str(e)}")
            return False
    
    def _validate_cache_format(self, cache_data: Any) -> bool:
        """验证缓存数据格式是否有效"""
        if not isinstance(cache_data, dict):
            return False
        
        required_keys = ["timestamp", "apps"]
        for key in required_keys:
            if key not in cache_data:
                return False
        
        if not isinstance(cache_data["timestamp"], (int, float)):
            return False
        
        if not isinstance(cache_data["apps"], dict):
            return False
        
        # 验证应用数据格式
        for app_name, app_info in cache_data["apps"].items():
            if not isinstance(app_info, dict) or "path" not in app_info:
                return False
        
        return True
    
    def _is_cache_expired(self, timestamp: float) -> bool:
        """判断缓存是否过期"""
        if self.cache_valid_hours <= 0:
            # 缓存永不过期
            return False
        
        current_time = time.time()
        cache_age = current_time - timestamp
        max_age = self.cache_valid_hours * 3600  # 转换为秒
        
        return cache_age > max_age
    
    def _merge_cache(self, old_cache: Dict[str, Dict[str, Any]], new_cache: Dict[str, Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
        """
        合并新旧缓存（增量更新）
        
        Args:
            old_cache: 旧缓存
            new_cache: 新扫描的缓存
            
        Returns:
            dict: 合并后的缓存
        """
        self.logger.info("执行增量更新缓存")
        
        # 以新缓存为主，保留旧缓存中的附加信息
        merged_cache = {}
        
        # 处理新缓存中的应用
        for app_name, new_info in new_cache.items():
            if app_name in old_cache:
                # 合并旧缓存中的附加信息
                merged_info = new_info.copy()
                merged_info.update(old_cache[app_name])
                merged_cache[app_name] = merged_info
            else:
                # 直接添加新应用
                merged_cache[app_name] = new_info
        
        # 记录变化
        new_apps_count = len(new_cache) - len(set(new_cache.keys()) & set(old_cache.keys()))
        removed_apps_count = len(old_cache) - len(set(new_cache.keys()) & set(old_cache.keys()))
        
        self.logger.info(f"增量更新完成: 新增 {new_apps_count} 个应用，移除 {removed_apps_count} 个应用")
        
        return merged_cache

# 初始化缓存管理器实例
global_cache_manager = AppCacheManager()

# 提供模块级函数接口
def load_cache() -> bool:
    """加载缓存文件"""
    return global_cache_manager.load_cache()

def save_cache(apps_index: Dict[str, Dict[str, Any]]) -> bool:
    """保存应用索引到缓存文件"""
    return global_cache_manager.save_cache(apps_index)

def update_cache() -> bool:
    """更新缓存"""
    return global_cache_manager.update_cache()

def get_apps_index() -> Dict[str, Dict[str, Any]]:
    """获取当前加载的应用索引表"""
    return global_cache_manager.get_apps_index()

def clear_cache() -> bool:
    """清除缓存文件"""
    return global_cache_manager.clear_cache()

# 测试代码
if __name__ == "__main__":
    # 尝试加载缓存
    if load_cache():
        print(f"缓存加载成功，包含 {len(get_apps_index())} 个应用")
    else:
        print("缓存加载失败，更新缓存...")
        if update_cache():
            print(f"缓存更新成功，包含 {len(get_apps_index())} 个应用")
        else:
            print("缓存更新失败")