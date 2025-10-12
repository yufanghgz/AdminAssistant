#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
本地应用智能打开模块 - MCP工具接口
提供统一的对外接口函数，整合所有模块功能
"""
import os
import sys
import time
from typing import Dict, List, Any, Optional
from .config import CONFIG
from .app_scanner import scan_installed_apps
from .app_cache import load_cache, update_cache, get_apps_index, clear_cache
from .nlp_parser import parse_query, add_custom_alias
from .app_executor import execute_app, is_app_running
from .app_usage import record_usage, get_most_used_apps, get_recently_used_apps
from ..mcp_logger import MCPToolLogger

logger = MCPToolLogger("open_app_tool")

# 全局标志，指示是否已初始化
g_initialized = False

# 初始化函数
def initialize_apps() -> Dict[str, Any]:
    """
    启动时扫描并加载缓存
    
    Returns:
        dict: 初始化结果，包含成功状态和消息
    """
    global g_initialized
    logger.info("开始初始化应用索引")
    
    start_time = time.time()
    
    try:
        # 尝试加载缓存
        cache_loaded = load_cache()
        
        # 如果缓存加载成功，检查是否有足够的应用
        if cache_loaded:
            apps_index = get_apps_index()
            if len(apps_index) > 0:
                g_initialized = True
                elapsed_time = time.time() - start_time
                logger.info(f"应用索引初始化成功（从缓存加载），共 {len(apps_index)} 个应用，耗时 {elapsed_time:.2f} 秒")
                return {
                    "success": True,
                    "message": f"应用索引初始化成功（从缓存加载），共 {len(apps_index)} 个应用",
                    "app_count": len(apps_index),
                    "from_cache": True
                }
        
        # 缓存加载失败或应用数量不足，执行扫描
        logger.info("缓存加载失败或应用数量不足，执行应用扫描")
        
        # 更新缓存
        cache_updated = update_cache()
        
        if cache_updated:
            apps_index = get_apps_index()
            g_initialized = True
            elapsed_time = time.time() - start_time
            logger.info(f"应用索引初始化成功（执行扫描），共 {len(apps_index)} 个应用，耗时 {elapsed_time:.2f} 秒")
            return {
                "success": True,
                "message": f"应用索引初始化成功（执行扫描），共 {len(apps_index)} 个应用",
                "app_count": len(apps_index),
                "from_cache": False
            }
        else:
            # 扫描失败，尝试直接扫描但不保存缓存
            logger.warning("更新缓存失败，尝试直接扫描应用")
            apps_index = scan_installed_apps()
            if len(apps_index) > 0:
                g_initialized = True
                elapsed_time = time.time() - start_time
                logger.info(f"应用索引初始化成功（直接扫描），共 {len(apps_index)} 个应用，耗时 {elapsed_time:.2f} 秒")
                return {
                    "success": True,
                    "message": f"应用索引初始化成功（直接扫描），共 {len(apps_index)} 个应用",
                    "app_count": len(apps_index),
                    "from_cache": False
                }
            else:
                elapsed_time = time.time() - start_time
                logger.error(f"应用索引初始化失败，耗时 {elapsed_time:.2f} 秒")
                return {
                    "success": False,
                    "message": "应用索引初始化失败，无法扫描到应用",
                    "app_count": 0,
                    "from_cache": False
                }
    except Exception as e:
        elapsed_time = time.time() - start_time
        logger.error(f"应用索引初始化异常: {str(e)}，耗时 {elapsed_time:.2f} 秒")
        return {
            "success": False,
            "message": f"应用索引初始化异常: {str(e)}",
            "app_count": 0,
            "from_cache": False
        }

def list_apps(sort_by: str = "name", limit: int = 100) -> Dict[str, Any]:
    """
    返回已检测到的所有应用信息
    
    Args:
        sort_by: 排序方式，可选值: "name"（按名称）, "usage"（按使用频率）, "recent"（按最近使用）
        limit: 返回的最大应用数量
        
    Returns:
        dict: 应用列表及元数据
    """
    logger.info(f"获取应用列表，排序方式: {sort_by}，限制数量: {limit}")
    
    try:
        # 确保已初始化
        if not g_initialized:
            logger.warning("应用索引尚未初始化，自动执行初始化")
            init_result = initialize_apps()
            if not init_result["success"]:
                return {
                    "success": False,
                    "message": "应用索引尚未初始化",
                    "apps": []
                }
        
        # 获取应用索引
        apps_index = get_apps_index()
        
        # 构建应用列表
        apps_list = []
        for app_name, app_info in apps_index.items():
            app_data = {
                "name": app_name,
                "path": app_info.get("path", ""),
                "platform": app_info.get("platform", "unknown"),
                "scanned_from": app_info.get("scanned_from", "unknown"),
                "aliases": app_info.get("aliases", []),
                "is_running": is_app_running(app_name) if app_name else False
            }
            apps_list.append(app_data)
        
        # 排序应用列表
        if sort_by == "name":
            apps_list.sort(key=lambda x: x["name"].lower())
        elif sort_by == "usage":
            # 按使用频率排序
            most_used_apps = {app[0]: app[1] for app in get_most_used_apps(len(apps_list))}
            apps_list.sort(key=lambda x: most_used_apps.get(x["name"], {}).get("count", 0), reverse=True)
        elif sort_by == "recent":
            # 按最近使用排序
            recent_apps = {app[0]: app[1] for app in get_recently_used_apps(len(apps_list))}
            apps_list.sort(key=lambda x: recent_apps.get(x["name"], {}).get("last_used_timestamp", 0), reverse=True)
        
        # 限制返回数量
        limited_apps = apps_list[:limit]
        
        logger.info(f"成功获取应用列表，返回 {len(limited_apps)} 个应用")
        return {
            "success": True,
            "message": f"成功获取应用列表",
            "total_count": len(apps_list),
            "returned_count": len(limited_apps),
            "apps": limited_apps
        }
    except Exception as e:
        logger.error(f"获取应用列表失败: {str(e)}")
        return {
            "success": False,
            "message": f"获取应用列表失败: {str(e)}",
            "apps": []
        }

def search_app(query: str, max_results: int = 5) -> Dict[str, Any]:
    """
    根据自然语言查找目标应用
    
    Args:
        query: 自然语言查询字符串
        max_results: 最大返回结果数量
        
    Returns:
        dict: 搜索结果，包含匹配的应用信息
    """
    logger.info(f"搜索应用，查询: {query}，最大结果数: {max_results}")
    
    try:
        # 参数验证
        if not query:
            return {
                "success": False,
                "message": "查询字符串不能为空",
                "query": query,
                "results": []
            }
        
        # 确保已初始化
        if not g_initialized:
            logger.warning("应用索引尚未初始化，自动执行初始化")
            init_result = initialize_apps()
            if not init_result["success"]:
                return {
                    "success": False,
                    "message": "应用索引尚未初始化",
                    "query": query,
                    "results": []
                }
        
        # 获取应用索引
        apps_index = get_apps_index()
        
        # 使用NLP解析器解析查询
        parse_result = parse_query(query, apps_index)
        
        # 构建搜索结果
        search_results = []
        for app_name, confidence in parse_result.get("candidates", [])[:max_results]:
            app_info = apps_index.get(app_name, {})
            result_item = {
                "name": app_name,
                "path": app_info.get("path", ""),
                "confidence": float(confidence),  # 转换为标准Python float
                "is_running": is_app_running(app_name) if app_name else False
            }
            search_results.append(result_item)
        
        # 检查是否有最佳匹配
        best_match = parse_result.get("matched_app")
        
        logger.info(f"搜索完成，找到 {len(search_results)} 个匹配结果")
        return {
            "success": True,
            "message": f"搜索完成，找到 {len(search_results)} 个匹配结果",
            "query": query,
            "best_match": best_match,
            "results": search_results
        }
    except Exception as e:
        logger.error(f"搜索应用失败: {str(e)}")
        return {
            "success": False,
            "message": f"搜索应用失败: {str(e)}",
            "query": query,
            "results": []
        }

def open_app(app_name: str) -> Dict[str, Any]:
    """
    执行打开目标应用
    
    Args:
        app_name: 应用名称
        
    Returns:
        dict: 执行结果
    """
    logger.info(f"尝试打开应用: {app_name}")
    
    try:
        # 参数验证
        if not app_name:
            return {
                "success": False,
                "message": "应用名称不能为空",
                "app_name": app_name
            }
        
        # 确保已初始化
        if not g_initialized:
            logger.warning("应用索引尚未初始化，自动执行初始化")
            init_result = initialize_apps()
            if not init_result["success"]:
                return {
                    "success": False,
                    "message": "应用索引尚未初始化",
                    "app_name": app_name
                }
        
        # 获取应用索引
        apps_index = get_apps_index()
        
        # 检查应用是否存在
        if app_name not in apps_index:
            # 尝试搜索应用（模糊匹配）
            logger.info(f"应用 '{app_name}' 不存在于索引中，尝试搜索")
            search_result = search_app(app_name, max_results=1)
            
            if search_result["success"] and search_result["results"]:
                # 使用搜索到的第一个结果
                best_match = search_result["results"][0]["name"]
                logger.info(f"找到模糊匹配: {best_match}")
                app_name = best_match
            else:
                logger.error(f"未找到应用: {app_name}")
                return {
                    "success": False,
                    "message": f"未找到应用: {app_name}",
                    "app_name": app_name
                }
        
        # 获取应用信息
        app_info = apps_index[app_name]
        
        # 执行打开操作
        execute_result = execute_app(app_name, app_info)
        
        # 记录使用情况
        if execute_result["success"]:
            record_usage(app_name)
            logger.info(f"成功打开应用: {app_name}")
        else:
            logger.error(f"打开应用失败: {app_name}, 错误: {execute_result.get('message')}")
        
        # 丰富返回结果
        return {
            "success": execute_result["success"],
            "message": execute_result.get("message", ""),
            "app_name": app_name,
            "app_path": app_info.get("path", ""),
            "is_running": is_app_running(app_name) if execute_result["success"] else False
        }
    except Exception as e:
        logger.error(f"打开应用时发生异常: {str(e)}")
        return {
            "success": False,
            "message": f"打开应用时发生异常: {str(e)}",
            "app_name": app_name
        }

def reload_apps() -> Dict[str, Any]:
    """
    手动刷新扫描结果
    
    Returns:
        dict: 刷新结果，包含新索引表信息
    """
    global g_initialized
    logger.info("手动刷新应用索引")
    
    try:
        # 清除缓存
        clear_cache()
        
        # 重新扫描
        start_time = time.time()
        apps_index = scan_installed_apps()
        
        # 保存到缓存
        update_cache()
        
        g_initialized = True
        elapsed_time = time.time() - start_time
        
        logger.info(f"应用索引刷新成功，共 {len(apps_index)} 个应用，耗时 {elapsed_time:.2f} 秒")
        return {
            "success": True,
            "message": f"应用索引刷新成功，共 {len(apps_index)} 个应用",
            "app_count": len(apps_index),
            "elapsed_time": elapsed_time
        }
    except Exception as e:
        logger.error(f"应用索引刷新失败: {str(e)}")
        return {
            "success": False,
            "message": f"应用索引刷新失败: {str(e)}",
            "app_count": 0
        }

def get_tool_status() -> Dict[str, Any]:
    """
    获取工具当前状态
    
    Returns:
        dict: 工具状态信息
    """
    try:
        # 获取应用索引信息
        apps_index = get_apps_index()
        app_count = len(apps_index)
        
        # 获取使用情况信息
        most_used = get_most_used_apps(5)
        recent_used = get_recently_used_apps(5)
        
        return {
            "success": True,
            "initialized": g_initialized,
            "app_count": app_count,
            "cache_file": CONFIG["CACHE_FILE_PATH"],
            "cache_exists": os.path.exists(CONFIG["CACHE_FILE_PATH"]),
            "usage_tracking_enabled": CONFIG["ENABLE_USAGE_TRACKING"],
            "most_used_apps": [app[0] for app in most_used],
            "recently_used_apps": [app[0] for app in recent_used]
        }
    except Exception as e:
        logger.error(f"获取工具状态失败: {str(e)}")
        return {
            "success": False,
            "message": f"获取工具状态失败: {str(e)}"
        }

# 模块初始化函数
def _init_module():
    """模块初始化函数，在导入时执行"""
    global g_initialized
    
    # 如果配置了自动初始化，则在导入时执行初始化
    if CONFIG.get("AUTO_INITIALIZE", False):
        initialize_apps()
    
    logger.info("本地应用智能打开工具模块已加载")

# 模块导入时自动执行初始化
# _init_module()  # 注释掉自动初始化，避免导入时耗时过长

# 测试代码
if __name__ == "__main__":
    print("本地应用智能打开工具测试")
    
    # 初始化
    print("\n1. 初始化应用索引...")
    init_result = initialize_apps()
    print(f"初始化结果: {init_result}")
    
    # 查看工具状态
    print("\n2. 查看工具状态...")
    status = get_tool_status()
    print(f"工具状态: {status}")
    
    # 列出应用
    print("\n3. 列出前5个应用...")
    apps_list = list_apps(limit=5)
    print(f"应用列表: {apps_list}")
    
    # 搜索应用
    print("\n4. 搜索浏览器...")
    search_result = search_app("浏览器", max_results=3)
    print(f"搜索结果: {search_result}")
    
    # 注意：实际打开应用的测试被注释掉，避免意外打开应用
    # print("\n5. 尝试打开第一个搜索结果...")
    # if search_result['success'] and search_result['results']:
    #     app_to_open = search_result['results'][0]['name']
    #     open_result = open_app(app_to_open)
    #     print(f"打开结果: {open_result}")
    
    print("\n测试完成")