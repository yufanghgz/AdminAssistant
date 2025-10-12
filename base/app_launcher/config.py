#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
本地应用智能打开模块 - 配置管理
定义全局配置项及优化参数
"""
import os

# 全局配置
CONFIG = {
    # 缓存相关
    "CACHE_VALID_HOURS": 24,
    "ENABLE_INCREMENTAL_SCAN": True,
    "CACHE_FILE_PATH": "~/.mcp/apps_cache.json",
    
    # 模糊与语义匹配
    "ENABLE_ALIAS_EXPANSION": True,
    "ENABLE_FUZZY_MATCH": True,
    "ENABLE_SEMANTIC_INTENT": True,
    "MAX_CANDIDATES": 3,
    "FUZZY_MATCH_THRESHOLD": 0.7,
    
    # 安全控制
    "ENABLE_USAGE_TRACKING": True,
    "ENABLE_USAGE_PRIORITY": True,
    "USAGE_FILE_PATH": "~/.mcp/app_usage.json",
    
    # 用户体验
    "ENABLE_PROGRESS_BAR": True,
    "LANGUAGE": "zh",
    "LOG_LEVEL": "info"
}

# 平台特定配置
PLATFORM_CONFIG = {
    "windows": {
        "scan_paths": [
            "C:\\Program Files",
            "C:\\Program Files (x86)",
            os.path.expanduser("~\\AppData\\Local\\Microsoft\\WindowsApps")
        ],
        "extensions": [".exe", ".lnk"],
        "registry_paths": [
            "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths",
            "SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\App Paths"
        ]
    },
    "macos": {
        "scan_paths": [
            "/Applications",
            os.path.expanduser("~/Applications")
        ],
        "extensions": [".app"],
        "spotlight_cmd": "mdfind 'kMDItemContentType == com.apple.application-bundle'"
    }
}

# 应用别名映射表
APP_ALIASES = {
    "chrome": ["谷歌浏览器", "浏览器", "上网"],
    "firefox": ["火狐浏览器", "火狐"],
    "safari": ["苹果浏览器"],
    "edge": ["微软浏览器", "edge浏览器"],
    "word": ["文字处理", "文档编辑"],
    "excel": ["电子表格", "表格处理"],
    "powerpoint": ["演示文稿", "幻灯片"],
    "outlook": ["邮件客户端"],
    "vscode": ["代码编辑器", "vs code"],
    "pycharm": ["python编辑器"],
    "intellij": ["idea"],
    "photoshop": ["ps", "图片编辑"],
    "illustrator": ["ai", "矢量图"],
    "spotify": ["音乐播放器"],
    "itunes": ["苹果音乐"],
    "vlc": ["视频播放器"],
    "zoom": ["视频会议"],
    "teams": ["团队协作"],
    "wechat": ["微信"],
    "qq": ["腾讯qq"],
    "whatsapp": ["聊天工具"]
}

import os
import sys

# 动态获取当前操作系统
def get_platform():
    """获取当前操作系统类型"""
    if sys.platform.startswith('win'):
        return "windows"
    elif sys.platform.startswith('darwin'):
        return "macos"
    else:
        # 不支持的操作系统
        return "unknown"

# 初始化时加载平台特定配置
CURRENT_PLATFORM = get_platform()
PLATFORM_SPECIFIC_CONFIG = PLATFORM_CONFIG.get(CURRENT_PLATFORM, {})

# 确保配置路径正确解析
def resolve_path(path):
    """解析路径中的用户目录符号"""
    if path.startswith('~'):
        return os.path.expanduser(path)
    return path

# 解析缓存和使用记录文件路径
CONFIG["CACHE_FILE_PATH"] = resolve_path(CONFIG["CACHE_FILE_PATH"])
CONFIG["USAGE_FILE_PATH"] = resolve_path(CONFIG["USAGE_FILE_PATH"])

# 确保缓存目录存在
def ensure_cache_dir():
    """确保缓存目录存在"""
    cache_dir = os.path.dirname(CONFIG["CACHE_FILE_PATH"])
    if not os.path.exists(cache_dir):
        os.makedirs(cache_dir, exist_ok=True)
    
    usage_dir = os.path.dirname(CONFIG["USAGE_FILE_PATH"])
    if not os.path.exists(usage_dir):
        os.makedirs(usage_dir, exist_ok=True)

# 初始化
ensure_cache_dir()