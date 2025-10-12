#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
本地应用智能打开模块 - 应用扫描器
系统启动时检测本地已安装应用，生成索引表
"""
import os
import sys
import json
import platform
import subprocess
import re
from typing import List, Dict, Any
from .config import CONFIG, PLATFORM_SPECIFIC_CONFIG, CURRENT_PLATFORM
from ..mcp_logger import MCPToolLogger

logger = MCPToolLogger("app_scanner")

class AppScanner:
    """应用扫描器类，负责检测系统已安装的应用程序"""
    
    def __init__(self):
        """初始化应用扫描器"""
        self.logger = logger
        self.apps_index = {}
        self.scan_time = None
    
    def scan_installed_apps(self) -> Dict[str, Dict[str, Any]]:
        """
        扫描系统中已安装的应用程序
        
        Returns:
            dict: 应用程序索引表，格式为 {app_name: {path: xxx, aliases: [], ...}}
        """
        self.logger.info("开始扫描系统已安装应用")
        
        if CURRENT_PLATFORM == "windows":
            self.apps_index = self._scan_windows_apps()
        elif CURRENT_PLATFORM == "macos":
            self.apps_index = self._scan_macos_apps()
        else:
            self.logger.warning(f"不支持的操作系统: {CURRENT_PLATFORM}")
            self.apps_index = {}
        
        self.scan_time = os.path.getmtime(__file__) if os.path.exists(__file__) else os.path.getctime(__file__)
        self.logger.info(f"扫描完成，共发现 {len(self.apps_index)} 个应用")
        
        return self.apps_index
    
    def _scan_windows_apps(self) -> Dict[str, Dict[str, Any]]:
        """扫描Windows系统中的应用程序"""
        apps = {}
        
        # 从注册表扫描
        try:
            apps.update(self._scan_windows_registry())
        except Exception as e:
            self.logger.error(f"从注册表扫描应用失败: {str(e)}")
        
        # 从Program Files目录扫描
        try:
            apps.update(self._scan_windows_program_files())
        except Exception as e:
            self.logger.error(f"从Program Files目录扫描应用失败: {str(e)}")
        
        return apps
    
    def _scan_windows_registry(self) -> Dict[str, Dict[str, Any]]:
        """从Windows注册表扫描应用程序"""
        apps = {}
        
        # 这里简化实现，实际需要使用winreg模块读取注册表
        try:
            import winreg
            
            # 扫描普通应用路径
            reg_paths = PLATFORM_SPECIFIC_CONFIG.get("registry_paths", [])
            
            for reg_path in reg_paths:
                try:
                    key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path)
                    num_subkeys = winreg.QueryInfoKey(key)[0]
                    
                    for i in range(num_subkeys):
                        try:
                            app_name = winreg.EnumKey(key, i)
                            app_key = winreg.OpenKey(key, app_name)
                            
                            try:
                                # 获取应用路径
                                path = winreg.QueryValue(app_key, None)
                                if path and os.path.exists(path):
                                    apps[app_name] = {
                                        "path": path,
                                        "aliases": [],
                                        "platform": "windows",
                                        "scanned_from": "registry"
                                    }
                            except Exception:
                                pass
                            finally:
                                winreg.CloseKey(app_key)
                        except Exception:
                            continue
                    
                    winreg.CloseKey(key)
                except Exception:
                    continue
        except ImportError:
            self.logger.warning("无法导入winreg模块，跳过注册表扫描")
        
        return apps
    
    def _scan_windows_program_files(self) -> Dict[str, Dict[str, Any]]:
        """从Windows Program Files目录扫描应用程序"""
        apps = {}
        
        scan_paths = PLATFORM_SPECIFIC_CONFIG.get("scan_paths", [])
        extensions = PLATFORM_SPECIFIC_CONFIG.get("extensions", [".exe"])
        
        for scan_path in scan_paths:
            if not os.path.exists(scan_path):
                continue
            
            for root, _, files in os.walk(scan_path, topdown=True):
                # 限制目录深度，避免扫描过慢
                depth = len(os.path.normpath(root).split(os.sep)) - len(os.path.normpath(scan_path).split(os.sep))
                if depth > 4:
                    continue
                
                for file in files:
                    if any(file.lower().endswith(ext) for ext in extensions):
                        app_path = os.path.join(root, file)
                        app_name = os.path.splitext(file)[0]
                        
                        # 避免重复添加
                        if app_name not in apps:
                            apps[app_name] = {
                                "path": app_path,
                                "aliases": [],
                                "platform": "windows",
                                "scanned_from": "program_files"
                            }
        
        return apps
    
    def _scan_macos_apps(self) -> Dict[str, Dict[str, Any]]:
        """扫描macOS系统中的应用程序"""
        apps = {}
        
        # 使用Spotlight搜索
        try:
            apps.update(self._scan_macos_with_spotlight())
        except Exception as e:
            self.logger.error(f"使用Spotlight扫描应用失败: {str(e)}")
        
        # 从Applications目录扫描
        try:
            apps.update(self._scan_macos_applications_folder())
        except Exception as e:
            self.logger.error(f"从Applications目录扫描应用失败: {str(e)}")
        
        return apps
    
    def _scan_macos_with_spotlight(self) -> Dict[str, Dict[str, Any]]:
        """使用Spotlight搜索macOS应用程序"""
        apps = {}
        
        try:
            cmd = PLATFORM_SPECIFIC_CONFIG.get("spotlight_cmd", "mdfind 'kMDItemContentType == com.apple.application-bundle'")
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
            
            if result.returncode == 0:
                app_paths = result.stdout.strip().split('\n')
                
                for app_path in app_paths:
                    if os.path.exists(app_path) and app_path.endswith('.app'):
                        app_name = os.path.basename(app_path).replace('.app', '')
                        apps[app_name] = {
                            "path": app_path,
                            "aliases": [],
                            "platform": "macos",
                            "scanned_from": "spotlight"
                        }
        except Exception as e:
            self.logger.error(f"执行Spotlight命令失败: {str(e)}")
        
        return apps
    
    def _scan_macos_applications_folder(self) -> Dict[str, Dict[str, Any]]:
        """从macOS Applications目录扫描应用程序"""
        apps = {}
        
        scan_paths = PLATFORM_SPECIFIC_CONFIG.get("scan_paths", ["/Applications"])
        extensions = PLATFORM_SPECIFIC_CONFIG.get("extensions", [".app"])
        
        for scan_path in scan_paths:
            if not os.path.exists(scan_path):
                continue
            
            for item in os.listdir(scan_path):
                item_path = os.path.join(scan_path, item)
                if os.path.isdir(item_path) and any(item.lower().endswith(ext) for ext in extensions):
                    app_name = os.path.splitext(item)[0]
                    apps[app_name] = {
                        "path": item_path,
                        "aliases": [],
                        "platform": "macos",
                        "scanned_from": "applications_folder"
                    }
        
        return apps
    
    def get_apps_index(self) -> Dict[str, Dict[str, Any]]:
        """获取应用索引表"""
        return self.apps_index
    
    def get_scan_time(self) -> float:
        """获取扫描时间"""
        return self.scan_time

# 初始化扫描器实例
global_scanner = AppScanner()

# 提供模块级函数接口
def scan_installed_apps() -> Dict[str, Dict[str, Any]]:
    """扫描系统已安装的应用程序"""
    return global_scanner.scan_installed_apps()

# 测试代码
if __name__ == "__main__":
    print("开始扫描系统应用...")
    apps = scan_installed_apps()
    print(f"扫描完成，发现 {len(apps)} 个应用")
    
    # 打印前10个应用
    for i, (app_name, app_info) in enumerate(apps.items()):
        print(f"{i+1}. {app_name}: {app_info['path']}")
        if i >= 9:
            break