#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
本地应用智能打开模块 - 应用执行器
校验路径安全并执行打开操作
"""
import os
import sys
import subprocess
import platform
from typing import Dict, Any, Optional
from .config import CURRENT_PLATFORM
from ..mcp_logger import MCPToolLogger

logger = MCPToolLogger("app_executor")

class AppExecutor:
    """应用执行器类，负责安全地执行应用程序"""
    
    def __init__(self):
        """初始化应用执行器"""
        self.logger = logger
        self.executed_apps = []  # 记录已执行的应用
    
    def execute_app(self, app_name: str, app_info: Dict[str, Any]) -> Dict[str, Any]:
        """
        执行打开应用程序的操作
        
        Args:
            app_name: 应用名称
            app_info: 应用信息，包含路径等
            
        Returns:
            dict: 执行结果，包含成功状态、消息等
        """
        self.logger.info(f"尝试执行应用: {app_name}")
        
        # 参数验证
        if not app_name or not app_info:
            return {
                "success": False,
                "app_name": app_name,
                "message": "应用名称或信息为空"
            }
        
        # 获取应用路径
        app_path = app_info.get("path")
        if not app_path:
            return {
                "success": False,
                "app_name": app_name,
                "message": "应用路径未提供"
            }
        
        # 安全检查
        if not self._security_check(app_path):
            return {
                "success": False,
                "app_name": app_name,
                "message": "应用路径安全检查失败"
            }
        
        try:
            # 执行打开操作
            if CURRENT_PLATFORM == "windows":
                result = self._execute_on_windows(app_path)
            elif CURRENT_PLATFORM == "macos":
                result = self._execute_on_macos(app_path)
            else:
                result = {
                    "success": False,
                    "app_name": app_name,
                    "message": f"不支持的操作系统: {CURRENT_PLATFORM}"
                }
            
            # 更新执行记录
            if result["success"]:
                self.executed_apps.append({
                    "app_name": app_name,
                    "app_path": app_path,
                    "timestamp": os.path.getmtime(app_path) if os.path.exists(app_path) else os.path.getctime(app_path)
                })
                self.logger.info(f"成功执行应用: {app_name} (路径: {app_path})")
            else:
                self.logger.error(f"执行应用失败: {app_name}, 错误: {result.get('message')}")
            
            return result
        except Exception as e:
            self.logger.error(f"执行应用时发生异常: {str(e)}")
            return {
                "success": False,
                "app_name": app_name,
                "message": f"执行过程中发生异常: {str(e)}"
            }
    
    def _security_check(self, app_path: str) -> bool:
        """
        安全检查应用路径
        
        Args:
            app_path: 应用路径
            
        Returns:
            bool: 是否通过安全检查
        """
        # 1. 检查路径是否存在
        if not os.path.exists(app_path):
            self.logger.warning(f"应用路径不存在: {app_path}")
            return False
        
        # 2. 检查路径是否是绝对路径
        if not os.path.isabs(app_path):
            self.logger.warning(f"应用路径不是绝对路径: {app_path}")
            return False
        
        # 3. 检查路径是否在安全目录内
        safe_dirs = []
        if CURRENT_PLATFORM == "windows":
            safe_dirs = ["C:\\Program Files", "C:\\Program Files (x86)", os.path.expanduser("~\\AppData")]
        elif CURRENT_PLATFORM == "macos":
            safe_dirs = ["/Applications", os.path.expanduser("~/Applications"), "/System/Applications"]
        
        if safe_dirs:
            in_safe_dir = False
            for safe_dir in safe_dirs:
                if app_path.startswith(safe_dir):
                    in_safe_dir = True
                    break
            
            # 如果不在安全目录内，记录警告但允许执行（可能是用户自定义安装位置）
            if not in_safe_dir:
                self.logger.warning(f"应用路径不在安全目录内: {app_path}")
        
        # 4. 检查文件是否可执行
        if CURRENT_PLATFORM != "windows":  # Windows不使用此检查
            if not os.access(app_path, os.X_OK):
                self.logger.warning(f"应用文件不可执行: {app_path}")
                # 对于macOS的.app文件夹，我们仍然允许尝试打开
                if not (CURRENT_PLATFORM == "macos" and app_path.endswith('.app')):
                    return False
        
        # 5. 检查是否包含危险字符（如命令注入）
        dangerous_patterns = [';', '&&', '||', '|', '&', '`', '$', '\\', '>', '<']
        for pattern in dangerous_patterns:
            if pattern in app_path:
                self.logger.warning(f"应用路径包含危险字符: {pattern}")
                return False
        
        self.logger.debug(f"应用路径通过安全检查: {app_path}")
        return True
    
    def _execute_on_windows(self, app_path: str) -> Dict[str, Any]:
        """在Windows系统上执行应用程序"""
        try:
            # 使用subprocess.Popen启动应用
            # shell=True表示使用shell执行，这对于.lnk文件很重要
            shell = app_path.endswith('.lnk')
            subprocess.Popen([app_path], shell=shell, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            
            return {
                "success": True,
                "app_name": os.path.basename(app_path),
                "message": f"已成功在Windows上启动应用: {app_path}"
            }
        except Exception as e:
            self.logger.error(f"Windows上启动应用失败: {str(e)}")
            return {
                "success": False,
                "app_name": os.path.basename(app_path),
                "message": f"Windows上启动应用失败: {str(e)}"
            }
    
    def _execute_on_macos(self, app_path: str) -> Dict[str, Any]:
        """在macOS系统上执行应用程序"""
        try:
            # 使用open命令打开应用
            if app_path.endswith('.app'):
                # 对于.app文件夹，使用-a参数
                subprocess.run(['open', '-a', app_path], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            else:
                # 对于其他可执行文件，直接打开
                subprocess.run(['open', app_path], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            
            return {
                "success": True,
                "app_name": os.path.basename(app_path),
                "message": f"已成功在macOS上启动应用: {app_path}"
            }
        except Exception as e:
            self.logger.error(f"macOS上启动应用失败: {str(e)}")
            return {
                "success": False,
                "app_name": os.path.basename(app_path),
                "message": f"macOS上启动应用失败: {str(e)}"
            }
    
    def is_app_running(self, app_name: str) -> Optional[bool]:
        """
        检查应用是否正在运行（简化实现）
        
        Args:
            app_name: 应用名称
            
        Returns:
            bool: 应用是否正在运行，None表示无法确定
        """
        try:
            if CURRENT_PLATFORM == "windows":
                # Windows上检查进程
                tasklist = subprocess.run(['tasklist'], capture_output=True, text=True).stdout
                return app_name.lower() in tasklist.lower()
            elif CURRENT_PLATFORM == "macos":
                # macOS上检查进程
                ps = subprocess.run(['ps', '-ax'], capture_output=True, text=True).stdout
                return app_name.lower() in ps.lower()
            else:
                self.logger.warning(f"无法检查应用运行状态，不支持的操作系统: {CURRENT_PLATFORM}")
                return None
        except Exception as e:
            self.logger.error(f"检查应用运行状态失败: {str(e)}")
            return None
    
    def get_executed_apps(self) -> list:
        """获取已执行的应用列表"""
        return self.executed_apps.copy()

# 初始化应用执行器实例
global_app_executor = AppExecutor()

# 提供模块级函数接口
def execute_app(app_name: str, app_info: Dict[str, Any]) -> Dict[str, Any]:
    """执行打开应用程序的操作"""
    return global_app_executor.execute_app(app_name, app_info)

def is_app_running(app_name: str) -> Optional[bool]:
    """检查应用是否正在运行"""
    return global_app_executor.is_app_running(app_name)

def get_executed_apps() -> list:
    """获取已执行的应用列表"""
    return global_app_executor.get_executed_apps()

# 测试代码
if __name__ == "__main__":
    # 注意：这个测试会实际打开应用，请谨慎运行
    print("应用执行器测试")
    
    # 创建测试应用信息
    test_app_info = {}
    
    if CURRENT_PLATFORM == "windows":
        # Windows测试应用
        notepad_path = "C:\\Windows\\System32\\notepad.exe"
        if os.path.exists(notepad_path):
            test_app_info = {"path": notepad_path}
    elif CURRENT_PLATFORM == "macos":
        # macOS测试应用
        textedit_path = "/Applications/TextEdit.app"
        if os.path.exists(textedit_path):
            test_app_info = {"path": textedit_path}
    
    if test_app_info:
        print(f"尝试打开测试应用: {test_app_info['path']}")
        result = execute_app("测试应用", test_app_info)
        print(f"执行结果: {result}")
        
        # 检查应用是否运行（可能需要一些时间才能启动）
        import time
        time.sleep(1)
        running = is_app_running("TextEdit" if CURRENT_PLATFORM == "macos" else "notepad")
        print(f"应用是否运行: {running}")
    else:
        print("未找到适合测试的应用")