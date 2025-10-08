#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MCP工具专用日志系统
为每个MCP工具提供独立的日志功能
"""
import os
import logging
import datetime
from typing import Optional

class MCPToolLogger:
    """MCP工具专用日志记录器"""
    
    def __init__(self, tool_name: str, log_dir: Optional[str] = None):
        """
        初始化MCP工具日志记录器
        
        Args:
            tool_name: 工具名称
            log_dir: 日志目录，默认为用户目录下的logs/mcp_tools
        """
        self.tool_name = tool_name
        
        # 设置日志目录
        if log_dir is None:
            user_home = os.path.expanduser("~")
            log_dir = os.path.join(user_home, "logs", "mcp_tools")
        
        os.makedirs(log_dir, exist_ok=True)
        
        # 创建日志文件名（按日期和工具名）
        today = datetime.datetime.now().strftime('%Y%m%d')
        log_filename = f"{tool_name}_{today}.log"
        log_file_path = os.path.join(log_dir, log_filename)
        
        # 创建logger
        self.logger = logging.getLogger(f"mcp_tool_{tool_name}")
        self.logger.setLevel(logging.DEBUG)
        
        # 清除已有的处理器，避免重复日志
        if self.logger.handlers:
            self.logger.handlers.clear()
        
        # 创建文件处理器
        file_handler = logging.FileHandler(log_file_path, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        
        # 创建控制台处理器（仅输出到stderr，避免影响MCP通信）
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        
        # 创建格式化器
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)
        
        # 添加处理器
        self.logger.addHandler(file_handler)
        self.logger.addHandler(console_handler)
        
        # 记录工具启动
        self.logger.info(f"MCP工具 '{tool_name}' 日志系统已启动")
    
    def info(self, message: str, **kwargs):
        """记录信息日志"""
        self.logger.info(f"[{self.tool_name}] {message}", **kwargs)
    
    def debug(self, message: str, **kwargs):
        """记录调试日志"""
        self.logger.debug(f"[{self.tool_name}] {message}", **kwargs)
    
    def warning(self, message: str, **kwargs):
        """记录警告日志"""
        self.logger.warning(f"[{self.tool_name}] {message}", **kwargs)
    
    def error(self, message: str, **kwargs):
        """记录错误日志"""
        self.logger.error(f"[{self.tool_name}] {message}", **kwargs)
    
    def critical(self, message: str, **kwargs):
        """记录严重错误日志"""
        self.logger.critical(f"[{self.tool_name}] {message}", **kwargs)
    
    def log_tool_call(self, tool_name: str, arguments: dict, start_time: datetime.datetime = None):
        """记录工具调用开始"""
        if start_time is None:
            start_time = datetime.datetime.now()
        
        self.info(f"工具调用开始: {tool_name}")
        self.debug(f"调用参数: {arguments}")
        self.debug(f"调用时间: {start_time.strftime('%Y-%m-%d %H:%M:%S.%f')}")
    
    def log_tool_result(self, tool_name: str, result: any, end_time: datetime.datetime = None, duration: float = None):
        """记录工具调用结果"""
        if end_time is None:
            end_time = datetime.datetime.now()
        
        self.info(f"工具调用完成: {tool_name}")
        self.debug(f"完成时间: {end_time.strftime('%Y-%m-%d %H:%M:%S.%f')}")
        
        if duration is not None:
            self.info(f"执行耗时: {duration:.2f}秒")
        
        # 记录结果摘要
        if isinstance(result, list) and len(result) > 0:
            if hasattr(result[0], 'text'):
                # TextContent对象
                result_text = result[0].text
                if len(result_text) > 200:
                    result_summary = result_text[:200] + "..."
                else:
                    result_summary = result_text
                self.debug(f"返回结果摘要: {result_summary}")
            else:
                self.debug(f"返回结果类型: {type(result[0])}")
        else:
            self.debug(f"返回结果: {result}")
    
    def log_tool_error(self, tool_name: str, error: Exception, end_time: datetime.datetime = None):
        """记录工具调用错误"""
        if end_time is None:
            end_time = datetime.datetime.now()
        
        self.error(f"工具调用失败: {tool_name}")
        self.error(f"错误类型: {type(error).__name__}")
        self.error(f"错误信息: {str(error)}")
        self.error(f"失败时间: {end_time.strftime('%Y-%m-%d %H:%M:%S.%f')}")
        
        # 记录详细错误信息
        import traceback
        error_traceback = traceback.format_exc()
        self.error(f"错误堆栈: {error_traceback}")

# 全局日志记录器字典
_tool_loggers = {}

def get_tool_logger(tool_name: str) -> MCPToolLogger:
    """
    获取指定工具的日志记录器
    
    Args:
        tool_name: 工具名称
        
    Returns:
        MCPToolLogger实例
    """
    if tool_name not in _tool_loggers:
        _tool_loggers[tool_name] = MCPToolLogger(tool_name)
    
    return _tool_loggers[tool_name]

def cleanup_loggers():
    """清理所有日志记录器"""
    global _tool_loggers
    for logger in _tool_loggers.values():
        for handler in logger.logger.handlers:
            handler.close()
    _tool_loggers.clear()
