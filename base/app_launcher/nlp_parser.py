#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
本地应用智能打开模块 - 语义匹配器
解析自然语言输入，实现别名映射与模糊匹配
"""
import re
import difflib
from typing import Dict, List, Tuple, Optional, Any
from .config import CONFIG, APP_ALIASES
from ..mcp_logger import MCPToolLogger

logger = MCPToolLogger("nlp_parser")

class NlpParser:
    """自然语言解析器类，负责解析用户输入并匹配应用"""
    
    def __init__(self):
        """初始化自然语言解析器"""
        self.logger = logger
        self.app_aliases = APP_ALIASES.copy()
        self.max_candidates = CONFIG["MAX_CANDIDATES"]
        self.fuzzy_threshold = CONFIG["FUZZY_MATCH_THRESHOLD"]
    
    def parse_query(self, query: str, apps_index: Dict[str, Dict[str, Any]]) -> Dict[str, Any]:
        """
        解析用户输入的自然语言查询，匹配目标应用
        
        Args:
            query: 用户输入的查询字符串
            apps_index: 应用索引表
            
        Returns:
            dict: 解析结果，包含匹配的应用信息和置信度
        """
        if not query or not apps_index:
            return {
                "query": query,
                "matched_app": None,
                "confidence": 0.0,
                "candidates": []
            }
        
        self.logger.info(f"解析查询: {query}")
        
        # 预处理查询
        processed_query = self._preprocess_query(query)
        
        # 提取关键词
        keywords = self._extract_keywords(processed_query)
        self.logger.debug(f"提取的关键词: {keywords}")
        
        # 匹配应用
        candidates = self._match_apps(keywords, apps_index)
        
        # 排序候选项
        sorted_candidates = self._sort_candidates(candidates)
        
        # 选择最佳匹配
        if sorted_candidates:
            best_match = sorted_candidates[0]
            result = {
                "query": query,
                "matched_app": best_match[0],
                "confidence": best_match[1],
                "candidates": sorted_candidates[:self.max_candidates]
            }
        else:
            result = {
                "query": query,
                "matched_app": None,
                "confidence": 0.0,
                "candidates": []
            }
        
        self.logger.info(f"解析结果: {result}")
        return result
    
    def _preprocess_query(self, query: str) -> str:
        """预处理查询字符串"""
        # 转换为小写
        query = query.lower()
        
        # 移除常见的前缀动词（如"打开"、"启动"、"运行"等）
        action_words = ["打开", "启动", "运行", "开启", "启动", "打开", "运行", "launch", "open", "start", "run"]
        for word in action_words:
            if query.startswith(word):
                query = query[len(word):].strip()
                break
        
        # 移除标点符号和多余空格
        query = re.sub(r'[，。！？,!.?]', ' ', query)
        query = re.sub(r'\s+', ' ', query).strip()
        
        return query
    
    def _extract_keywords(self, query: str) -> List[str]:
        """从查询中提取关键词"""
        if not query:
            return []
        
        # 简单分词（适用于中文和英文）
        # 中文分词：按字符切分
        # 英文分词：按空格切分
        
        # 检查是否包含中文
        has_chinese = bool(re.search(r'[\u4e00-\u9fa5]', query))
        
        keywords = []
        
        if has_chinese:
            # 中文分词（简单实现）
            # 首先尝试匹配完整的应用别名
            for app_name, aliases in self.app_aliases.items():
                for alias in aliases:
                    if alias in query:
                        keywords.append(alias)
                        query = query.replace(alias, ' ')
            
            # 然后按字符切分剩余部分
            keywords.extend([char for char in query if char.strip()])
        else:
            # 英文分词：按空格切分
            keywords = query.split()
        
        return list(set(keywords))  # 去重
    
    def _match_apps(self, keywords: List[str], apps_index: Dict[str, Dict[str, Any]]) -> List[Tuple[str, float]]:
        """根据关键词匹配应用"""
        if not keywords or not apps_index:
            return []
        
        candidates = []
        
        # 遍历所有应用
        for app_name, app_info in apps_index.items():
            confidence = 0.0
            
            # 1. 尝试精确匹配
            if self._exact_match(keywords, app_name, app_info):
                confidence = 1.0
            # 2. 尝试别名匹配
            elif CONFIG["ENABLE_ALIAS_EXPANSION"] and self._alias_match(keywords, app_name):
                confidence = 0.9
            # 3. 尝试模糊匹配
            elif CONFIG["ENABLE_FUZZY_MATCH"]:
                confidence = self._fuzzy_match(keywords, app_name)
            
            # 如果匹配度超过阈值，则添加到候选项
            if confidence >= self.fuzzy_threshold:
                candidates.append((app_name, confidence))
        
        return candidates
    
    def _exact_match(self, keywords: List[str], app_name: str, app_info: Dict[str, Any]) -> bool:
        """检查是否精确匹配"""
        app_name_lower = app_name.lower()
        
        # 检查应用名称是否包含所有关键词
        for keyword in keywords:
            if keyword.lower() not in app_name_lower:
                return False
        
        return True
    
    def _alias_match(self, keywords: List[str], app_name: str) -> bool:
        """检查是否匹配别名"""
        # 获取应用的别名（如果有）
        app_lower = app_name.lower()
        app_aliases = []
        
        # 从预定义的别名表中查找
        for app_key, aliases in self.app_aliases.items():
            if app_key.lower() in app_lower or app_lower in app_key.lower():
                app_aliases.extend(aliases)
                break
        
        # 检查关键词是否匹配别名
        for keyword in keywords:
            if any(keyword in alias for alias in app_aliases):
                return True
        
        return False
    
    def _fuzzy_match(self, keywords: List[str], app_name: str) -> float:
        """计算关键词与应用名称的模糊匹配度"""
        max_score = 0.0
        
        for keyword in keywords:
            # 使用 difflib 计算相似度
            ratio = difflib.SequenceMatcher(None, keyword.lower(), app_name.lower()).ratio()
            max_score = max(max_score, ratio)
            
            # 如果已经找到较高的匹配度，可以提前返回
            if max_score >= 0.9:
                break
        
        return max_score
    
    def _sort_candidates(self, candidates: List[Tuple[str, float]]) -> List[Tuple[str, float]]:
        """排序候选应用列表"""
        # 按置信度降序排序
        return sorted(candidates, key=lambda x: x[1], reverse=True)
    
    def add_custom_alias(self, app_name: str, alias: str) -> bool:
        """
        添加自定义应用别名
        
        Args:
            app_name: 应用名称
            alias: 别名
            
        Returns:
            bool: 添加是否成功
        """
        try:
            # 规范化应用名称
            app_name_lower = app_name.lower()
            
            # 查找是否已存在该应用的别名
            found = False
            for app_key in self.app_aliases:
                if app_key.lower() == app_name_lower:
                    if alias not in self.app_aliases[app_key]:
                        self.app_aliases[app_key].append(alias)
                    found = True
                    break
            
            # 如果不存在，则创建新的应用别名条目
            if not found:
                self.app_aliases[app_name_lower] = [alias]
            
            self.logger.info(f"为应用 '{app_name}' 添加别名: '{alias}'")
            return True
        except Exception as e:
            self.logger.error(f"添加别名失败: {str(e)}")
            return False

# 初始化NLP解析器实例
global_nlp_parser = NlpParser()

# 提供模块级函数接口
def parse_query(query: str, apps_index: Dict[str, Dict[str, Any]]) -> Dict[str, Any]:
    """解析用户输入的自然语言查询，匹配目标应用"""
    return global_nlp_parser.parse_query(query, apps_index)

def add_custom_alias(app_name: str, alias: str) -> bool:
    """添加自定义应用别名"""
    return global_nlp_parser.add_custom_alias(app_name, alias)

# 测试代码
if __name__ == "__main__":
    # 创建测试应用索引
    test_apps_index = {
        "Google Chrome": {"path": "/Applications/Google Chrome.app"},
        "Microsoft Word": {"path": "/Applications/Microsoft Word.app"},
        "Visual Studio Code": {"path": "/Applications/Visual Studio Code.app"}
    }
    
    # 测试不同的查询
    test_queries = ["打开浏览器", "启动Word", "运行VSCode", "打开音乐播放器"]
    
    for query in test_queries:
        result = parse_query(query, test_apps_index)
        print(f"查询: {query}")
        print(f"匹配结果: {result}")
        print()