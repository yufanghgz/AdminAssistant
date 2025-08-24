#!/usr/bin/env python3
"""
待办事项列表 MCP 服务器示例
"""
import asyncio
import json
import os
from typing import List, Dict, Any
from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent, ImageContent

# 创建 Server 实例
server = Server("todo-server")

# 待办事项存储文件
TODO_FILE = "todos.json"

def load_todos() -> List[Dict[str, Any]]:
    """从文件加载待办事项"""
    if os.path.exists(TODO_FILE):
        try:
            with open(TODO_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, FileNotFoundError):
            return []
    return []

def save_todos(todos: List[Dict[str, Any]]):
    """保存待办事项到文件"""
    with open(TODO_FILE, 'w', encoding='utf-8') as f:
        json.dump(todos, f, ensure_ascii=False, indent=2)

@server.list_tools()
async def list_tools() -> List[Tool]:
    """列出服务器提供的所有工具"""
    return [
        Tool(
            name="add_todo",
            description="添加一个新的待办事项",
            inputSchema={
                "type": "object",
                "properties": {
                    "task": {
                        "type": "string",
                        "description": "待办事项的描述"
                    },
                    "priority": {
                        "type": "string",
                        "enum": ["high", "medium", "low"],
                        "description": "优先级（high/medium/low）",
                        "default": "medium"
                    }
                },
                "required": ["task"]
            }
        ),
        Tool(
            name="list_todos",
            description="列出所有的待办事项",
            inputSchema={
                "type": "object",
                "properties": {}
            }
        ),
        Tool(
            name="complete_todo",
            description="标记一个待办事项为完成状态",
            inputSchema={
                "type": "object",
                "properties": {
                    "index": {
                        "type": "integer",
                        "description": "待办事项的索引（从1开始）"
                    }
                },
                "required": ["index"]
            }
        ),
        Tool(
            name="clear_completed",
            description="清除所有已完成的待办事项",
            inputSchema={
                "type": "object",
                "properties": {}
            }
        )
    ]

@server.call_tool()
async def call_tool(name: str, arguments: dict) -> List[TextContent]:
    """处理工具调用"""
    todos = load_todos()
    
    if name == "add_todo":
        new_todo = {
            "task": arguments["task"],
            "priority": arguments.get("priority", "medium"),
            "completed": False,
            "id": len(todos) + 1
        }
        todos.append(new_todo)
        save_todos(todos)
        return [TextContent(
            type="text",
            text=f"✅ 已添加待办事项: {arguments['task']} (优先级: {new_todo['priority']})"
        )]
    
    elif name == "list_todos":
        if not todos:
            return [TextContent(type="text", text="📝 当前没有待办事项")]
        
        result = "📝 待办事项列表:\n\n"
        for i, todo in enumerate(todos, 1):
            status = "✅" if todo["completed"] else "⏳"
            result += f"{i}. {status} [{todo['priority'].upper()}] {todo['task']}\n"
        
        return [TextContent(type="text", text=result)]
    
    elif name == "complete_todo":
        index = arguments["index"] - 1
        if 0 <= index < len(todos):
            todos[index]["completed"] = True
            save_todos(todos)
            return [TextContent(
                type="text", 
                text=f"🎉 已完成: {todos[index]['task']}"
            )]
        else:
            return [TextContent(
                type="text", 
                text=f"❌ 无效的索引: {arguments['index']}。当前有 {len(todos)} 个待办事项。"
            )]
    
    elif name == "clear_completed":
        remaining_todos = [todo for todo in todos if not todo["completed"]]
        removed_count = len(todos) - len(remaining_todos)
        save_todos(remaining_todos)
        return [TextContent(
            type="text", 
            text=f"🧹 已清除 {removed_count} 个已完成的事项"
        )]
    
    return [TextContent(type="text", text="❌ 未知的工具")]

async def main():
    """主函数"""
    # 使用 stdio 传输层创建服务器
    async with stdio_server() as (read_stream, write_stream):
        # 运行服务器
        await server.run(
            read_stream,
            write_stream,
            server.create_initialization_options()
        )

if __name__ == "__main__":
    # 启动服务器
    asyncio.run(main())