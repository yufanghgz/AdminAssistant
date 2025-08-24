#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
发票OCR和邮件下载 MCP 服务器示例
"""
import asyncio
import json
import os
import sys
import datetime
from typing import List, Dict, Any
from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent, ImageContent
from base.invoice_ocr import InvoiceRecognizer
from base.email_attachment_downloader import EmailAttachmentDownloader

# 设置默认编码为UTF-8
sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

# 创建 Server 实例
server = Server("invoice-ocr-server")

# 初始化发票识别器和邮件下载器
recognizer = InvoiceRecognizer()
email_downloader = EmailAttachmentDownloader()


@server.list_tools()
async def list_tools() -> List[Tool]:
    """列出服务器提供的所有工具"""
    return [
        Tool(
            name="ocr_invoice",
            description="识别单个发票",
            inputSchema={
                "type": "object",
                "properties": {
                    "filepath": {
                        "type": "string",
                        "description": "发票文件路径"
                    }
                },
                "required": ["filepath"]
            }
        ),
        Tool(
            name="ocr_invoices",
            description="批量识别目录下的所有发票",
            inputSchema={
                "type": "object",
                "properties": {
                    "folder_path": {
                        "type": "string",
                        "description": "包含发票的目录路径"
                    }
                },
                "required": ["folder_path"]
            }
        ),
        Tool(
            name="download_emails",
            description="下载指定日期范围内的邮件附件",
            inputSchema={
                "type": "object",
                "properties": {
                    "imap_server": {
                        "type": "string",
                        "description": "IMAP服务器地址"
                    },
                    "username": {
                        "type": "string",
                        "description": "邮箱用户名"
                    },
                    "password": {
                        "type": "string",
                        "description": "邮箱密码或授权码"
                    },
                    "save_dir": {
                        "type": "string",
                        "description": "保存附件的目录路径"
                    },
                    "days_ago": {
                        "type": "integer",
                        "description": "下载多少天前的邮件，默认为7天",
                        "default": 7
                    },
                    "file_extensions": {
                        "type": "array",
                        "items": {
                            "type": "string"
                        },
                        "description": "要下载的文件扩展名列表，如['pdf', 'doc']，为None时下载所有附件"
                    }
                },
                "required": ["imap_server", "username", "password", "save_dir"]
            }
        )
    ]

@server.call_tool()
async def call_tool(name: str, arguments: dict) -> List[TextContent]:
    """处理工具调用"""
    if name == "ocr_invoice":
        filepath = arguments.get("filepath")
        if not filepath or not os.path.exists(filepath):
            return [TextContent(
                type="text",
                text=f"文件路径不存在: {filepath}"
            )]
        
        try:
            # 确保文件路径使用UTF-8编码
            filepath = str(filepath)
            result = recognizer.recognize_invoice(filepath)
            # 格式化结果为易读的文本
            text_result = f"发票识别结果:\n\n"
            text_result += f"文件路径: {result['file_path']}\n"
            text_result += f"日期: {result['date']}\n"
            text_result += f"金额: {result['amount']}\n"
            text_result += f"内容: {result['content']}\n"
            
            return [TextContent(type="text", text=text_result)]
        except Exception as e:
            return [TextContent(
                type="text",
                text=f"识别发票失败: {str(e)}"
            )]
    
    elif name == "ocr_invoices":
        folder_path = arguments.get("folder_path")
        if not folder_path or not os.path.exists(folder_path):
            return [TextContent(
                type="text",
                text=f"目录路径不存在: {folder_path}"
            )]
        
        try:
            # 确保文件夹路径使用UTF-8编码
            folder_path = str(folder_path)
            results = recognizer.batch_recognize_invoices(folder_path)
            if not results:
                return [TextContent(
                    type="text",
                    text=f"未找到可识别的发票文件在目录: {folder_path}"
                )]
            
            # 格式化结果为易读的文本
            text_result = f"批量发票识别结果 (共 {len(results)} 张发票):\n\n"
            for i, result in enumerate(results, 1):
                text_result += f"=== 发票 {i} ===\n"
                text_result += f"文件路径: {result['file_path']}\n"
                text_result += f"日期: {result['date']}\n"
                text_result += f"金额: {result['amount']}\n"
                # 确保内容使用UTF-8编码
                content = str(result['content'])[:50]
                text_result += f"内容: {content}...\n\n"
            
            return [TextContent(type="text", text=text_result)]
        except Exception as e:
            return [TextContent(
                type="text",
                text=f"批量识别发票失败: {str(e)}"
            )]
    
    elif name == "download_emails":
        print("herer",arguments)
        imap_server = arguments.get("imap_server")
        username = arguments.get("username")
        password = arguments.get("password")
        save_dir = arguments.get("save_dir")
        days_ago = arguments.get("days_ago", 7)
        file_extensions = arguments.get("file_extensions")
 
        try:
            # 连接到邮箱服务器
            if not email_downloader.connect(imap_server, username, password):
                return [TextContent(
                    type="text",
                    text=f"无法连接到邮箱服务器: {imap_server}"
                )]

            # 计算指定天数前的日期
            date_since = datetime.datetime.now() - datetime.timedelta(days=days_ago)

            # 搜索邮件
            email_ids = email_downloader.search_emails(date_since=date_since)
            if not email_ids:
                return [TextContent(
                    type="text",
                    text=f"未找到{days_ago}天内的邮件"
                )]

            # 下载附件
            downloaded_count = email_downloader.download_attachments(
                email_ids, save_dir, file_extensions
            )

            return [TextContent(
                type="text",
                text=f"成功下载 {downloaded_count} 个附件到目录: {save_dir}"
            )]
        except Exception as e:
            return [TextContent(
                type="text",
                text=f"下载邮件附件失败: {str(e)}"
            )]
        finally:
            # 确保断开连接
            email_downloader.disconnect()

    return [TextContent(type="text", text=f"未知的工具: {name}")]

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