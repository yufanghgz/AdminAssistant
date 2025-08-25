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
from base.image_to_pdf import images_to_pdf

# 设置默认编码为UTF-8
sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

# 创建 Server 实例
server = Server("invoice-ocr-server")

# 初始化发票识别器、邮件下载器和添加PDF转图像所需的导入
from base.batch_pdf_to_image import pdf_to_images
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
                    "date": {
                        "type": "string",
                        "description": "指定单日，格式 YYYY-MM-DD，可与 days_ago 互斥"
                    },
                    "start_date": {
                        "type": "string",
                        "description": "起始日期，格式 YYYY-MM-DD"
                    },
                    "end_date": {
                        "type": "string",
                        "description": "结束日期，格式 YYYY-MM-DD（含当日）"
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
        ),
        Tool(
            name="convert_pdf_to_image",
            description="将单个PDF文件转换为图像文件",
            inputSchema={
                "type": "object",
                "properties": {
                    "pdf_path": {
                        "type": "string",
                        "description": "PDF文件路径"
                    },
                    "output_folder": {
                        "type": "string",
                        "description": "图像输出目录"
                    },
                    "resolution": {
                        "type": "integer",
                        "description": "图像分辨率 (dpi)，默认为300",
                        "default": 300
                    },
                    "image_format": {
                        "type": "string",
                        "description": "图像格式 (png/jpg/jpeg)，默认为png",
                        "default": "png"
                    }
                },
                "required": ["pdf_path", "output_folder"]
            }
        ),
        Tool(
            name="batch_convert_pdfs",
            description="批量将文件夹下的PDF文件转换为图像文件",
            inputSchema={
                "type": "object",
                "properties": {
                    "input_dir": {
                        "type": "string",
                        "description": "包含PDF文件的目录"
                    },
                    "output_dir": {
                        "type": "string",
                        "description": "图片输出目录"
                    },
                    "resolution": {
                        "type": "integer",
                        "description": "图像分辨率 (dpi)，默认为300",
                        "default": 300
                    },
                    "image_format": {
                        "type": "string",
                        "description": "图像格式 (png/jpg/jpeg)，默认为png",
                        "default": "png"
                    }
                },
                "required": ["input_dir", "output_dir"]
            }
        ),
        Tool(
            name="images_to_pdf",
            description="将文件夹中的多张图片合并为PDF文件",
            inputSchema={
                "type": "object",
                "properties": {
                    "input_folder": {
                        "type": "string",
                        "description": "包含图片的文件夹路径"
                    },
                    "output_file": {
                        "type": "string",
                        "description": "输出的PDF文件路径"
                    }
                },
                "required": ["input_folder", "output_file"]
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
    
    elif name == "convert_pdf_to_image":
        pdf_path = arguments.get("pdf_path")
        output_folder = arguments.get("output_folder")
        resolution = arguments.get("resolution", 300)
        image_format = arguments.get("image_format", "png")
        
        if not pdf_path or not os.path.exists(pdf_path):
            return [TextContent(
                type="text",
                text=f"PDF文件路径不存在: {pdf_path}"
            )]
        
        if not output_folder:
            return [TextContent(
                type="text",
                text="未指定图像输出目录"
            )]
        
        try:
            # 确保目录存在
            os.makedirs(output_folder, exist_ok=True)
            
            # 调用pdf_to_images函数进行转换
            success = pdf_to_images(pdf_path, output_folder, resolution, image_format)
            
            if success:
                # 获取生成的图像文件列表
                pdf_filename = os.path.basename(pdf_path)
                pdf_name = os.path.splitext(pdf_filename)[0]
                
                # 获取输出目录中的相关图像文件
                image_files = [f for f in os.listdir(output_folder) if f.startswith(pdf_name) and f.lower().endswith(image_format)]
                
                return [TextContent(
                    type="text",
                    text=f"PDF文件 '{pdf_filename}' 已成功转换为 {len(image_files)} 个图像文件，保存至 '{output_folder}'"
                )]
            else:
                return [TextContent(
                    type="text",
                    text=f"转换PDF文件 '{pdf_filename}' 时出错"
                )]
        except Exception as e:
            return [TextContent(
                type="text",
                text=f"处理PDF转换请求时出错: {str(e)}"
            )]
    
    elif name == "batch_convert_pdfs":
        input_dir = arguments.get("input_dir")
        output_dir = arguments.get("output_dir")
        resolution = arguments.get("resolution", 300)
        image_format = arguments.get("image_format", "png")
        
        if not input_dir or not os.path.exists(input_dir):
            return [TextContent(
                type="text",
                text=f"PDF目录路径不存在: {input_dir}"
            )]
        
        if not output_dir:
            return [TextContent(
                type="text",
                text="未指定图像输出目录"
            )]
        
        try:
            # 确保输出目录存在
            os.makedirs(output_dir, exist_ok=True)
            
            # 获取输入目录中的所有PDF文件
            pdf_files = [f for f in os.listdir(input_dir) if f.lower().endswith('.pdf')]
            
            if not pdf_files:
                return [TextContent(
                    type="text",
                    text=f"在目录 '{input_dir}' 中未找到PDF文件"
                )]
            
            # 遍历所有PDF文件并转换
            converted_count = 0
            for pdf_file in pdf_files:
                pdf_path = os.path.join(input_dir, pdf_file)
                success = pdf_to_images(pdf_path, output_dir, resolution, image_format)
                if success:
                    converted_count += 1
            
            return [TextContent(
                type="text",
                text=f"已成功转换 {converted_count} 个PDF文件，保存至 '{output_dir}'"
            )]
        except Exception as e:
            return [TextContent(
                type="text",
                text=f"批量转换PDF文件时出错: {str(e)}"
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
        print("here：",arguments)
        imap_server = arguments.get("imap_server")
        username = arguments.get("username")
        password = arguments.get("password")
        save_dir = arguments.get("save_dir")
        days_ago = arguments.get("days_ago", 7)
        single_date = arguments.get("date")
        start_date = arguments.get("start_date")
        end_date = arguments.get("end_date")
        file_extensions = arguments.get("file_extensions")
        use_precise_date = arguments.get("use_precise_date", True)  # 默认使用精准日期搜索
 
        try:
            # 连接到邮箱服务器
            if not email_downloader.connect(imap_server, username, password):
                return [TextContent(
                    type="text",
                    text=f"无法连接到邮箱服务器: {imap_server}"
                )]

            # 优先级：date > (start_date/end_date) > days_ago
            if single_date:
                email_ids = email_downloader.search_emails_by_date(single_date)
                date_info = f"指定日期 {single_date}"
            elif start_date or end_date:
                email_ids = email_downloader.search_emails_by_range(start_date, end_date)
                if start_date and end_date:
                    date_info = f"日期范围 {start_date} 至 {end_date}"
                elif start_date:
                    date_info = f"自 {start_date} 起"
                else:
                    date_info = f"至 {end_date} 止"
            elif use_precise_date and days_ago == 0:
                today = datetime.datetime.now().date()
                email_ids = email_downloader.search_emails_by_date(today)
                date_info = f"今天 ({today})"
            else:
                date_since = datetime.datetime.now() - datetime.timedelta(days=days_ago)
                email_ids = email_downloader.search_emails(date_since=date_since)
                date_info = f"最近{days_ago}天"

            if not email_ids:
                return [TextContent(
                    type="text",
                    text=f"未找到{date_info}的邮件"
                )]

            # 下载附件
            downloaded_count = email_downloader.download_attachments(
                email_ids, save_dir, file_extensions
            )

            return [TextContent(
                type="text",
                text=f"成功下载 {downloaded_count} 个{date_info}的附件到目录: {save_dir}"
            )]
        except Exception as e:
            return [TextContent(
                type="text",
                text=f"下载邮件附件失败: {str(e)}"
            )]
        finally:
            # 确保断开连接
            email_downloader.disconnect()

    elif name == "images_to_pdf":
        input_folder = arguments.get("input_folder")
        output_file = arguments.get("output_file")

        if not input_folder or not os.path.exists(input_folder):
            return [TextContent(
                type="text",
                text=f"图片文件夹路径不存在: {input_folder}"
            )]

        if not output_file:
            return [TextContent(
                type="text",
                text="未指定输出PDF文件路径"
            )]

        try:
            # 确保输出目录存在
            output_dir = os.path.dirname(output_file)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)

            # 调用images_to_pdf函数进行转换
            images_to_pdf(input_folder, output_file)

            return [TextContent(
                type="text",
                text=f"图片已成功合并为PDF文件: '{output_file}'"
            )]
        except Exception as e:
            return [TextContent(
                type="text",
                text=f"合并图片为PDF文件时出错: {str(e)}"
            )]

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