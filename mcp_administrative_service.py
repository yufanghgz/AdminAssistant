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
from base.worktime.excel_merger import read_and_merge_files, get_excel_files_from_dir
from base.worktime.worktime_processor import process_worktime_file
from base.attendance_processor import AttendanceProcessor

# 设置默认编码为UTF-8
sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

# 创建 Server 实例
server = Server("administrative-server")

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
            description="批量识别目录下的所有发票，并复制文件重命名（日期+金额+内容）到当前目录",
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
        ),
        Tool(
            name="generate_worktime_report",
            description="工时安排报表生成工具：合并多个Excel文件并生成工时展示图片",
            inputSchema={
                "type": "object",
                "properties": {
                    "input_files": {
                        "type": "array",
                        "items": {
                            "type": "string"
                        },
                        "description": "要合并的项目任务格式Excel文件路径列表"
                    },
                    "input_dir": {
                        "type": "string",
                        "description": "包含项目任务格式Excel文件的目录路径（与input_files互斥）"
                    },
                    "merged_excel_path": {
                        "type": "string",
                        "description": "合并后的Excel文件保存路径（可选，不指定则使用临时文件）"
                    },
                    "output_image_path": {
                        "type": "string",
                        "description": "输出工时展示图片文件路径"
                    },
                    "staff_file_path": {
                        "type": "string",
                        "description": "人员名单Excel文件路径（可选）",
                        "default": "/Users/heguangzhong/Work_Doc/11.项目管理/2025/工时计划/人员列表.xlsx"
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "要读取的工作表名称，默认为'下周工作计划'",
                        "default": "下周工作计划"
                    },
                    "use_default_sheet": {
                        "type": "boolean",
                        "description": "当指定的工作表不存在时，是否使用第一个工作表，默认为false",
                        "default": False
                    },
                    "title": {
                        "type": "string",
                        "description": "图表标题（可选）"
                    },
                    "keep_merged_file": {
                        "type": "boolean",
                        "description": "是否保留合并后的Excel文件，默认为false",
                        "default": False
                    }
                },
                "required": ["output_image_path"]
            }
        ),
        Tool(
            name="read_emails",
            description="邮件阅读工具：根据日期范围读取邮件内容",
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
                    "days_ago": {
                        "type": "integer",
                        "description": "读取多少天前的邮件，默认为7天",
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
                    "max_emails": {
                        "type": "integer",
                        "description": "最大读取邮件数量，默认为50封",
                        "default": 50
                    },
                    "include_body": {
                        "type": "boolean",
                        "description": "是否包含邮件正文，默认为true",
                        "default": True
                    },
                    "folder": {
                        "type": "string",
                        "description": "邮箱文件夹，默认为INBOX",
                        "default": "INBOX"
                    }
                },
                "required": ["imap_server", "username", "password"]
            }
        ),
        Tool(
            name="process_attendance",
            description="考勤处理工具：自动化处理考勤表数据，包括复制文件、更新字段、处理请假记录、添加入职离职人员等",
            inputSchema={
                "type": "object",
                "properties": {
                    "work_dir": {
                        "type": "string",
                        "description": "包含考勤相关Excel文件的工作目录路径"
                    }
                },
                "required": ["work_dir"]
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

    elif name == "generate_worktime_report":
        input_files = arguments.get("input_files")
        input_dir = arguments.get("input_dir")
        merged_excel_path = arguments.get("merged_excel_path")
        output_image_path = arguments.get("output_image_path")
        staff_file_path = arguments.get("staff_file_path", "/Users/heguangzhong/Work_Doc/11.项目管理/2025/工时计划/人员列表.xlsx")
        sheet_name = arguments.get("sheet_name", "下周工作计划")
        use_default_sheet = arguments.get("use_default_sheet", False)
        title = arguments.get("title")
        keep_merged_file = arguments.get("keep_merged_file", False)

        if not output_image_path:
            return [TextContent(
                type="text",
                text="未指定输出图片路径"
            )]

        # 确定输入文件列表
        file_paths = []
        if input_files and input_dir:
            return [TextContent(
                type="text",
                text="不能同时指定input_files和input_dir参数"
            )]
        elif input_files:
            file_paths = input_files
        elif input_dir:
            if not os.path.exists(input_dir):
                return [TextContent(
                    type="text",
                    text=f"输入目录路径不存在: {input_dir}"
                )]
            try:
                file_paths = get_excel_files_from_dir(input_dir)
                if not file_paths:
                    return [TextContent(
                        type="text",
                        text=f"在目录 {input_dir} 中没有找到Excel文件"
                    )]
            except Exception as e:
                return [TextContent(
                    type="text",
                    text=f"获取目录中的Excel文件时出错: {str(e)}"
                )]
        else:
            return [TextContent(
                type="text",
                text="必须指定input_files或input_dir参数之一"
            )]

        # 验证输入文件是否存在
        for file_path in file_paths:
            if not os.path.exists(file_path):
                return [TextContent(
                    type="text",
                    text=f"输入文件不存在: {file_path}"
                )]

        try:
            # 确定合并后的Excel文件路径
            if not merged_excel_path:
                # 使用临时文件
                import tempfile
                temp_dir = tempfile.gettempdir()
                merged_excel_path = os.path.join(temp_dir, f"merged_worktime_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            else:
                # 确保输出目录存在
                output_dir = os.path.dirname(merged_excel_path)
                if output_dir and not os.path.exists(output_dir):
                    os.makedirs(output_dir, exist_ok=True)

            # 确保图片输出目录存在
            image_output_dir = os.path.dirname(output_image_path)
            if image_output_dir and not os.path.exists(image_output_dir):
                os.makedirs(image_output_dir, exist_ok=True)

            # 检查合并文件是否存在，如果存在则删除
            if os.path.exists(merged_excel_path):
                try:
                    os.remove(merged_excel_path)
                except Exception as e:
                    return [TextContent(
                        type="text",
                        text=f"删除已存在的合并文件时出错: {str(e)}"
                    )]

            # 第一步：合并Excel文件
            merged_df = read_and_merge_files(file_paths, sheet_name=sheet_name, use_default_sheet=use_default_sheet)
            merged_df.to_excel(merged_excel_path, sheet_name=sheet_name, index=False)

            # 第二步：生成工时展示图片
            success = process_worktime_file(
                excel_file_path=merged_excel_path,
                output_path=output_image_path,
                staff_file_path=staff_file_path,
                sheet_name=sheet_name,
                title=title
            )

            result_text = f"工时安排报表生成完成！\n\n"
            result_text += f"📊 合并了 {len(file_paths)} 个Excel文件，共 {len(merged_df)} 行数据\n"
            result_text += f"📈 工时展示图片已保存到: {output_image_path}\n"
            
            if keep_merged_file:
                result_text += f"📁 合并后的Excel文件已保存到: {merged_excel_path}\n"
            else:
                # 删除临时文件
                try:
                    os.remove(merged_excel_path)
                    result_text += f"🗑️ 临时合并文件已清理\n"
                except:
                    pass

            if success:
                return [TextContent(type="text", text=result_text)]
            else:
                return [TextContent(
                    type="text",
                    text=f"工时展示图片生成失败，但Excel文件合并成功。请检查输入文件和参数。\n合并文件位置: {merged_excel_path}"
                )]

        except Exception as e:
            return [TextContent(
                type="text",
                text=f"工时安排报表生成时出错: {str(e)}"
            )]

    elif name == "read_emails":
        imap_server = arguments.get("imap_server")
        username = arguments.get("username")
        password = arguments.get("password")
        days_ago = arguments.get("days_ago", 7)
        single_date = arguments.get("date")
        start_date = arguments.get("start_date")
        end_date = arguments.get("end_date")
        max_emails = arguments.get("max_emails", 50)
        include_body = arguments.get("include_body", True)
        folder = arguments.get("folder", "INBOX")

        # 创建新的邮件下载器实例
        from base.email_attachment_downloader import EmailAttachmentDownloader
        email_reader = EmailAttachmentDownloader()

        try:
            # 连接到邮箱服务器
            if not email_reader.connect(imap_server, username, password):
                return [TextContent(
                    type="text",
                    text=f"无法连接到邮箱服务器: {imap_server}"
                )]

            # 搜索邮件
            if single_date:
                email_ids = email_reader.search_emails_by_date(single_date, folder)
                date_info = f"指定日期 {single_date}"
            elif start_date or end_date:
                email_ids = email_reader.search_emails_by_range(start_date, end_date, folder)
                if start_date and end_date:
                    date_info = f"日期范围 {start_date} 至 {end_date}"
                elif start_date:
                    date_info = f"自 {start_date} 起"
                else:
                    date_info = f"至 {end_date} 止"
            else:
                date_since = datetime.datetime.now() - datetime.timedelta(days=days_ago)
                email_ids = email_reader.search_emails(folder=folder, date_since=date_since)
                date_info = f"最近{days_ago}天"

            if not email_ids:
                return [TextContent(
                    type="text",
                    text=f"未找到{date_info}的邮件"
                )]

            # 读取邮件内容
            emails = email_reader.read_emails(email_ids, max_emails, include_body)

            if not emails:
                return [TextContent(
                    type="text",
                    text=f"成功连接到邮箱，但未能读取到邮件内容"
                )]

            # 格式化邮件信息
            result_text = f"📧 邮件阅读结果 ({date_info})\n\n"
            result_text += f"共找到 {len(email_ids)} 封邮件，成功读取 {len(emails)} 封\n\n"

            for i, email_info in enumerate(emails, 1):
                result_text += f"=== 邮件 {i} ===\n"
                result_text += f"📧 主题: {email_info['subject']}\n"
                result_text += f"👤 发件人: {email_info['sender']}\n"
                result_text += f"📅 日期: {email_info['date']}\n"
                result_text += f"📎 附件: {'是' if email_info['has_attachments'] else '否'}\n"
                
                if email_info['attachments']:
                    result_text += f"📁 附件列表: {', '.join(email_info['attachments'])}\n"
                
                if include_body and email_info['body']:
                    # 限制正文长度，避免输出过长
                    body_preview = email_info['body'][:200] + "..." if len(email_info['body']) > 200 else email_info['body']
                    result_text += f"📝 正文预览: {body_preview}\n"
                
                result_text += "\n"

            return [TextContent(type="text", text=result_text)]

        except Exception as e:
            return [TextContent(
                type="text",
                text=f"读取邮件时出错: {str(e)}"
            )]
        finally:
            # 确保断开连接
            email_reader.disconnect()

    elif name == "process_attendance":
        work_dir = arguments.get("work_dir")
        
        if not work_dir or not os.path.exists(work_dir):
            return [TextContent(
                type="text",
                text=f"工作目录不存在: {work_dir}"
            )]
        
        try:
            # 创建考勤处理器实例
            processor = AttendanceProcessor(work_dir)
            
            # 执行考勤处理流程
            result_text = "📊 考勤处理开始...\n\n"
            
            # 1. 复制文件
            processor.copy_attendance_file()
            result_text += "✅ 1. 复制上月考勤表为当月考勤表\n"
            
            # 2. 修改应出勤天数字段为22天
            processor.update_attendance_days()
            result_text += "✅ 2. 修改应出勤天数字段为22天\n"
            
            # 3. 去掉人员字段前的@符号
            processor.remove_at_symbols()
            result_text += "✅ 3. 去掉人员字段前的@符号\n"
            
            # 4. 重置各种假期字段为0
            processor.reset_leave_fields()
            result_text += "✅ 4. 重置各种假期字段为0\n"
            
            # 5. 删除备注字段里包含"离职"的人员记录
            processor.remove_departure_records()
            result_text += "✅ 5. 删除离职人员记录\n"
            
            # 6. 处理当月请假记录
            processor.process_leave_records()
            result_text += "✅ 6. 处理当月请假记录\n"
            
            # 7. 更新实际出勤天数字段
            processor.update_actual_attendance()
            result_text += "✅ 7. 更新实际出勤天数字段\n"
            
            # 8. 更新实际发工资天数字段
            processor.update_actual_salary_days()
            result_text += "✅ 8. 更新实际发工资天数字段\n"
            
            # 9. 清除备注、银行账号字段中的信息
            processor.clear_remarks_and_bank()
            result_text += "✅ 9. 清除备注、银行账号字段信息\n"
            
            # 10. 处理病假信息，添加待核查病假条备注
            processor.add_sick_leave_verification()
            result_text += "✅ 10. 处理病假信息\n"
            
            # 11. 处理离职人员信息
            processor.process_departure_info()
            result_text += "✅ 11. 处理离职人员信息\n"
            
            # 12. 添加入职人员信息
            processor.add_new_employees()
            result_text += "✅ 12. 添加入职人员信息\n"
            
            # 13. 修改月份字段为当月字段
            processor.update_month_field()
            result_text += "✅ 13. 修改月份字段为当月字段\n"
            
            result_text += "\n🎉 考勤处理完成！所有数据处理步骤已成功执行。"
            result_text += f"\n📁 处理结果保存在: {work_dir}/当月考勤表.xlsx"
            
            # 安全地获取日志文件路径
            try:
                if processor.logger.handlers and len(processor.logger.handlers) > 0:
                    log_file = processor.logger.handlers[0].baseFilename
                    result_text += f"\n📋 详细日志保存在: {log_file}"
            except:
                result_text += "\n📋 详细日志已记录到系统日志"
            
            return [TextContent(type="text", text=result_text)]
            
        except Exception as e:
            return [TextContent(
                type="text",
                text=f"考勤处理失败: {str(e)}"
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