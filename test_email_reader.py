#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试邮件阅读功能
"""
import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from base.email_attachment_downloader import EmailAttachmentDownloader
import datetime

def test_email_reader():
    # 创建邮件下载器实例
    email_reader = EmailAttachmentDownloader()
    
    try:
        # 连接到QQ邮箱
        print("正在连接到QQ邮箱...")
        if not email_reader.connect("imap.qq.com", "23512221@qq.com", "ifqdmucfxotlbigc"):
            print("❌ 无法连接到邮箱服务器")
            return
        
        print("✅ 成功连接到邮箱服务器")
        
        # 搜索最近7天的邮件
        print("正在搜索最近7天的邮件...")
        date_since = datetime.datetime.now() - datetime.timedelta(days=7)
        email_ids = email_reader.search_emails(folder="INBOX", date_since=date_since)
        
        print(f"找到 {len(email_ids)} 封邮件")
        
        if not email_ids:
            print("❌ 未找到邮件")
            return
        
        # 读取邮件内容
        print("正在读取邮件内容...")
        emails = email_reader.read_emails(email_ids[:5], max_emails=5, include_body=True)
        
        print(f"成功读取 {len(emails)} 封邮件")
        
        # 显示邮件信息
        for i, email_info in enumerate(emails, 1):
            print(f"\n=== 邮件 {i} ===")
            print(f"主题: {email_info['subject']}")
            print(f"发件人: {email_info['sender']}")
            print(f"日期: {email_info['date']}")
            print(f"附件: {'是' if email_info['has_attachments'] else '否'}")
            if email_info['attachments']:
                print(f"附件列表: {', '.join(email_info['attachments'])}")
            if email_info['body']:
                body_preview = email_info['body'][:100] + "..." if len(email_info['body']) > 100 else email_info['body']
                print(f"正文预览: {body_preview}")
        
    except Exception as e:
        print(f"❌ 测试过程中出错: {str(e)}")
    finally:
        # 断开连接
        email_reader.disconnect()
        print("已断开邮箱连接")

if __name__ == "__main__":
    test_email_reader()

