#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
读取最近几天的邮件并生成总结
"""
import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from base.email_attachment_downloader import EmailAttachmentDownloader
import datetime

def read_recent_emails():
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
        
        print(f"找到 {len(email_ids)} 封最近7天的邮件")
        
        if not email_ids:
            print("📭 最近7天没有收到邮件")
            return
        
        # 读取邮件内容
        print("正在读取邮件内容...")
        emails = email_reader.read_emails(email_ids, max_emails=20, include_body=True)
        
        print(f"成功读取 {len(emails)} 封邮件")
        
        # 生成邮件总结
        print("\n" + "="*60)
        print(f"📧 最近7天邮件总结")
        print("="*60)
        print(f"📊 总计收到 {len(emails)} 封邮件\n")
        
        # 按类型分类邮件
        invoices = []
        notifications = []
        others = []
        
        for email_info in emails:
            subject = email_info['subject'].lower()
            if '发票' in subject or 'invoice' in subject or '打车' in subject:
                invoices.append(email_info)
            elif '通知' in subject or 'notification' in subject or '理赔' in subject:
                notifications.append(email_info)
            else:
                others.append(email_info)
        
        # 发票类邮件
        if invoices:
            print("🧾 发票类邮件:")
            for i, email_info in enumerate(invoices, 1):
                print(f"  {i}. {email_info['subject']}")
                print(f"     发件人: {email_info['sender']}")
                print(f"     时间: {email_info['date']}")
                if email_info['attachments']:
                    print(f"     附件: {', '.join(email_info['attachments'])}")
                print()
        
        # 通知类邮件
        if notifications:
            print("🔔 通知类邮件:")
            for i, email_info in enumerate(notifications, 1):
                print(f"  {i}. {email_info['subject']}")
                print(f"     发件人: {email_info['sender']}")
                print(f"     时间: {email_info['date']}")
                if email_info['attachments']:
                    print(f"     附件: {', '.join(email_info['attachments'])}")
                print()
        
        # 其他邮件
        if others:
            print("📬 其他邮件:")
            for i, email_info in enumerate(others, 1):
                print(f"  {i}. {email_info['subject']}")
                print(f"     发件人: {email_info['sender']}")
                print(f"     时间: {email_info['date']}")
                if email_info['attachments']:
                    print(f"     附件: {', '.join(email_info['attachments'])}")
                print()
        
        # 统计信息
        print("📈 统计信息:")
        print(f"  • 发票类邮件: {len(invoices)} 封")
        print(f"  • 通知类邮件: {len(notifications)} 封")
        print(f"  • 其他邮件: {len(others)} 封")
        print(f"  • 带附件的邮件: {sum(1 for email in emails if email['has_attachments'])} 封")
        
        # 按日期分组
        print("\n📅 按日期分组:")
        date_groups = {}
        for email_info in emails:
            # 提取日期部分
            date_str = email_info['date'].split(',')[1].strip().split(' ')[0:3]
            date_key = ' '.join(date_str)
            if date_key not in date_groups:
                date_groups[date_key] = []
            date_groups[date_key].append(email_info)
        
        for date_key in sorted(date_groups.keys(), reverse=True):
            print(f"  {date_key}: {len(date_groups[date_key])} 封邮件")
        
    except Exception as e:
        print(f"❌ 处理过程中出错: {str(e)}")
    finally:
        # 断开连接
        email_reader.disconnect()
        print("\n已断开邮箱连接")

if __name__ == "__main__":
    read_recent_emails()

