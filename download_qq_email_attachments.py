#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
使用QQ邮箱配置下载今天的邮件附件
"""
import os
import json
import logging
from datetime import datetime, timedelta
from base.email_attachment_downloader import EmailAttachmentDownloader
import email.message
import email.header

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('qq_email_download.log', encoding='utf-8'),
        logging.StreamHandler()
    ],
    encoding='utf-8'
)
logger = logging.getLogger('QQEmailDownloader')

def load_qq_email_config():
    """加载QQ邮箱配置"""
    config_path = 'base/conf/qq-email.json'
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
            return config.get('email', {})
    except Exception as e:
        logger.error(f"读取QQ邮箱配置文件失败: {str(e)}")
        return None

def main():
    """主函数"""
    logger.info("开始下载QQ邮箱今天的邮件附件")
    
    try:
        # 加载QQ邮箱配置
        email_config = load_qq_email_config()
        if not email_config:
            logger.error("无法加载QQ邮箱配置，程序退出")
            return
        
        logger.info(f"成功加载QQ邮箱配置: {email_config.get('username')}")
        
        # 创建下载器实例
        downloader = EmailAttachmentDownloader(
            imap_server=email_config.get('imap_server'),
            username=email_config.get('username'),
            password=email_config.get('password'),
            port=email_config.get('port', 993),
            folder=email_config.get('folder', 'INBOX'),
            days_ago=1,  # 下载今天的邮件
            save_dir=email_config.get('save_dir'),
            max_attachment_size=email_config.get('max_attachment_size', 50)  # 默认50MB
        )
        
        # 连接邮箱
        logger.info("正在连接QQ邮箱...")
        if not downloader.connect():
            logger.error('无法连接到QQ邮箱，程序退出')
            return
        
        logger.info("成功连接到QQ邮箱")
        
        # 搜索今天的邮件
        today = datetime.now().date()
        logger.info(f"搜索 {today} 的邮件")
        
        # 使用新的精准日期搜索功能
        email_ids = downloader.search_emails_by_date(today)
        
        if not email_ids:
            logger.info('今天没有找到邮件')
            downloader.disconnect()
            return
        
        logger.info(f"找到 {len(email_ids)} 封今天的邮件")
        
        # 显示邮件详细信息
        logger.info("邮件详细信息:")
        for i, email_id in enumerate(email_ids, 1):
            try:
                # 获取邮件信息
                status, data = downloader.imap_server.fetch(email_id, '(RFC822.HEADER)')
                if status == 'OK':
                    email_message = email.message_from_bytes(data[0][1])
                    subject = email.header.decode_header(email_message['Subject'])
                    subject_str = ''
                    for part, encoding in subject:
                        if isinstance(part, bytes):
                            subject_str += part.decode(encoding or 'utf-8')
                        else:
                            subject_str += part
                    
                    date_str = email_message.get('Date', '未知日期')
                    logger.info(f"  {i}. {subject_str} - {date_str}")
            except Exception as e:
                logger.warning(f"无法获取邮件 {email_id} 的详细信息: {str(e)}")
        
        # 确保保存目录存在
        save_dir = email_config.get('save_dir', 'attachments')
        if not os.path.exists(save_dir):
            try:
                os.makedirs(save_dir)
                logger.info(f"创建保存目录: {save_dir}")
            except Exception as e:
                logger.error(f"创建保存目录失败: {str(e)}")
                return
        
        # 下载附件
        file_extensions = email_config.get('file_extensions', None)
        if file_extensions:
            logger.info(f"将下载以下类型的附件: {file_extensions}")
        else:
            logger.info("将下载所有类型的附件")
        
        downloaded_count = downloader.download_attachments(
            email_ids,
            save_dir,
            file_extensions=file_extensions,
            include_inline=True  # 包含内联附件
        )
        
        logger.info(f'附件下载完成，共下载 {downloaded_count} 个附件到目录: {save_dir}')
        
        # 显示下载的附件列表
        if downloaded_count > 0:
            logger.info("下载的附件列表:")
            try:
                for file in os.listdir(save_dir):
                    file_path = os.path.join(save_dir, file)
                    if os.path.isfile(file_path):
                        file_size = os.path.getsize(file_path)
                        logger.info(f"  - {file} ({file_size/1024:.1f} KB)")
            except Exception as e:
                logger.warning(f"无法列出附件文件: {str(e)}")
        
    except Exception as e:
        logger.error(f'程序执行出错: {str(e)}', exc_info=True)
    finally:
        # 确保断开连接
        if 'downloader' in locals() and downloader.connected:
            downloader.disconnect()
            logger.info("已断开QQ邮箱连接")

if __name__ == '__main__':
    main()
