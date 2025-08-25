#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试精准日期搜索功能
"""
import os
import json
import logging
from datetime import datetime, timedelta
from base.email_attachment_downloader import EmailAttachmentDownloader

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('test_precise_date_search.log', encoding='utf-8'),
        logging.StreamHandler()
    ],
    encoding='utf-8'
)
logger = logging.getLogger('TestPreciseDateSearch')

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

def test_precise_date_search():
    """测试精准日期搜索功能"""
    logger.info("开始测试精准日期搜索功能")
    
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
            save_dir=email_config.get('save_dir'),
            max_attachment_size=email_config.get('max_attachment_size', 50)
        )
        
        # 连接邮箱
        logger.info("正在连接QQ邮箱...")
        if not downloader.connect():
            logger.error('无法连接到QQ邮箱，程序退出')
            return
        
        logger.info("成功连接到QQ邮箱")
        
        # 测试不同日期的搜索
        test_dates = [
            datetime.now().date(),  # 今天
            (datetime.now() - timedelta(days=1)).date(),  # 昨天
            (datetime.now() - timedelta(days=7)).date(),  # 一周前
        ]
        
        for test_date in test_dates:
            logger.info(f"\n{'='*50}")
            logger.info(f"测试搜索 {test_date} 的邮件")
            logger.info(f"{'='*50}")
            
            # 使用新的精准日期搜索
            email_ids = downloader.search_emails_by_date(test_date)
            
            if email_ids:
                logger.info(f"找到 {len(email_ids)} 封 {test_date} 的邮件")
                
                # 显示邮件详情
                for i, email_id in enumerate(email_ids, 1):
                    try:
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
                            logger.info(f"  {i}. {subject_str}")
                            logger.info(f"     日期: {date_str}")
                    except Exception as e:
                        logger.warning(f"无法获取邮件 {email_id} 的详细信息: {str(e)}")
            else:
                logger.info(f"{test_date} 没有找到邮件")
        
        # 测试字符串日期格式
        logger.info(f"\n{'='*50}")
        logger.info("测试字符串日期格式")
        logger.info(f"{'='*50}")
        
        today_str = datetime.now().strftime('%Y-%m-%d')
        email_ids = downloader.search_emails_by_date(today_str)
        logger.info(f"使用字符串格式 '{today_str}' 找到 {len(email_ids)} 封邮件")
        
    except Exception as e:
        logger.error(f'测试过程中出错: {str(e)}', exc_info=True)
    finally:
        # 确保断开连接
        if 'downloader' in locals() and downloader.connected:
            downloader.disconnect()
            logger.info("已断开QQ邮箱连接")

def main():
    """主函数"""
    print("精准日期搜索功能测试")
    print("=" * 50)
    test_precise_date_search()

if __name__ == '__main__':
    main()
