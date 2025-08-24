import os
import logging
from datetime import datetime, timedelta
from base.email_attachment_downloader import EmailAttachmentDownloader

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('attachment_extractor_example.log', encoding='utf-8'),
        logging.StreamHandler()
    ],
    encoding='utf-8'
)
logger = logging.getLogger('AttachmentExtractorExample')

def main():
    try:
        # 创建下载器实例
        downloader = EmailAttachmentDownloader(
            save_dir=os.path.join(os.getcwd(), 'attachments'),
            max_attachment_size=10  # 限制附件大小为10MB
        )

        # 连接邮箱
        if not downloader.connect():
            logger.error('无法连接到邮箱，程序退出')
            return

        # 搜索最近7天的邮件
        date_since = datetime.now() - timedelta(days=7)
        email_ids = downloader.search_emails(date_since=date_since)

        if not email_ids:
            logger.info('没有找到符合条件的邮件')
            downloader.disconnect()
            return

        # 设置保存目录
        save_dir = os.path.join(os.getcwd(), 'attachments')
        os.makedirs(save_dir, exist_ok=True)
        logger.info(f'开始下载附件到目录: {save_dir}')

        # 下载附件（包括传统附件和内联附件）
        file_extensions = None  # 下载所有类型的附件
        downloaded_count = downloader.download_attachments(
            email_ids,
            save_dir,
            file_extensions=file_extensions,
            include_inline=True  # 包含内联附件
        )

        logger.info(f'附件下载完成，共下载 {downloaded_count} 个附件')

    except Exception as e:
        logger.error(f'程序执行出错: {str(e)}', exc_info=True)
    finally:
        # 确保断开连接
        if 'downloader' in locals() and downloader.connected:
            downloader.disconnect()

if __name__ == '__main__':
    main()

# 使用说明:
# 1. 确保已配置base/config.json文件，包含正确的邮箱信息
# 2. 运行此脚本将下载最近7天的所有邮件附件（包括内联附件）
# 3. 附件将保存在当前目录下的attachments文件夹中
# 4. 可以通过修改file_extensions参数来过滤特定类型的附件，如['pdf', 'docx']
# 5. 可以调整max_attachment_size参数来限制附件大小
# 6. include_inline参数设置为True表示同时下载内联附件