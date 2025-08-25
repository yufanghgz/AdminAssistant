import imaplib
import email
import os
import datetime
import json
import logging
import uuid
import re
import requests
from bs4 import BeautifulSoup

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('email_attachment_downloader.log', encoding='utf-8'),
        logging.StreamHandler()
    ],
    encoding='utf-8'
)
logger = logging.getLogger('EmailAttachmentDownloader')

class EmailAttachmentDownloader:
    def __init__(self, imap_server=None, username=None, password=None, port=993, folder='INBOX', days_ago=None, save_dir=None, max_attachment_size=None):
        logger.info("./base/conf/qq-email.json")
        try:
            # Read the configuration file (from the directory where the current script is located)
            config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'conf/qq-email.json')
            logger.info(f"config_path={config_path}")
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
                email_config = config.get('email', {})

            # Get parameters from the configuration, use the input parameters as fallback
            self.imap_server = imap_server if imap_server is not None else email_config.get('imap_server')
            self.username = username if username is not None else email_config.get('username')
            self.password = password if password is not None else email_config.get('password')
            self.port = port if port is not None else email_config.get('port', 993)
            self.folder = folder if folder is not None else email_config.get('folder', 'INBOX')
            self.days_ago = days_ago if days_ago is not None else email_config.get('days_ago')
            self.save_dir = save_dir if save_dir is not None else email_config.get('save_dir')
            self.max_attachment_size = max_attachment_size if max_attachment_size is not None else email_config.get('max_attachment_size')
            self.connected = False
            logger.info(f"self.imap_server={self.imap_server}, self.username={self.username}, self.password={self.password}, self.port={self.port}, self.folder={self.folder}, self.days_ago={self.days_ago}, self.save_dir={self.save_dir}, self.max_attachment_size={self.max_attachment_size}")
            # Validate configuration
            if not self.validate_config():
                logger.error("Invalid configuration, please check config.json")
        except Exception as e:
            logger.error(f"Error reading configuration file: {str(e)}")
            self.imap_server = imap_server
            self.username = username
            self.password = password
            self.port = port
            self.folder = folder
            self.days_ago = days_ago
            self.save_dir = save_dir
            self.max_attachment_size = max_attachment_size
            self.connected = False
            logger.info(f"config.json get success: imap_server={imap_server}, username={username}, save_dir={save_dir}")
    def validate_config(self):
        """
        验证配置是否有效
        :return: 配置是否有效
        """
        if not self.imap_server:
            logger.error("IMAP server is not configured")
            return False
        if not self.username:
            logger.error("Username is not configured")
            return False
        if not self.password:
            logger.error("Password is not configured")
            return False
        if not self.save_dir:
            logger.warning("Save directory is not configured, using default 'attachments'")
            self.save_dir = 'attachments'
        return True

    def connect(self, imap_server=None, username=None, password=None, port=993):
        """
        连接到IMAP邮箱服务器
        :param imap_server: IMAP服务器地址
        :param username: 邮箱用户名
        :param password: 邮箱密码或授权码
        :param port: 服务器端口，默认993（SSL）
        :return: 是否连接成功
        """
        # 如果没有提供参数，使用实例化时的参数
        
        imap_server = imap_server or self.imap_server
        username = username or self.username
        password = password or self.password
        port = port or self.port

        if not all([imap_server, username, password]):
            logger.error("Missing required connection parameters")
            return False

        try:
            # 连接服务器
            self.imap_server = imaplib.IMAP4_SSL(imap_server, port)
            # 登录账户
            self.imap_server.login(username, password)
            self.username = username
            self.password = password
            self.connected = True
            logger.info(f"成功连接到邮箱: {username}")
            return True
        except Exception as e:
            logger.error(f"连接邮箱失败: {str(e)}")
            self.connected = False
            return False

    def search_emails(self, folder='INBOX', criteria='ALL', date_since=None, date_until=None, target_date=None):
        """
        搜索邮件，支持精准日期过滤
        :param folder: 邮箱文件夹，默认收件箱(INBOX)
        :param criteria: 搜索条件，默认所有邮件(ALL)
                        其他条件示例: 'UNSEEN'（未读）, 'FROM "example@domain.com"'（来自特定发件人）
        :param date_since: 开始日期，datetime对象
        :param date_until: 结束日期，datetime对象
        :param target_date: 目标日期，datetime.date对象，只返回指定日期的邮件
        :return: 邮件ID列表
        """
        if not self.connected:
            logger.error("未连接到邮箱服务器")
            return []

        try:
            # 选择文件夹
            self.imap_server.select(folder)
            
            # 构建搜索条件（优先组合日期范围）
            search_criteria = criteria
            date_parts = []
            if date_since:
                imap_since = date_since.strftime('%d-%b-%Y')
                date_parts.append(f'SINCE {imap_since}')
            if date_until:
                # IMAP 的 BEFORE 不包含当日，因此将 end_date+1 天作为 BEFORE 的参数
                try:
                    next_day = (date_until + datetime.timedelta(days=1)) if isinstance(date_until, datetime.datetime) else (date_until + datetime.timedelta(days=1))
                except Exception:
                    next_day = None
                if next_day:
                    imap_before = next_day.strftime('%d-%b-%Y')
                    date_parts.append(f'BEFORE {imap_before}')

            if date_parts:
                search_criteria = f"({' '.join(date_parts)}) {search_criteria}"
                logger.info(f"搜索条件: {search_criteria}")

            # 搜索邮件
            result, data = self.imap_server.search(None, search_criteria)
            if result != 'OK':
                logger.error("搜索邮件失败")
                return []
            
            # 提取邮件ID
            email_ids = data[0].split()
            logger.info(f"初步搜索找到 {len(email_ids)} 封邮件")
            
            # 如果指定了目标日期或提供了日期范围，进行精准过滤
            if target_date or (date_since or date_until):
                logger.info(f"开始精准过滤 {target_date} 的邮件...")
                filtered_email_ids = []
                
                for email_id in email_ids:
                    try:
                        # 获取邮件头部信息
                        status, data = self.imap_server.fetch(email_id, '(RFC822.HEADER)')
                        if status == 'OK':
                            email_message = email.message_from_bytes(data[0][1])
                            
                            # 解析邮件日期
                            try:
                                from email.utils import parsedate_to_datetime
                                email_date = parsedate_to_datetime(email_message.get('Date'))
                                include = True
                                if target_date:
                                    include = include and (email_date and email_date.date() == target_date)
                                if date_since:
                                    include = include and (email_date and email_date >= (date_since if isinstance(date_since, datetime.datetime) else datetime.datetime.combine(date_since, datetime.datetime.min.time())))
                                if date_until:
                                    # 包含 end 当天
                                    end_dt = (date_until if isinstance(date_until, datetime.datetime) else datetime.datetime.combine(date_until, datetime.datetime.max.time()))
                                    include = include and (email_date and email_date <= end_dt)
                                if include:
                                    filtered_email_ids.append(email_id)
                                    logger.debug(f"邮件 {email_id} 日期匹配: {email_date}")
                                else:
                                    logger.debug(f"邮件 {email_id} 日期不匹配: {email_date}")
                            except Exception as e:
                                logger.warning(f"无法解析邮件 {email_id} 的日期: {str(e)}")
                                # 如果无法解析日期，暂时包含这封邮件
                                filtered_email_ids.append(email_id)
                    except Exception as e:
                        logger.warning(f"获取邮件 {email_id} 头部信息失败: {str(e)}")
                
                logger.info(f"精准过滤后找到 {len(filtered_email_ids)} 封 {target_date} 的邮件")
                return filtered_email_ids
            
            return email_ids
            
        except Exception as e:
            logger.error(f"搜索邮件时出错: {str(e)}")
            return []

    def search_emails_by_date(self, target_date, folder='INBOX', criteria='ALL'):
        """
        搜索指定日期的邮件（便捷方法）
        :param target_date: 目标日期，可以是datetime.date对象或字符串 'YYYY-MM-DD'
        :param folder: 邮箱文件夹，默认收件箱(INBOX)
        :param criteria: 搜索条件，默认所有邮件(ALL)
        :return: 邮件ID列表
        """
        from datetime import datetime, date
        
        # 处理日期参数
        if isinstance(target_date, str):
            try:
                target_date = datetime.strptime(target_date, '%Y-%m-%d').date()
            except ValueError:
                logger.error(f"日期格式错误: {target_date}，请使用 'YYYY-MM-DD' 格式")
                return []
        elif isinstance(target_date, datetime):
            target_date = target_date.date()
        elif not isinstance(target_date, date):
            logger.error(f"不支持的日期类型: {type(target_date)}")
            return []
        
        # 设置搜索范围（从目标日期开始）
        date_since = datetime.combine(target_date, datetime.min.time())
        
        logger.info(f"搜索 {target_date} 的邮件")
        return self.search_emails(folder=folder, criteria=criteria, date_since=date_since, target_date=target_date)

    def search_emails_by_range(self, start_date=None, end_date=None, folder='INBOX', criteria='ALL'):
        """
        按日期范围搜索邮件（包含端点）。
        :param start_date: 开始日期，datetime/date 或 'YYYY-MM-DD' 字符串
        :param end_date: 结束日期，datetime/date 或 'YYYY-MM-DD' 字符串
        :param folder: 邮箱文件夹
        :param criteria: 其他搜索条件
        :return: 邮件ID列表
        """
        from datetime import datetime, date

        def normalize(d, is_start=True):
            if d is None:
                return None
            if isinstance(d, str):
                d = datetime.strptime(d, '%Y-%m-%d').date()
            if isinstance(d, datetime):
                return d
            if isinstance(d, date):
                return datetime.combine(d, datetime.min.time() if is_start else datetime.max.time())
            raise ValueError('不支持的日期类型')

        start_dt = normalize(start_date, True)
        end_dt = normalize(end_date, False)

        # 如果只给了一个日期，就按单日搜索
        if start_dt and not end_dt:
            return self.search_emails(folder=folder, criteria=criteria, date_since=start_dt, target_date=start_dt.date())
        if end_dt and not start_dt:
            return self.search_emails(folder=folder, criteria=criteria, date_since=end_dt, target_date=end_dt.date())

        return self.search_emails(folder=folder, criteria=criteria, date_since=start_dt, date_until=end_dt)

    def _clean_filename(self, filename):
        """
        清理文件名，移除Windows系统不支持的特殊字符
        :param filename: 原始文件名
        :return: 清理后的文件名
        """
        import re
        # 替换Windows不支持的特殊字符为下划线
        cleaned_filename = re.sub(r'[\\/:*?"<>|]', '_', filename)
        # 移除多余的空格和点
        cleaned_filename = re.sub(r'\s+', ' ', cleaned_filename).strip()
        cleaned_filename = re.sub(r'\.+', '.', cleaned_filename)
        return cleaned_filename

    def download_attachments(self, email_ids, save_dir, file_extensions=None, include_inline=False):
        """
        下载指定邮件ID列表中的附件
        :param email_ids: 邮件ID列表
        :param save_dir: 附件保存目录
        :param file_extensions: 允许的文件扩展名列表，None表示允许所有
        :param include_inline: 是否下载内联附件
        :return: 成功下载的附件数量
        """
        if not self.connected:
            logger.error("未连接到邮箱服务器")
            return 0

        # 确保保存目录存在
        if not os.path.exists(save_dir):
            try:
                os.makedirs(save_dir)
                logger.info(f"创建保存目录: {save_dir}")
            except Exception as e:
                logger.error(f"创建保存目录失败: {str(e)}")
                return 0

        downloaded_count = 0

        for email_id in email_ids:
            try:
                # 获取邮件内容
                result, data = self.imap_server.fetch(email_id, '(RFC822)')
                if result != 'OK':
                    logger.error(f"获取邮件 {email_id} 失败")
                    continue

                # 解析邮件内容
                raw_email = data[0][1]
                email_message = email.message_from_bytes(raw_email)

                # 获取邮件主题
                subject = email.header.decode_header(email_message['Subject'])
                subject_str = ''
                for part, encoding in subject:
                    if isinstance(part, bytes):
                        subject_str += part.decode(encoding or 'utf-8')
                    else:
                        subject_str += part

                logger.info(f"处理邮件: {subject_str} (ID: {email_id})")

                # 遍历邮件的各个部分
                for part in email_message.walk():
                    # 检查是否是附件或内联附件
                    content_disposition = part.get_content_disposition()
                    is_attachment = content_disposition == 'attachment'
                    is_inline = content_disposition == 'inline' or (content_disposition is None and part.get('Content-ID') is not None)

                    if is_attachment or (include_inline and is_inline):
                        # 获取附件文件名
                        filename = part.get_filename()
                        if filename:
                            # 解码文件名
                            decoded_filename = email.header.decode_header(filename)
                            filename_str = ''
                            for part_name, encoding in decoded_filename:
                                if isinstance(part_name, bytes):
                                    filename_str += part_name.decode(encoding or 'utf-8')
                                else:
                                    filename_str += part_name

                            # 检查文件扩展名
                            if file_extensions:
                                ext = os.path.splitext(filename_str.lower())[1][1:]
                                if ext not in file_extensions:
                                    logger.debug(f"跳过附件 {filename_str}，扩展名 {ext} 不在白名单中")
                                    continue

                            # 检查文件大小
                            attachment_size = len(part.get_payload(decode=True))
                            if self.max_attachment_size and attachment_size > self.max_attachment_size * 1024 * 1024:
                                logger.warning(f"跳过附件 {filename_str}，文件大小 {attachment_size/1024/1024:.2f}MB 超过限制 {self.max_attachment_size}MB")
                                continue

                            # 处理文件名冲突
                            # 清理文件名
                            cleaned_filename = self._clean_filename(filename_str)
                            save_path = os.path.join(save_dir, cleaned_filename)
                            base, ext = os.path.splitext(cleaned_filename)
                            counter = 1
                            while os.path.exists(save_path):
                                save_path = os.path.join(save_dir, f"{base}_{counter}{ext}")
                                counter += 1

                            # 保存附件
                            try:
                                with open(save_path, 'wb') as f:
                                    f.write(part.get_payload(decode=True))
                                logger.info(f"已下载附件: {filename_str} 到 {save_path}")
                                downloaded_count += 1
                            except Exception as e:
                                logger.error(f"保存附件 {filename_str} 失败: {str(e)}")
            except Exception as e:
                logger.error(f"处理邮件 {email_id} 时出错: {str(e)}")

        logger.info(f"共下载 {downloaded_count} 个附件")
        return downloaded_count

    def disconnect(self):
        """
        断开与邮箱服务器的连接
        """
        if self.connected and self.imap_server:
            try:
                self.imap_server.close()
                self.imap_server.logout()
                self.connected = False
                logger.info("已断开与邮箱服务器的连接")
            except Exception as e:
                logger.error(f"断开连接时出错: {str(e)}")
                self.connected = False

    def download_attachments_from_body(self, email_ids, save_dir, file_extensions=None):
        """
        从邮件正文提取链接并下载附件
        :param email_ids: 邮件ID列表
        :param save_dir: 附件保存目录
        :param file_extensions: 允许的文件扩展名列表，None表示允许所有
        :return: 成功下载的附件数量
        """
        if not self.connected:
            logger.error("未连接到邮箱服务器")
            return 0

        # 确保保存目录存在
        if not os.path.exists(save_dir):
            try:
                os.makedirs(save_dir)
                logger.info(f"创建保存目录: {save_dir}")
            except Exception as e:
                logger.error(f"创建保存目录失败: {str(e)}")
                return 0

        downloaded_count = 0

        for email_id in email_ids:
            try:
                # 获取邮件内容
                result, data = self.imap_server.fetch(email_id, '(RFC822)')
                if result != 'OK':
                    logger.error(f"获取邮件 {email_id} 失败")
                    continue

                # 解析邮件内容
                raw_email = data[0][1]
                email_message = email.message_from_bytes(raw_email)

                # 获取邮件主题
                subject = email.header.decode_header(email_message['Subject'])
                subject_str = ''
                for part, encoding in subject:
                    if isinstance(part, bytes):
                        subject_str += part.decode(encoding or 'utf-8')
                    else:
                        subject_str += part

                logger.info(f"处理邮件正文链接: {subject_str} (ID: {email_id})")

                # 获取邮件正文
                body = ""
                if email_message.is_multipart():
                    for part in email_message.get_payload():
                        content_type = part.get_content_type()
                        if content_type == 'text/plain' or content_type == 'text/html':
                            charset = part.get_content_charset() or 'utf-8'
                            try:
                                body += part.get_payload(decode=True).decode(charset)
                            except:
                                # 如果解码失败，尝试使用utf-8
                                body += part.get_payload(decode=True).decode('utf-8', errors='replace')
                else:
                    content_type = email_message.get_content_type()
                    charset = email_message.get_content_charset() or 'utf-8'
                    try:
                        body += email_message.get_payload(decode=True).decode(charset)
                    except:
                        body += email_message.get_payload(decode=True).decode('utf-8', errors='replace')

                # 提取链接
                links = []
                # 尝试用BeautifulSoup解析HTML
                try:
                    soup = BeautifulSoup(body, 'html.parser')
                    for a_tag in soup.find_all('a', href=True):
                        links.append(a_tag['href'])
                except:
                    logger.debug(f"解析HTML失败，尝试用正则表达式提取链接")
                    # 用正则表达式提取链接
                    url_pattern = r'https?://\S+|www\.\S+'
                    links = re.findall(url_pattern, body)

                # 过滤链接
                for link in links:
                    # 清理链接
                    link = link.strip()
                    # 跳过javascript链接
                    if link.lower().startswith('javascript:'):
                        continue

                    # 检查文件扩展名
                    if file_extensions:
                        # 获取链接中的文件名
                        filename = os.path.basename(link.split('?')[0])
                        ext = os.path.splitext(filename.lower())[1][1:]
                        if ext not in file_extensions:
                            logger.debug(f"跳过链接 {link}，扩展名 {ext} 不在白名单中")
                            continue

                    try:
                        # 下载链接内容
                        logger.info(f"尝试下载链接: {link}")
                        response = requests.get(link, timeout=30)
                        if response.status_code == 200:
                            # 提取文件名
                            filename = os.path.basename(link.split('?')[0])
                            if not filename:
                                filename = f"attachment_{uuid.uuid4()}"

                            # 清理文件名
                            cleaned_filename = self._clean_filename(filename)
                            save_path = os.path.join(save_dir, cleaned_filename)
                            base, ext = os.path.splitext(cleaned_filename)
                            counter = 1
                            while os.path.exists(save_path):
                                save_path = os.path.join(save_dir, f"{base}_{counter}{ext}")
                                counter += 1

                            # 保存文件
                            with open(save_path, 'wb') as f:
                                f.write(response.content)
                            logger.info(f"已下载链接附件: {filename} 到 {save_path}")
                            downloaded_count += 1
                        else:
                            logger.error(f"下载链接失败: {link}, 状态码: {response.status_code}")
                    except Exception as e:
                        logger.error(f"下载链接 {link} 时出错: {str(e)}")
            except Exception as e:
                logger.error(f"处理邮件 {email_id} 正文链接时出错: {str(e)}")

        logger.info(f"共从邮件正文下载 {downloaded_count} 个附件")
        return downloaded_count


def main():
    """
    主函数，用于演示如何使用EmailAttachmentDownloader类
    """
    # 初始化下载器
    downloader = EmailAttachmentDownloader()

    try:
        # 连接到邮箱
        if not downloader.connect():
            logger.error("无法连接到邮箱，程序退出")
            return

        # 计算指定天数前的日期
        date_since = datetime.datetime.now() - datetime.timedelta(days=downloader.days_ago or 7)

        # 搜索邮件
        email_ids = downloader.search_emails(folder=downloader.folder, date_since=date_since)

        # 如果没有找到邮件，搜索所有邮件
        if not email_ids:
            logger.warning(f"没有找到{(downloader.days_ago or 7)}天内的邮件，搜索所有邮件")
            email_ids = downloader.search_emails(folder=downloader.folder)

        # 确保保存目录存在
        save_dir = os.path.join(os.getcwd(), downloader.save_dir)

        # 从配置中获取文件扩展名过滤
        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.json')
        file_extensions = None
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
                email_config = config.get('email', {})
                file_extensions = email_config.get('file_extensions')
        except Exception as e:
            logger.error(f"读取配置文件失败: {str(e)}")

        # 下载附件
        downloader.download_attachments(email_ids, save_dir, file_extensions=file_extensions)
        # 从邮件正文下载附件
        downloader.download_attachments_from_body(email_ids, save_dir, file_extensions=file_extensions)
    except Exception as e:
        logger.error(f"程序执行出错: {str(e)}")
    finally:
        # 断开连接
        downloader.disconnect()


if __name__ == '__main__':
    main()