import imaplib
import email
import os
import datetime
import json

class EmailAttachmentDownloader:
    def __init__(self):
        self.imap_server = None
        self.username = None
        self.password = None
        self.connected = False

    def connect(self, imap_server, username, password, port=993):
        """
        连接到IMAP邮箱服务器
        :param imap_server: IMAP服务器地址
        :param username: 邮箱用户名
        :param password: 邮箱密码或授权码
        :param port: 服务器端口，默认993（SSL）
        :return: 是否连接成功
        """
        try:
            # 连接服务器
            self.imap_server = imaplib.IMAP4_SSL(imap_server, port)
            # 登录账户
            self.imap_server.login(username, password)
            self.username = username
            self.password = password
            self.connected = True
            print(f"成功连接到邮箱: {username}")
            return True
        except Exception as e:
            print(f"连接邮箱失败: {str(e)}")
            self.connected = False
            return False

    def search_emails(self, folder='INBOX', criteria='ALL', date_since=None):
        """
        搜索邮件
        :param folder: 邮箱文件夹，默认收件箱(INBOX)
        :param criteria: 搜索条件，默认所有邮件(ALL)
                        其他条件示例: 'UNSEEN'（未读）, 'FROM "example@domain.com"'（来自特定发件人）
        :return: 邮件ID列表
        """
        if not self.connected:
            print("未连接到邮箱服务器")
            return []

        try:
            # 选择文件夹
            self.imap_server.select(folder)
            # 如果提供了日期参数，添加到搜索条件
            if date_since:
                # 格式化日期为IMAP格式 (DD-Month-YYYY)
                imap_date = date_since.strftime('%d-%b-%Y')
                criteria = f'(SINCE {imap_date}) {criteria}'
                print(f"搜索条件: {criteria}")

            # 搜索邮件
            result, data = self.imap_server.search(None, criteria)
            if result == 'OK':
                # 提取邮件ID
                email_ids = data[0].split()
                print(f"找到 {len(email_ids)} 封邮件")
                return email_ids
            else:
                print("搜索邮件失败")
                return []
        except Exception as e:
            print(f"搜索邮件时出错: {str(e)}")
            return []

    def download_attachments(self, email_ids, save_dir, file_extensions=None):
        """
        下载邮件附件
        :param email_ids: 邮件ID列表
        :param save_dir: 保存附件的目录
        :param file_extensions: 要下载的文件扩展名列表，如['pdf', 'doc']，为None时下载所有附件
        :return: 成功下载的附件数量
        """
        if not self.connected:
            print("未连接到邮箱服务器")
            return 0

        # 确保保存目录存在
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)
            print(f"创建保存目录: {save_dir}")

        downloaded_count = 0

        for email_id in email_ids:
            try:
                # 获取邮件内容
                result, data = self.imap_server.fetch(email_id, '(RFC822)')
                if result != 'OK':
                    print(f"获取邮件 {email_id} 失败")
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

                print(f"处理邮件: {subject_str} (ID: {email_id})")

                # 遍历邮件的各个部分
                for part in email_message.walk():
                    # 检查是否是附件
                    if part.get_content_disposition() == 'attachment':
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
                                    continue

                            # 保存附件
                            save_path = os.path.join(save_dir, filename_str)
                            with open(save_path, 'wb') as f:
                                f.write(part.get_payload(decode=True))
                            print(f"已下载附件: {filename_str}")
                            downloaded_count += 1
            except Exception as e:
                print(f"处理邮件 {email_id} 时出错: {str(e)}")

        print(f"共下载 {downloaded_count} 个附件")
        return downloaded_count

    def disconnect(self):
        """
        断开与邮箱服务器的连接
        """
        if self.connected and self.imap_server:
            self.imap_server.close()
            self.imap_server.logout()
            self.connected = False
            print("已断开与邮箱服务器的连接")


def main():
    # 示例用法
    # 初始化下载器
    downloader = EmailAttachmentDownloader()

    try:
        # 读取配置文件（从当前脚本所在目录）
        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.json')
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
            email_config = config.get('email', {})

        # 从配置中获取参数
        imap_server = email_config.get('imap_server', 'imaphz.qiye.163.com')
        username = email_config.get('username', 'heguangzhong@cosmosource.com')
        password = email_config.get('password', 'Password01')
        port = email_config.get('port', 993)
        folder = email_config.get('folder', 'INBOX')
        days_ago = email_config.get('days_ago', 7)
        save_dir = email_config.get('save_dir', 'attachments')
        file_extensions = email_config.get('file_extensions', ['pdf', 'jpg', 'jpeg', 'png'])

        # 连接到邮箱
        if not downloader.connect(imap_server, username, password, port):
            print("无法连接到邮箱，程序退出")
            return

        # 计算指定天数前的日期
        date_since = datetime.datetime.now() - datetime.timedelta(days=days_ago)

        # 搜索邮件
        email_ids = downloader.search_emails(folder=folder, date_since=date_since)

        # 如果没有找到邮件，搜索所有邮件
        if not email_ids:
            print(f"没有找到{days_ago}天内的邮件，搜索所有邮件")
            email_ids = downloader.search_emails(folder=folder)

        # 确保保存目录存在
        save_dir = os.path.join(os.getcwd(), save_dir)

        # 下载附件
        downloader.download_attachments(email_ids, save_dir, file_extensions=file_extensions)
    except Exception as e:
        print(f"程序执行出错: {str(e)}")
    finally:
        # 断开连接
        downloader.disconnect()


if __name__ == '__main__':
    main()