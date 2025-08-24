import os
import re
import sys
import pdfplumber
from PIL import Image
import cv2
import numpy as np
import easyocr
import datetime
import shutil

# 设置默认编码为UTF-8
sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

class InvoiceRecognizer:
    def __init__(self):
        # 初始化EasyOCR读取器
        self.reader = easyocr.Reader(['ch_sim', 'en'])  # 中文简体和英文
        # 确保文件路径使用UTF-8编码
        self.file_path_encoding = 'utf-8'

    def extract_text_from_pdf(self, pdf_path):
        """
        从PDF文件中提取文本
        :param pdf_path: PDF文件路径
        :return: 提取的文本
        """
        text = ''
        try:
            # 确保文件路径使用UTF-8编码
            pdf_path = str(pdf_path)
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    # 提取文本
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + '\n'
        except Exception as e:
            print(f"从PDF提取文本时出错: {str(e)}")
        return text

    def extract_text_from_image(self, image_path):
        """
        从图片中提取文本
        :param image_path: 图片文件路径
        :return: 提取的文本
        """
        try:
            # 确保文件路径使用UTF-8编码
            image_path = str(image_path)
            # 使用EasyOCR识别文本
            result = self.reader.readtext(image_path)
            # 合并识别结果
            text = ' '.join([item[1] for item in result])
            return text
        except Exception as e:
            print(f"从图片提取文本时出错: {str(e)}")
            return ''

    def recognize_invoice(self, file_path):
        """
        识别发票信息
        :param file_path: 发票文件路径（PDF或图片）
        :return: 识别结果字典
        """
        # 确保文件路径使用UTF-8编码
        file_path = str(file_path)
        result = {
            'file_path': file_path,
            'date': '',
            'amount': '',
            'all_amounts': [],  # 存储所有找到的金额
            'content': ''
        }

        # 根据文件扩展名选择不同的处理方式
        ext = os.path.splitext(file_path.lower())[1]

        if ext == '.pdf':
            text = self.extract_text_from_pdf(file_path)
        elif ext in ['.jpg', '.jpeg', '.png', '.bmp']:
            text = self.extract_text_from_image(file_path)
        else:
            print(f"不支持的文件格式: {ext}")
            return result

        # 提取日期（匹配常见日期格式）
        date_patterns = [
            r'\d{4}[年|/|\-]\d{1,2}[月|/|\-]\d{1,2}[日]?',  # 2023年10月1日 或 2023/10/1 或 2023-10-1
            r'\d{2}[年|/|\-]\d{1,2}[月|/|\-]\d{1,2}[日]?'   # 23年10月1日 或 23/10/1 或 23-10-1
        ]

        for pattern in date_patterns:
            date_match = re.search(pattern, text)
            if date_match:
                result['date'] = date_match.group(0)
                break

        # 提取金额（匹配常见金额格式，特别优化了"小写"字符右边的金额识别）
        amount_patterns = [
            r'(?:小写|金额)[\s]*[:：]?[\s]*¥?[\s]*([\d,]+\.\d{2})',  # 小写: ¥1,234.56 或 金额: 1,234.56
            r'(?:小写|金额)[\s]*([\d,]+\.\d{2})',                    # 小写1,234.56 或 金额1,234.56
            r'¥([\d,]+\.\d{2})',                                    # ¥1,234.56
            r'([\d,]+\.\d{2})元'                                    # 1,234.56元
        ]

        # 收集所有匹配的金额
        amounts = []
        for pattern in amount_patterns:
            # 使用findall找到所有匹配项
            matches = re.findall(pattern, text)
            amounts.extend(matches)

        # 找出最大金额
        if amounts:
            # 转换为浮点数进行比较
            max_amount = max(float(amount.replace(',', '')) for amount in amounts)
            # 格式化为保留2位小数的字符串
            result['amount'] = f"{max_amount:.2f}"
            # 保存所有找到的金额
            result['all_amounts'] = amounts

        # 提取内容（优先提取以*号开头的内容）
        # 尝试提取以*号开头并以下一个*号结束的内容
        star_pattern = r'\*(.*?)\*'
        star_match = re.search(star_pattern, text, re.DOTALL)
        if star_match:
            result['content'] = star_match.group(1).strip()
        else:
            # 如果没有找到以*号开头的内容，尝试提取"项目名称"后的内容
            content_pattern = r'(?:项目名称|项目|品名)[\s]*[:：]?[\s]*(.*?)(?:[\n]|金额|合计|数量|单位|单价|\Z)'
            content_match = re.search(content_pattern, text, re.DOTALL | re.IGNORECASE)
            if content_match:
                result['content'] = content_match.group(1).strip()
            else:
                # 如果没有找到"项目名称"相关内容，尝试提取"货物或应税劳务、服务名称"后的内容
                fallback_pattern = r'[货物|服务][\s]*[或|及][\s]*[应税劳务|服务][\s]*[名称]?[\s]*[:：]?[\s]*(.*?)(?:[\n]|\s|金额|合计)'
                fallback_match = re.search(fallback_pattern, text, re.DOTALL)
                if fallback_match:
                    result['content'] = fallback_match.group(1).strip()
                else:
                    # 如果没有找到特定模式，就取文本的前200个字符作为内容
                    result['content'] = text[:200].strip()

        return result

    def batch_recognize_invoices(self, input_dir):
        """
        批量识别目录下的所有发票，并将识别后的信息连接起来，将原有发票文件复制一份存入以当前时间命名的文件夹，并用连接起来的字串重新命名
        :param input_dir: 包含发票的目录
        :return: 识别结果列表
        """
        import shutil
        import datetime
        import os
        import re
        
        results = []

        # 确保目录路径使用UTF-8编码
        input_dir = str(input_dir)
        print(f"处理目录: {input_dir}")
        
        # 检查输入目录是否存在
        if not os.path.exists(input_dir):
            print(f"错误: 目录 '{input_dir}' 不存在。")
            return results

        # 创建以当前时间命名的文件夹
        current_time = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        output_dir = os.path.join(input_dir, current_time)
        os.makedirs(output_dir, exist_ok=True)
        print(f"已创建输出目录: {output_dir}")

        # 获取目录中的所有文件（确保正确处理中文文件名）
        try:
            # 使用scandir代替listdir，更好地处理编码问题
            with os.scandir(input_dir) as entries:
                files = [entry.name for entry in entries if entry.is_file()]
        except Exception as e:
            print(f"获取目录文件时出错: {str(e)}")
            # 回退到listdir
            files = os.listdir(input_dir)

        print(f"找到 {len(files)} 个文件")

        for file in files:
            file_path = os.path.join(input_dir, file)
            # 只处理PDF和图片文件
            ext = os.path.splitext(file.lower())[1]
            if ext in ['.pdf', '.jpg', '.jpeg', '.png', '.bmp']:
                print(f"识别发票: {file}")
                result = self.recognize_invoice(file_path)
                results.append(result)
                
                # 连接识别信息作为新文件名
                if result and 'date' in result and 'amount' in result and 'content' in result:
                    # 构建新文件名，去掉特殊字符
                    date_str = result['date'].replace('/', '').replace('年', '').replace('月', '').replace('日', '')
                    amount_str = str(result['amount']).replace(',', '')
                    content_str = result['content'][:10]  # 取内容前10个字
                    
                    # 组合文件名
                    combined_str = f"{date_str}_{amount_str}_{content_str}{ext}"
                    
                    # 过滤文件名中的非法字符
                    invalid_chars = '<>:/\|?*"'  # 完整的Windows文件名非法字符
                    valid_filename = re.sub(f'[{re.escape(invalid_chars)}]', '', combined_str)
                    
                    # 复制文件到新目录并重命名
                    new_file_path = os.path.join(output_dir, valid_filename)
                    try:
                        print(f"准备复制文件: {file_path} -> {new_file_path}")
                        shutil.copy2(file_path, new_file_path)
                        print(f"已复制并重命名文件: {new_file_path}")
                    except Exception as e:
                        print(f"复制文件时出错: {str(e)}")
                        # 尝试使用shutil.copy作为备选
                        try:
                            print(f"尝试使用备用方法复制文件...")
                            shutil.copy(file_path, new_file_path)
                            print(f"已使用备用方法复制文件: {new_file_path}")
                        except Exception as e2:
                            print(f"备用方法也失败: {str(e2)}")
                else:
                    print(f"发票识别结果不完整，跳过文件复制: {file}")
            else:
                print(f"跳过非支持文件类型: {file}")

        print(f"批量处理完成，共识别 {len(results)} 张发票")
        return results

        return results

    def get_max_amount_invoice(self, invoice_results):
        """
        获取发票识别结果中金额最大的发票
        :param invoice_results: 发票识别结果列表
        :return: 金额最大的发票信息字典，如果没有有效金额则返回None
        """
        if not invoice_results:
            return None

        max_amount = -1
        max_invoice = None

        for invoice in invoice_results:
            try:
                # 去除金额中的逗号并转换为浮点数
                amount = float(invoice['amount'].replace(',', ''))
                if amount > max_amount:
                    max_amount = amount
                    max_invoice = invoice
            except (ValueError, TypeError):
                # 忽略无法转换为数字的金额
                continue

        return max_invoice

    def create_timestamp_folder(self, base_dir='.'):
        """
        以当前时间创建文件夹
        :param base_dir: 基础目录
        :return: 创建的文件夹路径
        """
        # 获取当前时间并格式化
        current_time = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        # 创建文件夹路径
        folder_path = os.path.join(base_dir, current_time)
        # 创建文件夹
        os.makedirs(folder_path, exist_ok=True)
        print(f"创建文件夹: {folder_path}")
        return folder_path

    def copy_and_rename_invoices(self, invoice_results, target_folder):
        """
        将发票复制到目标文件夹并重命名
        :param invoice_results: 发票识别结果列表
        :param target_folder: 目标文件夹路径
        """
        for result in invoice_results:
            try:
                # 获取原文件路径
                source_path = result['file_path']
                # 获取文件扩展名
                ext = os.path.splitext(source_path)[1]
                # 拼接新文件名（日期+内容+金额+扩展名）
                combined_str = f"{result['date']}{result['content']}{result['amount']}"
                # 移除文件名中不允许的字符
                valid_filename = re.sub(r'[<>:"/\\|?*]', '', combined_str)
                # 完整目标路径
                target_path = os.path.join(target_folder, f"{valid_filename}{ext}")
                # 复制文件
                shutil.copy2(source_path, target_path)
                print(f"已复制并重命名: {source_path} -> {target_path}")
            except Exception as e:
                print(f"复制文件时出错: {str(e)}")



def main():
    # 示例用法
    recognizer = InvoiceRecognizer()

    # 批量识别
    input_dir = r'd:\PyProject\NounParse\yufang发票'
    results = recognizer.batch_recognize_invoices(input_dir)

    # 打印结果
    print("\n批量识别结果:")
    for i, result in enumerate(results, 1):
        print(f"发票 {i}:")
        print(f"文件路径: {result['file_path']}")
        print(f"日期: {result['date']}")
        print(f"金额: {result['amount']}")
        print(f"内容: {result['content']}")
        # 将日期、内容、金额拼成一个字符串并打印
        combined_str = f"{result['date']}{result['content']}{result['amount']}"
        print(f"拼接结果: {combined_str}\n")

    # 创建以当前时间命名的文件夹，与原始电子发票在同一个位置
    target_folder = recognizer.create_timestamp_folder(r'd:\PyProject\NounParse\yufang发票')

    # 复制并重命名发票到目标文件夹
    recognizer.copy_and_rename_invoices(results, target_folder)

if __name__ == '__main__':
    main()