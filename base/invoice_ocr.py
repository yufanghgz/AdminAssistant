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
import logging
import pandas as pd
import os

# 设置默认编码为UTF-8
sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

# 设置日志记录器，避免调试信息输出到标准输出
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

# 创建用户logs目录
user_home = os.path.expanduser("~")
logs_dir = os.path.join(user_home, "logs")
os.makedirs(logs_dir, exist_ok=True)

# 创建文件处理器，将日志写入用户logs目录
if not logger.handlers:
    log_file = os.path.join(logs_dir, 'invoice_ocr.log')
    handler = logging.FileHandler(log_file, encoding='utf-8')
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)

class InvoiceRecognizer:
    def __init__(self):
        # 延迟初始化EasyOCR读取器，避免启动时卡住
        self.reader = None
        # 确保文件路径使用UTF-8编码
        self.file_path_encoding = 'utf-8'
    
    def _get_reader(self):
        """延迟初始化EasyOCR读取器"""
        if self.reader is None:
            try:
                logger.info("正在初始化EasyOCR读取器...")
                self.reader = easyocr.Reader(['ch_sim', 'en'], gpu=False)  # 禁用GPU加速，避免兼容性问题
                logger.info("EasyOCR读取器初始化完成")
            except Exception as e:
                logger.error(f"EasyOCR初始化失败: {str(e)}")
                raise e
        return self.reader

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
        import signal
        import time
        
        def timeout_handler(signum, frame):
            raise TimeoutError("OCR处理超时")
        
        try:
            # 确保文件路径使用UTF-8编码
            image_path = str(image_path)
            logger.info(f"开始处理图片: {image_path}")
            
            # 设置超时机制（60秒）
            signal.signal(signal.SIGALRM, timeout_handler)
            signal.alarm(60)
            
            # 获取EasyOCR读取器
            reader = self._get_reader()
            
            # 使用EasyOCR识别文本
            result = reader.readtext(image_path)
            
            # 取消超时
            signal.alarm(0)
            
            # 合并识别结果
            text = ' '.join([item[1] for item in result])
            logger.info(f"图片处理完成，识别到 {len(result)} 个文本块")
            return text
            
        except TimeoutError:
            logger.error(f"图片OCR处理超时: {image_path}")
            return ''
        except Exception as e:
            logger.error(f"从图片提取文本时出错: {str(e)}")
            return ''
        finally:
            # 确保取消超时
            signal.alarm(0)

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

        # 提取金额（匹配常见金额格式，特别优化了运输服务发票的金额识别）
        amount_patterns = [
            # 价税合计（通常是最重要的金额）
            r'(?:价税合计|合计|总计|金额合计)[\s]*[:：]?[\s]*¥?[\s]*([\d,]+\.\d{2})',
            r'(?:价税合计|合计|总计|金额合计)[\s]*([\d,]+\.\d{2})',
            # 小写金额
            r'(?:小写|金额)[\s]*[:：]?[\s]*¥?[\s]*([\d,]+\.\d{2})',
            r'(?:小写|金额)[\s]*([\d,]+\.\d{2})',
            # 餐饮服务特有的金额模式
            r'(?:餐饮服务|餐饮费|餐费|消费金额)[\s]*[:：]?[\s]*¥?[\s]*([\d,]+\.\d{2})',
            r'(?:餐饮服务|餐饮费|餐费|消费金额)[\s]*([\d,]+\.\d{2})',
            # 运输服务特有的金额模式（优先识别）
            r'(?:运费|运输费|服务费|车费)[\s]*[:：]?[\s]*¥?[\s]*([\d,]+\.\d{2})',
            r'(?:运费|运输费|服务费|车费)[\s]*([\d,]+\.\d{2})',
            # 税额
            r'(?:税额|增值税)[\s]*[:：]?[\s]*¥?[\s]*([\d,]+\.\d{2})',
            r'(?:税额|增值税)[\s]*([\d,]+\.\d{2})',
            # 文件名中的金额（运输服务发票常见）
            r'(\d+\.\d{2})元',
            # 金额（通用）
            r'¥([\d,]+\.\d{2})',
            r'([\d,]+\.\d{2})元',
            # 大写金额对应的数字（通常在发票底部）
            r'(?:大写|金额大写)[\s]*[:：]?[\s]*[^\d]*([\d,]+\.\d{2})'
        ]

        # 收集所有匹配的金额及其类型
        amounts_with_type = []
        for i, pattern in enumerate(amount_patterns):
            matches = re.findall(pattern, text)
            for match in matches:
                amounts_with_type.append((match, i))

        # 从文件名中提取金额（运输服务发票的特殊处理）
        filename = os.path.basename(file_path)
        filename_amount_match = re.search(r'(\d+\.\d{2})元', filename)
        if filename_amount_match:
            filename_amount = filename_amount_match.group(1)
            amounts_with_type.append((filename_amount, 10))  # 文件名金额优先级最高

        # 智能选择金额：优先选择价税合计，其次是小写金额，最后是其他金额
        if amounts_with_type:
            # 按优先级排序：文件名金额(8) > 价税合计(0-1) > 小写金额(2-3) > 餐饮服务金额(4-5) > 运输服务金额(6-7) > 税额(8-9) > 其他(10+)
            priority_order = []
            
            # 首先添加文件名中的金额
            for amount, pattern_idx in amounts_with_type:
                if pattern_idx == 10:  # 文件名金额
                    priority_order.append((amount, 0))
            
            # 然后添加价税合计
            for amount, pattern_idx in amounts_with_type:
                if pattern_idx <= 1:  # 价税合计模式
                    priority_order.append((amount, 1))
            
            # 然后添加小写金额
            for amount, pattern_idx in amounts_with_type:
                if 2 <= pattern_idx <= 3:  # 小写金额模式
                    priority_order.append((amount, 2))
            
            # 然后添加餐饮服务金额
            for amount, pattern_idx in amounts_with_type:
                if 4 <= pattern_idx <= 5:  # 餐饮服务金额模式
                    priority_order.append((amount, 3))
            
            # 然后添加运输服务金额
            for amount, pattern_idx in amounts_with_type:
                if 6 <= pattern_idx <= 7:  # 运输服务金额模式
                    priority_order.append((amount, 4))
            
            # 然后添加税额
            for amount, pattern_idx in amounts_with_type:
                if 8 <= pattern_idx <= 9:  # 税额模式
                    priority_order.append((amount, 5))
            
            # 最后添加其他金额
            for amount, pattern_idx in amounts_with_type:
                if pattern_idx >= 10:  # 其他模式
                    priority_order.append((amount, 6))
            
            # 按优先级选择金额
            if priority_order:
                # 选择优先级最高的金额（数字越小优先级越高）
                selected_amount, priority = min(priority_order, key=lambda x: x[1])
                
                # 在相同优先级的情况下，选择最大的金额
                same_priority_amounts = [amount for amount, p in priority_order if p == priority]
                if len(same_priority_amounts) > 1:
                    # 转换为浮点数进行比较
                    float_amounts = []
                    for amount in same_priority_amounts:
                        try:
                            float_amounts.append(float(amount.replace(',', '')))
                        except ValueError:
                            continue
                    
                    if float_amounts:
                        max_amount = max(float_amounts)
                        selected_amount = f"{max_amount:.2f}"
                        print(f"  在相同优先级({priority})中选择最大金额: {selected_amount}")
                
                result['amount'] = selected_amount
                result['all_amounts'] = [amount for amount, _ in amounts_with_type]
                result['amount_priority'] = priority  # 0=文件名, 1=价税合计, 2=小写金额, 3=餐饮服务, 4=运输服务, 5=税额, 6=其他
                
                # 添加调试信息
                print(f"发票金额识别调试信息:")
                print(f"  找到的所有金额: {result['all_amounts']}")
                print(f"  优先级排序: {priority_order}")
                print(f"  选择的金额: {selected_amount} (优先级: {priority})")
                priority_names = {0: "文件名", 1: "价税合计", 2: "小写金额", 3: "餐饮服务", 4: "运输服务", 5: "税额", 6: "其他"}
                print(f"  优先级说明: {priority_names.get(priority, '未知')}")
                
                # 如果没有找到高优先级金额，尝试选择最大的金额
                if priority >= 3 and len(amounts_with_type) > 1:
                    # 转换为浮点数进行比较
                    float_amounts = []
                    for amount, _ in amounts_with_type:
                        try:
                            float_amounts.append(float(amount.replace(',', '')))
                        except ValueError:
                            continue
                    
                    if float_amounts:
                        max_amount = max(float_amounts)
                        # 如果最大金额明显大于当前选择的金额，使用最大金额
                        current_amount = float(selected_amount.replace(',', ''))
                        if max_amount > current_amount * 1.05:  # 如果最大金额比当前金额大5%以上
                            result['amount'] = f"{max_amount:.2f}"
                            print(f"  调整选择最大金额: {result['amount']} (原选择: {selected_amount})")

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
        logger.info(f"处理目录: {input_dir}")
        logger.info(f"开始处理目录: {input_dir}")
        
        # 检查输入目录是否存在
        if not os.path.exists(input_dir):
            error_msg = f"错误: 目录 '{input_dir}' 不存在。"
            logger.error(error_msg)
            logger.error(error_msg)
            return results

        # 创建以当前时间命名的文件夹
        current_time = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        output_dir = os.path.join(input_dir, current_time)
        os.makedirs(output_dir, exist_ok=True)
        logger.info(f"已创建输出目录: {output_dir}")

        # 获取目录中的所有文件（确保正确处理中文文件名）
        try:
            # 使用scandir代替listdir，更好地处理编码问题
            with os.scandir(input_dir) as entries:
                files = [entry.name for entry in entries if entry.is_file()]
        except Exception as e:
            logger.error(f"获取目录文件时出错: {str(e)}")
            # 回退到listdir
            files = os.listdir(input_dir)

        logger.info(f"找到 {len(files)} 个文件")

        processed_count = 0
        total_files = len([f for f in files if os.path.splitext(f.lower())[1] in ['.pdf', '.jpg', '.jpeg', '.png', '.bmp']])
        
        for file in files:
            file_path = os.path.join(input_dir, file)
            # 只处理PDF和图片文件
            ext = os.path.splitext(file.lower())[1]
            if ext in ['.pdf', '.jpg', '.jpeg', '.png', '.bmp']:
                processed_count += 1
                logger.info(f"识别发票 ({processed_count}/{total_files}): {file}")
                # 使用logger而不是print，避免MCP客户端JSON解析错误
                logger.info(f"Processing ({processed_count}/{total_files}): {file}")
                
                result = None
                try:
                    result = self.recognize_invoice(file_path)
                    results.append(result)
                    logger.info(f"发票识别完成: {file}")
                except Exception as e:
                    logger.error(f"识别发票失败 {file}: {str(e)}")
                    # 添加失败的结果
                    result = {
                        'file_path': file_path,
                        'date': '',
                        'amount': '',
                        'all_amounts': [],
                        'content': '',
                        'error': str(e)
                    }
                    results.append(result)
                
                # 连接识别信息作为新文件名
                if result and 'date' in result and 'amount' in result and 'content' in result and 'error' not in result:
                    # 构建新文件名，去掉特殊字符
                    date_str = result['date'].replace('/', '').replace('年', '').replace('月', '').replace('日', '')
                    amount_str = str(result['amount']).replace(',', '')
                    content_str = result['content'][:10]  # 取内容前10个字
                    
                    # 组合文件名
                    combined_str = f"{date_str}_{amount_str}_{content_str}{ext}"
                    
                    # 过滤文件名中的非法字符
                    invalid_chars = '<>:/\\|?*"'  # 完整的Windows文件名非法字符
                    valid_filename = re.sub(f'[{re.escape(invalid_chars)}]', '', combined_str)
                    
                    # 复制文件到新目录并重命名
                    new_file_path = os.path.join(output_dir, valid_filename)
                    try:
                        logger.info(f"准备复制文件: {file_path} -> {new_file_path}")
                        shutil.copy2(file_path, new_file_path)
                        logger.info(f"已复制并重命名文件: {new_file_path}")
                    except Exception as e:
                        logger.error(f"复制文件时出错: {str(e)}")
                        # 尝试使用shutil.copy作为备选
                        try:
                            logger.info(f"尝试使用备用方法复制文件...")
                            shutil.copy(file_path, new_file_path)
                            logger.info(f"已使用备用方法复制文件: {new_file_path}")
                        except Exception as e2:
                            logger.error(f"备用方法也失败: {str(e2)}")
                else:
                    logger.warning(f"发票识别结果不完整，跳过文件复制: {file}")
            else:
                logger.info(f"跳过非支持文件类型: {file}")

        success_count = len([r for r in results if 'error' not in r])
        error_count = len([r for r in results if 'error' in r])
        
        completion_msg = f"批量处理完成！共处理 {len(results)} 张发票，成功 {success_count} 张，失败 {error_count} 张"
        logger.info(completion_msg)
        logger.info(completion_msg)
        
        if error_count > 0:
            logger.warning("失败的文件已记录在日志中，请查看 ~/logs/invoice_ocr.log")
        
        # 生成Excel统计表
        try:
            excel_file_path = self.generate_excel_report(results, output_dir)
            if excel_file_path:
                logger.info(f"发票统计Excel表已生成: {excel_file_path}")
                print(f"发票统计Excel表已生成: {excel_file_path}")
            else:
                logger.warning("无法生成发票统计Excel表")
        except Exception as e:
            logger.error(f"生成Excel表时发生异常: {str(e)}")
        
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

    def generate_excel_report(self, invoice_results, output_dir=None):
        """
        生成Excel统计表，记录每张发票的相关信息并统计总金额
        :param invoice_results: 发票识别结果列表
        :param output_dir: 输出目录路径，如果为None则使用当前时间文件夹
        :return: 生成的Excel文件路径
        """
        if not invoice_results:
            logger.warning("没有发票识别结果，无法生成Excel报表")
            return None
            
        # 准备数据
        data = []
        total_amount = 0.0
        
        for i, result in enumerate(invoice_results, 1):
            # 跳过处理失败的发票
            if 'error' in result:
                continue
                
            # 提取发票信息
            invoice_id = f"发票{i}"
            date = result.get('date', '')
            amount = result.get('amount', '')
            content = result.get('content', '')
            file_path = result.get('file_path', '')
            file_name = os.path.basename(file_path)
            
            # 添加到数据列表
            data.append({
                '发票号码': invoice_id,
                '日期': date,
                '内容': content,
                '金额': amount,
                '文件名': file_name,
                '文件路径': file_path
            })
            
            # 累加总金额
            try:
                if amount:
                    total_amount += float(amount.replace(',', ''))
            except ValueError:
                logger.warning(f"无法将金额 '{amount}' 转换为数字")
                
        # 如果没有成功识别的发票，返回None
        if not data:
            logger.warning("没有成功识别的发票，无法生成Excel报表")
            return None
            
        # 创建DataFrame
        df = pd.DataFrame(data)
        
        # 添加总计行
        total_row = pd.DataFrame({
            '发票号码': ['总计'],
            '日期': [''],
            '内容': [''],
            '金额': [f"{total_amount:.2f}"],
            '文件名': [''],
            '文件路径': ['']
        })
        df = pd.concat([df, total_row], ignore_index=True)
        
        # 确定输出目录
        if output_dir is None:
            # 使用当前时间创建文件夹
            current_time = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            # 获取第一个发票的目录作为基础目录
            base_dir = os.path.dirname(invoice_results[0]['file_path'])
            output_dir = os.path.join(base_dir, current_time)
            os.makedirs(output_dir, exist_ok=True)
        
        # 创建Excel文件路径
        excel_file_name = f"发票统计_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        excel_file_path = os.path.join(output_dir, excel_file_name)
        
        try:
            # 将DataFrame写入Excel文件
            with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='发票统计')
                
                # 格式化工作表
                worksheet = writer.sheets['发票统计']
                
                # 设置列宽
                worksheet.column_dimensions['A'].width = 10  # 发票ID
                worksheet.column_dimensions['B'].width = 20  # 日期
                worksheet.column_dimensions['C'].width = 30  # 内容
                worksheet.column_dimensions['D'].width = 15  # 金额
                worksheet.column_dimensions['E'].width = 30  # 文件名
                worksheet.column_dimensions['F'].width = 50  # 文件路径
                
                # 高亮总计行
                for cell in worksheet[f'{worksheet.dimensions.split(":")[0].split("1")[0]}{len(df)}']:
                    cell.font = cell.font.copy(bold=True)
                
            logger.info(f"已生成Excel统计表: {excel_file_path}")
            return excel_file_path
        except Exception as e:
            logger.error(f"生成Excel统计表时出错: {str(e)}")
            return None



def main():
    # 示例用法
    recognizer = InvoiceRecognizer()

    # 批量识别（注意：实际使用时请修改为您的发票目录路径）
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
    
    # 注意：由于batch_recognize_invoices方法已经自动调用generate_excel_report生成Excel表，
    # 所以这里不需要再次调用。Excel表会保存在与发票重命名后的文件相同的时间戳文件夹中。

if __name__ == '__main__':
    main()