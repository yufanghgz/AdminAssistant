#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
8月份发票识别脚本
识别D:\8月份发票目录下的所有电子发票
"""

import os
import sys
from base.invoice_ocr import InvoiceRecognizer

def main():
    """
    主函数：识别8月份发票目录下的所有电子发票
    """
    # 设置控制台编码为UTF-8
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
    
    # 目标目录
    target_dir = r'D:\8月份发票'
    
    print("=" * 60)
    print("8月份发票识别系统")
    print("=" * 60)
    
    # 检查目录是否存在
    if not os.path.exists(target_dir):
        print(f"错误：目录 '{target_dir}' 不存在！")
        print("请确保D盘存在'8月份发票'文件夹")
        return
    
    print(f"开始处理目录: {target_dir}")
    
    # 创建发票识别器
    recognizer = InvoiceRecognizer()
    
    try:
        # 批量识别发票
        print("\n开始识别发票...")
        results = recognizer.batch_recognize_invoices(target_dir)
        
        if not results:
            print("未找到任何可识别的发票文件")
            return
        
        # 显示识别结果
        print(f"\n识别完成！共识别 {len(results)} 张发票")
        print("-" * 60)
        
        for i, result in enumerate(results, 1):
            print(f"发票 {i}:")
            print(f"  文件路径: {os.path.basename(result['file_path'])}")
            print(f"  日期: {result['date']}")
            print(f"  金额: {result['amount']}")
            print(f"  内容: {result['content'][:50]}...")  # 只显示前50个字符
            print()
        
        # 查找金额最大的发票
        max_invoice = recognizer.get_max_amount_invoice(results)
        if max_invoice:
            print("=" * 60)
            print("金额最大的发票:")
            print(f"  文件: {os.path.basename(max_invoice['file_path'])}")
            print(f"  日期: {max_invoice['date']}")
            print(f"  金额: {max_invoice['amount']}")
            print(f"  内容: {max_invoice['content'][:50]}...")
        
        print("\n处理完成！")
        print(f"识别后的文件已保存在: {target_dir}")
        
    except Exception as e:
        print(f"处理过程中出现错误: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    main()

