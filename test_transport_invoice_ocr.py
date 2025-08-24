#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
运输服务电子发票金额识别测试脚本
用于测试和验证运输服务发票的金额识别是否正确
"""

import os
import sys
from base.invoice_ocr import InvoiceRecognizer

def test_transport_invoice_recognition():
    """
    测试运输服务发票识别功能
    """
    # 设置控制台编码为UTF-8
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
    
    print("=" * 60)
    print("运输服务电子发票金额识别测试")
    print("=" * 60)
    
    # 创建发票识别器
    recognizer = InvoiceRecognizer()
    
    # 测试目录 - 请根据实际情况修改
    test_dirs = [
        r'D:\PyProject\AdminAssistent\attachments',  # 8月份发票目录

    ]
    
    for test_dir in test_dirs:
        if os.path.exists(test_dir):
            print(f"\n测试目录: {test_dir}")
            print("-" * 40)
            
            try:
                # 批量识别发票
                results = recognizer.batch_recognize_invoices(test_dir)
                
                if not results:
                    print("未找到任何可识别的发票文件")
                    continue
                
                print(f"识别完成！共识别 {len(results)} 张发票")
                print()
                
                # 显示每张发票的详细信息
                for i, result in enumerate(results, 1):
                    print(f"发票 {i}:")
                    print(f"  文件: {os.path.basename(result['file_path'])}")
                    print(f"  日期: {result['date']}")
                    print(f"  金额: {result['amount']}")
                    print(f"  内容: {result['content'][:50]}...")
                    
                    # 显示金额识别详情
                    if 'all_amounts' in result and result['all_amounts']:
                        print(f"  所有找到的金额: {result['all_amounts']}")
                    if 'amount_priority' in result:
                        priority_names = {0: "价税合计", 1: "小写金额", 2: "其他"}
                        print(f"  金额类型: {priority_names.get(result['amount_priority'], '未知')}")
                    print()
                
                # 查找金额最大的发票
                max_invoice = recognizer.get_max_amount_invoice(results)
                if max_invoice:
                    print("=" * 40)
                    print("金额最大的发票:")
                    print(f"  文件: {os.path.basename(max_invoice['file_path'])}")
                    print(f"  日期: {max_invoice['date']}")
                    print(f"  金额: {max_invoice['amount']}")
                    print(f"  内容: {max_invoice['content'][:50]}...")
                    print("=" * 40)
                
                break  # 找到有效目录后退出循环
                
            except Exception as e:
                print(f"处理目录时出错: {str(e)}")
                import traceback
                traceback.print_exc()
        else:
            print(f"目录不存在: {test_dir}")
    
    print("\n测试完成！")

def test_single_invoice(file_path):
    """
    测试单张发票识别
    :param file_path: 发票文件路径
    """
    if not os.path.exists(file_path):
        print(f"文件不存在: {file_path}")
        return
    
    print(f"\n测试单张发票: {file_path}")
    print("-" * 40)
    
    recognizer = InvoiceRecognizer()
    result = recognizer.recognize_invoice(file_path)
    
    print(f"识别结果:")
    print(f"  日期: {result['date']}")
    print(f"  金额: {result['amount']}")
    print(f"  内容: {result['content']}")
    
    if 'all_amounts' in result and result['all_amounts']:
        print(f"  所有找到的金额: {result['all_amounts']}")
    if 'amount_priority' in result:
        priority_names = {0: "价税合计", 1: "小写金额", 2: "其他"}
        print(f"  金额类型: {priority_names.get(result['amount_priority'], '未知')}")

if __name__ == '__main__':
    # 测试批量识别
    test_transport_invoice_recognition()
    
    # 如果有命令行参数，测试单张发票
    if len(sys.argv) > 1:
        test_single_invoice(sys.argv[1])
