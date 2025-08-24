#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
直接测试PDF转图像功能，不通过MCP客户端
"""
import os
from base.batch_pdf_to_image import pdf_to_images

def test_convert_pdf_to_image():
    # 确保测试目录存在
    test_output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_output")
    os.makedirs(test_output_dir, exist_ok=True)
    
    # 查找一个PDF文件进行测试
    test_pdf_path = None
    project_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 递归搜索项目目录中的所有PDF文件
    for root, dirs, files in os.walk(project_dir):
        for file in files:
            if file.lower().endswith('.pdf'):
                test_pdf_path = os.path.join(root, file)
                break
        if test_pdf_path:
            break
    
    if not test_pdf_path:
        print("未找到测试用的PDF文件，请确保项目目录中有PDF文件")
        return
    
    print(f"使用测试文件: {test_pdf_path}")
    print(f"输出目录: {test_output_dir}")
    
    # 直接调用转换函数
    success = pdf_to_images(test_pdf_path, test_output_dir, resolution=300, image_format='png')
    
    # 检查结果
    if success:
        # 获取生成的图像文件列表
        pdf_filename = os.path.basename(test_pdf_path)
        pdf_name = os.path.splitext(pdf_filename)[0]
        
        # 获取输出目录中的相关图像文件
        image_files = [f for f in os.listdir(test_output_dir) if f.startswith(pdf_name) and f.lower().endswith('.png')]
        
        print(f"PDF文件 '{pdf_filename}' 已成功转换为 {len(image_files)} 个图像文件，保存至 '{test_output_dir}'")
        print(f"生成的图像文件列表: {', '.join(image_files)}")
    else:
        print(f"转换PDF文件 '{pdf_filename}' 时出错")

if __name__ == "__main__":
    test_convert_pdf_to_image()