import pdfplumber
import os
import argparse
from pathlib import Path


def pdf_to_images(pdf_path, output_folder, resolution=600, image_format='png'):
    """
    将单个PDF文件转换为图片
    :param pdf_path: PDF文件路径
    :param output_folder: 图片输出目录
    :param resolution: 图片分辨率 (默认 300 dpi)
    :param image_format: 图片格式 (默认 'png')
    """
    # 确保输出目录存在
    os.makedirs(output_folder, exist_ok=True)
    
    try:
        # 打开PDF文件
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            pdf_filename = os.path.basename(pdf_path)
            pdf_name = os.path.splitext(pdf_filename)[0]
            print(f"\n处理文件: {pdf_filename} (共 {total_pages} 页)")
            
            # 遍历每一页
            for page_num, page in enumerate(pdf.pages, start=1):
                # 转换页面为图片
                img = page.to_image(resolution=resolution)
                
                # 生成图片文件名
                if total_pages > 1:
                    # 多页PDF，文件名包含页码
                    image_filename = f"{pdf_name}_page_{page_num}.{image_format}"
                else:
                    # 单页PDF，直接使用原文件名
                    image_filename = f"{pdf_name}.{image_format}"
                
                image_path = os.path.join(output_folder, image_filename)
                
                # 保存图片
                img.save(image_path)
                print(f"已保存: {image_filename}")
        
        return True
    
    except Exception as e:
        print(f"处理文件 {pdf_filename} 时出错: {str(e)}")
        return False


def batch_convert_pdfs(input_dir, output_dir, resolution=600, image_format='png'):
    """
    批量转换目录下的所有PDF文件
    :param input_dir: 包含PDF文件的目录
    :param output_dir: 图片输出目录
    :param resolution: 图片分辨率
    :param image_format: 图片格式
    """
    # 确保输入目录存在
    if not os.path.exists(input_dir):
        print(f"错误: 找不到输入目录 '{input_dir}'")
        return
    
    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)
    
    # 获取输入目录中的所有PDF文件
    pdf_files = [f for f in os.listdir(input_dir) if f.lower().endswith('.pdf')]
    
    if not pdf_files:
        print(f"在目录 '{input_dir}' 中未找到PDF文件")
        return
    
    print(f"找到 {len(pdf_files)} 个PDF文件，开始转换...")
    
    # 遍历所有PDF文件
    for pdf_file in pdf_files:
        pdf_path = os.path.join(input_dir, pdf_file)
        pdf_to_images(pdf_path, output_dir, resolution, image_format)
    
    print(f"\n批量转换完成！所有图片已保存到: {output_dir}")


def main():
    # 解析命令行参数
    parser = argparse.ArgumentParser(description='批量将PDF文件转换为图片')
    parser.add_argument('-i', '--input', help='包含PDF文件的目录', default='.')
    parser.add_argument('-o', '--output', help='图片输出目录', default='pdf_images')
    parser.add_argument('-r', '--resolution', type=int, help='图片分辨率 (dpi)', default=600)
    parser.add_argument('-f', '--format', help='图片格式 (png/jpg)', default='png')
    args = parser.parse_args()
    
    # 检查图片格式
    if args.format.lower() not in ['png', 'jpg', 'jpeg']:
        print("错误: 不支持的图片格式。请使用 'png'、'jpg' 或 'jpeg'")
        return
    
    # 执行批量转换
    batch_convert_pdfs(args.input, args.output, args.resolution, args.format.lower())


if __name__ == "__main__":
    main()