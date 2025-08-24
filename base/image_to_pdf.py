import os
import sys
import logging
from PIL import Image
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def images_to_pdf(input_folder, output_file):
    """
    将文件夹中的图片转换为PDF文件，每页包含两张图片（垂直排列）
    :param input_folder: 包含图片的文件夹路径
    :param output_file: 输出的PDF文件路径
    """
    # 检查输入文件夹是否存在
    if not os.path.exists(input_folder):
        raise FileNotFoundError(f"文件夹 '{input_folder}' 不存在。")

    # 获取输入文件夹中的所有图片文件
    image_extensions = ['.jpg', '.jpeg', '.png', '.bmp', '.gif']
    image_files = [f for f in os.listdir(input_folder) 
                   if os.path.splitext(f.lower())[1] in image_extensions]
    image_files.sort()  # 按文件名排序

    if not image_files:
        raise ValueError(f"在文件夹 '{input_folder}' 中未找到图片文件。")

    # 创建PDF画布
    c = canvas.Canvas(output_file, pagesize=A4)
    width, height = A4

    # 每页放置两张图片（垂直排列）
    for i in range(0, len(image_files), 2):
        # 创建新页面
        c.showPage()

        # 计算图片位置和大小
        # 页面边距
        margin = 1 * cm
        # 图片之间的垂直间距
        v_spacing = 0.5 * cm
        # 可用宽度
        available_width = width - 2 * margin
        # 可用高度（减去边距和间距后除以2）
        available_height = (height - 2 * margin - v_spacing) / 2
        # 每张图片的宽度（几乎占满页面宽度）
        image_width = available_width
        # 图片高度（保持比例，这里设置为可用高度）
        image_height = available_height

        # 第一张图片位置 (上方)
        x1 = margin
        y1 = height - margin - image_height

        # 第二张图片位置 (下方)
        x2 = margin
        y2 = height - margin - image_height - v_spacing - image_height

        # 处理第一张图片
        if i < len(image_files):
            image_path = os.path.join(input_folder, image_files[i])
            try:
                img = Image.open(image_path)
                # 计算实际宽度和高度（保持原图比例）
                img_width, img_height = img.size
                aspect_ratio = img_width / img_height
                
                # 根据可用空间调整图片大小
                if aspect_ratio > image_width / image_height:
                    # 图片太宽，按宽度缩放
                    actual_width = image_width
                    actual_height = image_width / aspect_ratio
                else:
                    # 图片太高，按高度缩放
                    actual_height = image_height
                    actual_width = image_height * aspect_ratio
                
                # 计算水平居中位置
                actual_x1 = margin + (available_width - actual_width) / 2
                # 绘制图片
                c.drawImage(image_path, actual_x1, y1, width=actual_width, height=actual_height)
            except Exception as e:
                logger.warning(f"无法打开图片 '{image_files[i]}': {str(e)}")

        # 处理第二张图片
        if i + 1 < len(image_files):
            image_path = os.path.join(input_folder, image_files[i + 1])
            try:
                img = Image.open(image_path)
                # 计算实际宽度和高度（保持原图比例）
                img_width, img_height = img.size
                aspect_ratio = img_width / img_height
                
                # 根据可用空间调整图片大小
                if aspect_ratio > image_width / image_height:
                    # 图片太宽，按宽度缩放
                    actual_width = image_width
                    actual_height = image_width / aspect_ratio
                else:
                    # 图片太高，按高度缩放
                    actual_height = image_height
                    actual_width = image_height * aspect_ratio
                
                # 计算水平居中位置
                actual_x2 = margin + (available_width - actual_width) / 2
                # 绘制图片
                c.drawImage(image_path, actual_x2, y2, width=actual_width, height=actual_height)
            except Exception as e:
                logger.warning(f"无法打开图片 '{image_files[i + 1]}': {str(e)}")

    # 保存PDF文件
    c.save()
    logger.info(f"图片已合成为PDF文件 '{output_file}'")


if __name__ == '__main__':
    if len(sys.argv) != 3:
        print("用法: python image_to_pdf.py <包含图片的文件夹路径> <输出PDF文件名>")
        print("示例: python image_to_pdf.py C:\\Users\\username\\Desktop\\images output.pdf")
    else:
        input_folder = sys.argv[1]
        output_file = sys.argv[2]
        images_to_pdf(input_folder, output_file)