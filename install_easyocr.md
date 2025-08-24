# EasyOCR 安装说明

## 自动安装

EasyOCR会自动安装所有必要的依赖，包括：
- PyTorch
- OpenCV
- 其他必要的库

## 安装命令

```bash
pip install easyocr
```

## 验证安装

```python
import easyocr
reader = easyocr.Reader(['en'])
print("EasyOCR安装成功！")
```

## 优势

相比Tesseract，EasyOCR具有以下优势：
1. 安装更简单，无需额外配置
2. 识别准确率更高
3. 支持多种语言
4. 更好的文本检测能力
5. 内置图像预处理

## 使用说明

在ERP登录脚本中，EasyOCR会自动：
1. 识别验证码图片中的文本
2. 清理和验证识别结果
3. 如果识别失败，会提示手动输入

## 性能优化

- 使用CPU模式（gpu=False）以提高兼容性
- 只初始化一次Reader对象以提高性能
- 支持批量处理多个验证码

