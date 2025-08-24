# Tesseract OCR 安装说明

## Windows 安装步骤

1. 下载 Tesseract OCR：
   - 访问：https://github.com/UB-Mannheim/tesseract/wiki
   - 下载适合您系统的版本（推荐64位版本）

2. 安装 Tesseract：
   - 运行下载的安装程序
   - 建议安装到默认路径：`C:\Program Files\Tesseract-OCR`
   - 确保勾选"Additional language data"以支持中文

3. 配置环境变量：
   - 将 `C:\Program Files\Tesseract-OCR` 添加到系统PATH环境变量
   - 或者在代码中指定Tesseract路径

4. 验证安装：
   - 打开命令提示符
   - 运行：`tesseract --version`
   - 如果显示版本信息，说明安装成功

## 在代码中指定Tesseract路径

如果Tesseract没有添加到PATH，可以在代码中指定路径：

```python
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
```

## 替代方案

如果不想安装Tesseract，可以：
1. 使用在线OCR服务
2. 使用其他OCR库如EasyOCR
3. 暂时使用手动输入验证码

