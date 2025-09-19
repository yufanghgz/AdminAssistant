# AdminAssistant 安装指南

## 环境要求

- Python 3.7+ (推荐 Python 3.8+)
- 操作系统：Windows, macOS, Linux

## 安装方式

### 方式一：使用 requirements.txt（推荐）

```bash
# 克隆项目
git clone <repository-url>
cd AdminAssistant

# 创建虚拟环境（推荐）
python -m venv venv

# 激活虚拟环境
# Windows:
venv\Scripts\activate
# macOS/Linux:
source venv/bin/activate

# 安装依赖
pip install -r requirements.txt
```

### 方式二：使用最小依赖

```bash
# 仅安装核心依赖
pip install -r requirements-minimal.txt
```

### 方式三：开发环境

```bash
# 安装开发环境依赖（包含测试和代码质量工具）
pip install -r requirements-dev.txt
```

## 手动安装核心依赖

如果遇到问题，可以手动安装核心依赖：

```bash
pip install pandas numpy matplotlib openpyxl Pillow opencv-python pdfplumber reportlab easyocr imaplib2 beautifulsoup4 requests mcp
```

## 验证安装

运行以下命令验证安装是否成功：

```bash
python -c "import pandas, numpy, matplotlib, openpyxl, PIL, cv2, pdfplumber, reportlab, easyocr, imaplib2, bs4, requests, mcp; print('所有依赖安装成功！')"
```

## 常见问题

### 1. easyocr 安装失败
```bash
# 如果 easyocr 安装失败，可能需要先安装系统依赖
# Ubuntu/Debian:
sudo apt-get install libgl1-mesa-glx libglib2.0-0

# macOS:
brew install libgl

# Windows: 通常不需要额外依赖
```

### 2. opencv-python 安装失败
```bash
# 如果 opencv-python 安装失败，可以尝试：
pip install opencv-python-headless
```

### 3. 权限问题
```bash
# 如果遇到权限问题，使用 --user 参数：
pip install --user -r requirements.txt
```

## 配置

安装完成后，需要配置邮箱设置：

1. 复制配置文件模板：
```bash
cp base/conf/qq-email.json.example base/conf/qq-email.json
```

2. 编辑配置文件，填入你的邮箱信息

## 使用

安装完成后，可以运行：

```bash
# 运行 MCP 服务器
python mcp_administrative_service.py

# 或运行具体功能模块
python base/invoice_ocr.py
python base/email_attachment_downloader.py
```

## 卸载

```bash
# 如果使用虚拟环境，直接删除虚拟环境目录
rm -rf venv

# 或者卸载所有包
pip freeze | xargs pip uninstall -y
```
