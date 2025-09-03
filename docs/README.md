# AdminAssistent - 管理助手工具集

## 项目概述

AdminAssistent 是一个基于 MCP (Model Context Protocol) 的管理助手工具集，提供发票识别、邮件下载、文档处理和工时管理等功能。

## 主要功能

### 1. 发票识别工具
- **单个发票识别** (`ocr_invoice`): 识别单个发票文件，提取日期、金额、内容等信息
- **批量发票识别** (`ocr_invoices`): 批量识别目录下的所有发票，并自动重命名

### 2. 邮件下载工具
- **邮件附件下载** (`download_emails`): 下载指定日期范围内的邮件附件
- 支持多种邮箱配置（QQ邮箱、飞书邮箱、企业邮箱等）
- 支持按日期范围、文件类型筛选

### 3. 文档处理工具
- **PDF转图片** (`convert_pdf_to_image`, `batch_convert_pdfs`): 将PDF文件转换为图片
- **图片合并PDF** (`images_to_pdf`): 将多张图片合并为PDF文件
- 支持高分辨率输出（300-600 DPI）

### 4. Excel处理工具
- **Excel文件合并** (`merge_excel_files`): 合并多个项目任务格式的Excel文件
- **工时处理** (`process_worktime`): 处理工时数据并生成可视化展示图片

## 工具详细说明

### Excel合并工具
专门用于合并项目任务格式的Excel文件，支持：
- 自动添加项目名称列
- 指派人名称拆分
- 日期格式统一
- 工时自动计算

详细说明请参考：[Excel合并MCP工具使用说明.md](./Excel合并MCP工具使用说明.md)

### 工时处理工具
用于处理项目任务数据并生成工时展示图片，支持：
- 智能人员分组
- 可视化展示
- 项目分配情况显示
- 工作状态区分

详细说明请参考：[工时处理MCP工具使用说明.md](./工时处理MCP工具使用说明.md)

## 项目结构

```
AdminAssistent/
├── base/                          # 核心功能模块
│   ├── invoice_ocr.py            # 发票识别模块
│   ├── email_attachment_downloader.py  # 邮件下载模块
│   ├── image_to_pdf.py           # 图片转PDF模块
│   ├── batch_pdf_to_image.py     # PDF转图片模块
│   └── worktime/                 # 工时处理模块
│       ├── excel_merger.py       # Excel合并功能
│       └── worktime_processor.py # 工时处理功能
├── base/conf/                    # 配置文件
│   ├── qq-email.json            # QQ邮箱配置
│   ├── feishu-email.json        # 飞书邮箱配置
│   └── cosmo-email.json         # 企业邮箱配置
├── docs/                         # 说明文档
│   ├── README.md                # 项目总览
│   ├── Excel合并MCP工具使用说明.md
│   └── 工时处理MCP工具使用说明.md
├── attachments/                  # 附件存储目录
└── mcp_administrative_service.py # MCP服务主文件
```

## 使用方式

### 1. 启动MCP服务
```bash
python mcp_administrative_service.py
```

### 2. 调用工具
通过MCP客户端调用相应的工具，例如：

```json
{
  "name": "merge_excel_files",
  "arguments": {
    "input_dir": "/path/to/excel/files",
    "output_file": "/path/to/merged.xlsx"
  }
}
```

## 配置说明

### 邮箱配置
在 `base/conf/` 目录下配置各种邮箱的IMAP设置：
- `imap_server`: IMAP服务器地址
- `username`: 邮箱用户名
- `password`: 邮箱密码或授权码
- `save_dir`: 附件保存目录
- `file_extensions`: 要下载的文件类型

## 技术特点

- **模块化设计**: 每个功能独立模块，便于维护和扩展
- **MCP协议**: 基于标准MCP协议，支持多种客户端
- **错误处理**: 完善的错误处理和日志记录
- **中文支持**: 全面支持中文显示和处理
- **高分辨率输出**: 支持高分辨率图片和PDF输出

## 依赖要求

- Python 3.7+
- pandas
- matplotlib
- openpyxl
- Pillow
- reportlab
- easyocr
- imaplib2

## 更新日志

### v1.0.0
- 初始版本发布
- 支持发票识别、邮件下载、文档处理功能
- 支持Excel合并和工时处理功能
- 完整的MCP工具集成

## 许可证

本项目采用 MIT 许可证。

## 贡献

欢迎提交 Issue 和 Pull Request 来改进这个项目。

## 联系方式

如有问题或建议，请通过 Issue 联系。
