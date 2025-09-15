# AdminAssistent - 管理助手工具集

## 📋 项目概述

AdminAssistent 是一个基于 MCP (Model Context Protocol) 的智能管理助手工具集，专为企业和个人用户提供发票识别、邮件处理、文档管理、工时统计和考勤分析等全方位管理功能。

## ✨ 核心特性

- 🧾 **智能发票识别** - 支持OCR识别和批量处理
- 📧 **邮件自动化** - 智能下载和阅读邮件附件
- 📊 **工时管理** - 可视化工时统计和报表生成
- 👥 **考勤分析** - 员工考勤数据处理和报告
- 📄 **文档处理** - PDF转图片、图片合并等
- 🔧 **MCP集成** - 标准MCP协议，支持多种客户端

## 🛠️ 功能模块

### 📊 文档处理类

- **Excel合并工具** - 合并多个项目任务格式的Excel文件
- **工时处理工具** - 生成工时展示图片和可视化报表
- **工时报表生成** - 合并Excel文件并生成专业工时报表

### 🧾 发票识别类

- **单个发票识别** - 识别单个发票文件，提取关键信息
- **批量发票识别** - 批量识别目录下的所有发票，自动重命名

### 📧 邮件处理类

- **邮件附件下载** - 下载指定日期范围内的邮件附件，支持按日期、天数、文件类型筛选
- **邮件内容阅读** - 读取邮件内容，支持多种筛选条件和日期范围

### 🖼️ 图片处理类

- **批量PDF转图片** - 批量转换目录下的所有PDF文件
- **图片合并PDF** - 将多张图片合并为PDF文件

### 👥 考勤管理类

- **考勤数据分析** - 分析员工考勤数据
- **考勤报告生成** - 生成考勤统计报告
- **员工信息管理** - 添加和管理员工信息
- **离职数据处理** - 处理离职员工相关数据

## 📁 项目结构

```
AdminAssistent/
├── base/                          # 核心功能模块
│   ├── conf/                      # 配置文件
│   │   ├── qq-email.json         # QQ邮箱配置
│   │   ├── feishu-email.json     # 飞书邮箱配置
│   │   └── cosmo-email.json      # 企业邮箱配置
│   ├── worktime/                  # 工时处理模块
│   │   ├── attendance/           # 考勤管理
│   │   │   ├── add_departure_info.py
│   │   │   ├── attendance_analyzer.py
│   │   │   ├── attendance_processor.py
│   │   │   ├── attendance_report_generator.py
│   │   │   └── attendance_visualizer.py
│   │   ├── excel_merger.py       # Excel合并功能
│   │   ├── process.py            # 工时处理主程序
│   │   └── worktime_processor.py # 工时处理器
│   ├── batch_pdf_to_image.py     # PDF转图片工具
│   ├── email_attachment_downloader.py # 邮件下载工具
│   ├── image_to_pdf.py           # 图片合并PDF工具
│   └── invoice_ocr.py            # 发票识别工具
├── docs/                          # 文档目录
│   ├── INDEX.md                  # 文档索引
│   ├── README.md                 # 项目说明
│   ├── Excel合并MCP工具使用说明.md
│   ├── 工时处理MCP工具使用说明.md
│   ├── 工时安排报表生成工具使用说明.md
│   └── 邮件阅读工具使用说明.md
├── temp/                          # 临时文件目录
├── mcp_administrative_service.py  # MCP服务主文件
└── email_attachment_downloader.log # 邮件下载日志
```

## 🚀 快速开始

### 1. 环境要求

- Python 3.7+
- 依赖包：pandas, matplotlib, openpyxl, Pillow, reportlab, easyocr, imaplib2

### 2. 安装依赖

```bash
pip install pandas matplotlib openpyxl Pillow reportlab easyocr imaplib2
```

### 3. 配置邮箱

复制示例配置文件并编辑 `base/conf/` 目录下的邮箱配置文件：

```bash
# 复制示例配置文件
cp base/conf/qq-email.json.example base/conf/qq-email.json
cp base/conf/feishu-email.json.example base/conf/feishu-email.json
cp base/conf/cosmo-email.json.example base/conf/cosmo-email.json
```

然后编辑相应的配置文件：

- `qq-email.json` - QQ邮箱配置
- `feishu-email.json` - 飞书邮箱配置
- `cosmo-email.json` - 企业邮箱配置

**注意**: 这些配置文件包含敏感信息，不会被提交到Git仓库中。

### 4. 启动服务

#### 方式一：直接启动

```bash
python mcp_administrative_service.py
```

#### 方式二：在Cursor中配置MCP服务器

1. **打开Cursor设置**

   - 按 `Cmd+,` (Mac) 或 `Ctrl+,` (Windows/Linux) 打开设置
   - 搜索 "MCP" 或 "Model Context Protocol"
2. **添加MCP服务器配置**
   在Cursor的MCP配置中添加以下配置(具体目录根据实际情况而定)：

   ```json
   {
     "mcpServers": {
       "admin-assistent": {
         "command": "python",
         "args": ["/Volumes/SD_1.5T/github/AdminAssistent/mcp_administrative_service.py"],
         "env": {
           "PYTHONPATH": "/Volumes/SD_1.5T/github/AdminAssistent"
         }
       }
     }
   }
   ```
3. **配置说明**

   - `command`: 使用Python解释器
   - `args`: 指向MCP服务文件的完整路径
   - `env`: 设置Python路径环境变量
   - 请根据实际安装路径调整路径
4. **重启Cursor**

   - 保存配置后重启Cursor
   - MCP服务器将自动启动并连接
5. **验证连接**

   - 在Cursor中打开AI聊天面板
   - 查看是否显示可用的MCP工具
   - 工具列表应包含：发票识别、邮件下载、文档处理等

## 📖 使用指南

### 在Cursor中使用MCP工具

配置完成后，您可以在Cursor的AI聊天中直接使用所有MCP工具：

#### 基本使用方式

1. **打开AI聊天面板** - 在Cursor中按 `Cmd+L` (Mac) 或 `Ctrl+L` (Windows/Linux)
2. **直接描述需求** - 例如："下载今天的邮件附件"、"合并Excel文件"等
3. **AI自动调用工具** - Cursor会自动识别需求并调用相应的MCP工具
4. **查看执行结果** - 工具执行结果会直接显示在聊天中

#### 常用命令示例

- "下载QQ邮箱最近3天的邮件附件到指定目录"
- "合并项目目录下的所有Excel文件"
- "生成工时报表并保存为图片"
- "识别发票目录下的所有发票文件"
- "将PDF文件转换为图片"

### 典型工作流程

#### 1. 项目管理流程

```
数据准备 → Excel合并 → 工时处理 → 报表生成 → 结果查看
```

#### 2. 邮件处理流程

```
配置邮箱 → 下载附件 → 文档处理 → 发票识别 → 结果整理
```

#### 3. 考勤管理流程

```
数据导入 → 数据处理 → 信息补充 → 报告生成 → 结果分析
```

### 工具调用示例

#### Excel文件合并

```json
{
  "name": "merge_excel_files",
  "arguments": {
    "input_dir": "/path/to/excel/files",
    "output_file": "/path/to/merged.xlsx"
  }
}
```

#### 工时报表生成

```json
{
  "name": "generate_worktime_report",
  "arguments": {
    "input_dir": "/path/to/project/files",
    "output_image_path": "/path/to/report.png",
    "title": "项目工时报表"
  }
}
```

#### 邮件附件下载

```json
{
  "name": "download_emails",
  "arguments": {
    "imap_server": "imap.qq.com",
    "username": "your@email.com",
    "password": "your_password",
    "save_dir": "/path/to/save",
    "days_ago": 7,
    "date": "2024-12-19",
    "start_date": "2024-12-01",
    "end_date": "2024-12-31",
    "file_extensions": ["pdf", "jpg", "png"]
  }
}
```

## ⚙️ 配置说明

### 邮箱配置格式

```json
{
  "email": {
    "imap_server": "imap.qq.com",
    "username": "your@email.com",
    "password": "your_password",
    "port": 993,
    "folder": "INBOX",
    "save_dir": "/path/to/save",
    "file_extensions": ["pdf", "jpg", "jpeg", "png", "xml", "eml"]
  }
}
```

### 文件格式支持

- **Excel文件**: .xlsx, .xls
- **图片文件**: .png, .jpg, .jpeg
- **PDF文件**: 标准PDF格式
- **发票文件**: .pdf, .jpg, .png, .ofd

## 🔧 技术特点

- **模块化设计** - 每个功能独立模块，便于维护和扩展
- **MCP协议** - 基于标准MCP协议，支持多种客户端
- **智能识别** - 支持OCR发票识别和智能重命名
- **高分辨率输出** - 支持300-600 DPI高质量输出
- **中文支持** - 全面支持中文显示和处理
- **错误处理** - 完善的错误处理和日志记录
- **批量处理** - 支持批量文件处理，提高效率

## 📊 功能演示

### 发票识别效果

- 自动识别发票日期、金额、内容
- 智能重命名：`日期+金额+内容.pdf`
- 支持多种发票格式

### 工时报表效果

- 可视化工时分配图表
- 项目人员工作状态展示
- 专业报表格式输出

### 邮件处理效果

- 智能筛选和下载附件，支持多种日期范围选择
- 支持按指定日期、天数、日期区间下载
- 支持多邮箱同时处理
- 自动文件分类整理和重命名

## 🆘 常见问题

### Q: Excel合并后数据格式不正确？

A: 请确保输入文件包含标准的项目任务字段（任务名称、指派人名称、计划日期等）

### Q: 工时展示图中文显示异常？

A: 请确保系统安装了中文字体支持

### Q: 邮件下载失败？

A: 请检查邮箱配置和网络连接，确保IMAP设置正确

### Q: 发票识别准确率低？

A: 请确保发票图片清晰，避免模糊或倾斜

### Q: 在Cursor中无法使用MCP工具？

A: 请检查以下几点：

- 确认MCP服务器配置正确，路径无误
- 检查Python环境和依赖包是否安装完整
- 重启Cursor后再次尝试
- 查看Cursor的开发者工具中是否有错误信息

## 📈 更新日志

### v1.3.0 (2024-12-19)

- 新增考勤管理模块
- 完善工时报表生成功能
- 优化文档索引和说明
- 增强错误处理机制

### v1.2.0

- 新增工时报表生成工具
- 完善邮件阅读功能
- 优化图片处理工具

### v1.1.0

- 新增邮件处理和发票识别功能
- 支持多种邮箱配置
- 增强文档处理能力

### v1.0.0

- 初始版本发布
- 基础Excel合并和工时处理功能
- 完整的MCP工具集成

## 📄 许可证

本项目采用 MIT 许可证。

## 🤝 贡献

欢迎提交 Issue 和 Pull Request 来改进这个项目。

## 📞 技术支持

如有问题或建议，请通过以下方式联系：

- 提交 Issue
- 查看详细文档：[INDEX.md](./INDEX.md)
- 参考各工具使用说明

---

**AdminAssistent** - 让管理更智能，让工作更高效！
