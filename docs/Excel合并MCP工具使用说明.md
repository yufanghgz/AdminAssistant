# Excel 合并 MCP 工具使用说明

## 功能概述

新的 `merge_excel_files` MCP 工具专门用于合并特定项目任务格式的 Excel 文件，特别适用于工时计划、项目计划等场景。该工具针对包含项目任务信息的 Excel 文件进行了优化处理。

## 工具参数

### 必需参数
- `output_file`: 输出合并后的 Excel 文件路径

### 可选参数
- `input_files`: 要合并的 Excel 文件路径列表（数组）
- `input_dir`: 包含 Excel 文件的目录路径（与 input_files 互斥）
- `sheet_name`: 要读取的工作表名称，默认为 "下周工作计划"
- `use_default_sheet`: 当指定的工作表不存在时，是否使用第一个工作表，默认为 false

## 使用方式

### 方式1：指定文件列表
```json
{
  "input_files": [
    "/path/to/file1.xlsx",
    "/path/to/file2.xlsx",
    "/path/to/file3.xlsx"
  ],
  "output_file": "/path/to/merged_output.xlsx"
}
```

### 方式2：指定目录
```json
{
  "input_dir": "/path/to/excel/files/directory",
  "output_file": "/path/to/merged_output.xlsx"
}
```

### 方式3：自定义工作表名称
```json
{
  "input_dir": "/path/to/excel/files/directory",
  "output_file": "/path/to/merged_output.xlsx",
  "sheet_name": "我的工作表",
  "use_default_sheet": true
}
```

## 功能特性

1. **自动添加项目名称列**: 根据文件名自动添加 "项目名称" 列
2. **指派人名称拆分**: 自动将多个指派人（用逗号、分号等分隔）拆分为多行
3. **日期格式统一**: 自动将日期列格式化为 YYYY-MM-DD 格式
4. **工时自动计算**: 根据开始和结束日期自动计算工时数
5. **错误处理**: 当工作表不存在时，可选择使用第一个工作表
6. **项目任务格式优化**: 专门针对包含任务名称、指派人、计划日期、工时等字段的项目任务 Excel 文件

## 支持的 Excel 格式
- .xlsx
- .xls

## 适用的文件格式

该工具专门针对项目任务格式的 Excel 文件，典型的文件结构应包含以下字段：

### 标准字段
- **任务名称**: 具体的工作任务描述
- **指派人名称**: 负责执行任务的人员（支持多人，用逗号、分号等分隔）
- **计划开始日期**: 任务的计划开始时间
- **计划结束日期**: 任务的计划结束时间
- **计划工时数(h)**: 预计的工作时间（小时）
- **任务描述**: 任务的详细说明（可选）
- **父级任务名称**: 上级任务名称（可选）

### 典型应用场景
- 项目工时计划表
- 周/月度工作计划
- 任务分配表
- 项目进度跟踪表

## 注意事项

1. 不能同时指定 `input_files` 和 `input_dir` 参数
2. 必须指定 `output_file` 参数
3. 如果输出文件已存在，会被自动覆盖
4. 所有输入文件必须包含相同的工作表名称（除非使用 `use_default_sheet`）
5. **适用文件类型**: 该工具专门针对项目任务格式的 Excel 文件，包含以下典型字段：
   - 任务名称
   - 指派人名称
   - 计划开始日期
   - 计划结束日期
   - 计划工时数
   - 任务描述等

## 示例输出

合并完成后会返回类似以下信息：
```
Excel文件合并完成，结果已保存到: /path/to/merged_output.xlsx
共合并了 3 个文件，12 行数据
```
