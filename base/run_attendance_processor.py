#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
考勤处理器运行脚本
使用示例
"""

from attendance_processor import AttendanceProcessor
import os

def main():
    """运行考勤处理器"""
    # 设置工作目录 - 请根据实际情况修改
    work_dir = "/Users/heguangzhong/Desktop/Temp/9月考勤"
    
    print("考勤表处理程序")
    print("=" * 50)
    print(f"工作目录: {work_dir}")
    print()
    
    # 检查工作目录是否存在
    if not os.path.exists(work_dir):
        print(f"错误: 工作目录不存在: {work_dir}")
        print("请确保目录存在并包含以下文件:")
        print("- 上月考勤表.xlsx")
        print("- 当月请假记录.xlsx (可选)")
        print("- 当月入职人员信息表.xlsx (可选)")
        print("- 当月离职人员信息表.xlsx (可选)")
        return
    
    # 检查必需文件
    required_files = {
        "上月考勤表.xlsx": os.path.join(work_dir, "上月考勤表.xlsx")
    }
    
    missing_files = []
    for file_name, file_path in required_files.items():
        if not os.path.exists(file_path):
            missing_files.append(file_name)
    
    if missing_files:
        print("错误: 缺少必需文件:")
        for file_name in missing_files:
            print(f"- {file_name}")
        print()
        print("请确保工作目录中包含上述文件")
        return
    
    # 显示可选文件状态
    optional_files = {
        "当月请假记录.xlsx": os.path.join(work_dir, "当月请假记录.xlsx"),
        "当月入职人员信息表.xlsx": os.path.join(work_dir, "当月入职人员信息表.xlsx"),
        "当月离职人员信息表.xlsx": os.path.join(work_dir, "当月离职人员信息表.xlsx")
    }
    
    print("文件检查:")
    for file_name, file_path in optional_files.items():
        status = "✓ 存在" if os.path.exists(file_path) else "⚠ 不存在"
        print(f"- {file_name}: {status}")
    print()
    
    # 自动继续处理（无需用户确认）
    print("自动开始处理...")
    
    # 创建处理器并执行处理
    try:
        processor = AttendanceProcessor(work_dir)
        processor.process_attendance()
        print("\n🎉 考勤表处理完成！")
    except Exception as e:
        print(f"\n❌ 处理失败: {str(e)}")
        return

if __name__ == "__main__":
    main()
