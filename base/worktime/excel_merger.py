import pandas as pd
import os
import argparse


def get_excel_files_from_dir(dir_path):
    """
    获取目录下所有Excel文件
    
    参数:
        dir_path: 目录路径
    
    返回:
        目录下所有Excel文件的路径列表
    """
    if not os.path.isdir(dir_path):
        raise ValueError(f"不是有效的目录: {dir_path}")
        
    excel_files = []
    for root, _, files in os.walk(dir_path):
        for file in files:
            if file.endswith(('.xlsx', '.xls')):
                excel_files.append(os.path.join(root, file))
    
    return excel_files


def read_and_merge_files(file_paths, sheet_name='下周工作计划', use_default_sheet=False):
    """
    读取多个Excel文件并合并成一个DataFrame
    
    参数:
        file_paths: 包含多个Excel文件路径的列表
        sheet_name: 要读取的工作表名称
        use_default_sheet: 当指定的工作表不存在时，是否使用第一个工作表
    
    返回:
        合并后的DataFrame
    """
    all_dfs = []
    for file_path in file_paths:
        if os.path.exists(file_path):
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                # 添加项目名称列，值为文件名
                # 获取不带后缀的文件名
                file_name = os.path.splitext(os.path.basename(file_path))[0]
                df['项目名称'] = file_name
                all_dfs.append(df)
            except ValueError as e:
                if 'not found' in str(e) and use_default_sheet:
                    print(f"警告: 在文件 {file_path} 中找不到工作表 '{sheet_name}'，使用第一个工作表")
                    df = pd.read_excel(file_path)
                    # 添加项目名称列，值为不带后缀的文件名
                    file_name = os.path.splitext(os.path.basename(file_path))[0]
                    df['项目名称'] = file_name
                    all_dfs.append(df)
                else:
                    raise ValueError(f"读取文件 {file_path} 失败: {str(e)}")
        else:
            print(f"警告: 文件不存在 - {file_path}")
    
    if not all_dfs:
        error_msg = "没有找到有效的Excel文件。请使用以下方式之一提供Excel文件：\n"
        error_msg += "1. 使用 --files 参数指定一个或多个Excel文件路径\n"
        error_msg += "2. 使用 --dir 参数指定包含Excel文件的目录\n"
        error_msg += "3. 将Excel文件放在当前工作目录下"
        raise ValueError(error_msg)
    
    # 合并所有DataFrame
    merged_df = pd.concat(all_dfs, ignore_index=True)
    
    # 处理指派人名称列，将多个人名拆分为多行
    assignee_col = '指派人名称'
    if assignee_col in merged_df.columns:
        # 定义可能的分隔符
        separators = [',', '，', '、', ';', '；']
        
        # 创建一个临时列表存储拆分后的行
        new_rows = []
        
        for idx, row in merged_df.iterrows():
            assignees = row[assignee_col]
            # 检查是否为字符串类型
            if isinstance(assignees, str):
                # 尝试用不同分隔符拆分
                split_assignees = [assignees]  # 默认不拆分
                for sep in separators:
                    if sep in assignees:
                        split_assignees = [a.strip() for a in assignees.split(sep) if a.strip()]
                        break
                
                # 如果拆分成多个人名，创建新行
                if len(split_assignees) > 1:
                    for name in split_assignees:
                        new_row = row.copy()
                        new_row[assignee_col] = name
                        new_rows.append(new_row)
                else:
                    new_rows.append(row)
            else:
                new_rows.append(row)
        
        # 创建新的DataFrame
        merged_df = pd.DataFrame(new_rows, columns=merged_df.columns)
        print(f"已处理'{assignee_col}'列，拆分多个人名为多行")
    else:
        print(f"警告: 未找到'{assignee_col}'列，跳过拆分处理")

    # 统一时间列格式
    date_columns = ['计划开始日期', '计划结束日期']
    for col in date_columns:
        if col in merged_df.columns:
            # 转换为日期时间类型
            merged_df[col] = pd.to_datetime(merged_df[col], errors='coerce')
            # 格式化为'YYYY-MM-DD'
            merged_df[col] = merged_df[col].dt.strftime('%Y-%m-%d')

    # 处理计划工时数列，如果为空则补充为8小时
    # 尝试精确匹配列名
    target_col = None
    for col in merged_df.columns:
        # 去除空格和特殊字符后比较
        cleaned_col = col.strip().replace(' ', '').replace('（', '(').replace('）', ')')
        if cleaned_col == '计划工时数(h)':
            target_col = col
            break

    if target_col:
        # 先确保日期列为datetime类型
        start_date_col = '计划开始日期'
        end_date_col = '计划结束日期'
        has_dates = start_date_col in merged_df.columns and end_date_col in merged_df.columns
        
        if has_dates:
            # 转换日期列为datetime类型
            merged_df[start_date_col] = pd.to_datetime(merged_df[start_date_col], errors='coerce')
            merged_df[end_date_col] = pd.to_datetime(merged_df[end_date_col], errors='coerce')
            
            # 计算天数差 + 1
            days_diff = (merged_df[end_date_col] - merged_df[start_date_col]).dt.days + 1
            
            # 计算工时数：天数 * 8小时
            calculated_hours = days_diff * 8
            
            # 仅填充空值
            merged_df[target_col] = merged_df[target_col].fillna(calculated_hours)
            print(f"已根据日期计算并填充空工时: '{target_col}'")
        else:
            # 没有日期列，仅填充空值为8
            merged_df[target_col] = merged_df[target_col].fillna(8)
            print(f"缺少日期列，已填充空工时为8: '{target_col}'")
        
        # 确保该列为数值类型
        merged_df[target_col] = pd.to_numeric(merged_df[target_col], errors='coerce')
    else:
        # 打印所有可用列名，帮助调试
        print(f"警告: 未找到'计划工时数（h）'列。可用列名: {', '.join(merged_df.columns)}")

    return merged_df


def parse_arguments():
    """
    解析命令行参数
    
    返回:
        解析后的参数
    """
    parser = argparse.ArgumentParser(description='合并多个Excel工时计划文件并可视化')
    parser.add_argument('--files', '-f', nargs='+', help='Excel文件路径列表')
    parser.add_argument('--dir', '-d', help='包含Excel文件的目录路径')
    parser.add_argument('--sheet', '-s', default='下周工作计划', help='工作表名称')
    parser.add_argument('--use-default-sheet', '-u', action='store_true', help='当指定的工作表不存在时，使用第一个工作表')
    return parser.parse_args()


def get_input_files(args):
    """
    根据命令行参数确定输入文件路径
    
    参数:
        args: 解析后的命令行参数
    
    返回:
        文件路径列表
    """
    if args.files and args.dir:
        raise ValueError("不能同时指定--files和--dir参数")
    elif args.files:
        return args.files
    elif args.dir:
        file_paths = get_excel_files_from_dir(args.dir)
        if not file_paths:
            raise ValueError(f"在目录 {args.dir} 中没有找到Excel文件")
        return file_paths
    else:
        # 默认使用当前工作目录
        current_dir = os.getcwd()
        # 尝试获取当前目录下的Excel文件
        file_paths = get_excel_files_from_dir(current_dir)
        if file_paths:
            return file_paths
        else:
            # 如果当前目录没有Excel文件，返回示例文件路径
            return [
                os.path.join(current_dir, 'example1.xlsx'),
                os.path.join(current_dir, 'example2.xlsx'),
            ]
            # 在这里添加更多默认文件路径


if __name__ == '__main__':
    # 仅当作为脚本直接运行时执行
    args = parse_arguments()
    file_paths = get_input_files(args)
    
    # 确定输出文件路径（保存在源文件夹中）
    if args.dir:
        # 如果指定了目录，保存在该目录下
        output_dir = args.dir
    elif args.files:
        # 如果指定了文件列表，保存在第一个文件所在的目录
        output_dir = os.path.dirname(args.files[0])
    else:
        # 如果使用默认文件，保存在当前工作目录
        output_dir = os.getcwd()
    
    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)
    
    # 确定输出文件路径
    output_file = os.path.join(output_dir, 'merged_excel_files.xlsx')
    
    # 检查文件是否存在，如果存在则删除
    if os.path.exists(output_file):
        try:
            os.remove(output_file)
            print(f"已删除已存在的文件: {output_file}")
        except Exception as e:
            print(f"删除文件时出错: {str(e)}")
    
    # 读取并合并文件
    df = read_and_merge_files(file_paths, sheet_name=args.sheet, use_default_sheet=args.use_default_sheet)
    
    # 保存合并后的DataFrame到新文件
    df.to_excel(output_file, sheet_name='下周工作计划', index=False)
    print(f"合并完成，结果已保存到 {output_file}")

    # 将数据存储到MySQL数据库
    try:
        from db.db_manager import DBManager
        # 创建数据库管理器实例
        db_manager = DBManager()
        # 连接数据库
        if db_manager.connect():
            # 创建表
            db_manager.create_table()
            # 插入数据
            db_manager.insert_data(df)
            print("数据已成功存储到MySQL数据库")
            # 断开连接
            db_manager.disconnect()
        else:
            print("无法连接到数据库，数据未存储")
    except ImportError:
        print("警告: 无法导入DBManager类，数据未存储到数据库")
    except Exception as e:
        print(f"存储数据到数据库时出错: {str(e)}")
