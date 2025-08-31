import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches  # 新增导入用于创建图例

from datetime import datetime
plt.rcParams['font.sans-serif'] = ['Arial Unicode MS']  # 设置字体为Heiti SC
plt.rcParams['axes.unicode_minus'] = False  # 用于正常显示负号
# 读取Excel文件
merged_file_path = 'D:/PyProject/NounParse/merged_excel_files.xlsx'
df = pd.read_excel(merged_file_path, sheet_name='下周工作计划')

# 读取所有人员名称的Excel文件
staff_file_path = 'D:/Work_Doc/11.项目管理/2025/工时计划/人员列表.xlsx'

# 添加错误处理，以防人员名单文件不存在
try:
    staff_df = pd.read_excel(staff_file_path)
    # 提取第一列'人员名称'作为所有人员的列表
    all_staff = staff_df.iloc[:, 0].unique().tolist()
    print(f"成功读取人员名单，共 {len(all_staff)} 人")
except FileNotFoundError:
    print(f"警告: 人员名单文件不存在 - {staff_file_path}")
    # 如果人员名单文件不存在，使用任务表中的人员名单
    all_staff = df['指派人名称'].unique().tolist()
    print(f"使用任务表中的人员名单，共 {len(all_staff)} 人")
except Exception as e:
    print(f"读取人员名单时出错: {str(e)}")
    # 使用任务表中的人员名单作为备选
    all_staff = df['指派人名称'].unique().tolist()
    print(f"使用任务表中的人员名单，共 {len(all_staff)} 人")

# 日期处理
df['计划开始日期'] = pd.to_datetime(df['计划开始日期']).dt.date
df['计划结束日期'] = pd.to_datetime(df['计划结束日期']).dt.date

# 生成日期范围 - 确保只使用有效日期
valid_dates_start = df['计划开始日期'].dropna()
valid_dates_end = df['计划结束日期'].dropna()

if valid_dates_start.empty or valid_dates_end.empty:
    print("警告: 没有找到有效的日期数据，使用默认日期范围")
    # 使用当前日期前后7天作为默认范围
    today = datetime.now().date()
    start_date = (datetime.now() - pd.Timedelta(days=7)).date()
    end_date = (datetime.now() + pd.Timedelta(days=7)).date()
else:
    start_date = valid_dates_start.min()
    end_date = valid_dates_end.max()

date_range = pd.date_range(start=start_date, end=end_date).date.tolist()

# 按项目分组人员
# 1. 创建项目到人员的映射
project_to_staff = {}
no_project_staff = []

for assignee in all_staff:
    # 获取该人员参与的所有项目
    assignee_tasks = df[df['指派人名称'] == assignee]
    if not assignee_tasks.empty:
        projects = assignee_tasks['项目名称'].unique().tolist()
        # 使用第一个项目作为主要项目进行分组
        main_project = projects[0]
        if main_project not in project_to_staff:
            project_to_staff[main_project] = []
        if assignee not in project_to_staff[main_project]:
            project_to_staff[main_project].append(assignee)
    else:
        # 没有项目的人员
        no_project_staff.append(assignee)

# 2. 按项目名称排序
sorted_projects = sorted(project_to_staff.keys())

# 3. 构建按项目分组的人员列表
assignees = []
for project in sorted_projects:
    assignees.extend(project_to_staff[project])

# 4. 添加没有项目的人员
assignees.extend(no_project_staff)

# 打印分组信息
print("按项目分组的人员列表:")
for project in sorted_projects:
    print(f"项目 '{project}': {', '.join(project_to_staff[project])}")
if no_project_staff:
    print(f"无项目人员: {', '.join(no_project_staff)}")

# 初始化状态表和项目信息表（包含所有人员）
status_data = np.zeros((len(assignees), len(date_range)))
project_data = [[[] for _ in range(len(date_range))] for _ in range(len(assignees))]

# 填充状态数据和项目信息
for idx, assignee in enumerate(assignees):
    # 检查该人员是否有任务
    assignee_tasks = df[df['指派人名称'] == assignee]
    if not assignee_tasks.empty:
        for _, task in assignee_tasks.iterrows():
            project_name = task.get('项目名称', '无项目信息')  # 获取项目名称
            task_dates = pd.date_range(task['计划开始日期'], task['计划结束日期']).date
            for date in task_dates:
                if date in date_range:
                    col = date_range.index(date)
                    status_data[idx, col] = 1
                    if project_name not in project_data[idx][col]:
                        project_data[idx][col].append(project_name)
    # 没有任务的人员会保持状态为0，项目信息为空列表

# 统计人数
num_people = len(assignees)

# 可视化设置 - 增加图表高度以容纳项目信息
fig, ax = plt.subplots(figsize=(20, len(assignees)*0.8))

# 直接创建一个全为绿色的背景
background = np.ones((len(assignees), len(date_range), 3)) * np.array([0.56, 0.93, 0.56])  # #90EE90的RGB值
ax.imshow(background, aspect='auto')

# 在单元格中添加项目信息
for i in range(len(assignees)):
    for j in range(len(date_range)):
        if status_data[i, j] == 1 and project_data[i][j]:
            projects = '\n'.join(project_data[i][j])
            ax.text(j, i, projects, ha='center', va='center', fontsize=9,
                    color='black', weight='bold',
                    bbox=dict(facecolor='none', alpha=0, edgecolor='black', boxstyle='round,pad=0.3'))
        elif status_data[i, j] == 0:
            # 对于未安排的单元格，覆盖为白色
            rect = plt.Rectangle((j-0.5, i-0.5), 1, 1, facecolor='white', edgecolor='none')
            ax.add_patch(rect)
            ax.text(j, i, '未安排', ha='center', va='center', fontsize=8,
                    color='gray',
                    bbox=dict(facecolor='none', alpha=0, edgecolor='none', boxstyle='round,pad=0.2'))

# 添加调试信息
print(f"状态数据中1的数量: {np.sum(status_data == 1)}")
print(f"状态数据中0的数量: {np.sum(status_data == 0)}")
print(f"背景颜色设置为: #90EE90 (浅绿色)")

# 设置坐标轴
ax.set_xticks(np.arange(len(date_range)))
ax.set_xticklabels([d.strftime('%m-%d') for d in date_range], rotation=45)
ax.set_yticks(np.arange(len(assignees)))
ax.set_yticklabels(assignees)

# 添加网格线
ax.set_xticks(np.arange(len(date_range))-0.5, minor=True)
ax.set_yticks(np.arange(len(assignees))-0.5, minor=True)
ax.grid(which="minor", color="gray", linestyle='-', linewidth=0.5)
ax.tick_params(which="minor", size=0)

plt.title(f'2025年8月部门工作状态与项目分配可视化 (共{num_people}人)', pad=20)

# 添加颜色图例
green_patch = mpatches.Patch(color='#90EE90', label='已安排工作')
white_patch = mpatches.Patch(color='white', edgecolor='gray', label='未安排工作')  # 带边框以便在白色背景上可见
plt.legend(handles=[green_patch, white_patch], loc='lower center', bbox_to_anchor=(0.5, -0.05), ncol=2, fontsize=10)

plt.tight_layout(rect=[0, 0.05, 1, 0.95])  # 调整布局以容纳图例
plt.show()
#





# import matplotlib.font_manager as fm
#
# # 查找字体路径
# font_paths = fm.findSystemFonts(fontpaths=None, fontext=ttf')
#
# # 打印所有字体路径
# for path in font_paths:
#     print(path)