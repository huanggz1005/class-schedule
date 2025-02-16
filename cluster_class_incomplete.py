'''
Version1 2024.5.13 Guozheng Huang
#分班不完整小班；只需要不完整小班助教时间和学生匹配
Input: student_time_incomplete_class.xlsx, ta_time_incomplete_class.xlsx
Output: incomplete_class_schedule.xlsx
'''
import numpy as np
import pandas as pd
student_time_path = 'data/student_time_incomplete_class.xlsx'
ta_time_path = 'data/ta_time_incomplete_class.xlsx'
output_path = 'result/incomplete_class_schedule.xlsx'

# 加载数据
student_data = pd.read_excel(student_time_path)
ta_data = pd.read_excel(ta_time_path)

# 数据预处理
ta_times = ta_data.iloc[:, 8:25].values
#按先后顺序去除重复学生数据；如果有同学填写多次问卷，以最后一次为准，前面数据删除
student_data = student_data.drop_duplicates(subset='2、你的学号', keep='last')
student_times = student_data.iloc[:, 8:25].values
student_preferences = student_data.iloc[:, 25:-2].values  # 假设最后一列不是评分
# 输出分配结果和后续步骤...
# 根据最优解生成课表
# 提取助教名和学生名
ta_names = ta_data.iloc[:, 6].values
student_names = student_data.iloc[:, 6].values
# 学生学号列
student_ids = student_data.iloc[:, 7].values
#建立上课时间字典
classtime = {0:'2-4', 1:'2-5',
             2:'2-6',
             3:'3-3', 4:'3-4', 5:'3-5', 6:'3-6',
             7:'4-3', 8:'4-4', 9:'4-5', 10:'4-6',
             11:'5-3', 12:'5-4', 13:'5-5', 14:'5-6',
             15:'6-6', 16:'7-6'}
def identify_critical_timeslots(student_times):
    """
    识别关键时间段，即至少有一部分学生只在这些时间段内有空。
    """
    num_timeslots = student_times.shape[1]
    critical_timeslots = set()
    for student_time in student_times:
        # 如果学生只在少数时间段有空，则这些时间段被视为关键时间段
        if np.sum(student_time) <= 2:  # 1为界限，可根据实际情况调整
            critical_timeslots.update(np.where(student_time)[0])
    return critical_timeslots

def find_unique_free_times_optimized(ta_times, critical_timeslots):
    num_tas = len(ta_times)
    ta_free_times = [-1] * num_tas

    # 首先尝试为关键时间段分配助教
    for timeslot in critical_timeslots:
        for ta_index, ta_time in enumerate(ta_times):
            if ta_time[timeslot] == 1 and ta_free_times[ta_index] == -1:
                ta_free_times[ta_index] = timeslot
                break

    # 然后为剩余的助教选择时间段
    for ta_index, ta_time in enumerate(ta_times):
        if ta_free_times[ta_index] == -1:  # 如果这位助教还没有被分配时间段
            for timeslot in range(len(ta_time)):
                #尽量选择其他助教没有选择的时间
                if timeslot not in ta_free_times and ta_time[timeslot] == 1:
                    ta_free_times[ta_index] = timeslot
                    break
    return ta_free_times


# 根据学生时间识别关键时间段
critical_timeslots = identify_critical_timeslots(student_times)
#print(critical_timeslots)
# 使用优化后的方法为助教选择空闲时间
ta_free_times_optimized = find_unique_free_times_optimized(ta_times, critical_timeslots)
#print(ta_free_times_optimized)
#ta_free_times_optimized= [14,11,0,2]

# 接下来可以使用`assign_students_to_tas_with_fixed_times`函数为学生分配助教
def assign_students_to_tas_with_fixed_times(ta_times, student_times, ta_free_times, max_students_per_ta=8):
    num_students = len(student_times)
    num_tas = len(ta_times)
    assignments = [-1] * num_students  # 初始时所有学生未分配
    ta_student_count = [0] * num_tas  # 每位助教分配的学生数

    for student_index, student_time in enumerate(student_times):
        for ta_index in range(num_tas):
            # 获取当前助教的空闲时间段
            ta_free_time_index = ta_free_times[ta_index]
            # 如果当前助教没有选定的空闲时间段，跳过
            if ta_free_time_index == -1:
                continue
            # 检查学生是否在这个时间段有空
            if student_time[ta_free_time_index] and ta_student_count[ta_index] < max_students_per_ta:
                assignments[student_index] = ta_index
                ta_student_count[ta_index] += 1
                break

    # 检查是否有学生未被分配
    unassigned_students = [index for index, assignment in enumerate(assignments) if assignment == -1]
    if unassigned_students:
        # print(f"有{len(unassigned_students)}个学生未被分配")
        # print(unassigned_students)
        # print(student_times[unassigned_students])
        #强行分配，将剩余学生分配给可分配时间的助教中学生最少的
        for student_index in unassigned_students:
            #找到该学生的空闲时间
            student_time = student_times[student_index]
            #找到所有有空的助教
            free_ta_indices = [ta_index for ta_index, free_time_index in enumerate(ta_free_times) if free_time_index != -1 and student_time[free_time_index]]
            #找到学生最少的助教
            if(len(free_ta_indices) == 0):
                # print(f"学生{student_index}没有找到合适的助教")
                # print(student_names[student_index])
                # print(student_ids[student_index])
                continue
            else:
                min_student_ta_index = min(free_ta_indices, key=lambda ta_index: ta_student_count[ta_index])
                #分配学生
                assignments[student_index] = min_student_ta_index
                ta_student_count[min_student_ta_index] += 1
    return assignments


assignments = assign_students_to_tas_with_fixed_times(ta_times, student_times, ta_free_times_optimized)

# 基于最优解和新的限制条件，为每位助教和其学生找到共同的时间并输出排课表

# 重新定义一个函数来寻找每位助教和其学生的共同有空时间，并记录下来
def find_common_time_for_ta_and_students(chromosome):
    # 分配学生到助教
    ta_to_students = {ta_index: [] for ta_index in range(len(ta_times))}
    for student_index, ta_index in enumerate(chromosome):
        if(ta_index != -1):
            ta_to_students[ta_index].append(student_index)

    # 为每位助教找到一个共同有空的时间段
    ta_common_times = {}
    for ta_index, students in ta_to_students.items():
        if students:  # 如果助教有学生
            students_times = np.all(student_times[students], axis=0)
            common_times = students_times & ta_times[ta_index]
            # 记录下每位助教的共同时间段索引
            common_time_index = np.where(common_times)[0]
            if common_time_index.size > 0:
                ta_common_times[ta_index] = common_time_index[0]  # 记录第一个共同时间段的索引
            else:
                ta_common_times[ta_index] = None  # 如果没有共同时间，则记录为None

    return ta_common_times

# 使用最优解找到每位助教的共同时间段
ta_common_times = find_common_time_for_ta_and_students(assignments)
# 根据新格式重新组织课表数据，包括助教的最终选定时间
schedule_list_with_time = []
for student_index, ta_index in enumerate(assignments):
    if ta_index != -1:
        ta_name = ta_names[ta_index]
        student_name = student_names[student_index]
        student_id = student_ids[student_index]
        common_time_index = ta_common_times[ta_index]
        common_time = f"{common_time_index}" if common_time_index is not None else "无共同时间"
        schedule_list_with_time.append([ta_name, student_name, student_id, classtime[int(common_time)]])


# 转换为DataFrame
schedule_df_with_time = pd.DataFrame(schedule_list_with_time,
                                     columns=['助教名', '学生名', '学生学号', '上课时间'])

#去除重复的学生
schedule_df_with_time = schedule_df_with_time.drop_duplicates(subset='学生名', keep='first')
#按助教名排序
schedule_df_with_time = schedule_df_with_time.sort_values(by='助教名').reset_index(drop=True)

# 创建一个空的 DataFrame 用于存储结果
result_df = pd.DataFrame(columns=schedule_df_with_time.columns)

# 遍历原 DataFrame 并插入空行
for i in range(len(schedule_df_with_time)):
    result_df = pd.concat([result_df, schedule_df_with_time.iloc[[i]]], ignore_index=True)
    if i < len(schedule_df_with_time) - 1 and schedule_df_with_time.loc[i, '助教名'] != schedule_df_with_time.loc[i + 1, '助教名']:
        result_df = pd.concat([result_df, pd.DataFrame([None] * len(schedule_df_with_time.columns)).T], ignore_index=True)
#去掉后面四列
result_df = result_df.iloc[:, :4]
result_df.to_excel(output_path, index=False)
