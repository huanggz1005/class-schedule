'''
Version1 2024.5.14 Guozheng Huang
#输出全部学生中既不在完整分班表也不在不完整分班表的学生姓名,存为表格
Input: student_info.xlsx，complete_class_schedule.xlsx，incomplete_class_schedule.xlsx
Output: unassigned_student_info.xlsx
'''
import numpy as np
import pandas as pd
studentinfopath = 'data/student_info.xlsx'
complete_schedule_path = 'result/complete_class_schedule.xlsx'
incomplete_schedule_path = 'result/incomplete_class_schedule.xlsx'
output_path = 'result/unassigned_student_info.xlsx'

# 加载全部学生数据
student_data = pd.read_excel(studentinfopath)
# 提取学生名
student_names = student_data['姓名'].values


#加载分班表
complete_class = pd.read_excel(complete_schedule_path)
# 提取学生名
complete_student_names = complete_class['学生名'].values
# 去除nan值
complete_student_names = complete_student_names[~pd.isnull(complete_student_names)]

#加载不完整分班表
incomplete_class = pd.read_excel(incomplete_schedule_path)
# 提取填了两个问卷的学生名
incomplete_student_names = incomplete_class['学生名'].values
incomplete_student_names = incomplete_student_names[~pd.isnull(incomplete_student_names)]

#输出全部学生中既不在完整分班表也不在不完整分班表的学生姓名
unassigned_student_names = []
for student_name in student_names:
    if student_name not in complete_student_names and student_name not in incomplete_student_names:
        unassigned_student_names.append(student_name)
#在info中找到该学生所有信息
unassigned_student_info = student_data[student_data['姓名'].isin(unassigned_student_names)]
unassigned_student_info.to_excel(output_path, index=False)