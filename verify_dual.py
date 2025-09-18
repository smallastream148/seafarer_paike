import pandas as pd
from auto_schedule.data_model import TimetableData
import sys

# 读取最新导出文件 (默认 __verify_dual.xlsx 或命令行传入)
excel = '__verify_dual.xlsx'
if len(sys.argv) > 1:
    excel = sys.argv[1]

data = TimetableData()
# 找出双师课程列表
dual_courses = {name for name, c in data.COURSE_DATA.items() if c.get('is_two_teacher')}
practical_courses = {name for name, c in data.COURSE_DATA.items() if c.get('is_practical')}
print('双师课程:', dual_courses)
print('实操(非理论)课程:', practical_courses)

try:
    df = pd.read_excel(excel, sheet_name='排课明细')
except Exception as e:
    print('读取结果文件失败:', e)
    sys.exit(1)

issues = []
rows = []
for _, r in df.iterrows():
    course = str(r['课程'])
    t1 = str(r['教师1']) if pd.notna(r['教师1']) else ''
    t2 = str(r['教师2']) if pd.notna(r['教师2']) else ''
    is_dual = course in dual_courses
    is_prac = course in practical_courses
    if is_dual:
        if (not t2) or (t1 == t2):
            issues.append(f"缺第二教师或重复: 班级{r['班级ID']} {course} {r['日期']} {r['时段']} t1={t1} t2={t2}")
    rows.append({'班级ID': r['班级ID'], '课程': course, 't1': t1, 't2': t2, 'is_dual': is_dual, 'is_practical': is_prac})

out_df = pd.DataFrame(rows)
print('\n排课中双师/实操课程分配概览(前50行):')
print(out_df.head(50).to_string(index=False))

if issues:
    print('\n发现问题条目:')
    for i in issues:
        print(i)
else:
    print('\n双师课程全部具备两个不同教师。')
