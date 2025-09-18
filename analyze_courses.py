import pandas as pd

def normalize_two(v):
    return str(v).strip().lower() in {'y','yes','true','1','双','2','two'}

def main():
    df = pd.read_excel('排课数据.xlsx', sheet_name='课程数据')
    print('列:', list(df.columns))
    rows = []
    for _, r in df.iterrows():
        name = r.get('课程名称')
        teachers_raw = str(r.get('available_teachers',''))
        teachers = [t.strip() for t in teachers_raw.replace('，',',').split(',') if t.strip()]
    is_two = normalize_two(r.get('is_two_teacher',''))
    # 派生: 双师=实操; 单师=理论
    is_practical = is_two
    is_theory = not is_two
    rows.append((name, len(teachers), teachers, is_theory, is_practical, is_two))
    print('课程概览:')
    for name, tcnt, teachers, th, pr, two in rows:
        print(f"{name} | 教师数:{tcnt} {teachers} | 理论:{th} 实操:{pr} 双师:{two}")
    print('\n统计:')
    print('理论课程数:', sum(1 for r in rows if r[3]))
    print('实操课程数:', sum(1 for r in rows if r[4]))
    print('双师课程数:', sum(1 for r in rows if r[5]))
    # 在新规则下 理论+双师 不会出现; 实操+双师 == 全部双师

if __name__ == '__main__':
    main()
