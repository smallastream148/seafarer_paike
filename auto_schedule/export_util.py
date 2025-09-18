"""自动排课导出工具

提供: export_schedule(individual, data, path)
与 GA 引擎中逻辑保持一致，独立出来便于复用 (如 sweep 时多次导出)。
"""
from __future__ import annotations

import datetime
import pandas as pd
from typing import List, Tuple

from .data_model import TimetableData

def export_schedule(individual, data: TimetableData, excel_out: str):
    rows = []
    for slot in individual:
        class_id, course, teacher1, teacher2, time_idx = slot
        class_info = data.CLASSES[class_id]
        if time_idx is None or time_idx < 0:
            date = None
            period = None
        else:
            date = class_info['start_date'] + datetime.timedelta(days=time_idx // 2)
            period = '上午' if time_idx % 2 == 0 else '下午'
        rows.append({'班级ID': class_id, '课程': course, '教师1': teacher1, '教师2': teacher2, '日期': date, '时段': period})
    df = pd.DataFrame(rows)
    try:
        df.sort_values(['班级ID', '日期', '时段'], inplace=True)
    except Exception:
        pass
    teacher_hours = {}
    course_progress = {}
    for slot in individual:
        class_id, course, t1, t2, time_idx = slot
        if time_idx is None or time_idx < 0:
            continue
        if t1:
            teacher_hours[t1] = teacher_hours.get(t1, 0) + 1
        if t2:
            teacher_hours[t2] = teacher_hours.get(t2, 0) + 1
        course_progress.setdefault((class_id, course), 0)
        course_progress[(class_id, course)] += 1
    teacher_df = pd.DataFrame([
        {'教师': k, '已排课时': v} for k, v in sorted(teacher_hours.items(), key=lambda x: (-x[1], x[0]))
    ])
    course_rows = []
    for class_id, info in data.CLASSES.items():
        for c in info['courses']:
            required = data.COURSE_DATA[c]['blocks']
            scheduled = course_progress.get((class_id, c), 0)
            course_rows.append({
                '班级ID': class_id,
                '课程': c,
                '需求块数': required,
                '已排块数': scheduled,
                '完成率%': round(scheduled / required * 100, 2) if required else 0.0
            })
    course_df = pd.DataFrame(course_rows)
    with pd.ExcelWriter(excel_out) as writer:
        df.to_excel(writer, sheet_name='排课明细', index=False)
        teacher_df.to_excel(writer, sheet_name='教师课时', index=False)
        course_df.to_excel(writer, sheet_name='课程进度', index=False)
    return excel_out

__all__ = ['export_schedule']
