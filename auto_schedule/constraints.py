"""约束与评分模块

包含:
1. build_absolute: 将个体相对索引转换为绝对日期/时段表示
2. hard_penalties: 计算硬性冲突与缺失罚分 (仅返回总硬罚, 细节在自检里做)
3. soft_adjust: 计算软约束调整 (连排奖励 / 理论前置奖励 / 非理论后置 / 先修顺序 / 教师负载均衡 / 教师切换)

注意: 软约束字典 key 必须与手动排课 manual_soft/evaluate_soft 输出保持一致, 以便统一展示。
"""
from __future__ import annotations

from typing import List, Tuple, Dict, Any
import datetime

from .config import CONFIG
from .data_model import TimetableData

__all__ = [
    'build_absolute',
    'hard_penalties',
    'soft_adjust',
]


def build_absolute(individual, data: TimetableData):
    """个体 -> [(class_id, course, t1, t2, date, period_idx, is_two), ...]"""
    absolute = []
    for slot in individual:
        class_id, course, t1, t2, time_slot_idx = slot
        class_info = data.CLASSES[class_id]
        is_two = data.COURSE_DATA[course].get('is_two_teacher', False)
        if time_slot_idx is None or time_slot_idx < 0:
            absolute.append((class_id, course, t1, t2, None, None, is_two))
            continue
        date = class_info['start_date'] + datetime.timedelta(days=time_slot_idx // 2)
        period_idx = time_slot_idx % 2
        absolute.append((class_id, course, t1, t2, date, period_idx, is_two))
    return absolute


def hard_penalties(absolute, data: TimetableData) -> int:
    cfg = CONFIG
    HARD = cfg['HARD_PENALTY']
    miss_t = cfg['MISSING_TEACHER_PENALTY']
    miss_co = cfg['MISSING_CO_TEACHER_PENALTY']
    penalty = 0
    time_slot_map: Dict[tuple, Dict[str, set]] = {}
    scheduled_blocks = {(cid, c): 0 for cid in data.CLASSES for c in data.CLASSES[cid]['courses']}
    missing_blocks = 0

    for entry in absolute:
        class_id, course, t1, t2, date, period_idx, is_two = entry
        if date is None:
            missing_blocks += 1
            if t1 is None:
                penalty += miss_t
            if is_two:
                penalty += miss_co
            continue
        key_time = (date, period_idx)
        if key_time not in time_slot_map:
            time_slot_map[key_time] = {'teachers': set(), 'classes': set()}
        # teacher conflict
        if t1 in time_slot_map[key_time]['teachers']:
            penalty += HARD
        if t2 and t2 in time_slot_map[key_time]['teachers']:
            penalty += HARD
        # class conflict
        if class_id in time_slot_map[key_time]['classes']:
            penalty += HARD
        # register
        if t1:
            time_slot_map[key_time]['teachers'].add(t1)
        if t2:
            time_slot_map[key_time]['teachers'].add(t2)
        time_slot_map[key_time]['classes'].add(class_id)
        # teacher unavailable
        if t1 and t1 in data.TEACHER_UNAVAILABLE_SLOTS and (date, period_idx) in data.TEACHER_UNAVAILABLE_SLOTS[t1]:
            penalty += HARD
        if t2 and t2 in data.TEACHER_UNAVAILABLE_SLOTS and (date, period_idx) in data.TEACHER_UNAVAILABLE_SLOTS[t2]:
            penalty += HARD
        # count blocks
        scheduled_blocks[(class_id, course)] += 1
        # teacher missing
        if t1 is None:
            penalty += miss_t
        if is_two and (t2 is None or t2 == t1):
            penalty += miss_co
        if (not is_two) and t2 is not None:  # 单师课程不应有第二教师
            penalty += HARD

    # blocks mismatch
    for (cid, course), count in scheduled_blocks.items():
        required = data.COURSE_DATA[course]['blocks']
        if count != required:
            penalty += HARD * abs(required - count)
    if missing_blocks > 0:
        penalty += HARD * missing_blocks
    return penalty


def soft_adjust(absolute, data: TimetableData) -> Tuple[int, Dict[str, int]]:
    reward_seq = CONFIG.get('SOFT_REWARD_SEQUENCE', 2)
    prereq_pen = CONFIG['SOFT_PREREQ_PENALTY']
    theory_reward = CONFIG.get('THEORY_EARLY_REWARD', 5)
    non_theory_late_thr = CONFIG.get('NON_THEORY_LATE_THRESHOLD', 0.75)
    switch_pen = CONFIG.get('TEACHER_SWITCH_PENALTY', 20)
    theory_change_pen = CONFIG.get('THEORY_TEACHER_CHANGE_HARD', 2000)
    adjust = 0
    details: Dict[str, int] = {
        'consecutive_reward': 0,
        'theory_early_reward': 0,              # 负值奖励
        'non_theory_early_penalty': 0,         # 非理论课过早罚
        'prereq_violation_penalty': 0,         # 先修顺序违规
        'teacher_switch_penalty': 0,           # 非理论课教师切换
        'theory_teacher_inconsistent_penalty': 0,  # 理论课教师不唯一
        'teacher_balance_penalty': 0,          # 负载均衡
    }
    # 过滤出已排定块
    seq = sorted([x for x in absolute if x[4] is not None], key=lambda v: (v[4], v[5]))
    # 连排奖励（同班同课同日 上午->下午）
    for i in range(len(seq)-1):
        a, b = seq[i], seq[i+1]
        if a[0]==b[0] and a[1]==b[1] and a[4]==b[4] and a[5]==0 and b[5]==1:
            adjust -= reward_seq
            details['consecutive_reward'] -= reward_seq
    # 每班课程时间线
    per_class: Dict[str, list] = {}
    for cid, course, t1, t2, date, pidx, is_two in seq:
        per_class.setdefault(cid, []).append((date, pidx, course, t1))
    for cid, arr in per_class.items():
        arr.sort(key=lambda x:(x[0],x[1]))
        total = len(arr)
        if total == 0:
            continue
        # 统计每课程出现索引
        course_positions: Dict[str, list] = {}
        for idx, (_, _, c, _) in enumerate(arr):
            course_positions.setdefault(c, []).append(idx)
        last_index = {c: max(v) for c, v in course_positions.items()}
        # 遍历序列计算早/晚奖励或罚分与先修顺序
        for idx, (_, _, c, t) in enumerate(arr):
            info = data.COURSE_DATA.get(c, {})
            prereqs = info.get('prerequisites', [])
            is_theory = info.get('is_theory', False)
            # 先修顺序: 当前出现位置 <= 先修课程最后一次位置 -> 违规
            if prereqs:
                for p in prereqs:
                    if p in last_index and idx <= last_index[p]:
                        adjust += prereq_pen
                        details['prereq_violation_penalty'] += prereq_pen
                        break
            if is_theory:
                # 线性递减奖励 (越早奖励越多)
                reward = max(0, theory_reward * (total - idx) / total)
                if reward > 0:
                    r = int(reward)
                    adjust -= r
                    details['theory_early_reward'] -= r
            else:
                # 非理论：应靠后，若 idx < ideal_start 则罚 (差距越大罚越多: ideal_start-idx)
                ideal_start = int(total * non_theory_late_thr)
                if idx < ideal_start:
                    pen = (ideal_start - idx)
                    adjust += pen
                    details['non_theory_early_penalty'] += pen
        # 教师切换与理论课教师一致性
        by_course: Dict[str, list] = {}
        for _, _, c, t in arr:
            by_course.setdefault(c, []).append(t)
        for c, teachers in by_course.items():
            info = data.COURSE_DATA.get(c, {})
            is_theory = info.get('is_theory', False)
            switches = sum(1 for i in range(1, len(teachers)) if teachers[i] != teachers[i-1])
            if switches > 0:
                if is_theory:
                    adjust += theory_change_pen
                    details['theory_teacher_inconsistent_penalty'] += theory_change_pen
                else:
                    pen = switches * switch_pen
                    adjust += pen
                    details['teacher_switch_penalty'] += pen
    # 教师负载均衡
    teacher_load: Dict[str, int] = {}
    for cid, course, t1, t2, date, pidx, is_two in seq:
        if t1:
            teacher_load[t1] = teacher_load.get(t1, 0) + 1
        if t2:
            teacher_load[t2] = teacher_load.get(t2, 0) + 1
    if teacher_load:
        loads = list(teacher_load.values())
        spread = max(loads) - min(loads)
        balance_pen = int(spread * CONFIG.get('TEACHER_BALANCE_WEIGHT', 0))
        adjust += balance_pen
        details['teacher_balance_penalty'] = balance_pen
    return adjust, details
