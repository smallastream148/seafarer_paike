from typing import List, Dict, Tuple
from collections import defaultdict

try:
    from .manual_core import PlacedBlock, TimetableData  # 包模式
except ImportError:  # 脚本模式
    from manual_core import PlacedBlock, TimetableData  # type: ignore

"""手动排课软约束评估，需与自动 GA 版本软约束字段名称保持一致。

输出 details 中包含:
    consecutive_reward (负值奖励)
    theory_early_reward (负值奖励)
    non_theory_early_penalty
    prereq_violation_penalty
    teacher_switch_penalty
    theory_teacher_inconsistent_penalty
    teacher_balance_penalty
"""

CONFIG_SOFT = {
        'CONSECUTIVE_REWARD': 2,
        'SOFT_PREREQ_PENALTY': 2000,
        'THEORY_EARLY_REWARD': 5,
        'NON_THEORY_LATE_THRESHOLD': 0.75,
        'TEACHER_SWITCH_PENALTY': 20,
        'THEORY_TEACHER_CHANGE_HARD': 2000,
        'TEACHER_BALANCE_WEIGHT': 5,
}

def evaluate_soft(blocks: List[PlacedBlock], data: TimetableData) -> Tuple[int, Dict[str,int]]:
    if not blocks:
        return 0, {
            'consecutive_reward':0,
            'theory_early_reward':0,
            'non_theory_early_penalty':0,
            'prereq_violation_penalty':0,
            'teacher_switch_penalty':0,
            'theory_teacher_inconsistent_penalty':0,
            'teacher_balance_penalty':0,
        }
    seq = sorted(blocks, key=lambda b: (b.date, b.period))
    details = {
        'consecutive_reward':0,
        'theory_early_reward':0,
        'non_theory_early_penalty':0,
        'prereq_violation_penalty':0,
        'teacher_switch_penalty':0,
        'theory_teacher_inconsistent_penalty':0,
        'teacher_balance_penalty':0,
    }
    adjust = 0
    # 连排奖励
    for i in range(len(seq)-1):
        a,b = seq[i], seq[i+1]
        if a.class_id==b.class_id and a.course==b.course and a.date==b.date and a.period==0 and b.period==1:
            adjust -= CONFIG_SOFT['CONSECUTIVE_REWARD']
            details['consecutive_reward'] -= CONFIG_SOFT['CONSECUTIVE_REWARD']
    # 按班级
    per_class = defaultdict(list)
    for b in seq:
        per_class[b.class_id].append(b)
    prereq_pen = CONFIG_SOFT['SOFT_PREREQ_PENALTY']
    theory_reward = CONFIG_SOFT['THEORY_EARLY_REWARD']
    non_theory_late_thr = CONFIG_SOFT['NON_THEORY_LATE_THRESHOLD']
    switch_pen = CONFIG_SOFT['TEACHER_SWITCH_PENALTY']
    theory_change_pen = CONFIG_SOFT['THEORY_TEACHER_CHANGE_HARD']
    for cid, lst in per_class.items():
        lst.sort(key=lambda b:(b.date,b.period))
        total = len(lst)
        if total == 0:
            continue
        last_idx = {}
        for i,b in enumerate(lst):
            last_idx.setdefault(b.course, []).append(i)
        last_idx = {c:max(v) for c,v in last_idx.items()}
        # 遍历块
        for i,b in enumerate(lst):
            cinfo = data.courses[b.course]
            prereqs = cinfo.prerequisites
            is_theory = getattr(cinfo, 'is_theory', False)
            if prereqs:
                for p in prereqs:
                    if p in last_idx and i <= last_idx[p]:
                        adjust += prereq_pen
                        details['prereq_violation_penalty'] += prereq_pen
                        break
            if is_theory:
                reward = max(0, theory_reward * (total - i) / total)
                if reward>0:
                    r = int(reward)
                    adjust -= r
                    details['theory_early_reward'] -= r
            else:
                ideal_start = int(total * non_theory_late_thr)
                if i < ideal_start:
                    pen = (ideal_start - i)
                    adjust += pen
                    details['non_theory_early_penalty'] += pen
        # 教师切换统计
        by_course = defaultdict(list)
        for b in lst:
            by_course[b.course].append(b.teacher1)
        for course, t_seq in by_course.items():
            cinfo = data.courses[course]
            is_theory = getattr(cinfo, 'is_theory', False)
            switches = sum(1 for i in range(1,len(t_seq)) if t_seq[i]!=t_seq[i-1])
            if switches>0:
                if is_theory:
                    adjust += theory_change_pen
                    details['theory_teacher_inconsistent_penalty'] += theory_change_pen
                else:
                    pen = switches * switch_pen
                    adjust += pen
                    details['teacher_switch_penalty'] += pen
    # 教师负载均衡
    load = defaultdict(int)
    for b in seq:
        load[b.teacher1]+=1
        if b.teacher2:
            load[b.teacher2]+=1
    if load:
        spread = max(load.values()) - min(load.values())
        pen = spread * CONFIG_SOFT['TEACHER_BALANCE_WEIGHT']
        adjust += pen
        details['teacher_balance_penalty'] = pen
    return adjust, details
