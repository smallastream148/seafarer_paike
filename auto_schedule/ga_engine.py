"""GA 引擎

包含：
generate_individual / repair_individual / normalize_single_teacher / mutate_individual
evaluate_schedule / quick_self_check / run_scheduler

依赖 constraints.build_absolute, hard_penalties, soft_adjust
"""
from __future__ import annotations

import random
import datetime
from typing import Tuple, Dict, Any
from deap import base, creator, tools
import pandas as pd

from .config import CONFIG
from .data_model import TimetableData
from .constraints import build_absolute, hard_penalties, soft_adjust

__all__ = [
    'generate_individual', 'repair_individual', 'normalize_single_teacher', 'mutate_individual',
    'evaluate_schedule', 'quick_self_check', 'run_scheduler'
]


def set_random_seed(seed: int | None):
    if seed is None:
        return
    random.seed(seed)


def generate_individual(data: TimetableData):
    import itertools
    positions = []
    for class_id, info in data.CLASSES.items():
        for course in info['courses']:
            for _ in range(data.COURSE_DATA[course]['blocks']):
                positions.append((class_id, course))
    random.shuffle(positions)
    individual = []
    occupancy = {}

    def slot_to_date_idx(class_id, idx):
        class_info = data.CLASSES[class_id]
        date = class_info['start_date'] + datetime.timedelta(days=idx // 2)
        period = idx % 2
        return date, period

    for class_id, course in positions:
        base_indices = data.CLASS_SLOT_CACHE.get(class_id, [])
        if not base_indices:
            individual.append((class_id, course, None, None, -1))
            continue
        course_info = data.COURSE_DATA[course]
        # 现在 is_theory 已由 data_model 派生: 单师 => True
        is_theory = course_info.get('is_theory', False)
        # 生成候选时间索引：理论课尽量靠前，其它课程靠后(双师视为实操靠后)
        idx_candidates = base_indices[:]
        if is_theory:
            idx_candidates.sort()  # 前置
        else:
            idx_candidates.sort(reverse=True)  # 后置
        is_two = course_info.get('is_two_teacher', False)
        teachers = list(course_info['available_teachers'])
        # 若是理论课并且非双师，固定第一教师；若既是理论又是双师，则仍需两个教师
        if is_theory and not is_two:
            teachers = [teachers[0]]
        else:
            random.shuffle(teachers)
        assigned = False
        for idx in idx_candidates:
            date, period = slot_to_date_idx(class_id, idx)
            if class_id in data.CLASS_UNAVAILABLE_SLOTS and (date, period) in data.CLASS_UNAVAILABLE_SLOTS[class_id]:
                continue
            key = (date, period)
            occ = occupancy.setdefault(key, {'classes': set(), 'teachers': set()})
            if class_id in occ['classes']:
                continue
            if is_two:
                for t1, t2 in itertools.permutations(teachers, 2):
                    if t1 in data.TEACHER_UNAVAILABLE_SLOTS and (date, period) in data.TEACHER_UNAVAILABLE_SLOTS[t1]:
                        continue
                    if t2 in data.TEACHER_UNAVAILABLE_SLOTS and (date, period) in data.TEACHER_UNAVAILABLE_SLOTS[t2]:
                        continue
                    if t1 in occ['teachers'] or t2 in occ['teachers']:
                        continue
                    individual.append((class_id, course, t1, t2, idx))
                    occ['classes'].add(class_id)
                    occ['teachers'].add(t1)
                    occ['teachers'].add(t2)
                    assigned = True
                    break
                if assigned:
                    break
            else:
                for t in teachers:
                    if t in data.TEACHER_UNAVAILABLE_SLOTS and (date, period) in data.TEACHER_UNAVAILABLE_SLOTS[t]:
                        continue
                    if t in occ['teachers']:
                        continue
                    individual.append((class_id, course, t, None, idx))
                    occ['classes'].add(class_id)
                    occ['teachers'].add(t)
                    assigned = True
                    break
                if assigned:
                    break
        if not assigned:
            individual.append((class_id, course, teachers[0] if teachers else None, None, -1))
    individual = repair_individual(individual, data)
    individual = normalize_single_teacher(individual, data)
    return individual


def repair_individual(individual, data: TimetableData, max_pass=2):
    import itertools
    for _ in range(max_pass):
        occupancy = {}
        for i, (cid, course, t1, t2, idx) in enumerate(individual):
            if idx is None or idx < 0:
                continue
            date = data.CLASSES[cid]['start_date'] + datetime.timedelta(days=idx // 2)
            period = idx % 2
            key = (date, period)
            occ = occupancy.setdefault(key, {'classes': set(), 'teachers': set()})
            occ['classes'].add(cid)
            if t1:
                occ['teachers'].add(t1)
            if t2:
                occ['teachers'].add(t2)
        modified = False
        for i, (cid, course, t1, t2, idx) in enumerate(individual):
            if idx >= 0 and t1 is not None:
                continue
            base_indices = data.CLASS_SLOT_CACHE.get(cid, [])
            if not base_indices:
                continue
            random.shuffle(base_indices)
            is_two = data.COURSE_DATA[course].get('is_two_teacher', False)
            course_info = data.COURSE_DATA[course]
            teachers = list(course_info['available_teachers'])
            if course_info.get('is_theory', False) and not is_two:
                # 理论单师固定第一教师
                teachers = [teachers[0]]
            else:
                random.shuffle(teachers)
            placed = False
            for cand in base_indices:
                date = data.CLASSES[cid]['start_date'] + datetime.timedelta(days=cand // 2)
                period = cand % 2
                if cid in data.CLASS_UNAVAILABLE_SLOTS and (date, period) in data.CLASS_UNAVAILABLE_SLOTS[cid]:
                    continue
                key = (date, period)
                occ = occupancy.setdefault(key, {'classes': set(), 'teachers': set()})
                if cid in occ['classes']:
                    continue
                if is_two:
                    for ta, tb in itertools.permutations(teachers, 2):
                        if ta in data.TEACHER_UNAVAILABLE_SLOTS and (date, period) in data.TEACHER_UNAVAILABLE_SLOTS[ta]:
                            continue
                        if tb in data.TEACHER_UNAVAILABLE_SLOTS and (date, period) in data.TEACHER_UNAVAILABLE_SLOTS[tb]:
                            continue
                        if ta in occ['teachers'] or tb in occ['teachers']:
                            continue
                        individual[i] = (cid, course, ta, tb, cand)
                        occ['classes'].add(cid)
                        occ['teachers'].add(ta)
                        occ['teachers'].add(tb)
                        placed = True
                        modified = True
                        break
                    if placed:
                        break
                else:
                    for ta in teachers:
                        if ta in data.TEACHER_UNAVAILABLE_SLOTS and (date, period) in data.TEACHER_UNAVAILABLE_SLOTS[ta]:
                            continue
                        if ta in occ['teachers']:
                            continue
                        individual[i] = (cid, course, ta, None, cand)
                        occ['classes'].add(cid)
                        occ['teachers'].add(ta)
                        placed = True
                        modified = True
                        break
                    if placed:
                        break
        if not modified:
            break
    return individual


def normalize_single_teacher(individual, data: TimetableData):
    for i, (cid, course, t1, t2, idx) in enumerate(individual):
        if not data.COURSE_DATA[course].get('is_two_teacher', False) and t2 is not None:
            individual[i] = (cid, course, t1, None, idx)
    return individual


def mutate_individual(individual, data: TimetableData, indpb=0.05):
    for i in range(len(individual)):
        if random.random() < indpb:
            class_id, course, teacher1, teacher2, old_idx = individual[i]
            base_indices = data.CLASS_SLOT_CACHE.get(class_id, [])
            if not base_indices:
                continue
            new_idx = random.choice(base_indices)
            course_info = data.COURSE_DATA[course]
            is_theory = course_info.get('is_theory', False)
            is_two = course_info.get('is_two_teacher', False)
            # 变异策略：
            # 1) 理论+单师：强制第一教师，清空第二教师
            # 2) 理论+双师：保留原两个教师；若缺失某个教师则尝试自动补齐为不同教师
            # 3) 非理论：保持原教师对；若是双师且缺第二教师尝试补齐
            if is_theory and not is_two:
                first_teacher = course_info['available_teachers'][0]
                individual[i] = (class_id, course, first_teacher, None, new_idx)
            else:
                # 尝试补齐双师缺失
                if is_two:
                    teachers_all = list(course_info['available_teachers'])
                    # 如果 teacher1 缺失，用列表第一位
                    if teacher1 is None or teacher1 not in teachers_all:
                        teacher1 = teachers_all[0]
                    if (teacher2 is None) or (teacher2 == teacher1) or (teacher2 not in teachers_all):
                        # 选择与 teacher1 不同的另一教师；若不足两人则置 None 等待 repair
                        candidates = [t for t in teachers_all if t != teacher1]
                        teacher2 = candidates[0] if candidates else None
                # 写回（非理论单师 / 理论双师 / 非理论双师）
                individual[i] = (class_id, course, teacher1, teacher2 if is_two else (None if (is_theory and not is_two) else teacher2), new_idx)
    individual = repair_individual(individual, data, max_pass=1)
    individual = normalize_single_teacher(individual, data)
    return (individual,)


def evaluate_schedule(individual, data: TimetableData):
    absolute = build_absolute(individual, data)
    hard = hard_penalties(absolute, data)
    soft_total, _ = soft_adjust(absolute, data)
    return (hard + soft_total,)


def quick_self_check(individual, data: TimetableData):
    HARD = CONFIG['HARD_PENALTY']
    miss_t_pen = CONFIG['MISSING_TEACHER_PENALTY']
    miss_co_pen = CONFIG['MISSING_CO_TEACHER_PENALTY']
    absolute = build_absolute(individual, data)
    teacher_conflicts = class_conflicts = teacher_unavailable = 0
    missing_blocks = missing_teacher = missing_co_teacher = extra_second_teacher = 0
    block_mismatch = 0
    time_slot_map = {}
    scheduled_blocks = {(cid, c): 0 for cid in data.CLASSES for c in data.CLASSES[cid]['courses']}
    for entry in absolute:
        class_id, course, t1, t2, date, period_idx, is_two = entry
        if date is None:
            missing_blocks += 1
            if t1 is None:
                missing_teacher += 1
            if is_two:
                missing_co_teacher += 1
            continue
        key = (date, period_idx)
        if key not in time_slot_map:
            time_slot_map[key] = {'teachers': set(), 'classes': set()}
        if t1 in time_slot_map[key]['teachers']:
            teacher_conflicts += 1
        if t2 and t2 in time_slot_map[key]['teachers']:
            teacher_conflicts += 1
        if class_id in time_slot_map[key]['classes']:
            class_conflicts += 1
        if t1:
            time_slot_map[key]['teachers'].add(t1)
        if t2:
            time_slot_map[key]['teachers'].add(t2)
        time_slot_map[key]['classes'].add(class_id)
        if t1 and t1 in data.TEACHER_UNAVAILABLE_SLOTS and (date, period_idx) in data.TEACHER_UNAVAILABLE_SLOTS[t1]:
            teacher_unavailable += 1
        if t2 and t2 in data.TEACHER_UNAVAILABLE_SLOTS and (date, period_idx) in data.TEACHER_UNAVAILABLE_SLOTS[t2]:
            teacher_unavailable += 1
        scheduled_blocks[(class_id, course)] += 1
        if t1 is None:
            missing_teacher += 1
        if is_two and (t2 is None or t2 == t1):
            missing_co_teacher += 1
        if (not is_two) and t2 is not None:
            extra_second_teacher += 1
    for (cid, course), count in scheduled_blocks.items():
        required = data.COURSE_DATA[course]['blocks']
        if count != required:
            block_mismatch += abs(required - count)
    soft_adj, soft_details = soft_adjust(absolute, data)
    hard_components_penalty = (
        teacher_conflicts * HARD +
        class_conflicts * HARD +
        teacher_unavailable * HARD +
        missing_blocks * HARD +
        block_mismatch * HARD +
        missing_teacher * miss_t_pen +
        missing_co_teacher * miss_co_pen
    )
    total = hard_components_penalty + soft_adj
    return {
        'hard_ok': hard_components_penalty == 0,
        'teacher_conflicts': teacher_conflicts,
        'class_conflicts': class_conflicts,
        'teacher_unavailable': teacher_unavailable,
        'missing_blocks': missing_blocks,
        'block_mismatch': block_mismatch,
        'missing_teacher': missing_teacher,
        'missing_co_teacher': missing_co_teacher,
        'extra_second_teacher': extra_second_teacher,
        'hard_penalty': hard_components_penalty,
        'soft_adjust': soft_adj,
        'soft_details': soft_details,
        'total_fitness': total,
    }


def run_scheduler(pop_size=CONFIG['DEFAULT_POP'], ngen=CONFIG['DEFAULT_GEN'], excel_out=CONFIG['DEFAULT_OUTPUT'], seed=CONFIG['DEFAULT_SEED'], verbose=1, excel_path: str | None = None):
    from .constraints import build_absolute
    from deap import creator
    def log(msg, level='INFO'):
        if verbose >= 1 or level == 'ERROR':
            print(f"[{level}] {msg}")
    if verbose:
        print(f"[INFO] 启动 GA: pop={pop_size} gen={ngen} seed={seed}")
    set_random_seed(seed)
    data = TimetableData(excel_path or '排课数据.xlsx')
    if verbose:
        print('[INFO] 数据加载完成')
    try:
        creator.FitnessMin
    except Exception:
        creator.create('FitnessMin', base.Fitness, weights=(-1.0,))
    try:
        creator.Individual
    except Exception:
        creator.create('Individual', list, fitness=creator.FitnessMin)
    toolbox = base.Toolbox()
    toolbox.register('individual_gen', generate_individual, data)
    def create_individual():
        return creator.Individual(toolbox.individual_gen())
    toolbox.register('individual', create_individual)
    toolbox.register('population', tools.initRepeat, list, toolbox.individual)
    toolbox.register('evaluate', evaluate_schedule, data=data)
    def safe_cx(ind1, ind2):
        if len(ind1) < 2 or len(ind2) < 2:
            return ind1, ind2
        return tools.cxTwoPoint(ind1, ind2)
    toolbox.register('mate', safe_cx)
    toolbox.register('mutate', mutate_individual, data=data, indpb=0.08)
    toolbox.register('select', tools.selTournament, tournsize=3)
    pop = toolbox.population(n=pop_size)
    if verbose:
        print('[INFO] 初始种群生成完成')
    best = None
    best_fit = float('inf')
    patience = CONFIG.get('EARLY_STOP_PATIENCE', None)
    no_improve = 0
    for g in range(ngen):
        invalid = [ind for ind in pop if not ind.fitness.valid]
        fits = map(toolbox.evaluate, invalid)
        for ind, fit in zip(invalid, fits):
            ind.fitness.values = fit
        current_best = tools.selBest(pop, 1)[0]
        cur_fit = current_best.fitness.values[0]
        if cur_fit < best_fit:
            best_fit = cur_fit
            best = toolbox.clone(current_best)
            no_improve = 0
        else:
            no_improve += 1
        if verbose >= 2:
            print(f"[INFO] Gen {g} best={best_fit}")
        if patience is not None and no_improve >= patience:
            if verbose:
                print(f"[INFO] 早停: 连续 {patience} 代无改进")
            break
        offspring = toolbox.select(pop, len(pop))
        offspring = list(map(toolbox.clone, offspring))
        for c1, c2 in zip(offspring[::2], offspring[1::2]):
            if random.random() < 0.6:
                toolbox.mate(c1, c2)
                del c1.fitness.values
                del c2.fitness.values
        for mut in offspring:
            if random.random() < 0.2:
                toolbox.mutate(mut)
                del mut.fitness.values
        pop[:] = offspring
    if verbose:
        print('[INFO] 进化完成, 选择最佳个体')
    metrics = quick_self_check(best, data)
    if verbose:
        print('[INFO] 自检结果: ' + ', '.join([f"{k}={v}" for k,v in metrics.items() if k != 'total_fitness']))
        print('[INFO] 软约束细分: ' + ', '.join([f"{k}={v}" for k,v in metrics.get('soft_details', {}).items()]))
        # 双师课程统计
        dual_courses = [c for c,v in data.COURSE_DATA.items() if v.get('is_two_teacher')]
        if dual_courses:
            print(f"[INFO] 双师课程: {dual_courses}")
            # 统计最佳个体中是否有缺第二教师的双师块
            abs_best = build_absolute(best, data)
            missing_dual = [ (cid,course,date,period) for (cid,course,t1,t2,date,period,is_two) in abs_best if is_two and (t2 is None or t2==t1) ]
            if missing_dual:
                print(f"[WARN] 存在{len(missing_dual)}个双师块缺第二教师或重复教师, 这将被硬罚, 需检查数据 available_teachers 或增加教师可用性")
            else:
                print('[INFO] 双师课程全部分配两位不同教师')
    # 导出
    rows = []
    for slot in best:
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
    for slot in best:
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
    if verbose:
        print(f"[INFO] 已导出最优排课到 {excel_out}")
    return best, metrics
