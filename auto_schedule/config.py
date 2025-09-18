CONFIG = {
    'HARD_PENALTY': 10000,
    'SOFT_PREREQ_PENALTY': 2000,
    'PRACTICAL_EARLY_SOFT_PENALTY': 8,
    'CONSECUTIVE_REWARD': 2,
    'PRACTICAL_LATE_THRESHOLD': 0.85,
    'PRACTICAL_EARLY_WEIGHT_SCALE': 25,
    'PRACTICAL_WEIGHTED_ACTIVATE_RATIO': 0.6,
    'DEFAULT_POP': 80,
    'DEFAULT_GEN': 200,
    'DEFAULT_OUTPUT': '排课结果_GA.xlsx',
    'DEFAULT_SEED': 42,
    'MISSING_TEACHER_PENALTY': 5000,
    'MISSING_CO_TEACHER_PENALTY': 3000,
    'TEACHER_BALANCE_WEIGHT': 3,  # 降低负载均衡权重，避免过度牵制其它软目标
    'EARLY_STOP_PATIENCE': 30,
    # 新增：理论/非理论课程策略参数
    'THEORY_EARLY_REWARD': 6,              # 略增强理论前置驱动力
    'NON_THEORY_LATE_THRESHOLD': 0.75,      # 非理论课理想开始占比 (靠后)
    'TEACHER_SWITCH_PENALTY': 12,           # 下调教师切换罚分，减少其在总分中的占比
    'THEORY_TEACHER_CHANGE_HARD': 2000,     # 理论课教师不唯一时的高额罚分(软中体现, 不是硬冲突)
}

# --- 参数调优实验批次说明 ---
# Batch T1 (当前修改):
#   TEACHER_SWITCH_PENALTY: 20 -> 12  (期望降低 teacher_switch_penalty 在 soft_details 中的主导比例)
#   TEACHER_BALANCE_WEIGHT: 5  -> 3   (减小均衡强度, 让进化更自由形成局部聚类减少切换)
#   THEORY_EARLY_REWARD: 5 -> 6       (略加强理论课前置激励)
# 观察指标:
#   soft_details.teacher_switch_penalty 是否显著下降 (目标 < 250)
#   soft_details.teacher_balance_penalty 是否保持可控 (目标 < 160)
#   不破坏: hard_ok True 且 missing_co_teacher 仍为 0
# 后续若需要继续：可再试 SWITCH 10/8, BALANCE 2; 或增加同教师聚类启发(生成/修复阶段)。
CONFIG.setdefault('SOFT_REWARD_SEQUENCE', CONFIG.get('CONSECUTIVE_REWARD', 2))
