"""命令行入口

示例:
  python -m auto_schedule.cli --pop 60 --gen 150 --out 结果.xlsx
  python -m auto_schedule.cli --sweep_scales 5,10,20
"""
from __future__ import annotations

import argparse
from .config import CONFIG
from .ga_engine import run_scheduler


def build_parser():
    p = argparse.ArgumentParser(description='遗传算法排课调度')
    p.add_argument('--pop', type=int, default=CONFIG['DEFAULT_POP'], help='种群大小')
    p.add_argument('--gen', type=int, default=CONFIG['DEFAULT_GEN'], help='迭代代数')
    p.add_argument('--out', type=str, default=CONFIG['DEFAULT_OUTPUT'], help='结果 Excel 路径')
    p.add_argument('--seed', type=int, default=CONFIG['DEFAULT_SEED'], help='随机种子')
    p.add_argument('--verbose', type=int, default=1, help='日志详细级别(0-静默,1-概要,2-详细)')
    p.add_argument('--practical_scale', type=float, help='加权提前罚分系数覆盖(旧参数, 若仍使用旧早置逻辑)')
    p.add_argument('--practical_late', type=float, help='后置阈值(0-1)覆盖(旧参数, 与非理论后置策略相关)')
    p.add_argument('--practical_activate', type=float, help='加权罚分激活区间比例(0-1)')
    p.add_argument('--sweep_scales', type=str, help='逗号分隔多个scale进行快速扫描')
    p.add_argument('--teacher_balance_weight', type=float, help='教师负载均衡罚分系数')
    p.add_argument('--early_stop', type=int, help='早停耐心代数')
    p.add_argument('--launch_manual', action='store_true', help='完成后启动手动界面并载入结果')
    return p


def apply_overrides(args):
    if args.practical_scale is not None:
        CONFIG['PRACTICAL_EARLY_WEIGHT_SCALE'] = args.practical_scale
    if args.practical_late is not None:
        CONFIG['PRACTICAL_LATE_THRESHOLD'] = args.practical_late
    if args.practical_activate is not None:
        CONFIG['PRACTICAL_WEIGHTED_ACTIVATE_RATIO'] = args.practical_activate
    if args.teacher_balance_weight is not None:
        CONFIG['TEACHER_BALANCE_WEIGHT'] = args.teacher_balance_weight
    if args.early_stop is not None:
        CONFIG['EARLY_STOP_PATIENCE'] = args.early_stop


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)
    apply_overrides(args)
    if args.sweep_scales:
        scales = [float(x) for x in args.sweep_scales.split(',') if x.strip()]
        orig_scale = CONFIG['PRACTICAL_EARLY_WEIGHT_SCALE']
        print('[SWEEP] scales=', scales)
        for sc in scales:
            CONFIG['PRACTICAL_EARLY_WEIGHT_SCALE'] = sc
            pop_size = max(6, min(12, args.pop))
            gens = max(2, min(5, args.gen))
            _, met = run_scheduler(pop_size=pop_size, ngen=gens, excel_out=args.out, seed=args.seed, verbose=0)
            sd = met['soft_details']
            print(f"scale={sc} weighted={sd.get('practical_early_weighted_penalty')} early={sd.get('practical_early_penalty')} consec={sd.get('consecutive_reward')} prereq={sd.get('prereq_violation_penalty')} hard_ok={met['hard_ok']}")
        CONFIG['PRACTICAL_EARLY_WEIGHT_SCALE'] = orig_scale
        return
    _, metrics = run_scheduler(pop_size=args.pop, ngen=args.gen, excel_out=args.out, seed=args.seed, verbose=args.verbose)
    if not metrics['hard_ok']:
        print('[ERROR] 排课结果不符合硬性条件')
    else:
        print('[INFO] 排课结果满足硬性条件')
    if args.launch_manual:
        # 尝试启动 streamlit, 传递查询参数
        import subprocess, sys, os
        app_path = os.path.join(os.path.dirname(__file__), '..', 'manual_schedule', 'app_manual.py')
        app_path = os.path.normpath(app_path)
        if not os.path.exists(app_path):
            print('[WARN] 未找到手动排课界面脚本, 跳过启动')
            return
        url_param = f"load_schedule={args.out}"
        print('[INFO] 启动手动界面, 自动载入', args.out)
        # Windows 下使用 python -m streamlit run ... --server.headless=true
        cmd = [sys.executable, '-m', 'streamlit', 'run', app_path, '--', f'--{url_param}']
        try:
            subprocess.Popen(cmd)
        except Exception as e:
            print('[WARN] 启动 Streamlit 失败:', e)


if __name__ == '__main__':  # pragma: no cover
    main()
