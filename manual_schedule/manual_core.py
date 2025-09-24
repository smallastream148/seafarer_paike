import datetime
from dataclasses import dataclass
from typing import List, Dict, Tuple, Optional
import pandas as pd
import sys, pathlib, re

# 现已优先使用 auto_schedule.data_model.TimetableData；若在 manual_schedule 目录直接运行需补 parent 路径。
try:
    from auto_schedule.data_model import TimetableData as _AutoTimetableData  # type: ignore
except ImportError:  # 尝试将父目录加入 sys.path 再试
    try:
        parent = pathlib.Path(__file__).resolve().parents[1]
        if str(parent) not in sys.path:
            sys.path.insert(0, str(parent))
        from auto_schedule.data_model import TimetableData as _AutoTimetableData  # type: ignore
    except ImportError:
        _AutoTimetableData = None  # 最终失败，后续走 legacy 路径

@dataclass
class CourseInfo:
    name: str
    blocks: int
    teachers: List[str]
    is_two: bool
    prerequisites: List[str]
    is_practical: bool
    is_theory: bool = False

@dataclass
class ClassInfo:
    class_id: str
    courses: List[str]
    start_date: datetime.date
    end_date: datetime.date

@dataclass
class PlacedBlock:
    class_id: str
    course: str
    teacher1: str
    teacher2: Optional[str]
    date: datetime.date
    period: int  # 0 上午 1 下午

class TimetableData:
    """手动排课适配数据模型: 封装 auto_schedule 的 TimetableData.

    提供属性:
      courses: 名称 -> CourseInfo
      classes: 班级ID -> ClassInfo
      teacher_unavailable / class_unavailable: 与旧接口保持一致
    """
    def __init__(self, excel_file_path='排课数据.xlsx'):
        import os, glob
        # 若传入了绝对路径且存在，直接使用（避免跨会话串改）
        if excel_file_path and os.path.isabs(excel_file_path) and os.path.exists(excel_file_path):
            self._excel_file_path = excel_file_path
            auto = None
            if _AutoTimetableData:
                try:
                    auto = _AutoTimetableData(excel_file_path)
                except Exception:
                    # 若自动数据模型校验失败（例如不可用时间包含未知班级），回退到兼容旧版的解析方式
                    auto = None
            if auto is None:
                self._legacy_load(excel_file_path)
                return
            self._auto = auto
            # 下方构建 courses/classes 同原逻辑
            self.courses: Dict[str, CourseInfo] = {}
            for name, c in auto.COURSE_DATA.items():
                is_two = bool(c.get('is_two_teacher', False))
                is_practical = is_two
                is_theory = not is_two
                self.courses[name] = CourseInfo(
                    name=name,
                    blocks=int(c['blocks']),
                    teachers=list(c['available_teachers']),
                    is_two=is_two,
                    prerequisites=list(c.get('prerequisites', [])),
                    is_practical=is_practical,
                    is_theory=is_theory,
                )
            self.classes: Dict[str, ClassInfo] = {}
            for cid, info in auto.CLASSES.items():
                self.classes[cid] = ClassInfo(
                    class_id=cid,
                    courses=list(info['courses']),
                    start_date=info['start_date'],
                    end_date=info['end_date'],
                )
            self.teacher_unavailable = auto.TEACHER_UNAVAILABLE_SLOTS
            self.class_unavailable = auto.CLASS_UNAVAILABLE_SLOTS
            return
        # 否则：优先从 SEAFARER_UPLOAD_DIR、/mount/data/uploaded_data、项目根 uploaded_data 查找最新上传文件
        root_dir = os.path.dirname(os.path.dirname(__file__))
        search_dirs = []
        env_dir = os.environ.get('SEAFARER_UPLOAD_DIR')
        if env_dir:
            search_dirs.append(env_dir)
        search_dirs.append('/mount/data/uploaded_data')
        search_dirs.append(os.path.join(root_dir, 'uploaded_data'))
        latest_file = None
        latest_mtime = -1
        for d in search_dirs:
            try:
                if d and os.path.exists(d):
                    files = glob.glob(os.path.join(d, '*.xlsx'))
                    for f in files:
                        m = os.path.getmtime(f)
                        if m > latest_mtime:
                            latest_mtime = m
                            latest_file = f
            except Exception:
                continue
        if latest_file:
            excel_file_path = latest_file
        else:
            # 若是相对路径则锚定到仓库根，避免云端 CWD 与本地不同
            if not os.path.isabs(excel_file_path):
                candidate = os.path.join(root_dir, excel_file_path)
            else:
                candidate = excel_file_path
            excel_file_path = candidate if os.path.exists(candidate) else os.path.join(root_dir, '排课数据.xlsx')

        self._excel_file_path = excel_file_path  # Store the determined path internally

        if _AutoTimetableData is None:
            # Legacy 回退：直接解析 Excel（与早期版本逻辑类似）
            self._legacy_load(excel_file_path)
            return
        # 优先尝试使用自动数据模型；失败则回退到 legacy 解析，避免应用直接崩溃
        try:
            auto = _AutoTimetableData(excel_file_path)
        except Exception:
            self._legacy_load(excel_file_path)
            return
        self._auto = auto
        self.courses: Dict[str, CourseInfo] = {}
        for name, c in auto.COURSE_DATA.items():
            is_two = bool(c.get('is_two_teacher', False))
            # 派生: 双师=实操 非理论; 单师=理论 非实操
            is_practical = is_two
            is_theory = not is_two
            self.courses[name] = CourseInfo(
                name=name,
                blocks=int(c['blocks']),
                teachers=list(c['available_teachers']),
                is_two=is_two,
                prerequisites=list(c.get('prerequisites', [])),
                is_practical=is_practical,
                is_theory=is_theory,
            )
        self.classes: Dict[str, ClassInfo] = {}
        for cid, info in auto.CLASSES.items():
            self.classes[cid] = ClassInfo(
                class_id=cid,
                courses=list(info['courses']),
                start_date=info['start_date'],
                end_date=info['end_date'],
            )
        self.teacher_unavailable = auto.TEACHER_UNAVAILABLE_SLOTS
        self.class_unavailable = auto.CLASS_UNAVAILABLE_SLOTS

    @property
    def excel_file_path(self):
        """Provides read-only access to the excel file path."""
        return self._excel_file_path

    def _legacy_load(self, excel_file_path: str):
        self._auto = None
        self.courses = {}
        self.classes = {}
        self.teacher_unavailable = {}
        self.class_unavailable = {}

        # --- 工具: 表名与列名别名支持 ---
        def pick_sheet(xls: pd.ExcelFile, candidates):
            names = set(xls.sheet_names)
            for n in candidates:
                if n in names:
                    return n
            # 简单模糊匹配（去除空格）
            simplified = {re.sub(r"\s+", "", n): n for n in xls.sheet_names}
            for cand in candidates:
                key = re.sub(r"\s+", "", cand)
                if key in simplified:
                    return simplified[key]
            raise ValueError(f"未找到工作表: {candidates}")

        def normalize_columns(df: pd.DataFrame, alias_map: Dict[str, list], required: list[str]):
            rename = {}
            cols = list(df.columns)
            # 统一去掉首尾空白
            df.columns = [str(c).strip() for c in df.columns]
            for canon, aliases in alias_map.items():
                found = None
                for a in aliases:
                    if a in df.columns:
                        found = a
                        break
                if found:
                    if found != canon:
                        rename[found] = canon
                else:
                    if canon in df.columns:
                        # 已有同名，无需处理
                        pass
            if rename:
                df = df.rename(columns=rename)
            missing = [c for c in required if c not in df.columns]
            if missing:
                raise ValueError(f"缺少必要列: {missing}")
            return df

        # --- 打开 Excel，一次性选择各表 ---
        with pd.ExcelFile(excel_file_path) as xls:
            # 课程数据表
            course_sheet = pick_sheet(xls, ['课程数据', '课程', '课程信息', '课程表'])
            dfc = pd.read_excel(xls, sheet_name=course_sheet)
            dfc = normalize_columns(
                dfc,
                alias_map={
                    '课程名称': ['课程名称', '课程名', '名称', '课程'],
                    'blocks': ['blocks', '课时数', '总块数', '时长'],
                    'available_teachers': ['available_teachers', '教师', '可选教师', '授课教师'],
                    'is_two_teacher': ['is_two_teacher', '双师', '双师课程', '双教师', '是否双师'],
                    'prereq': ['prereq', '先修课', '前置课程', '先决条件'],
                },
                required=['课程名称', 'blocks', 'available_teachers']
            )

            has_prereq = 'prereq' in dfc.columns
            for _, r in dfc.iterrows():
                name = r['课程名称']
                blocks = int(r['blocks'])
                raw_teachers = str(r['available_teachers'])
                teachers = [t.strip() for t in re.split(r'[，,、;；/\\ ]+', raw_teachers) if t and t.strip()]
                raw_two = str(r.get('is_two_teacher', '')).strip().lower()
                is_two = raw_two in {'y', 'yes', 'true', '双', '2', 'two', '是', 'true'}
                prereqs = []
                if has_prereq:
                    prereqs = [p.strip() for p in re.split(r'[，,、;；/\\ ]+', str(r.get('prereq', ''))) if p.strip()]
                is_practical = is_two
                is_theory = (not is_two)
                self.courses[name] = CourseInfo(name, blocks, teachers, is_two, prereqs, is_practical, is_theory)

            # 班级数据表
            class_sheet = pick_sheet(xls, ['班级数据', '班级', '班级信息', '班级表'])
            dfcl = pd.read_excel(xls, sheet_name=class_sheet)
            dfcl = normalize_columns(
                dfcl,
                alias_map={
                    '班级ID': ['班级ID', '班级', '班级名称', '班级名', 'class_id'],
                    'courses': ['courses', '课程列表', '课程', '课程安排'],
                    'start_date': ['start_date', '开始日期', '开始', '起始日期', 'start'],
                    'end_date': ['end_date', '结束日期', '结束', '截止日期', 'end'],
                },
                required=['班级ID', 'courses', 'start_date', 'end_date']
            )
            for _, r in dfcl.iterrows():
                cid = str(r['班级ID']).strip()
                courses = [c.strip() for c in re.split(r'[，,、;；/\\ ]+', str(r['courses'])) if c.strip()]
                sd = pd.to_datetime(r['start_date']).date()
                ed = pd.to_datetime(r['end_date']).date()
                self.classes[cid] = ClassInfo(cid, courses, sd, ed)

            # 教师不可用
            try:
                tea_un_sheet = pick_sheet(xls, ['教师不可用时间', '教师不可用', '教师请假', '教师占用'])
                dft = pd.read_excel(xls, sheet_name=tea_un_sheet)
                dft = normalize_columns(
                    dft,
                    alias_map={
                        '教师姓名': ['教师姓名', '教师', '老师', '教师名', 'teacher'],
                        '日期': ['日期', 'date', 'day'],
                        '时间段': ['时间段', '时段', '上午/下午', 'period'],
                    },
                    required=['教师姓名', '日期', '时间段']
                )
                time_map = {'上午': 0, '下午': 1, 'am': 0, 'pm': 1}
                for _, r in dft.iterrows():
                    t = str(r['教师姓名']).strip()
                    date = pd.to_datetime(r['日期']).date()
                    p = time_map.get(str(r['时间段']).strip().lower())
                    if p is None:
                        continue
                    self.teacher_unavailable.setdefault(t, set()).add((date, p))
            except Exception:
                pass

            # 班级不可用
            try:
                cls_un_sheet = pick_sheet(xls, ['班级不可用时间', '班级不可用', '班级占用'])
                dfu = pd.read_excel(xls, sheet_name=cls_un_sheet)
                dfu = normalize_columns(
                    dfu,
                    alias_map={
                        '班级ID': ['班级ID', '班级', '班级名称', '班级名', 'class_id'],
                        '日期': ['日期', 'date', 'day'],
                        '时间段': ['时间段', '时段', '上午/下午', 'period'],
                    },
                    required=['班级ID', '日期', '时间段']
                )
                time_map = {'上午': 0, '下午': 1, 'am': 0, 'pm': 1}
                for _, r in dfu.iterrows():
                    cid = str(r['班级ID']).strip()
                    date = pd.to_datetime(r['日期']).date()
                    p = time_map.get(str(r['时间段']).strip().lower())
                    if p is None:
                        continue
                    self.class_unavailable.setdefault(cid, set()).add((date, p))
            except Exception:
                pass

    def iter_class_slots(self, class_id: str):
        info = self.classes[class_id]
        days = (info.end_date - info.start_date).days + 1
        for d in range(days):
            date = info.start_date + datetime.timedelta(days=d)
            for p in (0, 1):
                if class_id in self.class_unavailable and (date, p) in self.class_unavailable[class_id]:
                    continue
                yield date, p

class ManualScheduler:
    def __init__(self, data: TimetableData):
        self.data = data
        self.placed: List[PlacedBlock] = []
        self.history: List[Tuple[str, PlacedBlock]] = []  # ('add'/'del', block)

    # --- 硬性校验 ---
    def check_hard_violation(self, block: PlacedBlock) -> List[str]:
        errs = []
        # 课程存在性
        if block.course not in self.data.courses:
            errs.append('未知课程')
            return errs
        # 班级存在性
        if block.class_id not in self.data.classes:
            errs.append('未知班级')
            return errs
        cinfo = self.data.courses[block.course]
        # 教师合法
        if block.teacher1 not in cinfo.teachers:
            errs.append('教师1不在课程可选列表')
        if cinfo.is_two:
            if not block.teacher2:
                errs.append('双师缺第二教师')
            elif block.teacher2 == block.teacher1:
                errs.append('双师教师重复')
            elif block.teacher2 not in cinfo.teachers:
                errs.append('教师2不在课程可选列表')
        else:
            if block.teacher2:
                errs.append('单师课程不应有第二教师')
        # 时间合法范围和不可用
        cls = self.data.classes[block.class_id]
        if not (cls.start_date <= block.date <= cls.end_date):
            errs.append('日期超出班级范围')
        if block.class_id in self.data.class_unavailable and (block.date, block.period) in self.data.class_unavailable[block.class_id]:
            errs.append('班级该时段不可用')
        if block.teacher1 in self.data.teacher_unavailable and (block.date, block.period) in self.data.teacher_unavailable[block.teacher1]:
            errs.append('教师1该时段不可用')
        if block.teacher2 and block.teacher2 in self.data.teacher_unavailable and (block.date, block.period) in self.data.teacher_unavailable[block.teacher2]:
            errs.append('教师2该时段不可用')
        # 冲突：同时间教师 / 班级
        for b in self.placed:
            if b.date == block.date and b.period == block.period:
                if b.class_id == block.class_id:
                    errs.append('班级时间冲突')
                if b.teacher1 == block.teacher1 or (block.teacher2 and (b.teacher1 == block.teacher2)) or \
                   (b.teacher2 and (b.teacher2 == block.teacher1 or (block.teacher2 and b.teacher2 == block.teacher2))):
                    errs.append('教师时间冲突')
        # 已排块数超限
        existing = sum(1 for b in self.placed if b.class_id==block.class_id and b.course==block.course)
        if existing >= cinfo.blocks:
            errs.append('课程块数已达上限')
        # 理论课教师一致性
        if cinfo.is_theory:
            prev = [b for b in self.placed if b.class_id==block.class_id and b.course==block.course]
            if prev:
                base_t = prev[0].teacher1
                if block.teacher1 != base_t:
                    errs.append('理论课教师需保持一致')
        return errs

    def add_block(self, block: PlacedBlock) -> Tuple[bool, List[str]]:
        errs = self.check_hard_violation(block)
        if errs:
            return False, errs
        self.placed.append(block)
        self.history.append(('add', block))
        return True, []

    def remove_last(self) -> bool:
        if not self.history:
            return False
        act, blk = self.history.pop()
        if act == 'add':
            # 从 placed 删除最后对应对象
            for i in range(len(self.placed)-1, -1, -1):
                if self.placed[i] is blk:
                    self.placed.pop(i)
                    break
        elif act == 'del':
            # 撤销删除 -> 重新加入
            self.placed.append(blk)
        return True

    def delete_block(self, block_index: int) -> bool:
        """按 index 删除 placed 中的块, 并记录以便撤销。"""
        if 0 <= block_index < len(self.placed):
            blk = self.placed.pop(block_index)
            self.history.append(('del', blk))
            return True
        return False

    def remaining_blocks(self, class_id: str, course: str) -> int:
        cinfo = self.data.courses[course]
        # 计数逻辑：
        # 单师课程: 每个已放置块计 1。
        # 双师课程: 仅在该块拥有两个不同教师时才计 1 (缺第二教师视为未完成临时块)。
        used = 0
        for b in self.placed:
            if b.class_id == class_id and b.course == course:
                if cinfo.is_two:
                    if b.teacher1 and b.teacher2 and b.teacher1 != b.teacher2:
                        used += 1
                else:
                    used += 1
        return cinfo.blocks - used

    def export_rows(self):
        period_name = {0:'上午',1:'下午'}
        rows = []
        for b in self.placed:
            rows.append({'班级ID': b.class_id,'课程': b.course,'教师1': b.teacher1,'教师2': b.teacher2 or '',
                         '日期': b.date,'时段': period_name[b.period]})
        return rows

    def supplement_second_teacher(self, block_index: int, teacher2: str) -> Tuple[bool, str]:
        """为指定 index 的双师块补第二教师。
        规则:
          - 该块对应课程必须标记 is_two
          - 当前 teacher2 为空或无效
          - teacher2 在课程可选教师列表中且 != teacher1
          - 补齐后需再次通过基本时间冲突校验 (教师占用 & 不可用)；若冲突恢复原状返回 False
        """
        if not (0 <= block_index < len(self.placed)):
            return False, '索引不存在'
        blk = self.placed[block_index]
        cinfo = self.data.courses.get(blk.course)
        if not cinfo:
            return False, '课程不存在'
        if not cinfo.is_two:
            return False, '该课程非双师'
        if blk.teacher2 and blk.teacher2 != blk.teacher1:
            return False, '该块已具备两个教师'
        if teacher2 == blk.teacher1:
            return False, '第二教师需不同'
        if teacher2 not in cinfo.teachers:
            return False, '教师不在可选列表'
        # 冲突与不可用校验
        if teacher2 in self.data.teacher_unavailable and (blk.date, blk.period) in self.data.teacher_unavailable[teacher2]:
            return False, '教师该时段不可用'
        for i, other in enumerate(self.placed):
            if i == block_index:
                continue
            if other.date == blk.date and other.period == blk.period:
                if other.teacher1 == teacher2 or (other.teacher2 and other.teacher2 == teacher2):
                    return False, '教师该时段已被占用'
        # 通过
        old_teacher2 = blk.teacher2
        blk.teacher2 = teacher2
        self.history.append(('add', blk))  # 记录一条操作, 仍用 'add' 方便撤销 (撤销时移除最后变更)
        return True, '补齐成功'
