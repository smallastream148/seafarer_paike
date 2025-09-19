import datetime
import re
import pandas as pd
import os
import glob
from .config import CONFIG

class TimetableData:
    def __init__(self, excel_file_path='排课数据.xlsx'):
        # 优先从可写/云端目录查找最新上传文件，其次项目根 uploaded_data，最后落回默认文件
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
        self.excel_file_path = latest_file or excel_file_path
        
        self.COURSE_DATA = self._load_course_data()
        self.CLASSES = self._load_classes_data()
        self.TEACHER_UNAVAILABLE_SLOTS = self._load_teacher_availability()
        self.CLASS_UNAVAILABLE_SLOTS = self._load_class_availability()
        all_teachers_from_courses = set(t for c in self.COURSE_DATA.values() for t in c['available_teachers'])
        all_teachers_from_availability = set(self.TEACHER_UNAVAILABLE_SLOTS.keys())
        all_teachers = all_teachers_from_courses.union(all_teachers_from_availability)
        self.TEACHERS = {name: {} for name in all_teachers}
        self.TIMES_PER_DAY = ['上午','下午']
        self.validate()
        self.CLASS_SLOT_CACHE = self._precompute_class_slots()

    def _precompute_class_slots(self):
        cache = {}
        for class_id, info in self.CLASSES.items():
            days = (info['end_date'] - info['start_date']).days + 1
            indices = []
            for d in range(days):
                for p in range(2):
                    date = info['start_date'] + datetime.timedelta(days=d)
                    if class_id in self.CLASS_UNAVAILABLE_SLOTS and (date,p) in self.CLASS_UNAVAILABLE_SLOTS[class_id]:
                        continue
                    indices.append(d*2+p)
            cache[class_id] = indices
        return cache

    def _load_course_data(self):
        df = pd.read_excel(self.excel_file_path, sheet_name='课程数据')
        required = {'课程名称', 'blocks', 'available_teachers'}
        missing = required - set(df.columns)
        if missing:
            raise ValueError(f"课程数据缺列:{missing}")
        has_prereq = 'prereq' in df.columns
        # 简化：仅保留 is_two_teacher 列。派生规则：
        #   is_two_teacher = True => 视为 实操课 (practical), 非理论
        #   is_two_teacher = False => 视为 理论课 (theory), 非实操
        def norm_two(v):
            if isinstance(v, (int, float)):
                return int(v) == 2
            if isinstance(v, str):
                return v.strip().lower() in {'y', 'yes', 'true', '双', '2', 'two'}
            return bool(v)
        course_data: dict = {}
        for _, row in df.iterrows():
            name = row['课程名称']
            blocks = int(row['blocks'])
            if blocks <= 0:
                raise ValueError(f"课程 {name} blocks 必须>0")
            raw_teachers = str(row['available_teachers'])
            teachers = [t.strip() for t in re.split(r'[，,、;；/\\ ]+', raw_teachers) if t and t.strip()]
            if not teachers:
                raise ValueError(f"课程 {name} 缺少教师")
            prereqs = []
            if has_prereq:
                prereqs = [p.strip() for p in str(row.get('prereq', '')).split(',') if p.strip()]
            is_two = norm_two(row.get('is_two_teacher', ''))
            is_practical = is_two
            is_theory = (not is_two)
            course_data[name] = {
                'blocks': blocks,
                'available_teachers': teachers,
                'is_two_teacher': is_two,
                'prerequisites': prereqs,
                'is_practical': is_practical,
                'is_theory': is_theory,
            }
        return course_data

    def _load_classes_data(self):
        df = pd.read_excel(self.excel_file_path, sheet_name='班级数据')
        required = {'班级ID','courses','start_date','end_date'}
        missing = required - set(df.columns)
        if missing: raise ValueError(f"班级数据缺列:{missing}")
        out = {}
        for _,row in df.iterrows():
            cid = str(row['班级ID'])
            courses = [c.strip() for c in str(row['courses']).split(',') if c.strip()]
            if not courses: raise ValueError(f"班级 {cid} 无课程")
            sd = pd.to_datetime(row['start_date']).date()
            ed = pd.to_datetime(row['end_date']).date()
            if sd>ed: raise ValueError(f"班级 {cid} 日期非法")
            out[cid] = {'courses':courses,'start_date':sd,'end_date':ed}
        return out

    def _load_teacher_availability(self):
        try:
            df = pd.read_excel(self.excel_file_path, sheet_name='教师不可用时间')
        except Exception:
            return {}
        need={'教师姓名','日期','时间段'}
        if not need.issubset(df.columns):
            raise ValueError('教师不可用时间 缺列')
        out={}
        mp={'上午':0,'下午':1}
        for _,r in df.iterrows():
            t=str(r['教师姓名']).strip(); date=pd.to_datetime(r['日期']).date(); p=mp.get(str(r['时间段']).strip())
            if p is None: continue
            out.setdefault(t,set()).add((date,p))
        return out

    def _load_class_availability(self):
        try:
            df = pd.read_excel(self.excel_file_path, sheet_name='班级不可用时间')
        except Exception:
            return {}
        need={'班级ID','日期','时间段'}
        if not need.issubset(df.columns): raise ValueError('班级不可用时间 缺列')
        out={}
        mp={'上午':0,'下午':1}
        for _,r in df.iterrows():
            cid=str(r['班级ID']).strip(); date=pd.to_datetime(r['日期']).date(); p=mp.get(str(r['时间段']).strip())
            if p is None: continue
            out.setdefault(cid,set()).add((date,p))
        return out

    def validate(self):
        for cid,info in self.CLASSES.items():
            for c in info['courses']:
                if c not in self.COURSE_DATA:
                    raise ValueError(f"班级 {cid} 引用不存在课程 {c}")
        # 双师课程必须至少提供2个不同教师
        for cname, cinfo in self.COURSE_DATA.items():
            if cinfo.get('is_two_teacher') and len(set(cinfo['available_teachers'])) < 2:
                raise ValueError(f"课程 {cname} 标记双师但教师数量不足2")
        course_teachers = set(t for v in self.COURSE_DATA.values() for t in v['available_teachers'])
        extra = set(self.TEACHER_UNAVAILABLE_SLOTS.keys()) - course_teachers
        if extra:
            print(f"[警告] 不可用教师未在课程中: {sorted(extra)}")
        for cid in self.CLASS_UNAVAILABLE_SLOTS.keys():
            if cid not in self.CLASSES: raise ValueError(f"不可用时间未知班级 {cid}")
        for cid,info in self.CLASSES.items():
            days=(info['end_date']-info['start_date']).days+1
            capacity=days*2
            if cid in self.CLASS_UNAVAILABLE_SLOTS:
                capacity -= len(self.CLASS_UNAVAILABLE_SLOTS[cid])
            demand=sum(self.COURSE_DATA[c]['blocks'] for c in info['courses'])
            if demand>capacity:
                raise ValueError(f"班级 {cid} 需求 {demand} > 容量 {capacity}")
        print(f"[校验通过] 班级:{len(self.CLASSES)} 课程:{len(self.COURSE_DATA)} 教师:{len(self.TEACHERS)}")
