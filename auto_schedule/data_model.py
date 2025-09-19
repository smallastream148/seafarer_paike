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
        def _score_excel(path: str) -> tuple[int, float]:
            """对 Excel 文件进行打分以判断是否为输入源：
            +2: 包含课程输入表（课程数据/课程...）
            +1: 包含班级输入表（班级数据/班级...）
            -5: 包含导出表（排课明细）或文件名疑似结果
            次级排序使用修改时间（越新越优）。
            返回 (score, mtime)
            """
            score = 0
            try:
                import pandas as _pd
                with _pd.ExcelFile(path, engine='openpyxl') as xls:
                    names = set(xls.sheet_names)
                norm = lambda s: str(s).strip().lower().replace(' ', '')
                ns = {norm(n) for n in names}
                # 输入候选
                course_candidates = {'课程数据','课程','课程表','courses','course','课程设置'}
                class_candidates = {'班级数据','班级','classes','class','班级表'}
                if any(norm(c) in ns for c in course_candidates):
                    score += 2
                if any(norm(c) in ns for c in class_candidates):
                    score += 1
                # 导出特征
                if any(norm(x) in ns for x in {'排课明细','教师课时','课程进度'}):
                    score -= 5
            except Exception:
                # 无法读取，降低优先级
                score -= 10
            # 文件名暗示结果
            base = os.path.basename(path)
            if '__ui_auto_result' in base or base.startswith('schedule_'):
                score -= 5
            try:
                m = os.path.getmtime(path)
            except Exception:
                m = 0.0
            return score, m

        latest_file = None
        best_score = -9999
        best_mtime = -1
        for d in search_dirs:
            try:
                if d and os.path.exists(d):
                    files = glob.glob(os.path.join(d, '*.xlsx'))
                    for f in files:
                        s, m = _score_excel(f)
                        if (s > best_score) or (s == best_score and m > best_mtime):
                            best_score, best_mtime, latest_file = s, m, f
            except Exception:
                continue
        # 选择最终 Excel 路径：优先最新上传；否则将相对路径锚定到仓库根，避免云端 CWD 差异
        if latest_file:
            self.excel_file_path = latest_file
        else:
            # 若传入是相对路径，则基于项目根拼接；若没有则回退到根目录的同名文件
            if not os.path.isabs(excel_file_path):
                candidate = os.path.join(root_dir, excel_file_path)
            else:
                candidate = excel_file_path
            self.excel_file_path = candidate if os.path.exists(candidate) else os.path.join(root_dir, '排课数据.xlsx')
        
        try:
            self.COURSE_DATA = self._load_course_data()
        except ValueError as e:
            # 如果当前文件看起来是导出结果，给出更明确的指引
            try:
                import pandas as _pd
                with _pd.ExcelFile(self.excel_file_path, engine='openpyxl') as _xls:
                    _names = set(_xls.sheet_names)
                def _n(s): return str(s).strip().lower().replace(' ','')
                ns = {_n(n) for n in _names}
                if any(x in ns for x in {'排课明细','教师课时','课程进度'}):
                    raise ValueError(f"检测到正在尝试从导出结果文件导入({os.path.basename(self.excel_file_path)})；请上传或选择包含原始输入的Excel（需包含‘课程数据/班级数据’工作表）。原错误: {e}")
            except Exception:
                pass
            raise
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

    def _read_sheet(self, sheet_candidates, col_alias: dict, required_keys: set):
        """从 Excel 中解析符合条件的数据表。
        策略：
        1) 一次性读取全部工作表（sheet_name=None）避免名字匹配差异引发错误；
        2) 先按候选名进行规范化匹配（忽略空格/大小写），直接返回该表；
        3) 否则遍历所有表，按别名重命名后检测是否包含必需列，命中则返回；
        4) 否则报错并列出实际可用的表名。
        """
        def norm(s: str) -> str:
            return str(s).strip().lower().replace(' ', '')
        cand_norm = {norm(c): c for c in sheet_candidates}
        try:
            with pd.ExcelFile(self.excel_file_path, engine='openpyxl') as xls:
                names = list(xls.sheet_names)
                # 一次性读取所有表，避免逐表读取时的名字不一致问题
                all_sheets = pd.read_excel(xls, sheet_name=None, engine='openpyxl')
        except Exception as e:
            raise ValueError(f"无法打开Excel文件: {self.excel_file_path}; 错误: {e}")
        # 1) 候选名直接匹配
        for name in names:
            if norm(name) in cand_norm and name in all_sheets:
                return all_sheets[name]
        # 2) 遍历所有表，按别名重命名后检查必需列
        for name, tmp in all_sheets.items():
            try:
                cols = set(tmp.columns)
                ren = {}
                for std, alts in col_alias.items():
                    if std in cols:
                        continue
                    for a in alts:
                        if a in cols:
                            ren[a] = std
                            break
                if ren:
                    tmp2 = tmp.rename(columns=ren)
                    cols2 = set(tmp2.columns)
                else:
                    tmp2 = tmp
                    cols2 = cols
                if required_keys.issubset(cols2):
                    return tmp2
            except Exception:
                continue
        raise ValueError(f"未找到包含所需列的工作表；需要列: {sorted(required_keys)}；候选Sheet: {sheet_candidates}；实际: {names}")

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
        # 先尝试根据 sheet 候选与必需列自动识别
        alias = {
            '课程名称': ['课程', '课程名', 'course', '课程名称'],
            'blocks': ['课时', '块数', 'blocks', '总块数'],
            'available_teachers': ['可授教师', '教师', '教师名单', 'teachers', '可选教师'],
            'is_two_teacher': ['双师', '是否双师', '双师课程', 'two_teachers'],
            'prereq': ['先修', '先修课', '前置课程', 'prereq']
        }
        required = {'课程名称', 'blocks', 'available_teachers'}
        df = self._read_sheet(
            sheet_candidates=['课程数据', '课程', '课程表', 'courses', 'course', '课程设置'],
            col_alias=alias,
            required_keys=required
        )
        # 列同义映射，增强兼容性
        cols = set(df.columns)
        ren = {}
        for std, alts in alias.items():
            if std in cols:
                continue
            for a in alts:
                if a in cols:
                    ren[a] = std
                    break
        if ren:
            df = df.rename(columns=ren)
            cols = set(df.columns)
        missing = required - cols
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
            teachers = [t.strip() for t in re.split(r'[，,、;；/\\\s]+', raw_teachers) if t and t.strip()]
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
        alias = {
            '班级ID': ['班级', '班级编号', 'class_id', '班级名称'],
            'courses': ['课程列表', '课程', '课程安排', 'course_list'],
            'start_date': ['开始日期', '开课日期', 'start', '开始'],
            'end_date': ['结束日期', '结课日期', 'end', '结束']
        }
        required = {'班级ID','courses','start_date','end_date'}
        df = self._read_sheet(
            sheet_candidates=['班级数据', '班级', 'classes', 'class', '班级表'],
            col_alias=alias,
            required_keys=required
        )
        cols = set(df.columns)
        ren = {}
        for std, alts in alias.items():
            if std in cols:
                continue
            for a in alts:
                if a in cols:
                    ren[a] = std
                    break
        if ren:
            df = df.rename(columns=ren)
            cols = set(df.columns)
        missing = required - cols
        if missing:
            raise ValueError(f"班级数据缺列:{missing}")
        out = {}
        for _,row in df.iterrows():
            cid = str(row['班级ID'])
            raw_courses = str(row['courses'])
            courses = [c.strip() for c in re.split(r'[，,、;；/\\\s]+', raw_courses) if c.strip()]
            if not courses: raise ValueError(f"班级 {cid} 无课程")
            sd = pd.to_datetime(row['start_date']).date()
            ed = pd.to_datetime(row['end_date']).date()
            if sd>ed: raise ValueError(f"班级 {cid} 日期非法")
            out[cid] = {'courses':courses,'start_date':sd,'end_date':ed}
        return out

    def _load_teacher_availability(self):
        try:
            alias = {
            '教师姓名': ['教师', '老师', 'teacher', '教师名'],
            '日期': ['date', 'day'],
            '时间段': ['时段', '节次', 'period']
        }
            required = {'教师姓名','日期','时间段'}
            df = self._read_sheet(
                sheet_candidates=['教师不可用时间', '教师不可用', '教师不在', 'teacher_unavailable', '老师不可用'],
                col_alias=alias,
                required_keys=required
            )
        except Exception:
            return {}
        cols = set(df.columns)
        ren = {}
        for std, alts in alias.items():
            if std in cols:
                continue
            for a in alts:
                if a in cols:
                    ren[a] = std
                    break
        if ren:
            df = df.rename(columns=ren)
            cols = set(df.columns)
        need={'教师姓名','日期','时间段'}
        if not need.issubset(cols):
            raise ValueError('教师不可用时间 缺列')
        out={}
        mp={'上午':0,'下午':1,'AM':0,'PM':1,'0':0,'1':1,0:0,1:1}
        for _,r in df.iterrows():
            t=str(r['教师姓名']).strip(); date=pd.to_datetime(r['日期']).date(); p=mp.get(str(r['时间段']).strip())
            if p is None: continue
            out.setdefault(t,set()).add((date,p))
        return out

    def _load_class_availability(self):
        try:
            alias = {
            '班级ID': ['班级', '班级编号', 'class_id', '班级名称'],
            '日期': ['date', 'day'],
            '时间段': ['时段', '节次', 'period']
        }
            required = {'班级ID','日期','时间段'}
            df = self._read_sheet(
                sheet_candidates=['班级不可用时间', '班级不可用', '班级不在', 'class_unavailable'],
                col_alias=alias,
                required_keys=required
            )
        except Exception:
            return {}
        cols = set(df.columns)
        ren = {}
        for std, alts in alias.items():
            if std in cols:
                continue
            for a in alts:
                if a in cols:
                    ren[a] = std
                    break
        if ren:
            df = df.rename(columns=ren)
            cols = set(df.columns)
        need={'班级ID','日期','时间段'}
        if not need.issubset(cols): raise ValueError('班级不可用时间 缺列')
        out={}
        mp={'上午':0,'下午':1,'AM':0,'PM':1,'0':0,'1':1,0:0,1:1}
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
