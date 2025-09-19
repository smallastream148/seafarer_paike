import streamlit as st
import datetime
from pathlib import Path
import os
import shutil
import pandas as pd
import tempfile

# 兼容包/脚本两种运行方式
try:
    from manual_schedule.manual_state import ManualSession
except ModuleNotFoundError:
    from manual_state import ManualSession

st.set_page_config(page_title="船员培训智能排课系统", layout="centered", page_icon="⚓️")

# ============ 初始化 ============
@st.cache_resource
def get_session():
    """获取或创建会话状态对象"""
    return ManualSession()

session = get_session()
data = session.data
ASSET_DIR = Path(__file__).parent / 'assets'
ROOT_DIR = Path(__file__).resolve().parents[1]

def get_writable_upload_dir() -> Path:
    """选择一个可写的上传目录，兼容本地与云端。
    优先顺序：环境变量 SEAFARER_UPLOAD_DIR -> /mount/data/uploaded_data -> 项目根 uploaded_data -> 临时目录
    """
    candidates = []
    env_dir = os.environ.get('SEAFARER_UPLOAD_DIR')
    if env_dir:
        candidates.append(Path(env_dir))
    # Streamlit Cloud 持久化目录
    candidates.append(Path('/mount/data/uploaded_data'))
    # 项目根目录（本地）
    candidates.append(ROOT_DIR / 'uploaded_data')
    # 系统临时目录
    candidates.append(Path(tempfile.gettempdir()) / 'uploaded_data')
    for p in candidates:
        try:
            p.mkdir(parents=True, exist_ok=True)
            if os.access(str(p), os.W_OK):
                return p
        except Exception:
            continue
    # 兜底
    return ROOT_DIR / 'uploaded_data'

# ============ 侧边栏 (数据管理) ============
# 使用 session_state 来防止文件上传后无限循环刷新
if "file_uploader_key" not in st.session_state:
    st.session_state["file_uploader_key"] = 0

with st.sidebar:
    st.header("⚙️ 数据管理")
    
    upload_dir = get_writable_upload_dir()

    # 1. 上传数据
    uploaded_file = st.file_uploader(
        "上传新的排课数据",
        type=['xlsx'],
        help="上传后将自动覆盖现有上传数据并刷新页面",
        key=st.session_state["file_uploader_key"],
    )
    if uploaded_file is not None:
        if not upload_dir.exists():
            upload_dir.mkdir(parents=True, exist_ok=True)
        
        # 为了确保只使用最新的文件，先清空目录
        for f in upload_dir.glob('*.xlsx'):
            f.unlink()

        file_path = upload_dir / uploaded_file.name
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        # 通过增加key的值来重置file_uploader，避免循环
        st.session_state["file_uploader_key"] += 1
        st.toast(f"✅ 文件 '{uploaded_file.name}' 已上传。正在刷新...", icon="🎉")
        
        # 清除缓存并重新运行以加载新数据
        get_session.clear()
        st.rerun()

    # 2. 预览数据
    with st.expander("📄 预览当前数据", expanded=True):
        # 更健壮的文件选择：优先 data.excel_file_path，其次 uploaded_data 最新文件，最后根目录默认文件
        try:
            active_file = getattr(data, 'excel_file_path', None)
            if not active_file or not os.path.exists(active_file):
                latest = None
                if upload_dir.exists():
                    files = sorted(upload_dir.glob('*.xlsx'), key=lambda p: p.stat().st_mtime, reverse=True)
                    latest = str(files[0]) if files else None
                # 额外检查 /mount/data/uploaded_data（云端场景）
                if not latest and Path('/mount/data/uploaded_data').exists():
                    files = sorted(Path('/mount/data/uploaded_data').glob('*.xlsx'), key=lambda p: p.stat().st_mtime, reverse=True)
                    latest = str(files[0]) if files else None
                active_file = latest or str(ROOT_DIR / '排课数据.xlsx')
        except Exception:
            active_file = str(ROOT_DIR / '排课数据.xlsx')

        st.caption(f"当前使用文件: `{os.path.basename(active_file)}`")
        
        try:
            # 使用 with 语句确保文件在读取后被关闭
            with pd.ExcelFile(active_file) as xls:
                sheet_names = xls.sheet_names
                selected_sheet = st.selectbox("选择工作表预览", sheet_names, key="sheet_preview")
                if selected_sheet:
                    df = pd.read_excel(xls, sheet_name=selected_sheet)
                    st.dataframe(df.head(5), height=200)
        except FileNotFoundError:
            st.error("未找到数据文件，请在左侧上传 Excel 或将 `排课数据.xlsx` 放到仓库根目录。")
        except Exception as e:
            st.error(f"无法预览文件: {e}")

    # 2.5 环境与数据诊断
    with st.expander("🧪 环境与数据诊断", expanded=False):
        try:
            st.caption(f"cwd: {os.getcwd()}")
            st.caption(f"ROOT_DIR: {ROOT_DIR}")
            st.caption(f"上传目录: {upload_dir}")
            st.caption(f"环境变量 SEAFARER_UPLOAD_DIR: {os.environ.get('SEAFARER_UPLOAD_DIR')}")
            # 显示当前会话/数据层选用的文件
            active = getattr(data, 'excel_file_path', None)
            st.caption(f"data.excel_file_path: {active}")
            # 版本信息
            try:
                import pandas as _pd
                import openpyxl as _ox
                st.caption(f"pandas: {_pd.__version__}")
                st.caption(f"openpyxl: {_ox.__version__}")
            except Exception:
                pass
            # 罗列候选路径中的 xlsx
            candidates = []
            try:
                if upload_dir.exists():
                    candidates.extend([str(p) for p in upload_dir.glob('*.xlsx')])
            except Exception:
                pass
            try:
                mdir = Path('/mount/data/uploaded_data')
                if mdir.exists():
                    candidates.extend([str(p) for p in mdir.glob('*.xlsx')])
            except Exception:
                pass
            st.caption("检测到的Excel:")
            if candidates:
                for p in sorted(set(candidates)):
                    st.text(p)
            else:
                st.text("(未发现上传的 .xlsx 文件)")
            # 针对当前 active 文件，检查关键 sheet 的列与行数
            if active and os.path.exists(active):
                try:
                    with pd.ExcelFile(active) as _xls:
                        _sheets = set(_xls.sheet_names)
                        st.caption(f"工作表: {sorted(_sheets)}")
                    # 课程数据
                    req_course = {'课程名称', 'blocks', 'available_teachers'}
                    try:
                        _dc = pd.read_excel(active, sheet_name='课程数据')
                        st.caption(f"课程数据: {len(_dc)}行, 列={list(_dc.columns)}")
                        miss = req_course - set(_dc.columns)
                        if miss:
                            st.error(f"课程数据缺少列: {sorted(miss)}")
                    except Exception as e:
                        st.error(f"读取'课程数据'失败: {e}")
                    # 班级数据
                    req_class = {'班级ID','courses','start_date','end_date'}
                    try:
                        _dl = pd.read_excel(active, sheet_name='班级数据')
                        st.caption(f"班级数据: {len(_dl)}行, 列={list(_dl.columns)}")
                        miss = req_class - set(_dl.columns)
                        if miss:
                            st.error(f"班级数据缺少列: {sorted(miss)}")
                    except Exception as e:
                        st.error(f"读取'班级数据'失败: {e}")
                except Exception as e:
                    st.error(f"诊断读取失败: {e}")
        except Exception as e:
            st.error(f"诊断信息获取失败: {e}")

    # 3. 清除数据（按钮始终显示：无数据时禁用；清除后通过 session_state 给出提示）
    try:
        upload_dir.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass
    has_uploaded = any(upload_dir.glob('*.xlsx'))
    # 显示上一次清除后的成功提示
    if st.session_state.get('just_cleared'):
        st.success("✅ 已清除所有上传数据，已恢复默认数据。")
        del st.session_state['just_cleared']
    if st.button("🗑️ 清除上传数据", help="删除所有上传的数据，恢复使用默认数据", disabled=not has_uploaded):
        try:
            # 仅删除上传的 xlsx 文件，保留目录，避免按钮消失
            for item in upload_dir.glob('*.xlsx'):
                if item.is_file():
                    item.unlink()
            st.session_state['just_cleared'] = True
            get_session.clear()
            st.rerun()
        except PermissionError as e:
            st.error(f"清除失败：文件可能被占用。请关闭相关程序后重试。\n错误: {e}")
        except Exception as e:
            st.error(f"清除时发生未知错误: {e}")

# 初始化颜色映射
if 'course_color_map' not in st.session_state:
    st.session_state['course_color_map'] = {c: i % 10 for i, c in enumerate(data.courses.keys())}
course_color_map = st.session_state['course_color_map']

# ============ 工具函数 ============
def inject_css():
    """注入外部CSS样式"""
    try:
        # 基础样式
        with open(ASSET_DIR / 'style.css', 'r', encoding='utf-8') as f:
            base_css = f.read()
        st.markdown(f"<style>{base_css}</style>", unsafe_allow_html=True)
        
        # 暗黑模式
        if st.session_state.get('dark_mode', False):
            with open(ASSET_DIR / 'style_dark.css', 'r', encoding='utf-8') as f:
                dark_css = f.read()
            st.markdown(f"<style>{dark_css}</style>", unsafe_allow_html=True)
            st.markdown('<script>document.body.classList.add("dark-mode");</script>', unsafe_allow_html=True)
        else:
            st.markdown('<script>document.body.classList.remove("dark-mode");</script>', unsafe_allow_html=True)
    except Exception as e:
        st.error(f"样式加载失败: {e}")

def force_rerun():
    """强制 Streamlit 重新运行"""
    st.rerun()

def render_header():
    """渲染页面头部"""
    st.markdown("""
        <div style='text-align: center; padding: 1rem 0; border-bottom: 2px solid #e1e4e8;'>
            <h1 style='color: #2c3e50; margin: 0;'>⚓️ 船员培训智能排课系统</h1>
            <p style='color: #7f8c8d; margin: 0.5rem 0 0 0;'>智能排课，高效管理</p>
        </div>
    """, unsafe_allow_html=True)

def render_ga_section():
    """渲染自动排课部分"""
    with st.expander("🤖 自动排课 (遗传算法)", expanded=False):
        st.info("使用遗传算法自动生成完整排课方案，结果将覆盖当前已排课程")

        # 上次运行回显
        last = st.session_state.get('ga_last')
        if last:
            col_a, col_b, col_c = st.columns([2,1,2])
            with col_a:
                st.success(f"上次导入: {last.get('imported', 0)} 块")
            with col_b:
                met = last.get('metrics', {}) or {}
                st.metric("适应度", f"{met.get('total_fitness', 0):.2f}")
            with col_c:
                st.caption(f"结果文件: {last.get('path','-')}")
            with st.expander('📄 查看导出文件(排课明细)摘要', expanded=False):
                try:
                    p = last.get('path')
                    if p:
                        with pd.ExcelFile(p) as _xls:
                            if '排课明细' in _xls.sheet_names:
                                _df = pd.read_excel(_xls, sheet_name='排课明细')
                                st.caption(f"排课明细: {len(_df)} 行")
                                st.dataframe(_df.head(10), height=220, use_container_width=True)
                            else:
                                st.warning('导出文件中未找到“排课明细”工作表')
                except Exception as e:
                    st.error(f"预览导出文件失败: {e}")

        cols = st.columns(5)
        pop = cols[0].number_input('种群大小', 10, 500, 60, 10)
        gen = cols[1].number_input('迭代代数', 50, 2000, 200, 50)
        seed = cols[2].number_input('随机种子', 0, 999999, 42, 1)
        verbose = cols[3].selectbox('日志级别', [0, 1, 2], index=1)

        if cols[4].button('🚀 开始运行', type='primary', use_container_width=True):
            with st.spinner('正在运行遗传算法...'):
                try:
                    # 导出路径改为可写上传目录，避免云端根目录不可写
                    auto_result_path = str(get_writable_upload_dir() / '__ui_auto_result.xlsx')
                    # 兼容旧版引擎：在运行前将当前数据文件同步到项目根的默认文件名
                    try:
                        src_excel = getattr(data, 'excel_file_path', None)
                        if src_excel and os.path.exists(src_excel):
                            default_excel = str(ROOT_DIR / '排课数据.xlsx')
                            # 若源文件与目标不同路径，则拷贝覆盖
                            if os.path.abspath(src_excel) != os.path.abspath(default_excel):
                                shutil.copy2(src_excel, default_excel)
                    except Exception:
                        # 同步失败不阻断流程（新引擎会使用 excel_path）
                        pass
                    # 运行前做一次数据体检（容量与双师教师数）
                    fatal_msgs = []
                    # 容量 vs 需求
                    class_unavail = getattr(data, 'class_unavailable', {}) or {}
                    for cid, info in data.classes.items():
                        days = (info.end_date - info.start_date).days + 1
                        capacity = days * 2 - len(class_unavail.get(cid, set()))
                        demand = sum(data.courses[c].blocks for c in info.courses if c in data.courses)
                        if demand > capacity:
                            fatal_msgs.append(f"班级 {cid} 需求 {demand} > 容量 {capacity}")
                    # 双师课程教师数量
                    for cname, cinfo in data.courses.items():
                        if getattr(cinfo, 'is_two', False) and len(set(cinfo.teachers)) < 2:
                            fatal_msgs.append(f"课程 {cname} 标记双师但教师数量不足2")
                    if fatal_msgs:
                        raise RuntimeError('数据不可行：' + '；'.join(fatal_msgs))

                    from auto_schedule.ga_engine import run_scheduler, build_absolute
                    from auto_schedule.data_model import TimetableData as AutoData
                    from manual_schedule.manual_core import PlacedBlock as MBlock

                    try:
                        best, metrics = run_scheduler(
                            pop_size=int(pop),
                            ngen=int(gen),
                            excel_out=auto_result_path,
                            seed=int(seed),
                            verbose=int(verbose),
                            # 确保 GA 使用与界面相同的数据源（修复云端数据传输不一致）
                            excel_path=getattr(data, 'excel_file_path', None)
                        )
                    except TypeError as te:
                        # 兼容旧版本 run_scheduler 不支持 excel_path 的情况
                        if 'unexpected keyword argument' in str(te) and 'excel_path' in str(te):
                            st.info('检测到当前环境的排程引擎不支持 excel_path 参数，已回退到默认数据源逻辑。')
                            best, metrics = run_scheduler(
                                pop_size=int(pop),
                                ngen=int(gen),
                                excel_out=auto_result_path,
                                seed=int(seed),
                                verbose=int(verbose)
                            )
                        else:
                            raise

                    # 优先从导出的 Excel 回读，确保云端 rerun 后也能恢复状态
                    session.scheduler.placed.clear()
                    try:
                        imported = session.import_from_excel(auto_result_path)
                    except Exception:
                        # 回退到内存导入
                        auto_data = AutoData(getattr(session.data, 'excel_file_path', '排课数据.xlsx'))
                        abs_best = build_absolute(best, auto_data)
                        imported = 0
                        for cid, course, t1, t2, date, period_idx, _ in abs_best:
                            if date is None:
                                continue
                            blk = MBlock(cid, course, t1 or '', t2, date, period_idx)
                            session.scheduler.placed.append(blk)
                            imported += 1

                    st.success(f"✅ 自动排课完成！导入 {imported} 个课程块")
                    st.metric("硬约束满足", "是" if metrics['hard_ok'] else "否")
                    st.metric("适应度得分", f"{metrics['total_fitness']:.2f}")
                    # 记录到 session_state 以便刷新后仍能看到摘要
                    st.session_state['ga_last'] = {
                        'imported': imported,
                        'metrics': metrics,
                        'path': auto_result_path,
                    }
                    force_rerun()

                except Exception as e:
                    st.error(f"❌ 自动排课失败: {e}")

def render_legend():
    """渲染图例"""
    st.markdown("""
        <div class='legend-box'>
            <div class='legend-item'><div class='legend-color lg-theory'></div>理论课</div>
            <div class='legend-item'><div class='legend-color lg-practice'></div>实操课</div>
            <div class='legend-item'><div class='legend-color lg-dual'></div>双师课</div>
            <div class='legend-item'><div class='legend-color lg-done'></div>已完成</div>
        </div>
    """, unsafe_allow_html=True)

def compute_progress(class_id):
    """计算进度信息"""
    cls = data.classes[class_id]
    rows = []
    for c in cls.courses:
        need = data.courses[c].blocks
        remain = session.scheduler.remaining_blocks(cls.class_id, c)
        used = need - remain
        pct = 0 if need == 0 else used / need
        rows.append((c, need, used, remain, pct))
    
    total_remain = sum(r[3] for r in rows)
    finished = sum(1 for _, need, used, _, _ in rows if need > 0 and used >= need)
    return rows, total_remain, finished, len(rows)

def render_toolbar(class_id, finished, total_courses, total_remain):
    """渲染工具栏"""
    col1, col2, col3, col4, col5, col6, col7, col8 = st.columns([1.5, 1, 1, 0.8, 0.8, 0.8, 0.8, 1])
    
    with col1:
        st.markdown(f"<div class='summary-pill'>🏫 班级: <b>{class_id}</b></div>", unsafe_allow_html=True)
    with col2:
        st.markdown(f"<div class='summary-pill'>✅ 完成: {finished}/{total_courses}</div>", unsafe_allow_html=True)
    with col3:
        st.markdown(f"<div class='summary-pill'>📦 剩余: {total_remain}块</div>", unsafe_allow_html=True)
    with col4:
        if st.button('↩️ 撤销'):
            if session.undo():
                force_rerun()
            else:
                st.toast('无可撤销操作', icon='⚠️')
    with col5:
        st.checkbox('隐藏完成', key='hide_done')
    with col6:
        st.checkbox('仅未完成', key='unfinished_only')
    with col7:
        st.checkbox('🌙 夜间', key='dark_mode')
    with col8:
        if st.button('📊 进度详情'):
            st.session_state['show_progress'] = not st.session_state.get('show_progress', False)

def render_progress_panel(prog_rows, total_remain):
    """渲染进度面板"""
    if st.session_state.get('show_progress', False):
        with st.expander('📊 课程进度详情', expanded=True):
            for c, need, used, remain, pct in prog_rows:
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.progress(pct, text=f"{c}: {used}/{need} (剩余{remain})")
                with col2:
                    if pct >= 1:
                        st.success("已完成")
                    else:
                        st.warning(f"{pct*100:.0f}%")
            
            if total_remain > 0:
                est_days = (total_remain + 1) // 2
                st.info(f"📅 预计还需 {est_days} 天完成（按每天2块计算）")

def render_timetable(class_id):
    """渲染课表"""
    st.markdown("### 📅 课表视图")
    
    # 添加表格容器样式
    st.markdown("""
        <style>
        .timetable-container {
            background: white;
            border-radius: 8px;
            padding: 1rem;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        }
        .timetable-grid {
            display: grid;
            grid-template-columns: 100px 80px 1fr 1fr;
            gap: 2px;
            background: #e1e6eb;
            padding: 2px;
            border-radius: 6px;
        }
        .grid-cell-wrapper {
            background: white;
            min-height: 60px;
            padding: 8px;
            display: flex;
            flex-direction: column;
            justify-content: center;
        }
        .grid-header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            font-weight: 600;
            text-align: center;
            padding: 10px;
        }
        .grid-date-cell {
            background: #f8f9fa;
            font-weight: 600;
            text-align: center;
        }
        .grid-week-cell {
            background: #f8f9fa;
            text-align: center;
            color: #6c757d;
        }
        </style>
    """, unsafe_allow_html=True)
    
    # 获取数据
    rows = session.scheduler.export_rows()
    if not rows:
        st.info("暂无排课数据，请点击空白时段添加课程")
        return
    
    import pandas as pd
    df_all = pd.DataFrame(rows)
    class_df = df_all[df_all['班级ID'] == class_id].sort_values(['日期', '时段'])
    
    # 日期范围
    cls_info = data.classes[class_id]
    full_dates = [
        cls_info.start_date + datetime.timedelta(days=i) 
        for i in range((cls_info.end_date - cls_info.start_date).days + 1)
    ]
    
    # 过滤选项
    if st.session_state.get('unfinished_only', False):
        date_list = []
        for d in full_dates:
            day_blocks = class_df[class_df['日期'] == d]
            has_remaining = any(
                session.scheduler.remaining_blocks(class_id, c) > 0 
                for c in data.classes[class_id].courses
            )
            if has_remaining or not day_blocks.empty:
                date_list.append(d)
    else:
        date_list = full_dates
    
    # 使用容器包装整个课表
    with st.container():
        st.markdown('<div class="timetable-container">', unsafe_allow_html=True)
        
        # 渲染表头 - 使用固定比例的列宽
        header_cols = st.columns([1.2, 1, 4, 4])
        header_cols[0].markdown("<div class='grid-header'>📅 日期</div>", unsafe_allow_html=True)
        header_cols[1].markdown("<div class='grid-header'>星期</div>", unsafe_allow_html=True)
        header_cols[2].markdown("<div class='grid-header'>🌅 上午</div>", unsafe_allow_html=True)
        header_cols[3].markdown("<div class='grid-header'>🌆 下午</div>", unsafe_allow_html=True)
        
        week_map = ['一', '二', '三', '四', '五', '六', '日']
        
        # 渲染每一天 - 保持相同的列宽比例
        for d in date_list:
            row_cols = st.columns([1.2, 1, 4, 4])
            
            # 日期和星期
            row_cols[0].markdown(
                f"<div class='grid-date-cell grid-cell-wrapper'>{d.strftime('%m月%d日')}</div>", 
                unsafe_allow_html=True
            )
            row_cols[1].markdown(
                f"<div class='grid-week-cell grid-cell-wrapper'>周{week_map[d.weekday()]}</div>", 
                unsafe_allow_html=True
            )
            
            # 上午和下午时段
            for period, col_idx in [(0, 2), (1, 3)]:
                with row_cols[col_idx]:
                    has_course = not class_df[
                        (class_df['日期'] == d) & 
                        (class_df['时段'] == ('上午' if period == 0 else '下午'))
                    ].empty
                    
                    wrapper_class = "grid-cell-wrapper has-course" if has_course else "grid-cell-wrapper"
                    
                    st.markdown(f'<div class="{wrapper_class}">', unsafe_allow_html=True)
                    render_time_slot_improved(class_id, d, period, class_df)
                    st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown('</div>', unsafe_allow_html=True)

def render_time_slot_improved(class_id, date, period, class_df):
    """改进的时间段渲染"""
    slot_df = class_df[
        (class_df['日期'] == date) & 
        (class_df['时段'] == ('上午' if period == 0 else '下午'))
    ]
    
    container_key = f"cell_{class_id}_{date}_{period}"
    class_unavail = data.class_unavailable.get(class_id, set())
    slot_unavail = (date, period) in class_unavail
    
    # 检查编辑状态
    if 'editing_cell' not in st.session_state:
        st.session_state['editing_cell'] = None
    editing = st.session_state['editing_cell'] == container_key
    has_blocks = not slot_df.empty
    
    # 如果有课程且在编辑，取消编辑
    if has_blocks and editing:
        st.session_state['editing_cell'] = None
        editing = False
    
    # 判断是否可以添加
    remaining_any = any(
        session.scheduler.remaining_blocks(class_id, c) > 0 
        for c in data.classes[class_id].courses
    )
    can_add = (not slot_unavail) and remaining_any and (not has_blocks)
    
    # 不可用时段
    if slot_unavail and slot_df.empty:
        st.markdown("""
            <div style='
                background: linear-gradient(45deg, #f8f9fa 25%, transparent 25%, transparent 75%, #f8f9fa 75%, #f8f9fa),
                linear-gradient(45deg, #f8f9fa 25%, transparent 25%, transparent 75%, #f8f9fa 75%, #f8f9fa);
                background-size: 10px 10px;
                background-position: 0 0, 5px 5px;
                color: #adb5bd;
                text-align: center;
                padding: 15px;
                border-radius: 4px;
            '>🚫 不可排课</div>
        """, unsafe_allow_html=True)
        return
    
    # 显示已有课程
    if has_blocks:
        for _, r in slot_df.iterrows():
            render_course_chip_improved(r, class_id, date, period)
    
    # 添加课程按钮或表单
    elif can_add:
        if not editing:
            if st.button('➕ 添加课程', key=f"add_{container_key}", use_container_width=True):
                st.session_state['editing_cell'] = container_key
                force_rerun()
        else:
            render_add_form(class_id, date, period, container_key)
    else:
        # 无法添加时给出原因提示，便于排查
        reasons = []
        if slot_unavail:
            reasons.append('该时段不可用')
        if has_blocks:
            reasons.append('该时段已有课程')
        if not remaining_any:
            reasons.append('该班级所有课程块已排完')
        if reasons:
            st.caption('；'.join(reasons))

def render_course_chip_improved(row, class_id, date, period):
    """改进的课程卡片渲染"""
    # 查找对应的block索引
    idx_candidates = [
        i for i, b in enumerate(session.scheduler.placed)
        if b.class_id == row['班级ID'] and b.course == row['课程'] 
        and str(b.date) == str(row['日期'])
        and (0 if b.period == 0 else 1) == (0 if row['时段'] == '上午' else 1)
        and b.teacher1 == row['教师1'] and (b.teacher2 or '') == row['教师2']
    ]
    block_index = idx_candidates[0] if idx_candidates else -1
    
    # 课程信息
    course_info = data.courses[row['课程']]
    need = course_info.blocks
    remain = session.scheduler.remaining_blocks(row['班级ID'], row['课程'])
    used = need - remain
    progress = int((used / need) * 100) if need > 0 else 0
    
    # 颜色映射
    color_map = {
        0: '#e3f2fd', 1: '#f3e5f5', 2: '#fff3e0', 3: '#e8f5e9', 4: '#fff8e1',
        5: '#e1f5fe', 6: '#fce4ec', 7: '#e0f2f1', 8: '#f1f8e9', 9: '#fbe9e7'
    }
    color_idx = course_color_map.get(row['课程'], 0)
    bg_color = color_map.get(color_idx, '#f5f5f5')
    
    # 状态图标
    status_icon = "✅" if used >= need else ("⚡" if course_info.is_two else "📚")
    
    # 教师信息
    teachers = row['教师1']
    if row['教师2']:
        teachers += f" & {row['教师2']}"
    
    # 渲染卡片
    st.markdown(f"""
        <div style='
            background: {bg_color};
            border-left: 4px solid {"#4caf50" if used >= need else "#2196f3"};
            border-radius: 6px;
            /* padding: 8px; */ /* <-- 移除内联padding */
            /* margin: 4px 0; */ /* <-- 移除内联margin */
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            position: relative;
        '>
            <div style='display: flex; justify-content: space-between; align-items: center; width: 100%;'>
                <div style='display: flex; align-items: baseline; gap: 8px; flex-wrap: wrap;'>
                    <div style='font-weight: 600; color: #333; font-size: 16px;'>
                        {status_icon} {row['课程']}
                    </div>
                    <div style='color: #666; font-size: 14px;'>
                        👨‍🏫 {teachers}
                    </div>
                </div>
                <div style='text-align: right; flex-shrink: 0; margin-left: 8px;'>
                    <div style='font-size: 11px; color: #888;'>
                        {used}/{need} 块
                    </div>
                    <div style='
                        width: 50px;
                        height: 4px;
                        background: #e0e0e0;
                        border-radius: 2px;
                        margin-top: 4px;
                    '>
                        <div style='
                            width: {progress}%;
                            height: 100%;
                            background: {"#4caf50" if progress >= 100 else "#2196f3"};
                            border-radius: 2px;
                        '></div>
                    </div>
                </div>
            </div>
        </div>
    """, unsafe_allow_html=True)
    
    # 删除按钮
    if block_index >= 0:
        if st.button('🗑️ 删除', key=f"del_{block_index}", use_container_width=True):
            if session.delete_block(block_index):
                force_rerun()

def render_add_form(class_id, date, period, container_key):
    """渲染添加课程表单"""
    # 获取可选课程
    all_courses = data.classes[class_id].courses
    remain_map = {c: session.scheduler.remaining_blocks(class_id, c) for c in all_courses}
    options = [c for c in all_courses if remain_map[c] > 0]
    
    if not options:
        st.info('所有课程已完成')
        return
    
    # 选择课程
    course = st.selectbox(
        '选择课程',
        options,
        format_func=lambda x: f"{x} (剩余{remain_map[x]}块)",
        key=f"{container_key}_course"
    )
    
    if course:
        course_info = data.courses[course]
        
        # 获取可用教师
        occupied = {
            b.teacher1 for b in session.scheduler.placed 
            if b.date == date and b.period == period
        }
        occupied.update({
            b.teacher2 for b in session.scheduler.placed 
            if b.date == date and b.period == period and b.teacher2
        })
        
        teacher_unavail = data.teacher_unavailable
        
        def is_available(t):
            if t in occupied:
                return False
            if t in teacher_unavail and (date, period) in teacher_unavail[t]:
                return False
            return True
        
        available = [t for t in course_info.teachers if is_available(t)]
        
        if not available:
            st.warning('该时段无可用教师')
            available = course_info.teachers
        
        # 选择教师
        if course_info.is_theory:
            t1 = st.selectbox('教师', [available[0]], key=f"{container_key}_t1")
            t2 = None
        else:
            t1 = st.selectbox('教师1', available, key=f"{container_key}_t1")
            if course_info.is_two:
                avail2 = [t for t in course_info.teachers if t != t1 and is_available(t)]
                if not avail2:
                    avail2 = [t for t in course_info.teachers if t != t1]
                t2 = st.selectbox('教师2', avail2, key=f"{container_key}_t2")
            else:
                t2 = None
        
        # 按钮
        col1, col2 = st.columns(2)
        with col1:
            if st.button('✅ 保存', key=f"save_{container_key}", type='primary'):
                ok, errs = session.add_block(class_id, course, t1, t2, date, period)
                if ok:
                    st.session_state['editing_cell'] = None
                    force_rerun()
                else:
                    st.error('；'.join(errs))
        
        with col2:
            if st.button('❌ 取消', key=f"cancel_{container_key}"):
                st.session_state['editing_cell'] = None
                force_rerun()

def render_soft_constraints():
    """渲染软约束评估"""
    st.markdown("### 📊 软约束评估")
    
    soft_total, details = session.soft_report()
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("软约束总分", f"{soft_total:.1f}")
    
    with col2:
        for k, v in details.items():
            st.caption(f"{k}: {v}")

def render_export():
    """渲染导出部分"""
    st.markdown("### 💾 导出功能")
    
    # 教师统计
    if session.scheduler.placed:
        import pandas as pd
        teacher_blocks = []
        for b in session.scheduler.placed:
            teacher_blocks.append((b.teacher1, 1))
            if b.teacher2:
                teacher_blocks.append((b.teacher2, 1))
        
        df_load = pd.DataFrame(teacher_blocks, columns=['教师', '课时数'])
        df_load = df_load.groupby('教师').sum().reset_index()
        df_load = df_load.sort_values('课时数', ascending=False)
        
        with st.expander("👨‍🏫 教师课时统计", expanded=False):
            st.dataframe(df_load, use_container_width=True)
            dual_count = sum(1 for b in session.scheduler.placed if b.teacher2)
            st.caption(f"双师授课块数: {dual_count}")
    
    # 导出选项
    col1, col2, col3 = st.columns(3)
    with col1:
        only_current = st.checkbox('仅导出当前班级', value=True)
    with col2:
        export_mode = st.selectbox('导出方式', ['浏览器下载', '服务器保存'])
    with col3:
        if st.button('🔄 清除缓存'):
            get_session.clear()
            st.success('缓存已清除')
    
    # 文件名
    default_name = f"schedule_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    export_name = st.text_input('导出文件名', value=default_name)
    
    # 导出按钮
    if st.button('📥 导出Excel', type='primary', use_container_width=True):
        # 检查未完成
        incomplete = []
        for cid, info in data.classes.items():
            for c in info.courses:
                remain = session.scheduler.remaining_blocks(cid, c)
                if remain > 0:
                    incomplete.append(f"{cid}-{c} (剩{remain})")
        
        if incomplete:
            st.warning(f"⚠️ 未完成课程: {', '.join(incomplete[:5])}...")
        
        try:
            target_class = st.session_state.get('global_class') if only_current else None
            
            if export_mode == '浏览器下载':
                raw = session.export_excel_bytes(class_id=target_class)
                st.download_button(
                    '📥 点击下载',
                    data=raw,
                    file_name=export_name,
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
            else:
                path = session.export_excel(export_name, class_id=target_class)
                st.success(f"✅ 已保存到: {path}")
                
        except Exception as e:
            st.error(f"❌ 导出失败: {e}")

# ============ 主程序 ============
def main():
    # 注入CSS
    inject_css()
    
    # 渲染头部
    render_header()
    # 数据加载概览
    try:
        src_name = os.path.basename(getattr(data, 'excel_file_path', '未知'))
        st.caption(f"数据源: {src_name} | 班级: {len(getattr(data, 'classes', {}))} | 课程: {len(getattr(data, 'courses', {}))}")
    except Exception:
        pass
    
    # 班级选择
    st.markdown("---")
    class_keys = list(data.classes.keys())
    if not class_keys:
        st.error("未检测到班级数据，请上传有效的排课数据 Excel 文件，或检查数据格式。")
        return
    selected_class = st.selectbox(
        '🏫 选择班级',
        class_keys,
        key='global_class',
        index=0 if st.session_state.get('global_class') not in class_keys else class_keys.index(st.session_state['global_class'])
    )

    # 自动排课
    render_ga_section()

    # 图例
    render_legend()

    # 计算进度
    if selected_class is not None:
        prog_rows, total_remain, finished, total_courses = compute_progress(selected_class)
        # 工具栏
        render_toolbar(selected_class, finished, total_courses, total_remain)
        # 进度面板
        render_progress_panel(prog_rows, total_remain)
        # 课表
        render_timetable(selected_class)
    else:
        st.warning("请选择班级后查看课表和进度。"); return

    # 软约束
    render_soft_constraints()

    # 导出
    render_export()

    # 页脚
    st.markdown("---")
    st.caption("💡 提示：点击课程块可删除，点击空白时段可添加课程")

if __name__ == '__main__':
    main()
