import os
import sys
import io
import gc
import base64
import subprocess
import re
import math
from collections import Counter
from datetime import datetime
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# Excel / Openpyxl 非法控制字符正则过滤（ASCII 0-8, 11-12, 14-31 及 XML 违禁字符）

def make_b64_download_link(data_bytes, filename, mime_type, label, bg_color="#2563EB"):
    """生成前端 Base64 离线 Data-URI 下载链接，完全绕过 Streamlit MediaFileManager 与 Tornado 端点，100% 杜绝 404 及未发现文件错误"""
    b64 = base64.b64encode(data_bytes).decode()
    return f"""<a href="data:{mime_type};base64,{b64}" download="{filename}" style="display: block; width: 100%; text-align: center; background-color: {bg_color}; color: white; padding: 10px 16px; border-radius: 6px; font-weight: 600; font-size: 14px; text-decoration: none; margin-top: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.15);" target="_blank" rel="noopener noreferrer">{label}</a>"""

def save_file_to_local_disk(data_bytes, filename, target_folder_type="Downloads"):
    """本地桌面程序直接将文件写入用户磁盘，无需经过任何浏览器下载中转，100% 可靠"""
    user_home = os.path.expanduser("~")
    if target_folder_type == "Desktop":
        folder = os.path.join(user_home, "Desktop")
    elif target_folder_type == "Current":
        folder = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.getcwd()
    else:
        folder = os.path.join(user_home, "Downloads")
    
    if not os.path.exists(folder):
        try:
            os.makedirs(folder, exist_ok=True)
        except Exception:
            folder = os.getcwd()
            
    full_path = os.path.abspath(os.path.join(folder, filename))
    with open(full_path, "wb") as f:
        f.write(data_bytes)
    return full_path

def reveal_file_in_explorer(file_path):
    """在 Windows 资源管理器或 macOS Finder 中高亮选定文件"""
    try:
        norm_path = os.path.normpath(file_path)
        if sys.platform == 'win32':
            os.system(f'explorer /select,"{norm_path}"')
        elif sys.platform == 'darwin':
            subprocess.run(['open', '-R', norm_path])
        else:
            subprocess.run(['xdg-open', os.path.dirname(norm_path)])
    except Exception:
        pass

ILLEGAL_CHARACTERS_RE = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]|[\000-\010]|[\013-\014]|[\016-\037]')

# 核心防御：直接拦截 openpyxl 的 check_string，自动清洗非法控制字符而非抛出 IllegalCharacterError 异常
try:
    import openpyxl.cell.cell as ox_cell
    _orig_check_string = ox_cell.check_string
    def _safe_check_string(value):
        if value is None:
            return ""
        if not isinstance(value, str):
            value = str(value)
        # 自动剥离任何非打印控制字符，永远不触发异常
        return ILLEGAL_CHARACTERS_RE.sub('', value)
    ox_cell.check_string = _safe_check_string
except Exception:
    pass

def clean_excel_data(df):
    """
    深度递归清洗 DataFrame 中的非法字符（Excel/Openpyxl 禁止的 ASCII 0-31 控制字符）
    防止导出 Excel 时触发 openpyxl.utils.exceptions.IllegalCharacterError
    """
    if df is None or df.empty:
        return df
    
    clean_df = df.copy()
    # 1. 清洗列名
    clean_df.columns = [ILLEGAL_CHARACTERS_RE.sub('', str(c)) for c in clean_df.columns]
    
    def _clean_cell(val):
        if isinstance(val, str):
            return ILLEGAL_CHARACTERS_RE.sub('', val)
        elif isinstance(val, (bytes, bytearray)):
            try:
                s = val.decode('utf-8', errors='ignore')
                return ILLEGAL_CHARACTERS_RE.sub('', s)
            except Exception:
                return ''
        return val

    # 2. 全表单元格深度清洗（不依赖 dtype 判断，兼容 DataFrame.map 与 applymap）
    if hasattr(clean_df, 'map'):
        clean_df = clean_df.map(_clean_cell)
    elif hasattr(clean_df, 'applymap'):
        clean_df = clean_df.applymap(_clean_cell)
    else:
        for col in clean_df.columns:
            clean_df[col] = clean_df[col].map(_clean_cell)
            
    return clean_df

# ==========================================
# 1. 页面基本配置与自定义现代化样式
# ==========================================
st.set_page_config(
    page_title="FCT3 生产测试数据智能分析与 CPK 看板",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定义工业工程风 CSS 样式
st.markdown("""
<style>
    .main-title {
        font-size: 26px;
        font-weight: 700;
        color: #1E293B;
        margin-bottom: 2px;
    }
    .sub-title {
        font-size: 14px;
        color: #64748B;
        margin-bottom: 18px;
    }
    .metric-card {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 14px 18px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .metric-label {
        font-size: 13px;
        color: #64748B;
        font-weight: 500;
    }
    .metric-val {
        font-size: 24px;
        font-weight: 700;
        color: #0F172A;
        margin-top: 4px;
    }
    .cpk-badge-good {
        background-color: #DCFCE7;
        color: #166534;
        padding: 4px 10px;
        border-radius: 6px;
        font-weight: 600;
    }
    .cpk-badge-warn {
        background-color: #FEF9C3;
        color: #854D0E;
        padding: 4px 10px;
        border-radius: 6px;
        font-weight: 600;
    }
    .cpk-badge-danger {
        background-color: #FEE2E2;
        color: #991B1B;
        padding: 4px 10px;
        border-radius: 6px;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)


# ==========================================
# 2. 核心数据解析与智能规格限提取引擎
# ==========================================
def is_valid_numeric_limit(val):
    """判断规格限单元格是否为有效的连续数值（排除 0x 十六进制地址和纯文本）"""
    if pd.isna(val):
        return False
    val_str = str(val).strip()
    if not val_str or val_str.lower().startswith('0x'):
        return False
    try:
        float(val_str)
        return True
    except ValueError:
        return False


@st.cache_data(show_spinner=False)
def load_and_parse_csv(source_input):
    """
    智能解析含有前两行规格限的测试 CSV 文件
    自动过滤纯文本、状态列与十六进制地址列，仅提取包含有效上下限的连续数值测试项
    """
    if isinstance(source_input, str):
        with open(source_input, 'r', encoding='utf-8', errors='ignore') as f:
            lines = [f.readline() for _ in range(3)]
        raw_header = [col.strip() for col in lines[0].strip().split(',')]
        upper_limits_raw = [col.strip() for col in lines[1].strip().split(',')]
        lower_limits_raw = [col.strip() for col in lines[2].strip().split(',')]
        
        # 读取从第4行开始的正式数据
        df_data = pd.read_csv(
            source_input,
            skiprows=3,
            header=None,
            names=raw_header,
            encoding='utf-8',
            low_memory=False,
            on_bad_lines='skip'
        )
    else:
        # Streamlit 内存文件句柄
        source_input.seek(0)
        line1 = source_input.readline().decode('utf-8', errors='ignore')
        line2 = source_input.readline().decode('utf-8', errors='ignore')
        line3 = source_input.readline().decode('utf-8', errors='ignore')
        raw_header = [col.strip() for col in line1.strip().split(',')]
        upper_limits_raw = [col.strip() for col in line2.strip().split(',')]
        lower_limits_raw = [col.strip() for col in line3.strip().split(',')]
        
        source_input.seek(0)
        df_data = pd.read_csv(
            source_input,
            skiprows=3,
            header=None,
            names=raw_header,
            encoding='utf-8',
            low_memory=False,
            on_bad_lines='skip'
        )

    # 规范化列名长度
    n_cols = len(raw_header)
    upper_limits_raw += [''] * (n_cols - len(upper_limits_raw))
    lower_limits_raw += [''] * (n_cols - len(lower_limits_raw))

    # 智能提取包含有效规格上下限的连续数值项
    spec_dict = {}
    valid_test_items = []

    for i, col in enumerate(raw_header):
        u_val = upper_limits_raw[i] if i < len(upper_limits_raw) else ''
        l_val = lower_limits_raw[i] if i < len(lower_limits_raw) else ''
        
        has_usl = is_valid_numeric_limit(u_val)
        has_lsl = is_valid_numeric_limit(l_val)
        
        if has_usl or has_lsl:
            usl_float = float(u_val) if has_usl else None
            lsl_float = float(l_val) if has_lsl else None
            spec_dict[col] = {
                'usl': usl_float,
                'lsl': lsl_float,
                'usl_str': str(u_val) if has_usl else "None",
                'lsl_str': str(l_val) if has_lsl else "None"
            }
            valid_test_items.append(col)

    # 【核心内存优化1】：将数值测试项批量转为 float32（内存占用锐减85%）
    for col in valid_test_items:
        df_data[col] = pd.to_numeric(df_data[col], errors='coerce').astype('float32')

    # 仅清洗剩余非数值的元数据文本列
    meta_text_cols = [c for c in df_data.columns if c not in valid_test_items and df_data[c].dtype == object]
    for col in meta_text_cols:
        df_data[col] = df_data[col].map(lambda x: ILLEGAL_CHARACTERS_RE.sub('', x) if isinstance(x, str) else x)

    # 处理时间列格式
    if 'Time' in df_data.columns:
        df_data['Parsed_Time'] = pd.to_datetime(df_data['Time'], format='%m-%d-%Y %H:%M:%S', errors='coerce')
    else:
        df_data['Parsed_Time'] = pd.NaT

    gc.collect()
    return df_data, spec_dict, valid_test_items


# ==========================================
# 3. 统计学与 CPK 计算引擎
# ==========================================
def calculate_process_capability(series, lsl, usl, is_smaller_the_better=False):
    """
    计算制程能力指标：Mean, Std(总方差), Std_Within(组内方差), Min, Max, Cp, Cpu, Cpl, Cpk, Pp, Ppk
    支持双边公差、单边公差，以及望小特性（当LSL=0为物理零下界时仅考核USL）
    """
    numeric_s = pd.to_numeric(series, errors='coerce').dropna()
    n = len(numeric_s)
    if n < 2:
        return {
            'n': n,
            'mean': float('nan'), 'std': float('nan'), 'std_within': float('nan'),
            'min': float('nan'), 'max': float('nan'), 'median': float('nan'),
            'cp': float('nan'), 'cpu': float('nan'), 'cpl': float('nan'), 'cpk': float('nan'),
            'pp': float('nan'), 'ppk': float('nan'),
            'status': '数据不足'
        }

    mean = float(numeric_s.mean())
    std = float(numeric_s.std(ddof=1)) # 总体样本标准差
    val_min = float(numeric_s.min())
    val_max = float(numeric_s.max())
    median = float(numeric_s.median())

    # 移动极差估算组内标准差 (Within Sigma: MR-bar / d2, d2=1.128, JMP/Minitab 工业标准)
    # 【核心性能优化2】：NumPy 向量化差分计算，杜绝 Python 列表转换与临时对象堆积
    vals_arr = numeric_s.to_numpy(dtype=np.float64)
    if len(vals_arr) > 1:
        mr_avg = float(np.mean(np.abs(np.diff(vals_arr))))
        std_within = float(mr_avg / 1.128)
    else:
        std_within = std

    if std <= 1e-12:
        return {
            'n': n, 'mean': mean, 'std': 0.0, 'std_within': 0.0,
            'min': val_min, 'max': val_max, 'median': median,
            'cp': float('nan'), 'cpu': float('nan'), 'cpl': float('nan'), 'cpk': float('nan'),
            'pp': float('nan'), 'ppk': float('nan'),
            'status': '标准差极小/恒定'
        }

    sigma_w = std_within if std_within > 1e-12 else std

    # 若指定为望小特性（例如漏电流、噪声、低电平，0为物理零下界，越接近0越好），LSL 不作下限严苛惩罚
    effective_lsl = None if is_smaller_the_better else lsl

    # 1. 经典组内制程能力指数 Cpk (基于短期组内变异 sigma_w)
    cp = float('nan')
    cpu = float('nan')
    cpl = float('nan')
    cpk = float('nan')

    if usl is not None and not math.isnan(usl):
        cpu = (usl - mean) / (3.0 * sigma_w)
    if effective_lsl is not None and not math.isnan(effective_lsl):
        cpl = (mean - effective_lsl) / (3.0 * sigma_w)

    if usl is not None and effective_lsl is not None and not math.isnan(usl) and not math.isnan(effective_lsl):
        cp = (usl - effective_lsl) / (6.0 * sigma_w)
        cpk = min(cpu, cpl)
    elif usl is not None and not math.isnan(usl):
        cpk = cpu
    elif effective_lsl is not None and not math.isnan(effective_lsl):
        cpk = cpl

    # 2. 总体过程性能指数 Ppk (基于全样本长期方差 std)
    pp = float('nan')
    ppk = float('nan')
    if usl is not None and effective_lsl is not None and not math.isnan(usl) and not math.isnan(effective_lsl):
        pp = (usl - effective_lsl) / (6.0 * std)
        ppk = min((usl - mean) / (3.0 * std), (mean - effective_lsl) / (3.0 * std))
    elif usl is not None and not math.isnan(usl):
        ppk = (usl - mean) / (3.0 * std)
    elif effective_lsl is not None and not math.isnan(effective_lsl):
        ppk = (mean - effective_lsl) / (3.0 * std)

    # 等级评定标准
    eval_val = cpk if not math.isnan(cpk) else ppk
    if math.isnan(eval_val):
        status = "无规格限"
    elif eval_val < 1.0:
        status = "不足(危险)"
    elif eval_val < 1.33:
        status = "警告(合格)"
    elif eval_val < 1.67:
        status = "良好(稳定)"
    else:
        status = "卓越(极佳)"

    return {
        'n': n,
        'mean': mean,
        'std': std,
        'std_within': std_within,
        'cp': cp,
        'cpu': cpu,
        'cpl': cpl,
        'cpk': cpk,
        'pp': pp,
        'ppk': ppk,
        'min': val_min,
        'max': val_max,
        'median': median,
        'status': status
    }


# ==========================================
# 4. 数据源加载处理
# ==========================================
st.markdown('<div class="main-title">📊 FCT3 生产测试数据智能分析与制程能力看板</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">实时交互式良率统计、失效柏拉图、单项测试 CPK 能力评估、时序温漂分析与机台离散度对比</div>', unsafe_allow_html=True)

# 侧边栏配置
st.sidebar.header("📁 数据源与全局配置")

# 寻找默认数据集
default_csv_candidates = [
    os.path.join(os.path.dirname(__file__), "测试数据", "cateye havana FCT3 0701-0820.csv"),
    os.path.join(getattr(sys, "_MEIPASS", ""), "测试数据", "cateye havana FCT3 0701-0820.csv"),
    os.path.join(os.path.dirname(sys.executable), "测试数据", "cateye havana FCT3 0701-0820.csv"),
    os.path.join(os.path.dirname(__file__), "..", "4_测试与实验数据", "cateye havana FCT3 0701-0820.csv"),
    os.path.join(os.path.dirname(__file__), "4_测试与实验数据", "cateye havana FCT3 0701-0820.csv"),
    "/Users/J1407869/Downloads/4_测试与实验数据/cateye havana FCT3 0701-0820.csv",
    os.path.join(os.path.dirname(__file__), "cateye havana FCT3 0701-0820.csv"),
    "cateye havana FCT3 0701-0820.csv"
]
default_csv_path = None
for p in default_csv_candidates:
    if os.path.exists(p):
        default_csv_path = p
        break

uploaded_file = st.sidebar.file_uploader("上传测试数据 CSV 文件", type=["csv"], help="可直接上传包含 Upper/Lower Limit 规格限的原始 FCT 测试记录文件")

selected_data_source = None
if uploaded_file is not None:
    selected_data_source = uploaded_file
    st.sidebar.success("已成功加载上传文件")
elif default_csv_path is not None:
    selected_data_source = default_csv_path
    st.sidebar.info(f"已自动加载默认工程测试数据：\n`{os.path.basename(default_csv_path)}`")
else:
    st.warning("⚠️ 未找到默认测试数据文件，请通过左侧侧边栏上传 CSV 文件。")
    st.stop()

# 加载数据
with st.spinner("正在智能解析数据与测试规格限..."):
    df_raw, spec_dict, valid_test_items = load_and_parse_csv(selected_data_source)

# ==========================================
# 5. 数据清洗与筛选逻辑
# ==========================================
st.sidebar.markdown("---")
st.sidebar.subheader("⚙️ 数据清洗与重测处理")

dedup_retest = st.sidebar.checkbox(
    "按 SerialNumber 去重（仅保留最后一次测试记录）",
    value=True,
    help="勾选后将自动按测试时间排序，并针对每个序列号仅保留最终重测结果；不勾选则展示产线所有原始上电测试记录。"
)

# 执行去重逻辑
if dedup_retest and 'SerialNumber' in df_raw.columns:
    if 'Parsed_Time' in df_raw.columns and df_raw['Parsed_Time'].notna().any():
        df_working = df_raw.sort_values(by='Parsed_Time', ascending=True).drop_duplicates(subset=['SerialNumber'], keep='last').copy()
    else:
        df_working = df_raw.drop_duplicates(subset=['SerialNumber'], keep='last').copy()
else:
    df_working = df_raw.copy()

st.sidebar.markdown("---")
st.sidebar.subheader("🔍 全局多维数据切片")

# 日期过滤
if 'Parsed_Time' in df_working.columns and df_working['Parsed_Time'].notna().any():
    valid_dates = df_working['Parsed_Time'].dropna()
    min_datetime = valid_dates.min().date()
    max_datetime = valid_dates.max().date()
    
    date_range = st.sidebar.date_input(
        "测试日期区间 (Date Range)",
        value=(min_datetime, max_datetime),
        min_value=min_datetime,
        max_value=max_datetime,
        help="选择分析的时间起止范围"
    )
    if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
        start_d, end_d = date_range
        df_working = df_working[
            (df_working['Parsed_Time'].dt.date >= start_d) &
            (df_working['Parsed_Time'].dt.date <= end_d)
        ]

# 批次过滤 (Batch_ID)
if 'Batch_ID' in df_working.columns:
    all_batches = sorted([str(b) for b in df_working['Batch_ID'].dropna().unique()])
    selected_batches = st.sidebar.multiselect(
        "生产批次 (Batch_ID)",
        options=all_batches,
        default=all_batches,
        help="支持勾选一个或多个测试批次进行对比切片"
    )
    if selected_batches:
        df_working = df_working[df_working['Batch_ID'].astype(str).isin(selected_batches)]
    else:
        df_working = df_working.iloc[0:0]

# 机台过滤 (TesterName)
if 'TesterName' in df_working.columns:
    all_testers = sorted([str(t) for t in df_working['TesterName'].dropna().unique()])
    selected_testers = st.sidebar.multiselect(
        "测试机台 (TesterName)",
        options=all_testers,
        default=all_testers,
        help="支持按机台单独筛选或组合分析"
    )
    if selected_testers:
        df_working = df_working[df_working['TesterName'].astype(str).isin(selected_testers)]
    else:
        df_working = df_working.iloc[0:0]

# 治具/通道过滤 (CHECK_HOLDER，若存在)
if 'CHECK_HOLDER' in df_working.columns:
    all_holders = sorted([str(h) for h in df_working['CHECK_HOLDER'].dropna().unique() if str(h).strip()])
    if len(all_holders) > 1:
        selected_holders = st.sidebar.multiselect(
            "测试工位/治具夹具 (CHECK_HOLDER)",
            options=all_holders,
            default=all_holders
        )
        if selected_holders:
            df_working = df_working[df_working['CHECK_HOLDER'].astype(str).isin(selected_holders)]

# 当前全局指标计算
total_records = len(df_working)
if total_records == 0:
    st.error("⚠️ 当前筛选条件下暂无有效数据，请调整侧边栏的过滤条件。")
    st.stop()

unique_sns = df_working['SerialNumber'].nunique() if 'SerialNumber' in df_working.columns else total_records
pass_count = (df_working['TestResult'].str.upper() == 'PASS').sum() if 'TestResult' in df_working.columns else 0
fail_count = (df_working['TestResult'].str.upper() == 'FAIL').sum() if 'TestResult' in df_working.columns else 0
curr_yield = (pass_count / total_records * 100.0) if total_records > 0 else 0.0

# 顶部 KPI 指标看板
kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
with kpi1:
    st.metric("总测试量 (Total Tests)", f"{total_records:,}")
with kpi2:
    st.metric("独立产品序列号 (Unique SNs)", f"{unique_sns:,}")
with kpi3:
    st.metric("合格产品数 (PASS)", f"{pass_count:,}", delta=f"{curr_yield:.2f}%")
with kpi4:
    st.metric("不良产品数 (FAIL)", f"{fail_count:,}", delta=f"-{(100 - curr_yield):.2f}%", delta_color="inverse")
with kpi5:
    st.metric("综合良率 (Yield Rate)", f"{curr_yield:.2f}%")

st.markdown("---")

# ==========================================
# 6. 核心分析模块（Tab 形式多维呈现）
# ==========================================
tab_pareto, tab_cpk, tab_run, tab_box, tab_export = st.tabs([
    "📊 良率看板与不良柏拉图",
    "📈 单项测试数值分布与 CPK 计算",
    "⏱️ 时序漂移趋势图 (Run Chart)",
    "📦 机台与批次对比分析 (Boxplot)",
    "📥 数据明细与一键导出 (Excel)"
])

# ----------------------------------------------------
# Tab 1: 良率看板与不良柏拉图 (Pareto Chart)
# ----------------------------------------------------
with tab_pareto:
    st.subheader("良率分析与 Top 失效项柏拉图 (Pareto Analysis)")
    col_p_left, col_p_right = st.columns([1, 2])
    
    # 提取失效明细
    fail_df = df_working[df_working['TestResult'].astype(str).str.upper() == 'FAIL']
    fail_item_counter = Counter()
    
    if 'TestFailItem' in fail_df.columns:
        for items_str in fail_df['TestFailItem'].dropna():
            for item in str(items_str).split('|'):
                item_cleaned = item.strip()
                if item_cleaned:
                    fail_item_counter[item_cleaned] += 1

    if len(fail_item_counter) == 0:
        st.info("🎉 当前筛选条件下无失效记录（良率 100%）！")
    else:
        pareto_df = pd.DataFrame(fail_item_counter.most_common(), columns=['失效项目 (FailItem)', '频次 (Count)'])
        total_fails = pareto_df['频次 (Count)'].sum()
        pareto_df['占比 (%)'] = (pareto_df['频次 (Count)'] / total_fails * 100.0).round(2)
        pareto_df['累计占比 (%)'] = pareto_df['占比 (%)'].cumsum().round(2)

        with col_p_left:
            st.markdown("##### 📌 不良项频次与累计占比统计表")
            st.dataframe(
                pareto_df.style.format({
                    '频次 (Count)': '{:,}',
                    '占比 (%)': '{:.2f}%',
                    '累计占比 (%)': '{:.2f}%'
                }),
                height=420,
                use_container_width=True
            )

        with col_p_right:
            # 绘制双轴柏拉图 (Plotly Dual-Axis)
            fig_pareto = go.Figure()

            # 柱状图：不良频次
            fig_pareto.add_trace(go.Bar(
                x=pareto_df['失效项目 (FailItem)'],
                y=pareto_df['频次 (Count)'],
                name="失效频次",
                marker_color="#3B82F6",
                text=pareto_df['频次 (Count)'],
                textposition='outside'
            ))

            # 折线图：累计百分比
            fig_pareto.add_trace(go.Scatter(
                x=pareto_df['失效项目 (FailItem)'],
                y=pareto_df['累计占比 (%)'],
                name="累计百分比",
                yaxis="y2",
                mode="lines+markers",
                marker=dict(size=7, color="#EF4444"),
                line=dict(width=2, color="#EF4444")
            ))

            # 80% 关键截断线
            fig_pareto.add_hline(
                y=80, yref='y2',
                line_dash="dash",
                line_color="#F59E0B",
                annotation_text="80% 核心截断线 (Pareto 80/20)",
                annotation_position="top right"
            )

            fig_pareto.update_layout(
                title="Top 失效项柏拉图 (Pareto Chart)",
                xaxis=dict(title="失效测试项", tickangle=-30),
                yaxis=dict(title="失效频次 (次数)", showgrid=True),
                yaxis2=dict(
                    title="累计百分比 (%)",
                    overlaying="y",
                    side="right",
                    range=[0, 105],
                    showgrid=False
                ),
                legend=dict(x=0.8, y=1.1, orientation="h"),
                margin=dict(l=40, r=40, t=50, b=80),
                height=450
            )
            st.plotly_chart(
                fig_pareto,
                use_container_width=True,
                config={
                    'scrollZoom': True,
                    'displaylogo': False,
                    'modeBarButtonsToRemove': ['lasso2d'],
                    'responsive': True
                }
            )


# ----------------------------------------------------
# Tab 2: 数值分布直方图与 CPK 计算模块
# ----------------------------------------------------
with tab_cpk:
    st.subheader("单项测试数值分布与制程能力评估 (Histogram & Process Capability)")
    
    if not valid_test_items:
        st.warning("未提取到具有规格上下限的连续数值测试项。")
    else:
        col_scope, col_sel = st.columns([1.2, 1.8])
        with col_scope:
            sample_scope_mode = st.radio(
                "🎯 CPK 样本过滤范围：",
                options=[
                    "规格限内有效值 (In-Spec，剔除断电/0等噪点)",
                    "仅限整机合格品 (PASS Only)",
                    "全量数据 (All Records，含全部失效与0点)"
                ],
                index=0,
                help="【强烈建议选第一项】：产线坏品常因未供电、开路或短路记录为 0.0，强行计入会导致方差暴增、Cpk 严重虚假偏低。过滤开路坏品才能真实反映工站固有制造能力。"
            )
        with col_sel:
            selected_item = st.selectbox(
                "请选择分析测试项：",
                options=valid_test_items,
                index=valid_test_items.index('U0507-WC_Std_IR1') if 'U0507-WC_Std_IR1' in valid_test_items else 0,
                help="系统已智能过滤纯文本及十六进制寄存器列，下拉仅显示包含真实规格上下限的数值项"
            )
            
            spec_info = spec_dict[selected_item]
            usl_val = spec_info['usl']
            lsl_val = spec_info['lsl']
            
            # 检测是否为 LSL=0 的望小特性候选（例如低电平、漏电、噪声）
            is_zero_lsl = (lsl_val == 0.0 and (usl_val is not None and usl_val > 0))
            is_smaller_the_better = False
            if is_zero_lsl:
                is_smaller_the_better = st.checkbox(
                    "💡 该项 LSL=0 为自然物理下界（望小特性：低电平/漏电越小越好），仅考核上限 USL",
                    value=True,
                    help="对于非负物理量（如低电平电压、漏电流、失调），0为理想目标而非超差危险界限。若不勾选，靠近0的优秀良品会被双边公式误判为下限超差风险。"
                )
            
            st.markdown(f"""
            **规格限信息：**
            - 规格下限 (LSL): `{spec_info['lsl_str']}` {' (已设为望小下界)' if is_smaller_the_better else ''}
            - 规格上限 (USL): `{spec_info['usl_str']}`
            """)
            
            # 手动临时微调规格限（可选）
            with st.expander("🛠️ 规格限微调 (What-If 灵敏度模拟)"):
                custom_lsl = st.number_input("调整 LSL", value=float(lsl_val) if lsl_val is not None else 0.0, disabled=(lsl_val is None))
                custom_usl = st.number_input("调整 USL", value=float(usl_val) if usl_val is not None else 100.0, disabled=(usl_val is None))
                use_custom = st.checkbox("启用模拟规格限", value=False)
                if use_custom:
                    lsl_val = custom_lsl if lsl_val is not None else None
                    usl_val = custom_usl if usl_val is not None else None

        # 根据选择的样本过滤范围提取数据
        raw_series = pd.to_numeric(df_working[selected_item], errors='coerce').dropna()
        if "In-Spec" in sample_scope_mode:
            cond = pd.Series(True, index=raw_series.index)
            if lsl_val is not None and not is_smaller_the_better:
                cond = cond & (raw_series >= lsl_val)
            elif lsl_val is not None and is_smaller_the_better:
                cond = cond & (raw_series >= 0)
            if usl_val is not None:
                cond = cond & (raw_series <= usl_val)
            series_data = raw_series[cond]
            filtered_out_cnt = len(raw_series) - len(series_data)
            if filtered_out_cnt > 0:
                st.caption(f"💡 当前模式已自动过滤 {filtered_out_cnt} 处规格限外的坏品/未供电零值噪点，确保数据满足正态分布能力评估前提。")
        elif "PASS" in sample_scope_mode:
            pass_mask = (df_working['TestResult'].str.upper() == 'PASS')
            series_data = pd.to_numeric(df_working.loc[pass_mask, selected_item], errors='coerce').dropna()
            filtered_out_cnt = len(raw_series) - len(series_data)
            if filtered_out_cnt > 0:
                st.caption(f"💡 仅统计整机合格品，已排除 {filtered_out_cnt} 条 FAIL 板卡记录。")
        else:
            series_data = raw_series
            # 检测是否存在 0 噪点提示
            zero_cnt = (raw_series == 0).sum()
            if zero_cnt > 0 and ((lsl_val is not None and lsl_val > 10) or (usl_val is not None and usl_val > 10)):
                st.warning(f"⚠️ 提示：全量数据中包含 {zero_cnt} 处值为 0.0 的读数（通常为未上电/开路坏品），这会导致样本方差剧增、Cpk 计算结果被严重拉低！建议切换为上方【规格限内有效值】模式。")

        cpk_info = calculate_process_capability(series_data, lsl_val, usl_val, is_smaller_the_better=is_smaller_the_better)

        # 统计指标与卡片 (展示 Cpk 组内能力 与 Ppk 总体性能)
        m1, m2, m3, m4, m5, m6 = st.columns(6)
        m1.metric("样本总数 (N)", f"{cpk_info['n']:,}")
        m2.metric("均值 (Mean)", f"{cpk_info['mean']:.4f}")
        m3.metric("组内标准差 (σ_w)", f"{cpk_info['std_within']:.4f}", help="基于相邻移动极差 MR/d2 计算的组内固有标准差 (JMP/Minitab 短期能力基准)")
        m4.metric("总体标准差 (s)", f"{cpk_info['std']:.4f}", help="全样本长期标准差 (包含时间温漂与跨批次方差)")
        m5.metric("Cpk (组内能力)", f"{cpk_info['cpk']:.3f}" if not math.isnan(cpk_info['cpk']) else "N/A", help="经典制程能力指数 (基于组内标准差，排除长期漂移)")
        m6.metric("Ppk (总体性能)", f"{cpk_info['ppk']:.3f}" if not math.isnan(cpk_info['ppk']) else "N/A", help="过程性能指数 (基于总体样本标准差)")

        # 警报高亮展示
        cpk_num = cpk_info['cpk'] if not math.isnan(cpk_info['cpk']) else cpk_info['ppk']
        if not math.isnan(cpk_num):
            if cpk_num < 1.0:
                st.error(f"🚨 **制程能力不足严重预警**：当前测试项 `{selected_item}` 的 Cpk = **{cpk_num:.3f}** (< 1.0)。制程离散度过大或偏离中心，存在超差风险。")
            elif cpk_num < 1.33:
                st.warning(f"⚠️ **制程能力勉强达标提示**：当前测试项 `{selected_item}` 的 Cpk = **{cpk_num:.3f}** (1.0 ~ 1.33)。建议持续监控。")
            else:
                st.success(f"✅ **制程能力表现优良**：当前测试项 `{selected_item}` 的 Cpk = **{cpk_num:.3f}** (≥ 1.33)。制程波动在公差范围内受控良好。")

        # 绘制数值分布直方图 + 正态分布曲线 + USL/LSL/Mean 辅助线
        if len(series_data) > 0:
            fig_hist = go.Figure()

            # 实际频数直方图
            fig_hist.add_trace(go.Histogram(
                x=series_data,
                name="测量频数分布",
                histnorm='',
                marker=dict(color="#3B82F6", line=dict(color="#1D4ED8", width=1)),
                opacity=0.75
            ))

            # 正态拟合曲线
            if cpk_info['std'] and cpk_info['std'] > 0:
                x_axis = np.linspace(series_data.min(), series_data.max(), 200)
                # 计算对应于频数尺度的高斯正态分布曲线
                bin_width = (series_data.max() - series_data.min()) / 30.0 if series_data.max() != series_data.min() else 1.0
                gaussian_y = (len(series_data) * bin_width / (cpk_info['std'] * np.sqrt(2 * np.pi))) * np.exp(-0.5 * ((x_axis - cpk_info['mean']) / cpk_info['std'])**2)
                fig_hist.add_trace(go.Scatter(
                    x=x_axis,
                    y=gaussian_y,
                    mode='lines',
                    name='拟合正态分布曲线 (Gaussian)',
                    line=dict(color='#8B5CF6', width=2.5)
                ))

            # 叠加 LSL / USL / Mean 参考线
            if lsl_val is not None:
                fig_hist.add_vline(
                    x=lsl_val, line_width=2.5, line_dash="dash", line_color="#EF4444",
                    annotation_text=f"LSL: {lsl_val}", annotation_position="top left",
                    annotation_font=dict(color="#EF4444", size=12)
                )
            if usl_val is not None:
                fig_hist.add_vline(
                    x=usl_val, line_width=2.5, line_dash="dash", line_color="#EF4444",
                    annotation_text=f"USL: {usl_val}", annotation_position="top right",
                    annotation_font=dict(color="#EF4444", size=12)
                )
            if not np.isnan(cpk_info['mean']):
                fig_hist.add_vline(
                    x=cpk_info['mean'], line_width=2, line_dash="dot", line_color="#10B981",
                    annotation_text=f"Mean: {cpk_info['mean']:.3f}", annotation_position="bottom right",
                    annotation_font=dict(color="#10B981", size=12)
                )

            fig_hist.update_layout(
                title=f"测试项 [{selected_item}] 数值分布直方图与规格限叠加图",
                xaxis=dict(title=f"{selected_item} 测量数值"),
                yaxis=dict(title="频数 (Counts)", showgrid=True),
                legend=dict(x=0.75, y=0.98),
                height=480,
                margin=dict(l=40, r=40, t=50, b=50)
            )
            st.plotly_chart(
                fig_hist,
                use_container_width=True,
                config={
                    'scrollZoom': True,
                    'displaylogo': False,
                    'modeBarButtonsToRemove': ['lasso2d'],
                    'responsive': True
                }
            )


# ----------------------------------------------------
# Tab 3: 时序漂移趋势 (Run Chart)
# ----------------------------------------------------
with tab_run:
    st.subheader("传感器时序漂移与探针老化分析 (Run Chart / Time-Series Drift)")
    
    col_r1, col_r2, col_r3 = st.columns([1.8, 1.2, 1.2])
    with col_r1:
        run_item = st.selectbox(
            "选择时序分析测试项：",
            options=valid_test_items,
            index=valid_test_items.index(selected_item) if selected_item in valid_test_items else 0,
            key="run_chart_selector"
        )
    with col_r2:
        color_by_opt = st.selectbox("散点颜色分类维度：", ["TesterName (机台)", "Batch_ID (批次)", "TestResult (结果)"], index=0)
        color_col = "TesterName" if "TesterName" in color_by_opt else ("Batch_ID" if "Batch_ID" in color_by_opt else "TestResult")
    with col_r3:
        show_rolling = st.checkbox("叠加移动平均趋势线 (Rolling Mean)", value=True, help="计算移动均线可清晰过滤单颗噪点，观察整体温漂或探针接触阻抗渐变趋势")

    # 缩放与坐标轴高级交互控制栏
    c_ctl1, c_ctl2, c_ctl3, c_ctl4 = st.columns([1.2, 1.4, 1.2, 2.2])
    with c_ctl1:
        use_gl = st.checkbox("⚡ WebGL 硬件加速", value=True, help="【核心优化】：开启后由 GPU 硬件加速渲染散点与框选计算。拉框选择范围与局部缩放极速不卡顿；关闭则回退到 SVG 模式。")
    with c_ctl2:
        y_focus_spec = st.checkbox("🎯 纵轴自适应聚焦规格限", value=True, help="自动根据规格上限与下限适配纵轴初显视野，避免断电0值将纵轴压缩成一条扁平直线；拉框可随时二次自由放大")
    with c_ctl3:
        show_slider = st.checkbox("⏱️ 底部时间缩略滑块", value=False, help="开启后可在图表下方通过全景小条快速滑动时间窗口；关闭后主图表区可获得更纯粹的二维拉框自由度")
    with c_ctl4:
        st.caption("💡 **缩放提示**：已完全解锁横纵双向自由缩放！按住鼠标左键在图中**任意拉画矩形框**即可同时放大横轴时间与纵轴测试值；双击图表任意处即可恢复全景。")

    run_df = df_working[['SerialNumber', 'Time', 'Parsed_Time', 'TesterName', 'Batch_ID', 'TestResult', run_item]].dropna(subset=[run_item]).copy()
    run_df['Numeric_Value'] = pd.to_numeric(run_df[run_item], errors='coerce')
    run_df = run_df.dropna(subset=['Numeric_Value']).sort_values('Parsed_Time')

    if len(run_df) > 8000:
        c_ds1, c_ds2 = st.columns([3, 1])
        with c_ds1:
            st.caption(f"📊 当前筛选样本量较大 (共 {len(run_df):,} 点)。已开启 WebGL GPU 硬件加速，支持流畅拉框与无级放大。")
        with c_ds2:
            enable_decimate = st.checkbox("密集点云降采样预览", value=False, help="若前端电脑性能较弱，可开启抽样展示骨干形态；框选缩放到局部后可取消勾选查看全部点")
            if enable_decimate:
                step = max(1, len(run_df) // 4000)
                run_df = run_df.iloc[::step]

    if len(run_df) == 0:
        st.warning("所选测试项在当前过滤范围内暂无有效数值。")
    else:
        fig_run = go.Figure()
        ScatterCls = go.Scattergl if use_gl else go.Scatter

        # 散点轨迹 (使用 Scattergl 进行 GPU 渲染，避免数千个 SVG DOM 节点在框选时引起浏览器假死)
        for grp_name, grp_data in run_df.groupby(color_col):
            fig_run.add_trace(ScatterCls(
                x=grp_data['Parsed_Time'] if grp_data['Parsed_Time'].notna().any() else grp_data['Time'],
                y=grp_data['Numeric_Value'],
                mode='markers',
                name=f"{color_col}: {grp_name}",
                marker=dict(size=5, opacity=0.7),
                customdata=np.stack((grp_data['SerialNumber'], grp_data['TestResult'], grp_data['Batch_ID'], grp_data['TesterName']), axis=-1),
                hovertemplate="<b>SN: %{customdata[0]}</b><br>时间: %{x}<br>数值: %{y:.4f}<br>机台: %{customdata[3]}<br>批次: %{customdata[2]}<br>结果: %{customdata[1]}<extra></extra>"
            ))

        # 叠加移动平均线
        if show_rolling and len(run_df) >= 10:
            window_size = max(5, int(len(run_df) * 0.05))
            run_df['Rolling_Avg'] = run_df['Numeric_Value'].rolling(window=window_size, min_periods=3).mean()
            fig_run.add_trace(ScatterCls(
                x=run_df['Parsed_Time'] if run_df['Parsed_Time'].notna().any() else run_df['Time'],
                y=run_df['Rolling_Avg'],
                mode='lines',
                name=f'滑动平均趋势线 (窗口={window_size})',
                line=dict(color='#111827', width=2.5)
            ))

        # 辅助规格线
        item_spec = spec_dict.get(run_item, {})
        lsl_val = item_spec.get('lsl')
        usl_val = item_spec.get('usl')
        if lsl_val is not None:
            fig_run.add_hline(y=lsl_val, line_dash="dash", line_color="#EF4444", annotation_text=f"LSL={lsl_val}")
        if usl_val is not None:
            fig_run.add_hline(y=usl_val, line_dash="dash", line_color="#EF4444", annotation_text=f"USL={usl_val}")

        # 智能计算纵轴范围配置，彻底解除纵轴锁定 (fixedrange=False)
        y_axis_cfg = dict(
            title=f"{run_item} 实测值",
            showgrid=True,
            fixedrange=False  # 核心：解除纵轴锁定，允许鼠标自由拉框纵向放大与滚轮纵向缩放
        )
        if y_focus_spec and (lsl_val is not None or usl_val is not None):
            if lsl_val is not None and usl_val is not None:
                span = abs(usl_val - lsl_val)
                y_axis_cfg['range'] = [lsl_val - span * 0.15, usl_val + span * 0.15]
            elif usl_val is not None:
                v_s = run_df['Numeric_Value'].dropna()
                low_b = v_s.quantile(0.01) if len(v_s) > 0 else 0
                y_axis_cfg['range'] = [low_b, usl_val + (usl_val - low_b) * 0.15]
            elif lsl_val is not None:
                v_s = run_df['Numeric_Value'].dropna()
                high_b = v_s.quantile(0.99) if len(v_s) > 0 else lsl_val * 2
                y_axis_cfg['range'] = [lsl_val - (high_b - lsl_val) * 0.15, high_b]

        xaxis_cfg = dict(
            title="测试时间 (Time)",
            showgrid=True,
            fixedrange=False,  # 允许横轴缩放
            rangeselector=dict(
                buttons=list([
                    dict(count=1, label="1天", step="day", stepmode="backward"),
                    dict(count=3, label="3天", step="day", stepmode="backward"),
                    dict(count=7, label="7天", step="day", stepmode="backward"),
                    dict(step="all", label="全部")
                ])
            )
        )
        if show_slider:
            xaxis_cfg['rangeslider'] = dict(visible=True, thickness=0.08)
        else:
            xaxis_cfg['rangeslider'] = dict(visible=False)

        fig_run.update_layout(
            title=f"[{run_item}] 时序趋势散点图 (Run Chart {' - WebGL GPU 加速' if use_gl else ''})",
            dragmode='zoom',  # 默认激活 2D 双向矩形拉框放大工具
            xaxis=xaxis_cfg,
            yaxis=y_axis_cfg,
            legend=dict(x=0.01, y=0.99, orientation="h"),
            height=540,
            margin=dict(l=40, r=40, t=50, b=50)
        )
        st.plotly_chart(
            fig_run,
            use_container_width=True,
            config={
                'scrollZoom': True,
                'displaylogo': False,
                'modeBarButtonsToRemove': ['lasso2d'], # 剔除极卡且易误触的套索工具
                'responsive': True
            }
        )


# ----------------------------------------------------
# Tab 4: 机台与批次对比分析 (Boxplot)
# ----------------------------------------------------
with tab_box:
    st.subheader("机台差异与批次离散度箱线图对比 (Multi-Factor Boxplot Comparison)")
    
    col_b1, col_b2, col_b3 = st.columns([1.5, 1.5, 1])
    with col_b1:
        box_item = st.selectbox(
            "选择箱线图分析测试项：",
            options=valid_test_items,
            index=valid_test_items.index(selected_item) if selected_item in valid_test_items else 0,
            key="boxplot_selector"
        )
    with col_b2:
        group_dim = st.radio(
            "选择对比维度 (Group Dimension)：",
            options=["TesterName (机台间对比)", "Batch_ID (批次间对比)", "CHECK_HOLDER (工位/治具夹具)"],
            horizontal=True
        )
        split_field = "TesterName" if "TesterName" in group_dim else ("Batch_ID" if "Batch_ID" in group_dim else "CHECK_HOLDER")
    with col_b3:
        box_pts_opt = st.selectbox(
            "箱线图离群点显示：",
            ["仅离群点 (outliers)", "无散点/纯箱体 (大数据最快)", "全部点 (all)"],
            index=0,
            help="若单批次数据达数万行，选择'无散点/纯箱体'可消除散点 DOM，秒级完成渲染与框选"
        )
        points_setting = "outliers" if "outliers" in box_pts_opt else (False if "无散点" in box_pts_opt else "all")

    box_df = df_working[['SerialNumber', split_field, box_item, 'TestResult']].copy()
    box_df['Numeric_Value'] = pd.to_numeric(box_df[box_item], errors='coerce')
    box_df = box_df.dropna(subset=['Numeric_Value', split_field])

    if len(box_df) == 0:
        st.warning("当前维度下暂无有效对比数据。")
    else:
        fig_box = px.box(
            box_df,
            x=split_field,
            y='Numeric_Value',
            color=split_field,
            points=points_setting,
            title=f"不同 {split_field} 之间的 [{box_item}] 数据分布离散度与中位数箱线图",
            labels={'Numeric_Value': f'{box_item} 数值', split_field: split_field}
        )

        item_spec = spec_dict.get(box_item, {})
        if item_spec.get('lsl') is not None:
            fig_box.add_hline(y=item_spec['lsl'], line_dash="dash", line_color="#EF4444", annotation_text=f"LSL: {item_spec['lsl']}")
        if item_spec.get('usl') is not None:
            fig_box.add_hline(y=item_spec['usl'], line_dash="dash", line_color="#EF4444", annotation_text=f"USL: {item_spec['usl']}")

        fig_box.update_layout(height=480, margin=dict(l=40, r=40, t=50, b=50))
        st.plotly_chart(
            fig_box,
            use_container_width=True,
            config={
                'scrollZoom': True,
                'displaylogo': False,
                'modeBarButtonsToRemove': ['lasso2d'],
                'responsive': True
            }
        )

        # 分组统计表
        st.markdown(f"##### 📊 各 {split_field} 统计参数与良率明细")
        grp_summary = []
        for name, sub_df in box_df.groupby(split_field):
            vals = sub_df['Numeric_Value']
            pass_rate = (sub_df['TestResult'].str.upper() == 'PASS').mean() * 100.0 if 'TestResult' in sub_df.columns else np.nan
            grp_summary.append({
                split_field: str(name),
                "样本数 (N)": len(vals),
                "均值 (Mean)": vals.mean(),
                "标准差 (Std)": vals.std(ddof=1),
                "中位数 (Median)": vals.median(),
                "四分位距 (IQR)": vals.quantile(0.75) - vals.quantile(0.25),
                "最小值 (Min)": vals.min(),
                "最大值 (Max)": vals.max(),
                "分组良率 (%)": pass_rate
            })
        st.dataframe(
            pd.DataFrame(grp_summary).style.format({
                "均值 (Mean)": "{:.4f}",
                "标准差 (Std)": "{:.4f}",
                "中位数 (Median)": "{:.4f}",
                "四分位距 (IQR)": "{:.4f}",
                "最小值 (Min)": "{:.4f}",
                "最大值 (Max)": "{:.4f}",
                "分组良率 (%)": "{:.2f}%"
            }),
            use_container_width=True
        )


# ----------------------------------------------------
# Tab 5: 数据明细与一键 Excel 导出
# ----------------------------------------------------
with tab_export:
    st.subheader("清洗后数据集预览与一键 Excel 导出 (Data Export)")
    st.markdown("支持将经过时间切片、批次筛选、机台过滤以及 **重测记录清洗去重** 后的完整数据工程包一键导出为 Excel 格式，包含测试明细与全项 CPK 汇总。")

    # 全测试项 CPK 批量扫描表
    st.markdown("##### 🔍 筛选数据集全测试项制程能力汇总 (CPK Overview)")
    all_cpk_list = []
    for it in valid_test_items:
        raw_s = pd.to_numeric(df_working[it], errors='coerce').dropna()
        it_spec = spec_dict[it]
        it_lsl = it_spec['lsl']
        it_usl = it_spec['usl']
        
        # 默认采用规格限内过滤，排除未供电0读数等严重失真离群点
        cond = pd.Series(True, index=raw_s.index)
        is_zero_lsl = (it_lsl == 0.0 and (it_usl is not None and it_usl > 0))
        if it_lsl is not None and not is_zero_lsl:
            cond = cond & (raw_s >= it_lsl)
        elif it_lsl is not None and is_zero_lsl:
            cond = cond & (raw_s >= 0)
        if it_usl is not None:
            cond = cond & (raw_s <= it_usl)
            
        s_data = raw_s[cond] if len(raw_s[cond]) >= 2 else raw_s
        it_cpk = calculate_process_capability(s_data, it_lsl, it_usl, is_smaller_the_better=is_zero_lsl)
        
        all_cpk_list.append({
            "测试项目": it,
            "规格下限(LSL)": it_spec['lsl_str'],
            "规格上限(USL)": it_spec['usl_str'],
            "有效样本量": it_cpk['n'],
            "平均值": round(it_cpk['mean'], 4) if not math.isnan(it_cpk['mean']) else None,
            "组内标准差(σ_w)": round(it_cpk['std_within'], 4) if not math.isnan(it_cpk['std_within']) else None,
            "总体标准差(s)": round(it_cpk['std'], 4) if not math.isnan(it_cpk['std']) else None,
            "Cpk (组内能力)": round(it_cpk['cpk'], 3) if not math.isnan(it_cpk['cpk']) else "N/A",
            "Ppk (总体性能)": round(it_cpk['ppk'], 3) if not math.isnan(it_cpk['ppk']) else "N/A",
            "能力评定": it_cpk['status']
        })
    df_all_cpk = pd.DataFrame(all_cpk_list)
    st.dataframe(df_all_cpk, use_container_width=True, height=280)

    st.markdown("---")
    st.markdown("##### 📦 导出数据工程包 (多通道双重保障，彻底解决未发现文件问题)")

    # 导出状态持久化
    if "export_excel_bytes" not in st.session_state:
        st.session_state.export_excel_bytes = None
    if "export_csv_bytes" not in st.session_state:
        st.session_state.export_csv_bytes = None
    if "last_saved_path" not in st.session_state:
        st.session_state.last_saved_path = None

    ts_now = datetime.now().strftime('%Y%m%d_%H%M%S')
    default_excel_name = f"FCT3_Report_{ts_now}.xlsx"
    default_csv_name = f"FCT3_Cleaned_{ts_now}.csv"

    c_loc, c_gen = st.columns([1.5, 1.5])
    with c_loc:
        save_target = st.radio(
            "💾 本地直存目标路径：",
            ["Windows【下载】文件夹 (Downloads)", "Windows【桌面】(Desktop)", "程序当前所在文件夹"],
            index=0,
            horizontal=False,
            help="直接由本机 Python 保存到指定位置，速度极快且 100% 成功，彻底避免浏览器下载网络拦截与404报错"
        )
        folder_key = "Downloads" if "下载" in save_target else ("Desktop" if "桌面" in save_target else "Current")

    with c_gen:
        st.write("")
        st.write("")
        if st.button("🚀 生成 / 刷新导出数据包 (Excel + CSV)", type="primary", use_container_width=True):
            with st.spinner("正在生成 Excel 与 CSV 导出数据，请稍候..."):
                try:
                    # 1. 生成 CSV
                    st.session_state.export_csv_bytes = df_working.drop(columns=['Parsed_Time'], errors='ignore').to_csv(index=False).encode('utf-8-sig')
                    
                    # 2. 生成 Excel
                    buf = io.BytesIO()
                    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                        export_df = clean_excel_data(df_working.drop(columns=['Parsed_Time'], errors='ignore'))
                        export_df.to_excel(writer, sheet_name='测试清洗明细数据', index=False)
                        clean_excel_data(df_all_cpk).to_excel(writer, sheet_name='全项CPK与制程能力', index=False)
                        if len(fail_item_counter) > 0:
                            clean_excel_data(pareto_df).to_excel(writer, sheet_name='不良项频次统计', index=False)
                    st.session_state.export_excel_bytes = buf.getvalue()
                    st.success("🎉 数据包已就绪！请选择下方任意方式导出。")
                except Exception as e:
                    st.error(f"生成数据包遇到异常: {e}")

    # 预加载轻量 CSV 字节流
    if st.session_state.export_csv_bytes is None:
        st.session_state.export_csv_bytes = df_working.drop(columns=['Parsed_Time'], errors='ignore').to_csv(index=False).encode('utf-8-sig')

    st.markdown("---")
    col_opt1, col_opt2 = st.columns(2)

    with col_opt1:
        st.markdown("###### 方案 A：【💾 本地直接保存】(强烈推荐，100% 成功)")
        st.caption("由本软件直接写盘保存到您的电脑，零浏览器中转，永不报错。")
        
        btn_save_csv = st.button("💾 直接保存 CSV 到电脑", key="btn_save_csv_disk", use_container_width=True)
        if btn_save_csv and st.session_state.export_csv_bytes:
            saved_p = save_file_to_local_disk(st.session_state.export_csv_bytes, default_csv_name, folder_key)
            st.session_state.last_saved_path = saved_p
            st.success(f"✅ CSV 已成功保存到：\n`{saved_p}`")

        btn_save_excel = st.button("💾 直接保存 Excel 到电脑", key="btn_save_xlsx_disk", use_container_width=True)
        if btn_save_excel:
            if st.session_state.export_excel_bytes is None:
                with st.spinner("正在后台构建 Excel，请稍候..."):
                    buf = io.BytesIO()
                    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                        export_df = clean_excel_data(df_working.drop(columns=['Parsed_Time'], errors='ignore'))
                        export_df.to_excel(writer, sheet_name='测试清洗明细数据', index=False)
                        clean_excel_data(df_all_cpk).to_excel(writer, sheet_name='全项CPK与制程能力', index=False)
                        if len(fail_item_counter) > 0:
                            clean_excel_data(pareto_df).to_excel(writer, sheet_name='不良项频次统计', index=False)
                    st.session_state.export_excel_bytes = buf.getvalue()
            
            saved_p = save_file_to_local_disk(st.session_state.export_excel_bytes, default_excel_name, folder_key)
            st.session_state.last_saved_path = saved_p
            st.success(f"✅ Excel 已成功保存到：\n`{saved_p}`")

        if st.session_state.last_saved_path and os.path.exists(st.session_state.last_saved_path):
            if st.button("📂 在 Windows 文件夹中高亮查看文件", use_container_width=True):
                reveal_file_in_explorer(st.session_state.last_saved_path)

    with col_opt2:
        st.markdown("###### 方案 B：【🌐 浏览器离线无损直链下载】(Base64 模式)")
        st.caption("数据内嵌于网页中，不经过服务端临时端点，彻底杜绝 404 '未发现文件'。")

        # CSV 离线超链接
        csv_link_html = make_b64_download_link(
            st.session_state.export_csv_bytes,
            default_csv_name,
            "text/csv;charset=utf-8",
            "⬇️ 浏览器下载 CSV 文件 (.csv)",
            bg_color="#059669"
        )
        st.markdown(csv_link_html, unsafe_allow_html=True)

        # Excel 离线超链接
        if st.session_state.export_excel_bytes is not None:
            excel_link_html = make_b64_download_link(
                st.session_state.export_excel_bytes,
                default_excel_name,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "⬇️ 浏览器下载 Excel 工程包 (.xlsx)",
                bg_color="#2563EB"
            )
            st.markdown(excel_link_html, unsafe_allow_html=True)
        else:
            st.info("💡 如需通过浏览器下载完整 Excel，请先点击上方的【🚀 生成 / 刷新导出数据包】。")

    st.markdown("---\n")
    st.markdown("##### 📋 清洗后明细数据预览 (前 50 条)")
    st.dataframe(df_working.drop(columns=['Parsed_Time'], errors='ignore').head(50), use_container_width=True)

