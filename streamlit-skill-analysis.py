import os
import time
from datetime import datetime
from typing import List, Tuple

import pandas as pd
import streamlit as st
from streamlit_autorefresh import st_autorefresh
from streamlit_echarts import st_echarts
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots

# -------------------- 页面配置 --------------------
st.set_page_config(
    page_title="技能覆盖分析大屏", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------- 页面样式（完全重写） --------------------
PAGE_CSS = """
<style>
/* 重置所有元素的盒模型 */
* {
    box-sizing: border-box;
    margin: 0;
    padding: 0;
}

/* 主容器样式 */
.main .block-container {
    padding-top: 2rem !important;
    padding-bottom: 2rem !important;
    max-width: 100% !important;
}

/* 整体背景 - 深色渐变 */
.stApp {
    background: linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #0f172a 100%) !important;
    background-attachment: fixed !important;
    color: #f1f5f9 !important;
}

/* 标题样式 */
h1, h2, h3 {
    color: #e2e8f0 !important;
    font-weight: 700 !important;
    text-shadow: 0 2px 4px rgba(0, 0, 0, 0.3) !important;
}

h1 {
    background: linear-gradient(90deg, #60a5fa, #3b82f6) !important;
    -webkit-background-clip: text !important;
    -webkit-text-fill-color: transparent !important;
    background-clip: text !important;
    font-size: 2.5rem !important;
    margin-bottom: 1.5rem !important;
    border-bottom: 2px solid rgba(59, 130, 246, 0.3) !important;
    padding-bottom: 0.5rem !important;
}

/* 侧边栏 - 深色卡片效果 */
section[data-testid="stSidebar"] {
    background: rgba(15, 23, 42, 0.95) !important;
    backdrop-filter: blur(10px) !important;
    border-right: 1px solid rgba(148, 163, 184, 0.2) !important;
}

section[data-testid="stSidebar"] > div {
    background: transparent !important;
}

section[data-testid="stSidebar"] .stSelectbox,
section[data-testid="stSidebar"] .stMultiSelect,
section[data-testid="stSidebar"] .stRadio,
section[data-testid="stSidebar"] .stButton {
    margin-bottom: 1rem !important;
}

/* 侧边栏标题样式 */
.sidebar-title {
    color: #60a5fa !important;
    font-weight: 700 !important;
    font-size: 1.1rem !important;
    margin: 1.5rem 0 0.8rem 0 !important;
    padding-bottom: 0.5rem !important;
    border-bottom: 1px solid rgba(96, 165, 250, 0.3) !important;
    text-transform: uppercase !important;
    letter-spacing: 0.05em !important;
}

.sidebar-title:first-child {
    margin-top: 0.5rem !important;
}

/* 按钮样式 - 现代化设计 */
div.stButton > button {
    background: linear-gradient(135deg, #3b82f6 0%, #1d4ed8 100%) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    padding: 0.75rem 1.5rem !important;
    font-weight: 600 !important;
    font-size: 0.95rem !important;
    transition: all 0.3s ease !important;
    box-shadow: 0 4px 6px rgba(59, 130, 246, 0.25) !important;
    width: 100% !important;
    position: relative !important;
    overflow: hidden !important;
}

div.stButton > button:hover {
    background: linear-gradient(135deg, #2563eb 0%, #1e40af 100%) !important;
    transform: translateY(-2px) !important;
    box-shadow: 0 6px 12px rgba(59, 130, 246, 0.35) !important;
}

div.stButton > button:active {
    transform: translateY(0) !important;
}

/* 危险按钮样式 */
.danger-button div.stButton > button {
    background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%) !important;
    box-shadow: 0 4px 6px rgba(239, 68, 68, 0.25) !important;
}

.danger-button div.stButton > button:hover {
    background: linear-gradient(135deg, #dc2626 0%, #b91c1c 100%) !important;
    box-shadow: 0 6px 12px rgba(239, 68, 68, 0.35) !important;
}

/* 次要按钮样式 */
.secondary-button div.stButton > button {
    background: linear-gradient(135deg, #64748b 0%, #475569 100%) !important;
    box-shadow: 0 4px 6px rgba(100, 116, 139, 0.25) !important;
}

.secondary-button div.stButton > button:hover {
    background: linear-gradient(135deg, #475569 0%, #334155 100%) !important;
    box-shadow: 0 6px 12px rgba(100, 116, 139, 0.35) !important;
}

/* 卡片样式 - 玻璃态效果 */
.metric-card {
    background: rgba(30, 41, 59, 0.7) !important;
    backdrop-filter: blur(10px) !important;
    border: 1px solid rgba(255, 255, 255, 0.1) !important;
    border-radius: 16px !important;
    padding: 1.5rem !important;
    text-align: center !important;
    box-shadow: 0 8px 32px rgba(0, 0, 0, 0.2) !important;
    transition: all 0.3s ease !important;
    height: 100% !important;
    display: flex !important;
    flex-direction: column !important;
    justify-content: center !important;
}

.metric-card:hover {
    transform: translateY(-4px) !important;
    box-shadow: 0 12px 48px rgba(0, 0, 0, 0.3) !important;
    border-color: rgba(96, 165, 250, 0.3) !important;
}

.metric-value {
    font-size: 2.25rem !important;
    font-weight: 800 !important;
    background: linear-gradient(135deg, #60a5fa 0%, #3b82f6 100%) !important;
    -webkit-background-clip: text !important;
    -webkit-text-fill-color: transparent !important;
    background-clip: text !important;
    margin-bottom: 0.5rem !important;
    line-height: 1 !important;
}

.metric-label {
    font-size: 0.875rem !important;
    color: #94a3b8 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.1em !important;
    font-weight: 600 !important;
    margin-bottom: 0.25rem !important;
}

.metric-subtext {
    font-size: 0.75rem !important;
    color: #64748b !important;
    margin-top: 0.25rem !important;
}

/* 数据表格样式 */
[data-testid="stDataFrame"] {
    background: rgba(30, 41, 59, 0.7) !important;
    border-radius: 12px !important;
    border: 1px solid rgba(255, 255, 255, 0.1) !important;
    overflow: hidden !important;
}

/* 选择框和输入框样式 */
.stSelectbox > div > div,
.stMultiSelect > div > div,
.stTextInput > div > div {
    background: rgba(30, 41, 59, 0.9) !important;
    border: 1px solid rgba(148, 163, 184, 0.3) !important;
    border-radius: 8px !important;
    color: #f1f5f9 !important;
}

.stSelectbox > div > div:hover,
.stMultiSelect > div > div:hover {
    border-color: #60a5fa !important;
}

/* 单选按钮样式 */
.stRadio > div {
    background: rgba(30, 41, 59, 0.7) !important;
    border-radius: 10px !important;
    padding: 0.75rem !important;
    border: 1px solid rgba(148, 163, 184, 0.2) !important;
}

/* 警告和信息框样式 */
.stAlert {
    background: rgba(30, 41, 59, 0.8) !important;
    border: 1px solid rgba(148, 163, 184, 0.2) !important;
    border-radius: 10px !important;
    border-left: 4px solid !important;
}

.stAlert[data-testid="stSuccess"] {
    border-left-color: #10b981 !important;
}

.stAlert[data-testid="stWarning"] {
    border-left-color: #f59e0b !important;
}

.stAlert[data-testid="stError"] {
    border-left-color: #ef4444 !important;
}

.stAlert[data-testid="stInfo"] {
    border-left-color: #3b82f6 !important;
}

/* 分隔线 */
hr {
    border: none !important;
    height: 1px !important;
    background: linear-gradient(90deg, 
        transparent, 
        rgba(148, 163, 184, 0.3), 
        transparent) !important;
    margin: 2rem 0 !important;
}

/* 图表容器 */
[data-testid="stPlotlyChart"],
[data-testid="stECharts"] {
    background: rgba(30, 41, 59, 0.7) !important;
    border-radius: 16px !important;
    padding: 1rem !important;
    border: 1px solid rgba(255, 255, 255, 0.1) !important;
}

/* 页脚样式 */
footer {
    text-align: center !important;
    color: #64748b !important;
    font-size: 0.875rem !important;
    padding-top: 2rem !important;
    margin-top: 2rem !important;
    border-top: 1px solid rgba(148, 163, 184, 0.2) !important;
}

/* 加载动画 */
.stSpinner > div {
    border-color: #3b82f6 transparent transparent transparent !important;
}

/* 滚动条样式 */
::-webkit-scrollbar {
    width: 8px;
    height: 8px;
}

::-webkit-scrollbar-track {
    background: rgba(30, 41, 59, 0.5);
    border-radius: 4px;
}

::-webkit-scrollbar-thumb {
    background: linear-gradient(135deg, #3b82f6, #60a5fa);
    border-radius: 4px;
}

::-webkit-scrollbar-thumb:hover {
    background: linear-gradient(135deg, #2563eb, #3b82f6);
}

/* 工具提示 */
[data-tooltip] {
    position: relative !important;
}

[data-tooltip]:hover::before {
    content: attr(data-tooltip) !important;
    position: absolute !important;
    bottom: 100% !important;
    left: 50% !important;
    transform: translateX(-50%) !important;
    background: rgba(15, 23, 42, 0.95) !important;
    color: #f1f5f9 !important;
    padding: 0.5rem 1rem !important;
    border-radius: 6px !important;
    font-size: 0.875rem !important;
    white-space: nowrap !important;
    border: 1px solid rgba(148, 163, 184, 0.2) !important;
    z-index: 1000 !important;
}

/* 响应式设计 */
@media (max-width: 768px) {
    .main .block-container {
        padding: 1rem !important;
    }
    
    h1 {
        font-size: 2rem !important;
    }
    
    .metric-card {
        padding: 1rem !important;
    }
    
    .metric-value {
        font-size: 1.75rem !important;
    }
}
</style>
"""
st.markdown(PAGE_CSS, unsafe_allow_html=True)

# -------------------- 配色方案 --------------------
COLOR_SCHEME = {
    'primary': ['#3b82f6', '#2563eb', '#1d4ed8', '#1e40af'],  # 蓝色系
    'secondary': ['#10b981', '#059669', '#047857', '#065f46'],  # 绿色系
    'accent': ['#8b5cf6', '#7c3aed', '#6d28d9', '#5b21b6'],  # 紫色系
    'warning': ['#f59e0b', '#d97706', '#b45309', '#92400e'],  # 橙色系
    'danger': ['#ef4444', '#dc2626', '#b91c1c', '#991b1b'],  # 红色系
    'neutral': ['#64748b', '#475569', '#334155', '#1e293b'],  # 灰色系
}

# 图表配色序列
CHART_COLORS = [
    '#3b82f6', '#10b981', '#8b5cf6', '#f59e0b', '#ef4444',  # 主色
    '#06b6d4', '#84cc16', '#ec4899', '#f97316', '#6366f1',  # 辅色
    '#14b8a6', '#f43f5e', '#a855f7', '#eab308', '#22c55e',  # 点缀色
]

SAVE_FILE = "jixiao.xlsx"   # 固定保存的文件

# -------------------- 数据导入函数 --------------------
@st.cache_data
def load_sheets(file, ts=None) -> Tuple[List[str], dict]:
    try:
        xpd = pd.ExcelFile(file)
    except Exception as e:
        st.sidebar.error(f"❌ 无法读取Excel文件: {e}")
        return [], {}
    
    frames = {}
    for s in xpd.sheet_names:
        try:
            df0 = pd.read_excel(xpd, sheet_name=s, header=None, dtype=str)
            if df0.empty:
                continue

            # 判断是否是标准模板
            if "明细" in df0.iloc[0].astype(str).tolist() and df0.shape[0] > 1 and df0.iloc[1, 0] == "分组":
                df0.columns = df0.iloc[0].tolist()
                df0 = df0.drop(0).reset_index(drop=True)
            elif "明细" not in df0.columns and "明细" in df0.iloc[0].astype(str).tolist():
                df0.columns = df0.iloc[0].tolist()
                df0 = df0.drop(0).reset_index(drop=True)

            # 确保列名标准
            if not {"明细"}.issubset(df0.columns):
                continue

            # 检测分组行
            if df0.iloc[0, 0] == "分组":
                groups = df0.iloc[0, 1:].tolist()
                df0 = df0.drop(0).reset_index(drop=True)
                emp_cols = [c for c in df0.columns if c not in ["明细", "数量总和", "编号"]]
                group_map = {emp: groups[i] if i < len(groups) else None for i, emp in enumerate(emp_cols)}
                df_long = df0.melt(
                    id_vars=["明细", "数量总和"] if "数量总和" in df0.columns else ["明细"],
                    value_vars=emp_cols,
                    var_name="员工",
                    value_name="值"
                )
                df_long["值"] = pd.to_numeric(df_long["值"], errors='coerce').fillna(0)
                df_long["分组"] = df_long["员工"].map(group_map)
                df_long["时间点"] = s
                frames[s] = df_long
            else:
                if "时间点" not in df0.columns:
                    df0["时间点"] = s
                if "值" in df0.columns:
                    df0["值"] = pd.to_numeric(df0["值"], errors='coerce').fillna(0)
                frames[s] = df0
        except Exception as e:
            continue
    return xpd.sheet_names, frames

# -------------------- 文件读取 --------------------
sheets, sheet_frames = [], {}
try:
    if os.path.exists(SAVE_FILE):
        mtime = os.path.getmtime(SAVE_FILE)
        sheets, sheet_frames = load_sheets(SAVE_FILE, ts=mtime)
        st.sidebar.success(f"✅ 已加载数据文件")
    else:
        # 创建示例数据
        example_data = {
            "示例_2025_01": pd.DataFrame({
                "明细": ["任务A", "任务B", "任务C"],
                "数量总和": [3, 2, 5],
                "员工": ["张三", "李四", "王五"],
                "值": [1, 1, 1],
                "分组": ["A8", "B7", "VN"],
                "时间点": "示例_2025_01"
            })
        }
        with pd.ExcelWriter(SAVE_FILE, engine='openpyxl') as writer:
            for sheet_name, df0 in example_data.items():
                df0.to_excel(writer, sheet_name=sheet_name, index=False)
        
        sheets, sheet_frames = load_sheets(SAVE_FILE)
        st.sidebar.info("📁 创建了示例数据文件")

except Exception as e:
    st.sidebar.error(f"❌ 读取数据失败：{e}")

# -------------------- 删除功能 --------------------
st.sidebar.markdown('<div class="sidebar-title">🗑️ 删除时间点</div>', unsafe_allow_html=True)
if sheets:
    sheet_to_delete = st.sidebar.selectbox("选择要删除的时间点", sheets, key="delete_select", label_visibility="collapsed")
    
    col1, col2 = st.sidebar.columns(2)
    with col1:
        if st.button("🗑️ 删除", key="delete_btn", help="删除选中的时间点"):
            try:
                if not os.path.exists(SAVE_FILE):
                    st.sidebar.error("文件不存在")
                else:
                    xls = pd.ExcelFile(SAVE_FILE)
                    new_sheets = {}
                    
                    for sheet in xls.sheet_names:
                        if sheet != sheet_to_delete:
                            df0 = pd.read_excel(xls, sheet_name=sheet)
                            new_sheets[sheet] = df0
                    
                    with pd.ExcelWriter(SAVE_FILE, engine="openpyxl") as writer:
                        for sheet_name, df0 in new_sheets.items():
                            df0.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    st.cache_data.clear()
                    st.sidebar.success(f"✅ 已删除: {sheet_to_delete}")
                    time.sleep(1)
                    st.rerun()
            except Exception as e:
                st.sidebar.error(f"❌ 删除失败")
    
    with col2:
        if st.button("🔄 刷新", key="refresh_btn"):
            st.cache_data.clear()
            st.rerun()

# -------------------- 新增时间点功能 --------------------
st.sidebar.markdown('<div class="sidebar-title">📅 新增时间点</div>', unsafe_allow_html=True)
current_year = datetime.now().year
year = st.sidebar.selectbox("选择年份", list(range(current_year - 2, current_year + 2)), index=2, label_visibility="collapsed")
mode = st.sidebar.radio("时间类型", ["月份", "季度"], horizontal=True, label_visibility="collapsed")

if mode == "月份":
    month = st.sidebar.selectbox("选择月份", list(range(1, 13)), label_visibility="collapsed")
    new_sheet_name = f"{year}_{month:02d}"
else:
    quarter = st.sidebar.selectbox("选择季度", ["Q1", "Q2", "Q3", "Q4"], label_visibility="collapsed")
    new_sheet_name = f"{year}_{quarter}"

# 新增数据保存函数
def save_new_sheet(sheet_name, df_data):
    """安全保存新的sheet到Excel文件"""
    try:
        if os.path.exists(SAVE_FILE):
            from openpyxl import load_workbook
            wb = load_workbook(SAVE_FILE)
            
            if sheet_name in wb.sheetnames:
                st.sidebar.error(f"❌ 时间点已存在！")
                return False
            
            with pd.ExcelWriter(SAVE_FILE, engine='openpyxl') as writer:
                writer.book = wb
                writer.sheets = {ws.title: ws for ws in wb.worksheets}
                df_data.to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            with pd.ExcelWriter(SAVE_FILE, engine='openpyxl') as writer:
                df_data.to_excel(writer, sheet_name=sheet_name, index=False)
        
        return True
    except Exception as e:
        st.sidebar.error(f"❌ 保存失败")
        return False

if st.sidebar.button("🚀 创建新时间点", type="primary"):
    if new_sheet_name in sheets:
        st.sidebar.error(f"❌ 时间点已存在！")
    else:
        try:
            # 自动继承逻辑
            base_df = pd.DataFrame(columns=["明细", "数量总和", "员工", "值", "分组", "时间点"])
            
            if mode == "月份" and month == 12:
                old_dec_sheets = [s for s in sheets if s.endswith("_12")]
                for old_sheet in old_dec_sheets:
                    try:
                        xls = pd.ExcelFile(SAVE_FILE)
                        new_sheets_data = {}
                        for sheet in xls.sheet_names:
                            if sheet != old_sheet:
                                df0 = pd.read_excel(xls, sheet_name=sheet)
                                new_sheets_data[sheet] = df0
                        
                        with pd.ExcelWriter(SAVE_FILE, engine="openpyxl") as writer:
                            for sheet_name, df0 in new_sheets_data.items():
                                df0.to_excel(writer, sheet_name=sheet_name, index=False)
                    except:
                        pass

            prev_sheets = sorted([s for s in sheets if s.split("_")[0] == str(year) and s < new_sheet_name])
            
            if not prev_sheets:
                prev_years = sorted([int(s.split("_")[0]) for s in sheets if s.split("_")[0].isdigit()])
                if prev_years:
                    latest_prev_year = max(y for y in prev_years if y < year) if any(y < year for y in prev_years) else None
                    if latest_prev_year:
                        prev_sheets = sorted([s for s in sheets if s.startswith(str(latest_prev_year))])

            if prev_sheets:
                prev_name = prev_sheets[-1]
                base_df = sheet_frames.get(prev_name, base_df).copy()
                if "值" in base_df.columns:
                    base_df["值"] = 0
                    if "明细" in base_df.columns:
                        sum_df = (
                            base_df.groupby("明细", as_index=False)["值"].sum()
                            .rename(columns={"值": "数量总和"})
                        )
                        base_df = base_df.drop(columns=["数量总和"], errors="ignore")
                        base_df = base_df.merge(sum_df, on="明细", how="left")
                base_df["时间点"] = new_sheet_name
                st.sidebar.info(f"📋 已从 {prev_name} 继承结构")
            else:
                base_df = pd.DataFrame({
                    "明细": ["示例任务1", "示例任务2", "示例任务3"],
                    "数量总和": [0, 0, 0],
                    "员工": ["员工A", "员工B", "员工C"],
                    "值": [0, 0, 0],
                    "分组": ["分组A", "分组B", "分组C"],
                    "时间点": new_sheet_name
                })

            if save_new_sheet(new_sheet_name, base_df):
                st.cache_data.clear()
                st.sidebar.success(f"✅ 已创建: {new_sheet_name}")
                if mode == "月份" and month == 12:
                    st.sidebar.success("♻️ 已清理旧数据")
                time.sleep(1)
                st.rerun()

        except Exception as e:
            st.sidebar.error(f"❌ 创建失败")

# -------------------- 数据修复工具 --------------------
st.sidebar.markdown('<div class="sidebar-title">⚙️ 数据修复工具</div>', unsafe_allow_html=True)

if st.sidebar.button("🧮 一键更新所有数量总和", type="secondary"):
    try:
        if not os.path.exists(SAVE_FILE):
            st.sidebar.warning("未找到数据文件")
        else:
            xls = pd.ExcelFile(SAVE_FILE)
            updated_frames = {}
            
            with st.spinner("正在更新数据..."):
                for sheet_name in xls.sheet_names:
                    df0 = pd.read_excel(xls, sheet_name=sheet_name)
                    if "明细" in df0.columns and "值" in df0.columns:
                        sum_df = (
                            df0.groupby("明细", as_index=False)["值"].sum()
                            .rename(columns={"值": "数量总和"})
                        )
                        df0 = df0.drop(columns=["数量总和"], errors="ignore")
                        df0 = df0.merge(sum_df, on="明细", how="left")
                        if "时间点" not in df0.columns:
                            df0["时间点"] = sheet_name
                        updated_frames[sheet_name] = df0
                    else:
                        updated_frames[sheet_name] = df0

                with pd.ExcelWriter(SAVE_FILE, engine="openpyxl") as writer:
                    for sheet_name, df0 in updated_frames.items():
                        df0.to_excel(writer, sheet_name=sheet_name, index=False)

                st.cache_data.clear()
                st.sidebar.success("✅ 数量总和已更新！")
                time.sleep(1)
                st.rerun()

    except Exception as e:
        st.sidebar.error(f"❌ 更新失败")

# -------------------- 时间点选择 --------------------
st.sidebar.markdown('<div class="sidebar-title">📊 选择时间点</div>', unsafe_allow_html=True)

if sheets:
    all_time_points = sorted(sheets, reverse=True)
    time_choice = st.sidebar.multiselect(
        "选择月份/季度", 
        all_time_points, 
        default=all_time_points[:1] if all_time_points else [],
        key="time_select",
        label_visibility="collapsed"
    )
    
    if time_choice:
        dfs = []
        for t in time_choice:
            df0 = sheet_frames.get(t)
            if df0 is not None:
                dfs.append(df0)
        
        if dfs:
            combined_df = pd.concat(dfs, ignore_index=True)
            all_groups = combined_df["分组"].dropna().unique().tolist() if "分组" in combined_df.columns else []
            selected_groups = st.sidebar.multiselect(
                "选择分组", 
                all_groups, 
                default=all_groups,
                key="group_select",
                label_visibility="collapsed"
            )
        else:
            selected_groups = []
    else:
        selected_groups = []
else:
    time_choice = []
    selected_groups = []
    st.sidebar.warning("暂无数据")

# -------------------- 视图选择 --------------------
st.sidebar.markdown('<div class="sidebar-title">👁️ 视图选择</div>', unsafe_allow_html=True)
sections_names = [
    "人员完成任务数量排名",
    "任务对比（堆叠柱状图）",
    "任务掌握情况（热门任务）",
    "任务-人员热力图"
]
view = st.sidebar.radio(
    "切换视图", 
    ["编辑数据", "大屏轮播", "单页模式", "显示所有视图", "能力分析"],
    horizontal=False,
    key="view_select",
    label_visibility="collapsed"
)

# -------------------- 数据合并函数 --------------------
def get_merged_df(keys: List[str], groups: List[str]) -> pd.DataFrame:
    dfs = []
    for k in keys:
        df0 = sheet_frames.get(k)
        if df0 is not None and not df0.empty:
            if groups and "分组" in df0.columns and len(groups) > 0:
                df0 = df0[df0["分组"].isin(groups)]
            if "时间点" not in df0.columns:
                df0["时间点"] = k
            dfs.append(df0)
    
    if not dfs:
        return pd.DataFrame()
    
    merged_df = pd.concat(dfs, axis=0, ignore_index=True)
    
    if "值" in merged_df.columns:
        merged_df["值"] = pd.to_numeric(merged_df["值"], errors='coerce').fillna(0)
    
    return merged_df

df = get_merged_df(time_choice, selected_groups)

# -------------------- 图表函数 --------------------
def get_chart_color(idx):
    return CHART_COLORS[idx % len(CHART_COLORS)]

def chart_total(df0):
    if df0.empty:
        return go.Figure()
    
    if "明细" in df0.columns:
        df0 = df0[df0["明细"] != "分数总和"]
    
    if len(time_choice) > 1 and "时间点" in df0.columns:
        emp_time_stats = df0.groupby(["员工", "时间点"])["值"].sum().reset_index()
        fig = go.Figure()
        
        time_points = sorted(emp_time_stats["时间点"].unique())
        
        for i, time_point in enumerate(time_points):
            time_data = emp_time_stats[emp_time_stats["时间点"] == time_point]
            time_data = time_data.sort_values("值", ascending=False)
            
            fig.add_trace(go.Bar(
                x=time_data["员工"],
                y=time_data["值"],
                name=time_point,
                marker_color=get_chart_color(i),
                text=time_data["值"],
                textposition="outside",
                hovertemplate="员工: %{x}<br>时间点: %{customdata}<br>完成值: %{y}<extra></extra>",
                customdata=[time_point] * len(time_data)
            ))
        
        fig.update_layout(
            barmode='group',
            template="plotly_dark",
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='#e2e8f0',
            xaxis_title="员工",
            yaxis_title="完成总值",
            showlegend=True,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1,
                bgcolor='rgba(30, 41, 59, 0.8)',
                bordercolor='rgba(255, 255, 255, 0.2)',
                borderwidth=1
            )
        )
    else:
        emp_stats = df0.groupby("员工")["值"].sum().sort_values(ascending=False).reset_index()
        fig = go.Figure(go.Bar(
            x=emp_stats["员工"],
            y=emp_stats["值"],
            text=emp_stats["值"],
            textposition="outside",
            hovertemplate="员工: %{x}<br>完成总值: %{y}<extra></extra>",
            marker_color=CHART_COLORS[0],
            marker_line_width=0
        ))
        fig.update_layout(
            template="plotly_dark",
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='#e2e8f0',
            xaxis_title="员工",
            yaxis_title="完成总值",
            showlegend=False
        )
    
    return fig

def chart_stack(df0):
    if df0.empty:
        return go.Figure()
    
    if "明细" in df0.columns:
        df0 = df0[df0["明细"] != "分数总和"]
    
    if len(time_choice) > 1 and "时间点" in df0.columns:
        time_points = sorted(df0["时间点"].unique())
        
        if len(time_points) == 1:
            df_pivot = df0.pivot_table(index="明细", columns="员工", values="值", aggfunc="sum", fill_value=0)
            fig = go.Figure()
            for i, emp in enumerate(df_pivot.columns):
                fig.add_trace(go.Bar(
                    x=df_pivot.index, 
                    y=df_pivot[emp], 
                    name=emp,
                    marker_color=get_chart_color(i)
                ))
            fig.update_layout(
                barmode="stack", 
                template="plotly_dark",
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font_color='#e2e8f0',
                xaxis_title="任务", 
                yaxis_title="完成值",
                title=f"时间点: {time_points[0]}"
            )
        else:
            fig = make_subplots(
                rows=len(time_points), cols=1,
                subplot_titles=[f"时间点: {tp}" for tp in time_points],
                vertical_spacing=0.1
            )
            
            for i, tp in enumerate(time_points, 1):
                df_tp = df0[df0["时间点"] == tp]
                df_pivot = df_tp.pivot_table(index="明细", columns="员工", values="值", aggfunc="sum", fill_value=0)
                
                all_emps = df0["员工"].unique()
                
                for j, emp in enumerate(df_pivot.columns):
                    color_idx = list(all_emps).index(emp) % len(CHART_COLORS) if emp in all_emps else j
                    fig.add_trace(
                        go.Bar(
                            x=df_pivot.index, 
                            y=df_pivot[emp], 
                            name=emp,
                            marker_color=get_chart_color(color_idx),
                            showlegend=(i==1),
                            legendgroup=emp
                        ),
                        row=i, col=1
                    )
            
            fig.update_layout(
                barmode="stack", 
                template="plotly_dark",
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font_color='#e2e8f0',
                height=400*len(time_points),
                showlegend=True
            )
            fig.update_xaxes(title_text="任务", row=len(time_points), col=1)
            fig.update_yaxes(title_text="完成值", row=len(time_points)//2 + 1, col=1)
    else:
        df_pivot = df0.pivot_table(index="明细", columns="员工", values="值", aggfunc="sum", fill_value=0)
        fig = go.Figure()
        for i, emp in enumerate(df_pivot.columns):
            fig.add_trace(go.Bar(
                x=df_pivot.index, 
                y=df_pivot[emp], 
                name=emp,
                marker_color=get_chart_color(i)
            ))
        fig.update_layout(
            barmode="stack", 
            template="plotly_dark",
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='#e2e8f0',
            xaxis_title="任务", 
            yaxis_title="完成值"
        )
    
    return fig

def chart_hot(df0):
    if df0.empty:
        return {
            "backgroundColor": "transparent",
            "yAxis": {"type": "category", "data": [], "axisLabel": {"color": "#e2e8f0"}},
            "xAxis": {"type": "value", "axisLabel": {"color": "#e2e8f0"}},
            "series": [{"data": [], "type": "bar", "itemStyle": {"color": CHART_COLORS[3]}}]
        }
    
    if "明细" in df0.columns:
        df0 = df0[df0["明细"] != "分数总和"]
    
    if len(time_choice) > 1 and "时间点" in df0.columns:
        time_points = sorted(df0["时间点"].unique())
        tasks = df0["明细"].unique().tolist()[:15]
        
        option = {
            "backgroundColor": "transparent",
            "tooltip": {"trigger": "axis", "axisPointer": {"type": "shadow"}},
            "legend": {
                "data": time_points, 
                "textStyle": {"color": "#e2e8f0"},
                "top": "10px"
            },
            "grid": {"left": "3%", "right": "4%", "bottom": "3%", "containLabel": True},
            "xAxis": {
                "type": "value", 
                "axisLabel": {"color": "#e2e8f0"},
                "splitLine": {"lineStyle": {"color": "rgba(148, 163, 184, 0.2)"}}
            },
            "yAxis": {
                "type": "category", 
                "data": tasks, 
                "axisLabel": {"color": "#e2e8f0"},
                "axisLine": {"show": False},
                "axisTick": {"show": False}
            },
            "series": []
        }
        
        for i, tp in enumerate(time_points):
            df_tp = df0[df0["时间点"] == tp]
            ts = df_tp.groupby("明细")["员工"].nunique()
            ts_ordered = [ts.get(task, 0) for task in tasks]
            
            option["series"].append({
                "name": tp,
                "type": "bar",
                "data": ts_ordered,
                "itemStyle": {"color": get_chart_color(i)},
                "label": {"show": True, "position": "right", "color": "#e2e8f0"}
            })
    else:
        ts = df0.groupby("明细")["员工"].nunique().sort_values(ascending=False).head(15)
        option = {
            "backgroundColor": "transparent",
            "tooltip": {"trigger": "axis", "axisPointer": {"type": "shadow"}},
            "grid": {"left": "3%", "right": "4%", "bottom": "3%", "containLabel": True},
            "yAxis": {
                "type": "category", 
                "data": ts.index.tolist(), 
                "axisLabel": {"color": "#e2e8f0"},
                "axisLine": {"show": False},
                "axisTick": {"show": False}
            },
            "xAxis": {
                "type": "value", 
                "axisLabel": {"color": "#e2e8f0"},
                "splitLine": {"lineStyle": {"color": "rgba(148, 163, 184, 0.2)"}}
            },
            "series": [{
                "data": ts.tolist(), 
                "type": "bar", 
                "itemStyle": {"color": CHART_COLORS[3]},
                "label": {"show": True, "position": "right", "color": "#e2e8f0"}
            }]
        }
    
    return option

def chart_heat(df0):
    if df0.empty:
        return {
            "backgroundColor": "transparent",
            "tooltip": {"position": "top"},
            "xAxis": {"type": "category", "data": [], "axisLabel": {"color": "#e2e8f0"}},
            "yAxis": {"type": "category", "data": [], "axisLabel": {"color": "#e2e8f0"}},
            "visualMap": {
                "min": 0, 
                "max": 1, 
                "show": False, 
                "inRange": {"color": [CHART_COLORS[4], CHART_COLORS[1]]}
            },
            "series": [{"type": "heatmap", "data": []}]
        }
    
    if "明细" in df0.columns:
        df0 = df0[df0["明细"] != "分数总和"]
    
    if len(time_choice) > 1 and "时间点" in df0.columns:
        time_points = sorted(df0["时间点"].unique())
        
        option = {
            "baseOption": {
                "backgroundColor": "transparent",
                "tooltip": {"position": "top"},
                "visualMap": {
                    "min": 0, 
                    "max": 1, 
                    "show": True,
                    "orient": "vertical",
                    "left": "right",
                    "top": "center",
                    "textStyle": {"color": "#e2e8f0"},
                    "inRange": {"color": [CHART_COLORS[4], CHART_COLORS[1]]}
                },
                "timeline": {
                    "axisType": "category",
                    "autoPlay": False,
                    "playInterval": 2000,
                    "data": time_points,
                    "label": {"color": "#e2e8f0"},
                    "lineStyle": {"color": CHART_COLORS[0]},
                    "itemStyle": {"color": CHART_COLORS[0]},
                    "checkpointStyle": {"color": CHART_COLORS[0]},
                    "controlStyle": {"color": CHART_COLORS[0], "borderColor": CHART_COLORS[0]}
                },
                "series": [{"type": "heatmap"}],
                "title": {"text": "任务-人员热力图", "textStyle": {"color": "#e2e8f0"}}
            },
            "options": []
        }
        
        for tp in time_points:
            df_tp = df0[df0["时间点"] == tp]
            tasks = df_tp["明细"].unique().tolist()[:20]
            emps = df_tp["员工"].unique().tolist()[:20]
            data = []
            
            max_val = 0
            for i, t in enumerate(tasks):
                for j, e in enumerate(emps):
                    v = int(df_tp[(df_tp["明细"] == t) & (df_tp["员工"] == e)]["值"].sum())
                    data.append([j, i, v])
                    max_val = max(max_val, v)
            
            option["options"].append({
                "title": {"text": f"时间点: {tp}", "textStyle": {"color": "#e2e8f0"}},
                "xAxis": {
                    "type": "category", 
                    "data": emps, 
                    "axisLabel": {
                        "color": "#e2e8f0",
                        "rotate": 45,
                        "interval": 0
                    }
                },
                "yAxis": {
                    "type": "category", 
                    "data": tasks, 
                    "axisLabel": {"color": "#e2e8f0"}
                },
                "series": [{"type": "heatmap", "data": data}]
            })
        
        if max_val > 0:
            option["baseOption"]["visualMap"]["max"] = max_val
    else:
        tasks = df0["明细"].unique().tolist()[:20]
        emps = df0["员工"].unique().tolist()[:20]
        data = []
        
        max_val = 0
        for i, t in enumerate(tasks):
            for j, e in enumerate(emps):
                v = int(df0[(df0["明细"] == t) & (df0["员工"] == e)]["值"].sum())
                data.append([j, i, v])
                max_val = max(max_val, v)
        
        option = {
            "backgroundColor": "transparent",
            "tooltip": {"position": "top"},
            "xAxis": {
                "type": "category", 
                "data": emps, 
                "axisLabel": {
                    "color": "#e2e8f0",
                    "rotate": 45,
                    "interval": 0
                }
            },
            "yAxis": {
                "type": "category", 
                "data": tasks, 
                "axisLabel": {"color": "#e2e8f0"}
            },
            "visualMap": {
                "min": 0, 
                "max": max_val if max_val > 0 else 1, 
                "show": True,
                "orient": "vertical",
                "left": "right",
                "top": "center",
                "textStyle": {"color": "#e2e8f0"},
                "inRange": {"color": [CHART_COLORS[4], CHART_COLORS[1]]}
            },
            "series": [{"type": "heatmap", "data": data}]
        }
    
    return option

# -------------------- 卡片显示 --------------------
def show_cards(df0):
    if df0.empty:
        st.info("📭 暂无有效数据可展示")
        return
    
    if "明细" in df0.columns:
        df0 = df0[df0["明细"] != "分数总和"]
    
    total_tasks = df0["明细"].nunique()
    total_people = df0["员工"].nunique()
    ps = df0.groupby("员工")["值"].sum()
    top_person = ps.idxmax() if not ps.empty else ""
    top_value = ps.max() if not ps.empty else 0
    avg_score = round(ps.mean(), 1) if not ps.empty else 0
    
    time_points_display = ", ".join(time_choice) if time_choice else "未选择"
    
    c1, c2, c3, c4, c5 = st.columns(5)
    
    with c1:
        st.markdown(f"""
        <div class='metric-card'>
            <div class='metric-label'>任务数</div>
            <div class='metric-value'>{total_tasks}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with c2:
        st.markdown(f"""
        <div class='metric-card'>
            <div class='metric-label'>人数</div>
            <div class='metric-value'>{total_people}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with c3:
        st.markdown(f"""
        <div class='metric-card'>
            <div class='metric-label'>最高覆盖率</div>
            <div class='metric-value'>{top_person[:4] if len(top_person) > 4 else top_person}</div>
            <div class='metric-subtext'>值: {top_value}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with c4:
        st.markdown(f"""
        <div class='metric-card'>
            <div class='metric-label'>平均完成值</div>
            <div class='metric-value'>{avg_score}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with c5:
        st.markdown(f"""
        <div class='metric-card'>
            <div class='metric-label'>选择的时间点</div>
            <div style='font-size:1rem;margin:0.5rem 0;color:#60a5fa'>{len(time_choice)} 个</div>
            <div class='metric-subtext'>{time_points_display[:20]}{'...' if len(time_points_display) > 20 else ''}</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("<hr/>", unsafe_allow_html=True)

# -------------------- 主页面 --------------------
st.title("📊 技能覆盖分析大屏")

if view == "编辑数据":
    if not time_choice:
        st.warning("⚠️ 请在左侧选择时间点后再编辑数据")
    else:
        if len(time_choice) > 1:
            st.info(f"📅 当前编辑 {len(time_choice)} 个时间点")
        
        show_cards(df)
        
        if not df.empty:
            st.info("📝 直接编辑表格，完成后点击保存")
            
            edited_df = st.data_editor(
                df,
                num_rows="dynamic",
                use_container_width=True,
                column_config={
                    "值": st.column_config.NumberColumn(
                        "值",
                        help="任务完成值",
                        min_value=0,
                        max_value=100,
                        step=1,
                        format="%d"
                    ),
                    "时间点": st.column_config.TextColumn(
                        "时间点",
                        help="数据所属时间点",
                        disabled=True
                    )
                }
            )
            
            col1, col2 = st.columns([1, 3])
            with col1:
                if st.button("💾 保存修改", type="primary", use_container_width=True):
                    try:
                        if len(time_choice) == 1:
                            sheet_name = time_choice[0]
                            df_to_save = edited_df.drop(columns=["时间点"], errors="ignore")
                            
                            if "明细" in df_to_save.columns and "值" in df_to_save.columns:
                                sum_df = (
                                    df_to_save.groupby("明细", as_index=False)["值"].sum()
                                    .rename(columns={"值": "数量总和"})
                                )
                                df_to_save = df_to_save.drop(columns=["数量总和"], errors="ignore")
                                df_to_save = df_to_save.merge(sum_df, on="明细", how="left")
                            
                            with pd.ExcelWriter(SAVE_FILE, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                                df_to_save.to_excel(writer, sheet_name=sheet_name, index=False)
                            
                            st.success(f"✅ 已保存到 {sheet_name}")
                            st.cache_data.clear()
                            time.sleep(1)
                            st.rerun()
                    except Exception as e:
                        st.error(f"❌ 保存失败")
            
            with col2:
                if st.button("🔄 重置", type="secondary", use_container_width=True):
                    st.cache_data.clear()
                    st.rerun()
        else:
            st.info("📭 当前选择没有数据")

elif view == "大屏轮播":
    if not time_choice:
        st.warning("⚠️ 请在左侧选择时间点")
    else:
        st_autorefresh(interval=10000, key="aut")
        show_cards(df)
        
        if not df.empty:
            secs = [
                ("📊 完成排名", chart_total(df)),
                ("📈 任务对比", chart_stack(df)),
                ("🔥 热门任务", chart_hot(df)),
                ("🎨 热力图", chart_heat(df))
            ]
            idx = int(time.time() / 10) % len(secs)
            t, op = secs[idx]
            
            st.subheader(t)
            if isinstance(op, go.Figure):
                st.plotly_chart(op, use_container_width=True, theme="streamlit")
            else:
                st_echarts(op, height="600px", theme="dark")
        else:
            st.info("📭 当前选择没有数据")

elif view == "单页模式":
    if not time_choice:
        st.warning("⚠️ 请在左侧选择时间点")
    else:
        show_cards(df)
        
        if not df.empty:
            choice = st.sidebar.selectbox("单页查看", sections_names, index=0)
            mapping = {
                "人员完成任务数量排名": chart_total(df),
                "任务对比（堆叠柱状图）": chart_stack(df),
                "任务掌握情况（热门任务）": chart_hot(df),
                "任务-人员热力图": chart_heat(df)
            }
            chart_func = mapping.get(choice, chart_total(df))
            
            st.subheader(choice)
            if isinstance(chart_func, go.Figure):
                st.plotly_chart(chart_func, use_container_width=True, theme="streamlit")
            else:
                st_echarts(chart_func, height="600px", theme="dark")
        else:
            st.info("📭 当前选择没有数据")

elif view == "显示所有视图":
    if not time_choice:
        st.warning("⚠️ 请在左侧选择时间点")
    else:
        show_cards(df)
        
        if not df.empty:
            charts = [
                ("📊 完成排名", chart_total(df)),
                ("📈 任务对比（堆叠柱状图）", chart_stack(df)),
                ("🔥 热门任务", chart_hot(df)),
                ("🎨 热力图", chart_heat(df))
            ]
            for label, f in charts:
                st.subheader(label)
                if isinstance(f, go.Figure):
                    st.plotly_chart(f, use_container_width=True, theme="streamlit")
                else:
                    st_echarts(f, height="520px", theme="dark")
        else:
            st.info("📭 当前选择没有数据")

elif view == "能力分析":
    if not time_choice:
        st.warning("⚠️ 请在左侧选择时间点")
    else:
        st.subheader("📊 能力分析")
        
        if not df.empty:
            if len(time_choice) > 1:
                st.info(f"📊 当前分析 {len(time_choice)} 个时间点")
            
            employees = df["员工"].unique().tolist()
            selected_emps = st.sidebar.multiselect(
                "选择员工（图1显示）", 
                employees, 
                default=employees[:3] if employees else [],
                key="emp_select"
            )
            tasks = df["明细"].unique().tolist()
            
            fig1, fig2, fig3 = go.Figure(), go.Figure(), go.Figure()
            
            for idx, sheet in enumerate(time_choice):
                df_sheet = get_merged_df([sheet], selected_groups)
                if df_sheet.empty:
                    continue
                
                if "明细" in df_sheet.columns:
                    df_sheet = df_sheet[df_sheet["明细"] != "分数总和"]
                
                df_pivot = df_sheet.pivot_table(index="明细", columns="员工", values="值", fill_value=0)
                
                color = get_chart_color(idx)
                
                for emp in selected_emps:
                    if emp in df_pivot.columns:
                        fig1.add_trace(go.Scatter(
                            x=tasks, 
                            y=df_pivot[emp].reindex(tasks, fill_value=0),
                            mode="lines+markers", 
                            name=f"{sheet}-{emp}",
                            line=dict(color=color, width=2 if sheet == time_choice[-1] else 1),
                            opacity=0.7 if sheet != time_choice[-1] else 1,
                            showlegend=True if emp == selected_emps[0] else False,
                            legendgroup=sheet
                        ))
                
                task_sums = df_pivot.sum(axis=1).reindex(tasks, fill_value=0)
                fig2.add_trace(go.Scatter(
                    x=tasks, 
                    y=task_sums,
                    mode="lines+markers", 
                    name=sheet,
                    line=dict(color=color, width=3 if sheet == time_choice[-1] else 2),
                    marker=dict(size=8 if sheet == time_choice[-1] else 6)
                ))
                
                emp_sums = df_pivot.sum(axis=0)
                if not emp_sums.empty:
                    fig3.add_trace(go.Bar(
                        x=emp_sums.index,
                        y=emp_sums.values,
                        name=sheet,
                        marker_color=color,
                        opacity=0.7
                    ))
            
            for fig, title in [(fig1, "员工任务完成情况（多时间点对比）"), 
                              (fig2, "任务整体完成度趋势（多时间点对比）"), 
                              (fig3, "员工整体完成度对比（多时间点堆叠）")]:
                fig.update_layout(
                    title=title, 
                    template="plotly_dark",
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    font_color='#e2e8f0',
                    xaxis_title="员工" if fig == fig3 else "任务",
                    yaxis_title="完成值" + ("总和" if fig != fig1 else ""),
                    barmode='group' if (fig == fig3 and len(time_choice) > 1) else 'stack',
                    showlegend=True,
                    legend=dict(
                        bgcolor='rgba(30, 41, 59, 0.8)',
                        bordercolor='rgba(255, 255, 255, 0.2)',
                        borderwidth=1
                    )
                )
            
            st.plotly_chart(fig1, use_container_width=True, theme="streamlit")
            st.plotly_chart(fig2, use_container_width=True, theme="streamlit")
            st.plotly_chart(fig3, use_container_width=True, theme="streamlit")
        else:
            st.info("📭 当前选择没有数据")

# -------------------- 页脚 --------------------
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: #64748b; font-size: 0.875rem; padding: 1rem;'>
        <p>📊 技能覆盖分析大屏 | 数据文件: <code>{}</code></p>
        <p>最后更新时间: {}</p>
    </div>
    """.format(
        SAVE_FILE,
        datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ),
    unsafe_allow_html=True
)
