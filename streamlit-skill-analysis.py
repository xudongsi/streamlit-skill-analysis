import os
import time
from datetime import datetime
from typing import List, Tuple

import pandas as pd
import streamlit as st
from streamlit_autorefresh import st_autorefresh
from streamlit_echarts import st_echarts
import plotly.graph_objects as go
from plotly.subplots import make_subplots  # 添加缺失的导入

# -------------------- 页面配置 --------------------
st.set_page_config(
    page_title="技能覆盖分析大屏",
    layout="wide",
    page_icon="📊"
)

# -------------------- 页面样式 --------------------
PAGE_CSS = """
<style>
body, [data-testid="stAppViewContainer"]{
    background-color:#0d1b2a !important;
    color:#ffffff !important;
}
[data-testid="stSidebar"]{
    background-color:#1b263b !important;
    color:#ffffff !important;
}
div.stButton>button{
    background-color:#4cc9f0 !important;
    color:#000000 !important;
    border-radius:10px;
    height:40px;
    font-weight:700;
    margin:5px 0;
    width:100%;
}
div.stButton>button:hover{
    background-color:#4895ef !important;
    color:#ffffff !important;
}
.metric-card{
    background-color:#1b263b !important;
    padding:20px;
    border-radius:16px;
    text-align:center;
    box-shadow:0 0 15px rgba(0,0,0,0.4);
}
.metric-value{
    font-size:36px;
    font-weight:800;
    color:#4cc9f0 !important;
}
.metric-label{
    font-size:14px;
    color:#cccccc !important;
}
hr{
    border:none;
    border-top:1px solid rgba(255,255,255,.12);
    margin:16px 0;
}
.danger-button div.stButton>button{
    background-color:#ff4d4d !important;
    color:#ffffff !important;
}
.danger-button div.stButton>button:hover{
    background-color:#ff3333 !important;
}
</style>
"""
st.markdown(PAGE_CSS, unsafe_allow_html=True)

SAVE_FILE = "jixiao.xlsx"  # 固定保存的文件


# -------------------- 数据导入 --------------------
@st.cache_data  # 修复：删除重复装饰器
def load_sheets(file, ts=None) -> Tuple[List[str], dict]:
    try:
        xpd = pd.ExcelFile(file)
    except Exception as e:
        st.sidebar.error(f"❌ 无法读取Excel文件: {e}")
        return [], {}

    frames = {}
    for s in xpd.sheet_names:
        try:
            # ✅ 关键修复：不设 header，让我们手动检测"分组"行
            df0 = pd.read_excel(xpd, sheet_name=s, header=None)
            if df0.empty:
                continue

            # ✅ 判断是否是标准模板（第二行是分组）
            if "明细" in df0.iloc[0].tolist() and df0.shape[0] > 1 and df0.iloc[1, 0] == "分组":
                df0.columns = df0.iloc[0].tolist()
                df0 = df0.drop(0).reset_index(drop=True)
            elif "明细" not in df0.columns and "明细" in df0.iloc[0].tolist():
                # 兼容无"分组"行但首行为表头的表
                df0.columns = df0.iloc[0].tolist()
                df0 = df0.drop(0).reset_index(drop=True)

            # ✅ 确保列名标准
            if not {"明细"}.issubset(df0.columns):
                st.sidebar.warning(f"⚠️ 表 {s} 缺少 '明细' 列，已跳过。")
                continue

            # ✅ 检测"分组"行逻辑保持原样
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
                # 确保值为数值类型
                df_long["值"] = pd.to_numeric(df_long["值"], errors='coerce').fillna(0)
                df_long["分组"] = df_long["员工"].map(group_map)
                # ✅ 新增：添加时间点列
                df_long["时间点"] = s
                frames[s] = df_long
            else:
                # ✅ 新增：对于已有数据的表也添加时间点列
                if "时间点" not in df0.columns:
                    df0["时间点"] = s
                # 确保值为数值类型
                if "值" in df0.columns:
                    df0["值"] = pd.to_numeric(df0["值"], errors='coerce').fillna(0)
                frames[s] = df0
        except Exception as e:
            st.sidebar.error(f"❌ 读取 {s} 时出错: {e}")
    return xpd.sheet_names, frames


# -------------------- 文件读取 --------------------
sheets, sheet_frames = [], {}
try:
    if os.path.exists(SAVE_FILE):
        mtime = os.path.getmtime(SAVE_FILE)
        sheets, sheet_frames = load_sheets(SAVE_FILE, ts=mtime)
        st.sidebar.success(f"✅ 已加载库文件 {SAVE_FILE}")
    else:
        # 创建示例数据
        sheet_frames = {
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
            for sheet_name, df0 in sheet_frames.items():
                df0.to_excel(writer, sheet_name=sheet_name, index=False)

        sheets, sheet_frames = load_sheets(SAVE_FILE)
        st.sidebar.info("📁 创建了示例数据文件")

    # ---------- 🧠 自动检测并修复数量总和 ----------
    repaired_count = 0
    repaired_frames = {}
    for sheet_name, df0 in sheet_frames.items():
        if df0 is not None and not df0.empty and "明细" in df0.columns and "值" in df0.columns:
            # 检查数量总和列是否存在或是否为空
            if "数量总和" not in df0.columns or df0["数量总和"].isnull().any():
                repaired = True
            else:
                # 判断当前总和是否与真实值匹配
                true_sum = df0.groupby("明细")["值"].sum().reset_index()
                merged = df0.merge(true_sum, on="明细", how="left", suffixes=("", "_真实"))
                repaired = not merged["数量总和"].equals(merged["值_真实"])

            if repaired:
                repaired_count += 1
                sum_df = (
                    df0.groupby("明细", as_index=False)["值"].sum()
                    .rename(columns={"值": "数量总和"})
                )
                df0 = df0.drop(columns=["数量总和"], errors="ignore")
                df0 = df0.merge(sum_df, on="明细", how="left")
                # ✅ 确保时间点列存在
                if "时间点" not in df0.columns:
                    df0["时间点"] = sheet_name
                repaired_frames[sheet_name] = df0

    if repaired_frames:
        try:
            with pd.ExcelWriter(SAVE_FILE, engine="openpyxl") as writer:
                for sn in sheets:
                    if sn in repaired_frames:
                        repaired_df = repaired_frames[sn]
                        repaired_df.to_excel(writer, sheet_name=sn, index=False)
                        sheet_frames[sn] = repaired_df
                    elif sn in sheet_frames:
                        df0 = sheet_frames[sn]
                        # ✅ 确保时间点列存在
                        if "时间点" not in df0.columns:
                            df0["时间点"] = sn
                        df0.to_excel(writer, sheet_name=sn, index=False)

            st.cache_data.clear()
            if repaired_count > 0:
                st.sidebar.info(f"🔧 已自动修复 {repaired_count} 张表的数量总和列")
        except Exception as e:
            st.sidebar.error(f"❌ 修复数据时出错: {e}")

except Exception as e:
    st.sidebar.error(f"❌ 读取库文件失败：{e}")
    sheet_frames = {}
    sheets = []

# -------------------- 删除功能 --------------------
st.sidebar.markdown("### ❌ 删除时间点")
if sheets:
    sheet_to_delete = st.sidebar.selectbox("选择要删除的时间点", sheets, key="delete_select")

    col1, col2 = st.sidebar.columns(2)
    with col1:
        if st.button("🗑️ 删除", key="delete_btn", help="删除选中的时间点"):
            try:
                if not os.path.exists(SAVE_FILE):
                    st.sidebar.error("文件不存在")
                else:
                    # 读取所有sheet
                    xls = pd.ExcelFile(SAVE_FILE)
                    new_sheets = {}

                    for sheet in xls.sheet_names:
                        if sheet != sheet_to_delete:
                            df0 = pd.read_excel(xls, sheet_name=sheet)
                            new_sheets[sheet] = df0

                    # 重新写入Excel，跳过要删除的sheet
                    with pd.ExcelWriter(SAVE_FILE, engine="openpyxl") as writer:
                        for sheet_name, df0 in new_sheets.items():
                            df0.to_excel(writer, sheet_name=sheet_name, index=False)

                    st.cache_data.clear()
                    st.sidebar.success(f"✅ 已删除时间点: {sheet_to_delete}")
                    time.sleep(1)
                    st.rerun()
            except Exception as e:
                st.sidebar.error(f"❌ 删除失败: {str(e)[:100]}")

    with col2:
        if st.button("🔄 刷新", key="refresh_btn"):
            st.cache_data.clear()
            st.rerun()

# -------------------- 智能化新增月份/季度 --------------------
st.sidebar.markdown("### ➕ 新增数据时间点")
current_year = datetime.now().year
year = st.sidebar.selectbox("选择年份", list(range(current_year - 2, current_year + 2)), index=2)
mode = st.sidebar.radio("时间类型", ["月份", "季度"], horizontal=True)

if mode == "月份":
    month = st.sidebar.selectbox("选择月份", list(range(1, 13)))
    new_sheet_name = f"{year}_{month:02d}"
else:
    quarter = st.sidebar.selectbox("选择季度", ["Q1", "Q2", "Q3", "Q4"])
    new_sheet_name = f"{year}_{quarter}"

if st.sidebar.button("创建新的时间点"):
    if new_sheet_name in sheets:
        st.sidebar.error(f"❌ 时间点 {new_sheet_name} 已存在！")
    else:
        try:
            base_df = pd.DataFrame(columns=["明细", "数量总和", "员工", "值", "分组", "时间点"])

            # ---------- 🧠 智能自动继承 ----------
            # 如果创建的是12月，自动删除旧的12月数据
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

                        st.sidebar.info(f"♻️ 已自动删除旧的12月数据: {old_sheet}")
                    except Exception as e:
                        st.sidebar.warning(f"⚠️ 删除旧数据时出错: {str(e)[:50]}")

            # 筛选同年份中比当前时间点早的所有 sheet
            prev_sheets = sorted([s for s in sheets if s.split("_")[0] == str(year) and s < new_sheet_name])

            # 如果当年没有，就自动往前一年回溯
            if not prev_sheets:
                prev_years = sorted([int(s.split("_")[0]) for s in sheets if s.split("_")[0].isdigit()])
                if prev_years:
                    latest_prev_year = max(y for y in prev_years if y < year) if any(
                        y < year for y in prev_years) else None
                    if latest_prev_year:
                        prev_sheets = sorted([s for s in sheets if s.startswith(str(latest_prev_year))])

            if prev_sheets:
                prev_name = prev_sheets[-1]
                base_df = sheet_frames.get(prev_name, base_df).copy()
                # 清空"值"列，但保留其他结构
                if "值" in base_df.columns:
                    base_df["值"] = 0
                st.sidebar.info(f"📋 已从最近时间点 {prev_name} 自动继承结构")
            else:
                st.sidebar.info("🆕 未找到上期数据，创建空白模板")
                # 创建基本的示例数据
                base_df = pd.DataFrame({
                    "明细": ["示例任务1", "示例任务2", "示例任务3"],
                    "数量总和": [0, 0, 0],
                    "员工": ["员工A", "员工B", "员工C"],
                    "值": [0, 0, 0],
                    "分组": ["分组A", "分组B", "分组C"],
                    "时间点": new_sheet_name
                })

            # ---------- 写入 Excel ----------
            if os.path.exists(SAVE_FILE):
                with pd.ExcelWriter(SAVE_FILE, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
                    base_df.to_excel(writer, sheet_name=new_sheet_name, index=False)
            else:
                with pd.ExcelWriter(SAVE_FILE, engine="openpyxl") as writer:
                    base_df.to_excel(writer, sheet_name=new_sheet_name, index=False)

            st.cache_data.clear()
            st.sidebar.success(f"✅ 已创建新时间点: {new_sheet_name}")
            if mode == "月份" and month == 12:
                st.sidebar.success("♻️ 已自动清理旧的12月数据")

            time.sleep(1)
            st.rerun()

        except Exception as e:
            st.sidebar.error(f"❌ 创建失败：{str(e)[:100]}")

# -------------------- 🧮 一键更新所有数量总和 --------------------
st.sidebar.markdown("### ⚙️ 数据修复工具")

if st.sidebar.button("🧮 一键更新所有数量总和"):
    try:
        if not os.path.exists(SAVE_FILE):
            st.sidebar.warning("未找到文件 jixiao.xlsx")
        else:
            xls = pd.ExcelFile(SAVE_FILE)
            updated_frames = {}

            with st.spinner("正在更新数据..."):
                for sheet_name in xls.sheet_names:
                    df0 = pd.read_excel(xls, sheet_name=sheet_name)
                    if "明细" in df0.columns and "值" in df0.columns:
                        # 自动计算数量总和
                        sum_df = (
                            df0.groupby("明细", as_index=False)["值"].sum()
                            .rename(columns={"值": "数量总和"})
                        )
                        df0 = df0.drop(columns=["数量总和"], errors="ignore")
                        df0 = df0.merge(sum_df, on="明细", how="left")
                        # ✅ 确保时间点列存在
                        if "时间点" not in df0.columns:
                            df0["时间点"] = sheet_name
                        updated_frames[sheet_name] = df0
                    else:
                        updated_frames[sheet_name] = df0

                # 写回所有表
                with pd.ExcelWriter(SAVE_FILE, engine="openpyxl") as writer:
                    for sheet_name, df0 in updated_frames.items():
                        df0.to_excel(writer, sheet_name=sheet_name, index=False)

                st.cache_data.clear()
                st.sidebar.success("✅ 所有工作表的数量总和已重新计算并更新！")
                time.sleep(1)
                st.rerun()

    except Exception as e:
        st.sidebar.error(f"❌ 更新失败：{str(e)[:100]}")

# -------------------- 智能时间点选择 --------------------
# 允许跨年份选择
all_time_points = sorted(sheets, reverse=True)
time_choice = st.sidebar.multiselect(
    "选择月份/季度（可多选跨年）",
    all_time_points,
    default=all_time_points[:1] if all_time_points else [],
    key="time_select"
)

# 分组选择
if time_choice:
    # 合并选择的时间点数据
    dfs = []
    for t in time_choice:
        df0 = sheet_frames.get(t)
        if df0 is not None and not df0.empty:
            dfs.append(df0)

    if dfs:
        combined_df = pd.concat(dfs, ignore_index=True)
        all_groups = combined_df["分组"].dropna().unique().tolist() if "分组" in combined_df.columns else []
        selected_groups = st.sidebar.multiselect(
            "选择分组",
            all_groups,
            default=all_groups,
            key="group_select"
        )
    else:
        selected_groups = []
else:
    selected_groups = []
    if sheets:
        st.sidebar.warning("⚠️ 请选择时间点")

# -------------------- 视图选择 --------------------
sections_names = [
    "人员完成任务数量排名",
    "任务对比（堆叠柱状图）",
    "任务掌握情况（热门任务）",
    "任务-人员热力图"
]
view = st.sidebar.radio("切换视图", ["编辑数据", "大屏轮播", "单页模式", "显示所有视图", "能力分析"], key="view_select")


# -------------------- 数据合并（修复后） --------------------
def get_merged_df(keys: List[str], groups: List[str]) -> pd.DataFrame:
    dfs = []
    for k in keys:
        df0 = sheet_frames.get(k)
        if df0 is not None and not df0.empty:
            if groups and "分组" in df0.columns and len(groups) > 0:
                df0 = df0[df0["分组"].isin(groups)]
            # ✅ 确保时间点列存在
            if "时间点" not in df0.columns:
                df0["时间点"] = k
            dfs.append(df0)

    if not dfs:
        return pd.DataFrame()

    merged_df = pd.concat(dfs, axis=0, ignore_index=True)

    # 确保数值列类型正确
    if "值" in merged_df.columns:
        merged_df["值"] = pd.to_numeric(merged_df["值"], errors='coerce').fillna(0)

    return merged_df


df = get_merged_df(time_choice, selected_groups)


# -------------------- 图表函数（修复后） --------------------
def chart_total(df0):
    if df0 is None or df0.empty:
        return go.Figure()

    # 过滤分数总和
    if "明细" in df0.columns:
        df0 = df0[df0["明细"] != "分数总和"]

    # ✅ 修复：按员工和时间点分组，区分不同时间点
    if len(time_choice) > 1 and "时间点" in df0.columns:
        emp_time_stats = df0.groupby(["员工", "时间点"])["值"].sum().reset_index()

        # 创建分组柱状图
        fig = go.Figure()

        # 为每个时间点添加一个柱状图系列
        time_points = sorted(emp_time_stats["时间点"].unique())
        colors = ['#4cc9f0', '#4895ef', '#4361ee', '#3f37c9', '#3a0ca3']

        for i, time_point in enumerate(time_points):
            time_data = emp_time_stats[emp_time_stats["时间点"] == time_point]
            time_data = time_data.sort_values("值", ascending=False)

            fig.add_trace(go.Bar(
                x=time_data["员工"],
                y=time_data["值"],
                name=time_point,
                marker_color=colors[i % len(colors)],
                text=time_data["值"],
                textposition="outside",
                hovertemplate="员工: %{x}<br>时间点: %{customdata}<br>完成值: %{y}<extra></extra>",
                customdata=[time_point] * len(time_data)
            ))

        fig.update_layout(
            barmode='group',
            template="plotly_dark",
            xaxis_title="员工",
            yaxis_title="完成总值",
            showlegend=True,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            )
        )
    else:
        # 单个时间点的处理
        emp_stats = df0.groupby("员工")["值"].sum().sort_values(ascending=False).reset_index()
        fig = go.Figure(go.Bar(
            x=emp_stats["员工"],
            y=emp_stats["值"],
            text=emp_stats["值"],
            textposition="outside",
            hovertemplate="员工: %{x}<br>完成总值: %{y}<extra></extra>",
            marker_color='#4cc9f0'
        ))
        fig.update_layout(template="plotly_dark", xaxis_title="员工", yaxis_title="完成总值")

    return fig


def chart_stack(df0):
    if df0 is None or df0.empty:
        return go.Figure()

    if "明细" in df0.columns:
        df0 = df0[df0["明细"] != "分数总和"]

    # ✅ 修复：处理多个时间点的情况
    if len(time_choice) > 1 and "时间点" in df0.columns:
        # 使用子图显示不同时间点
        time_points = sorted(df0["时间点"].unique())

        if len(time_points) == 1:
            # 单个时间点
            df_pivot = df0.pivot_table(index="明细", columns="员工", values="值", aggfunc="sum", fill_value=0)
            fig = go.Figure()
            colors = ['#4cc9f0', '#4895ef', '#4361ee', '#3f37c9', '#3a0ca3']
            for i, emp in enumerate(df_pivot.columns):
                fig.add_trace(go.Bar(
                    x=df_pivot.index,
                    y=df_pivot[emp],
                    name=emp,
                    marker_color=colors[i % len(colors)]
                ))
            fig.update_layout(
                barmode="stack",
                template="plotly_dark",
                xaxis_title="任务",
                yaxis_title="完成值",
                title=f"时间点: {time_points[0]}"
            )
        else:
            # 多个时间点使用子图
            fig = make_subplots(
                rows=len(time_points), cols=1,
                subplot_titles=[f"时间点: {tp}" for tp in time_points],
                vertical_spacing=0.1
            )

            colors = ['#4cc9f0', '#4895ef', '#4361ee', '#3f37c9', '#3a0ca3',
                      '#7209b7', '#560bad', '#480ca8', '#3a0ca3', '#3f37c9']

            for i, tp in enumerate(time_points, 1):
                df_tp = df0[df0["时间点"] == tp]
                df_pivot = df_tp.pivot_table(index="明细", columns="员工", values="值", aggfunc="sum", fill_value=0)

                # 获取员工列表，确保颜色一致
                all_emps = df0["员工"].unique()

                for j, emp in enumerate(df_pivot.columns):
                    color_idx = list(all_emps).index(emp) % len(colors) if emp in all_emps else j
                    fig.add_trace(
                        go.Bar(
                            x=df_pivot.index,
                            y=df_pivot[emp],
                            name=emp,
                            marker_color=colors[color_idx],
                            showlegend=(i == 1),
                            legendgroup=emp
                        ),
                        row=i, col=1
                    )

            fig.update_layout(
                barmode="stack",
                template="plotly_dark",
                height=400 * len(time_points),
                showlegend=True
            )
            fig.update_xaxes(title_text="任务", row=len(time_points), col=1)
            fig.update_yaxes(title_text="完成值", row=len(time_points) // 2 + 1, col=1)
    else:
        # 原始逻辑（单个时间点）
        df_pivot = df0.pivot_table(index="明细", columns="员工", values="值", aggfunc="sum", fill_value=0)
        fig = go.Figure()
        colors = ['#4cc9f0', '#4895ef', '#4361ee', '#3f37c9', '#3a0ca3']
        for i, emp in enumerate(df_pivot.columns):
            fig.add_trace(go.Bar(
                x=df_pivot.index,
                y=df_pivot[emp],
                name=emp,
                marker_color=colors[i % len(colors)]
            ))
        fig.update_layout(barmode="stack", template="plotly_dark", xaxis_title="任务", yaxis_title="完成值")

    return fig


def chart_hot(df0):
    if df0 is None or df0.empty:
        return {
            "backgroundColor": "transparent",
            "yAxis": {"type": "category", "data": [], "axisLabel": {"color": "#fff"}},
            "xAxis": {"type": "value", "axisLabel": {"color": "#fff"}},
            "series": [{"data": [], "type": "bar", "itemStyle": {"color": "#ffb703"}}]
        }

    if "明细" in df0.columns:
        df0 = df0[df0["明细"] != "分数总和"]

    # ✅ 修复：处理多个时间点的情况
    if len(time_choice) > 1 and "时间点" in df0.columns:
        # 按时间点分组显示
        time_points = sorted(df0["时间点"].unique())
        tasks = df0["明细"].unique().tolist()[:15]  # 限制显示数量

        option = {
            "backgroundColor": "transparent",
            "tooltip": {"trigger": "axis", "axisPointer": {"type": "shadow"}},
            "legend": {"data": time_points, "textStyle": {"color": "#fff"}},
            "xAxis": {"type": "value", "axisLabel": {"color": "#fff"}},
            "yAxis": {"type": "category", "data": tasks, "axisLabel": {"color": "#fff"}},
            "series": []
        }

        colors = ['#ffb703', '#fb8500', '#ff006e', '#8338ec', '#3a86ff']

        for i, tp in enumerate(time_points):
            df_tp = df0[df0["时间点"] == tp]
            ts = df_tp.groupby("明细")["员工"].nunique()
            # 确保顺序一致
            ts_ordered = [ts.get(task, 0) for task in tasks]

            option["series"].append({
                "name": tp,
                "type": "bar",
                "data": ts_ordered,
                "itemStyle": {"color": colors[i % len(colors)]}
            })
    else:
        # 原始逻辑（单个时间点）
        ts = df0.groupby("明细")["员工"].nunique().sort_values(ascending=False).head(15)
        option = {
            "backgroundColor": "transparent",
            "yAxis": {"type": "category", "data": ts.index.tolist(), "axisLabel": {"color": "#fff"}},
            "xAxis": {"type": "value", "axisLabel": {"color": "#fff"}},
            "series": [{"data": ts.tolist(), "type": "bar", "itemStyle": {"color": "#ffb703"}}]
        }

    return option


def chart_heat(df0):
    if df0 is None or df0.empty:
        return {
            "backgroundColor": "transparent",
            "tooltip": {"position": "top"},
            "xAxis": {"type": "category", "data": [], "axisLabel": {"color": "#fff"}},
            "yAxis": {"type": "category", "data": [], "axisLabel": {"color": "#fff"}},
            "visualMap": {"min": 0, "max": 1, "show": False, "inRange": {"color": ["#ff4d4d", "#4caf50"]}},
            "series": [{"type": "heatmap", "data": []}]
        }

    if "明细" in df0.columns:
        df0 = df0[df0["明细"] != "分数总和"]

    # ✅ 修复：处理多个时间点的情况
    if len(time_choice) > 1 and "时间点" in df0.columns:
        # 使用下拉框选择时间点
        time_points = sorted(df0["时间点"].unique())

        option = {
            "backgroundColor": "transparent",
            "tooltip": {"position": "top"},
            "visualMap": {"min": 0, "max": 1, "show": False, "inRange": {"color": ["#ff4d4d", "#4caf50"]}},
            "series": [],
            "timeline": {
                "axisType": "category",
                "autoPlay": False,
                "playInterval": 2000,
                "data": time_points,
                "label": {"color": "#fff"},
                "lineStyle": {"color": "#4cc9f0"}
            },
            "options": []
        }

        for tp in time_points:
            df_tp = df0[df0["时间点"] == tp]
            tasks = df_tp["明细"].unique().tolist()[:20]  # 限制数量
            emps = df_tp["员工"].unique().tolist()[:20]  # 限制数量
            data = []

            max_val = 0
            for i, t in enumerate(tasks):
                for j, e in enumerate(emps):
                    v = int(df_tp[(df_tp["明细"] == t) & (df_tp["员工"] == e)]["值"].sum())
                    data.append([j, i, v])
                    max_val = max(max_val, v)

            option["options"].append({
                "title": {"text": f"时间点: {tp}", "textStyle": {"color": "#fff"}},
                "xAxis": {
                    "type": "category",
                    "data": emps,
                    "axisLabel": {"color": "#fff", "rotate": 45, "interval": 0}
                },
                "yAxis": {
                    "type": "category",
                    "data": tasks,
                    "axisLabel": {"color": "#fff"}
                },
                "series": [{"type": "heatmap", "data": data}]
            })

        # 更新visualMap的最大值
        if max_val > 0:
            option["visualMap"]["max"] = max_val
    else:
        # 原始逻辑（单个时间点）
        tasks = df0["明细"].unique().tolist()[:20]  # 限制数量
        emps = df0["员工"].unique().tolist()[:20]  # 限制数量
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
                "axisLabel": {"color": "#fff", "rotate": 45, "interval": 0}
            },
            "yAxis": {
                "type": "category",
                "data": tasks,
                "axisLabel": {"color": "#fff"}
            },
            "visualMap": {
                "min": 0,
                "max": max_val if max_val > 0 else 1,
                "show": True,
                "inRange": {"color": ["#ff4d4d", "#4caf50"]}
            },
            "series": [{"type": "heatmap", "data": data}]
        }

    return option


# -------------------- 卡片显示 --------------------
def show_cards(df0):
    if df0 is None or df0.empty:
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

    # ✅ 显示选择的时间点
    time_points_display = ", ".join(time_choice) if time_choice else "未选择"

    c1, c2, c3, c4, c5 = st.columns(5)

    # 使用更安全的HTML渲染
    card_html = f'''
    <div class="metric-card">
        <div class="metric-value">{total_tasks}</div>
        <div class="metric-label">任务数</div>
    </div>
    '''
    c1.markdown(card_html, unsafe_allow_html=True)

    c2.markdown(f'''
    <div class="metric-card">
        <div class="metric-value">{total_people}</div>
        <div class="metric-label">人数</div>
    </div>
    ''', unsafe_allow_html=True)

    c3.markdown(f'''
    <div class="metric-card">
        <div class="metric-value">{top_person[:8] if len(top_person) > 8 else top_person}</div>
        <div class="metric-label">覆盖率最高</div>
        <div style="font-size:12px;color:#94a3b8;margin-top:5px">值: {top_value}</div>
    </div>
    ''', unsafe_allow_html=True)

    c4.markdown(f'''
    <div class="metric-card">
        <div class="metric-value">{avg_score}</div>
        <div class="metric-label">平均完成值</div>
    </div>
    ''', unsafe_allow_html=True)

    c5.markdown(f'''
    <div class="metric-card">
        <div class="metric-label">选择的时间点</div>
        <div style="font-size:14px;margin-top:10px;color:#4cc9f0">{len(time_choice)} 个</div>
        <div style="font-size:12px;color:#94a3b8;margin-top:5px">{time_points_display[:30]}{'...' if len(time_points_display) > 30 else ''}</div>
    </div>
    ''', unsafe_allow_html=True)

    st.markdown("<hr/>", unsafe_allow_html=True)


# -------------------- 主页面 --------------------
st.title("📊 技能覆盖分析大屏")

if view == "编辑数据":
    if not time_choice:
        st.warning("⚠️ 请在左侧选择时间点（月或季）后再编辑数据")
    else:
        # ✅ 显示选择的时间点信息
        if len(time_choice) > 1:
            st.info(f"📅 当前编辑 {len(time_choice)} 个时间点: {', '.join(time_choice)}")
            st.warning("⚠️ 多时间点编辑模式下，请注意数据的时间点归属")

        show_cards(df)

        if not df.empty:
            st.info("你可以直接编辑下面的表格，修改完成后点击【保存】按钮。")

            # 编辑时显示时间点列
            edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)

            col1, col2 = st.columns([1, 3])
            with col1:
                if st.button("💾 保存修改到库里", type="primary", use_container_width=True):
                    try:
                        if len(time_choice) == 1:
                            # 单个时间点保存
                            sheet_name = time_choice[0]

                            # 自动计算数量总和
                            if "明细" in edited_df.columns and "值" in edited_df.columns:
                                sum_df = (
                                    edited_df.groupby("明细", as_index=False)["值"].sum()
                                    .rename(columns={"值": "数量总和"})
                                )
                                edited_df = edited_df.drop(columns=["数量总和"], errors="ignore")
                                edited_df = edited_df.merge(sum_df, on="明细", how="left")

                            # 移除时间点列（Excel中不需要）
                            df_to_save = edited_df.drop(columns=["时间点"], errors="ignore")

                            # 保存
                            with pd.ExcelWriter(SAVE_FILE, mode="a", if_sheet_exists="replace",
                                                engine="openpyxl") as writer:
                                df_to_save.to_excel(writer, sheet_name=sheet_name, index=False)

                            st.success(f"✅ 修改已保存到 {SAVE_FILE} ({sheet_name})")
                        else:
                            # 多个时间点保存 - 需要按时间点拆分
                            success_count = 0
                            for sheet_name in time_choice:
                                df_sheet = edited_df[edited_df["时间点"] == sheet_name]
                                if not df_sheet.empty:
                                    # 移除时间点列（Excel中不需要）
                                    df_to_save = df_sheet.drop(columns=["时间点"], errors="ignore")

                                    # 自动计算数量总和
                                    if "明细" in df_to_save.columns and "值" in df_to_save.columns:
                                        sum_df = (
                                            df_to_save.groupby("明细", as_index=False)["值"].sum()
                                            .rename(columns={"值": "数量总和"})
                                        )
                                        df_to_save = df_to_save.drop(columns=["数量总和"], errors="ignore")
                                        df_to_save = df_to_save.merge(sum_df, on="明细", how="left")

                                    # 保存
                                    with pd.ExcelWriter(SAVE_FILE, mode="a", if_sheet_exists="replace",
                                                        engine="openpyxl") as writer:
                                        df_to_save.to_excel(writer, sheet_name=sheet_name, index=False)
                                    success_count += 1

                            st.success(f"✅ 修改已保存到 {success_count} 个时间点")

                        st.cache_data.clear()
                        time.sleep(1)
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ 保存失败：{str(e)[:100]}")

            with col2:
                if st.button("🔄 重置修改", type="secondary", use_container_width=True):
                    st.cache_data.clear()
                    st.rerun()
        else:
            st.info("📭 当前选择没有数据，请先添加数据或选择其他时间点")

elif view == "大屏轮播":
    if not time_choice:
        st.warning("⚠️ 请在左侧选择时间点（月或季）后查看大屏轮播")
    else:
        st_autorefresh(interval=10000, key="aut")
        show_cards(df)

        if not df.empty:
            secs = [
                ("完成排名", chart_total(df)),
                ("任务对比", chart_stack(df)),
                ("热门任务", chart_hot(df)),
                ("热力图", chart_heat(df))
            ]
            idx = int(time.time() / 10) % len(secs)
            t, op = secs[idx]

            st.subheader(t)
            if isinstance(op, go.Figure):
                st.plotly_chart(op, use_container_width=True, theme="streamlit")
            else:
                st_echarts(op, height="600px", theme="dark")
        else:
            st.info("📭 当前选择没有数据，无法显示图表")

elif view == "单页模式":
    if not time_choice:
        st.warning("⚠️ 请在左侧选择时间点（月或季）后查看单页模式")
    else:
        show_cards(df)

        if not df.empty:
            choice = st.sidebar.selectbox("单页查看", sections_names, index=0, key="single_view_select")
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
            st.info("📭 当前选择没有数据，无法显示图表")

elif view == "显示所有视图":
    if not time_choice:
        st.warning("⚠️ 请在左侧选择时间点（月或季）后查看所有视图")
    else:
        show_cards(df)

        if not df.empty:
            charts = [
                ("完成排名", chart_total(df)),
                ("任务对比（堆叠柱状图）", chart_stack(df)),
                ("热门任务", chart_hot(df)),
                ("热图", chart_heat(df))
            ]
            for label, f in charts:
                st.subheader(label)
                if isinstance(f, go.Figure):
                    st.plotly_chart(f, use_container_width=True, theme="streamlit")
                else:
                    st_echarts(f, height="520px", theme="dark")
        else:
            st.info("📭 当前选择没有数据，无法显示图表")

elif view == "能力分析":
    if not time_choice:
        st.warning("⚠️ 请在左侧选择时间点（月或季）后查看能力分析")
    else:
        st.subheader("📊 能力分析")

        if not df.empty:
            # ✅ 显示选择的时间点
            if len(time_choice) > 1:
                st.info(f"📊 当前分析 {len(time_choice)} 个时间点: {', '.join(time_choice)}")

            employees = df["员工"].unique().tolist()
            selected_emps = st.sidebar.multiselect(
                "选择员工（图1显示）",
                employees,
                default=employees[:3] if employees else [],
                key="emp_select"
            )
            tasks = df["明细"].unique().tolist()

            fig1, fig2, fig3 = go.Figure(), go.Figure(), go.Figure()

            # ✅ 使用颜色区分不同时间点
            colors = ['#4cc9f0', '#4895ef', '#4361ee', '#3f37c9', '#3a0ca3',
                      '#7209b7', '#560bad', '#480ca8', '#3a0ca3', '#3f37c9']

            for idx, sheet in enumerate(time_choice):
                df_sheet = get_merged_df([sheet], selected_groups)
                if df_sheet is None or df_sheet.empty:
                    continue

                if "明细" in df_sheet.columns:
                    df_sheet = df_sheet[df_sheet["明细"] != "分数总和"]

                df_pivot = df_sheet.pivot_table(index="明细", columns="员工", values="值", fill_value=0)

                color = colors[idx % len(colors)]

                # 图1: 员工任务完成情况（多条线）
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

                # 图2: 任务整体完成度趋势
                task_sums = df_pivot.sum(axis=1).reindex(tasks, fill_value=0)
                fig2.add_trace(go.Scatter(
                    x=tasks,
                    y=task_sums,
                    mode="lines+markers",
                    name=sheet,
                    line=dict(color=color, width=3 if sheet == time_choice[-1] else 2),
                    marker=dict(size=8 if sheet == time_choice[-1] else 6)
                ))

                # 图3: 员工整体完成度对比
                emp_sums = df_pivot.sum(axis=0)
                if not emp_sums.empty:
                    fig3.add_trace(go.Bar(
                        x=emp_sums.index,
                        y=emp_sums.values,
                        name=sheet,
                        marker_color=color,
                        opacity=0.7
                    ))

            # 更新图表布局
            fig1.update_layout(
                title="员工任务完成情况（多时间点对比）",
                template="plotly_dark",
                xaxis_title="任务",
                yaxis_title="完成值",
                showlegend=True,
                legend=dict(
                    orientation="h",
                    yanchor="bottom",
                    y=1.02,
                    xanchor="right",
                    x=1
                )
            )

            fig2.update_layout(
                title="任务整体完成度趋势（多时间点对比）",
                template="plotly_dark",
                xaxis_title="任务",
                yaxis_title="完成值总和",
                showlegend=True
            )

            fig3.update_layout(
                title="员工整体完成度对比（多时间点堆叠）",
                template="plotly_dark",
                xaxis_title="员工",
                yaxis_title="完成值总和",
                barmode='group' if len(time_choice) > 1 else 'stack',
                showlegend=True if len(time_choice) > 1 else False
            )

            st.plotly_chart(fig1, use_container_width=True, theme="streamlit")
            st.plotly_chart(fig2, use_container_width=True, theme="streamlit")
            st.plotly_chart(fig3, use_container_width=True, theme="streamlit")
        else:
            st.info("📭 当前选择没有数据，无法进行分析")

# -------------------- 页脚 --------------------
st.markdown("---")
st.markdown(
    f"""
    <div style='text-align: center; color: #94a3b8; font-size: 0.875rem; padding: 1rem;'>
        <p>📊 技能覆盖分析大屏 | 数据文件: <code>{SAVE_FILE}</code></p>
        <p>最后更新时间: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>
    </div>
    """,
    unsafe_allow_html=True
)
