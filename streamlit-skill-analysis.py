import os
import time
from datetime import datetime
from typing import List, Tuple

import pandas as pd
import streamlit as st
from streamlit_autorefresh import st_autorefresh
from streamlit_echarts import st_echarts
import plotly.graph_objects as go

# -------------------- 页面配置 --------------------
st.set_page_config(page_title="技能覆盖分析大屏", layout="wide")

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
/* 热力图滚动容器样式 */
.heatmap-container {
    max-height: 700px;
    overflow-y: auto;
    overflow-x: auto;
    border-radius: 8px;
}
/* 滚动条美化 */
.heatmap-container::-webkit-scrollbar {
    width: 8px;
    height: 8px;
}
.heatmap-container::-webkit-scrollbar-thumb {
    background-color: #4cc9f0;
    border-radius: 4px;
}
.heatmap-container::-webkit-scrollbar-track {
    background-color: #1b263b;
}
/* 删除按钮样式 */
.delete-btn {
    background-color: #ff4d4d !important;
    color: white !important;
}
.delete-btn:hover {
    background-color: #ff1a1a !important;
}
</style>
"""
st.markdown(PAGE_CSS, unsafe_allow_html=True)

SAVE_FILE = r"C:\Users\128393112839311\Desktop\jixiao.xlsx"  # 固定保存的文件


# -------------------- 数据导入 --------------------
@st.cache_data(ttl=300)  # 缓存5分钟，避免频繁读取文件
def load_sheets(file) -> Tuple[List[str], dict]:
    """读取Excel所有工作表，返回工作表名列表和数据字典"""
    if not os.path.exists(file):
        return [], {}

    xpd = pd.ExcelFile(file)
    frames = {}
    for s in xpd.sheet_names:
        try:
            df0 = pd.read_excel(xpd, sheet_name=s)
            if df0.empty:
                continue
            if not {"明细", "员工", "值"}.issubset(df0.columns):
                st.sidebar.warning(f"⚠️ 表 {s} 缺少必要列，已跳过。")
                continue

            # 解析分组行
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
                df_long["分组"] = df_long["员工"].map(group_map)
                frames[s] = df_long
            else:
                frames[s] = df0
        except Exception as e:
            st.sidebar.error(f"❌ 读取 {s} 时出错: {e}")
    return xpd.sheet_names, frames


# -------------------- 删除工作表函数 --------------------
def delete_sheet(file_path, sheet_name):
    """删除指定工作表"""
    try:
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names
        if sheet_name not in sheet_names:
            return False, "工作表不存在"

        # 保留除要删除外的所有工作表
        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
            for sn in sheet_names:
                if sn != sheet_name:
                    df = pd.read_excel(xls, sheet_name=sn)
                    df.to_excel(writer, sheet_name=sn, index=False)

        return True, f"✅ 成功删除工作表: {sheet_name}"
    except Exception as e:
        return False, f"❌ 删除失败: {str(e)}"


# -------------------- 文件读取 --------------------
sheets, sheet_frames = load_sheets(SAVE_FILE)

# 初始化：文件不存在时创建空文件，不重置已有数据（解决问题1）
if not os.path.exists(SAVE_FILE):
    # 创建空Excel文件，避免后续报错
    with pd.ExcelWriter(SAVE_FILE, engine="openpyxl") as writer:
        pd.DataFrame(columns=["明细", "数量总和", "员工", "值", "分组"]).to_excel(
            writer, sheet_name="示例_2025_01", index=False
        )
    sheets, sheet_frames = load_sheets(SAVE_FILE)
    st.sidebar.success(f"✅ 已创建初始文件 {SAVE_FILE}")
elif not sheets:
    st.sidebar.warning("⚠️ 文件存在但无有效工作表，已创建示例数据")
    sheet_frames = {
        "示例_2025_01": pd.DataFrame({
            "明细": ["任务A", "任务B", "任务C"],
            "数量总和": [3, 2, 5],
            "员工": ["张三", "李四", "王五"],
            "值": [1, 1, 1],
            "分组": ["A8", "B7", "VN"]
        })
    }
    sheets = ["示例_2025_01"]
else:
    st.sidebar.success(f"✅ 已加载库文件 {SAVE_FILE}（共{len(sheets)}个工作表）")

# ---------- 🧠 自动检测并修复数量总和 ----------
repaired_count = 0
repaired_frames = {}
for sheet_name, df0 in sheet_frames.items():
    if "明细" in df0.columns and "值" in df0.columns:
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
            repaired_frames[sheet_name] = df0

if repaired_frames:
    with pd.ExcelWriter(SAVE_FILE, engine="openpyxl") as writer:
        for sn in sheets:
            if sn in repaired_frames:
                repaired_df = repaired_frames[sn]
                repaired_df.to_excel(writer, sheet_name=sn, index=False)
                sheet_frames[sn] = repaired_df
            else:
                # 保留原始数据
                df_original = pd.read_excel(SAVE_FILE, sheet_name=sn)
                df_original.to_excel(writer, sheet_name=sn, index=False)
    st.cache_data.clear()
    st.sidebar.info(f"🔧 已自动修复 {repaired_count} 张表的数量总和列")

# -------------------- 智能化新增月份/季度 --------------------
st.sidebar.markdown("### 📅 新增数据时间点")
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
            base_df = pd.DataFrame(columns=["明细", "数量总和", "员工", "值", "分组"])

            # ---------- 🧠 智能自动继承 ----------
            # 筛选所有比当前时间点早的sheet（跨年份）
            prev_sheets = sorted([s for s in sheets if "_" in s and s < new_sheet_name])

            if prev_sheets:
                prev_name = prev_sheets[-1]
                base_df = sheet_frames.get(prev_name, base_df).copy()
                st.sidebar.info(f"🔧 已从最近时间点 {prev_name} 自动继承数据")
            else:
                st.sidebar.info("🔧 未找到上期数据，创建空白模板")

            # ---------- 写入 Excel ----------
            with pd.ExcelWriter(SAVE_FILE, mode="a", engine="openpyxl") as writer:
                base_df.to_excel(writer, sheet_name=new_sheet_name, index=False)

            st.cache_data.clear()
            # 重新加载数据
            sheets, sheet_frames = load_sheets(SAVE_FILE)
            st.sidebar.success(f"✅ 已创建新时间点: {new_sheet_name}")

        except Exception as e:
            st.sidebar.error(f"❌ 创建失败：{e}")

# -------------------- 删除工作表功能（解决问题3） --------------------
st.sidebar.markdown("### 🗑️ 删除时间点")
if sheets:
    sheet_to_delete = st.sidebar.selectbox("选择要删除的时间点", sheets)
    # 防止删除最后一个工作表
    if len(sheets) == 1:
        st.sidebar.warning("⚠️ 至少保留一个工作表，无法删除")
    else:
        if st.sidebar.button("删除选中时间点", key="delete_btn", help="删除后不可恢复",
                             args=[{"key": "delete-btn"}]):
            success, msg = delete_sheet(SAVE_FILE, sheet_to_delete)
            st.sidebar.warning(msg)
            if success:
                st.cache_data.clear()
                sheets, sheet_frames = load_sheets(SAVE_FILE)

# -------------------- 🧮 一键更新所有数量总和 --------------------
st.sidebar.markdown("### 🔧 数据修复工具")

if st.sidebar.button("🧮 一键更新所有数量总和"):
    try:
        if not os.path.exists(SAVE_FILE):
            st.sidebar.warning("未找到文件 jixiao.xlsx")
        else:
            xls = pd.ExcelFile(SAVE_FILE)
            updated_frames = {}
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
                    updated_frames[sheet_name] = df0

            # 写回所有表
            with pd.ExcelWriter(SAVE_FILE, engine="openpyxl") as writer:
                for sheet_name, df0 in updated_frames.items():
                    df0.to_excel(writer, sheet_name=sheet_name, index=False)

            st.cache_data.clear()
            # 重新加载数据
            sheets, sheet_frames = load_sheets(SAVE_FILE)
            st.sidebar.success("✅ 所有工作表的数量总和已重新计算并更新！")

    except Exception as e:
        st.sidebar.error(f"❌ 更新失败：{e}")

# -------------------- 时间点选择优化（解决问题2） --------------------
st.sidebar.markdown("### 📋 数据筛选")
# 自动识别所有年份
years_available = sorted(list({s.split("_")[0] for s in sheets if "_" in s}))
# 新增"全部年份"选项
year_choice = st.sidebar.selectbox("筛选年份", ["全部年份"] + years_available)

# 根据年份筛选时间点（支持跨年份选择）
if year_choice == "全部年份":
    time_candidates = sorted(sheets)
else:
    time_candidates = sorted([s for s in sheets if s.startswith(year_choice)])

if not time_candidates:
    st.warning(f"⚠️ 暂无符合条件的数据，请先创建月份或季度。")
    time_choice = []
else:
    # 默认选择前2个时间点（方便跨年份对比）
    default_choice = time_candidates[:2] if len(time_candidates) >= 2 else time_candidates[:1]
    time_choice = st.sidebar.multiselect("选择时间点（支持跨年份对比）",
                                         time_candidates,
                                         default=default_choice)

# -------------------- 分组选择 --------------------
all_groups = pd.concat(sheet_frames.values())["分组"].dropna().unique().tolist() if sheet_frames else []
selected_groups = st.sidebar.multiselect("选择分组", all_groups, default=all_groups)

# -------------------- 视图选择 --------------------
sections_names = [
    "人员完成任务数量排名",
    "任务对比（堆叠柱状图）",
    "任务-人员热力图"
]
view = st.sidebar.radio("切换视图", ["编辑数据", "大屏轮播", "单页模式", "显示所有视图", "能力分析"])


# -------------------- 数据合并 --------------------
def get_merged_df(keys: List[str], groups: List[str]) -> pd.DataFrame:
    """合并选中时间点和分组的数据"""
    dfs = []
    for k in keys:
        df0 = sheet_frames.get(k)
        if df0 is not None:
            if groups and "分组" in df0.columns:
                df0 = df0[df0["分组"].isin(groups)]
            dfs.append(df0)
    if not dfs:
        st.warning("⚠️ 当前选择没有数据，请检查时间点或分组选择。")
        return pd.DataFrame()
    return pd.concat(dfs, axis=0, ignore_index=True)


df = get_merged_df(time_choice, selected_groups)


# -------------------- 图表函数 --------------------
def chart_total(df0):
    df0 = df0[df0["明细"] != "分数总和"]
    emp_stats = df0.groupby("员工")["值"].sum().sort_values(ascending=False).reset_index()
    fig = go.Figure(go.Bar(
        x=emp_stats["员工"],
        y=emp_stats["值"],
        text=emp_stats["值"],
        textposition="outside",
        hovertemplate="员工: %{x}<br>完成总值: %{y}<extra></extra>"
    ))
    fig.update_layout(template="plotly_dark", xaxis_title="员工", yaxis_title="完成总值")
    return fig


def chart_stack(df0):
    df0 = df0[df0["明细"] != "分数总和"]
    df_pivot = df0.pivot_table(index="明细", columns="员工", values="值", aggfunc="sum", fill_value=0)
    fig = go.Figure()
    for emp in df_pivot.columns:
        fig.add_trace(go.Bar(x=df_pivot.index, y=df_pivot[emp], name=emp))
    fig.update_layout(barmode="stack", template="plotly_dark", xaxis_title="任务", yaxis_title="完成值")
    return fig


def chart_heat(df0):
    df0 = df0[df0["明细"] != "分数总和"]
    tasks = df0["明细"].unique().tolist()
    emps = df0["员工"].unique().tolist()
    data = []
    for i, t in enumerate(tasks):
        for j, e in enumerate(emps):
            v = int(df0[(df0["明细"] == t) & (df0["员工"] == e)]["值"].sum())
            data.append([j, i, v])
    return {
        "backgroundColor": "transparent",
        "tooltip": {"position": "top"},
        "xAxis": {"type": "category", "data": emps, "axisLabel": {"color": "#fff", "rotate": 45}},
        "yAxis": {"type": "category", "data": tasks, "axisLabel": {"color": "#fff"}},
        "visualMap": {"min": 0, "max": max([d[2] for d in data]) if data else 1, "show": True,
                      "inRange": {"color": ["#ff4d4d", "#4caf50"]}, "textStyle": {"color": "#fff"}},
        "series": [{"type": "heatmap", "data": data, "emphasis": {"itemStyle": {"shadowBlur": 10}}}]
    }


# -------------------- 卡片显示 --------------------
def show_cards(df0):
    df0 = df0[df0["明细"] != "分数总和"]
    if df0.empty:
        return

    total_tasks = df0["明细"].nunique()
    total_people = df0["员工"].nunique()
    ps = df0.groupby("员工")["值"].sum()
    top_person = ps.idxmax() if not ps.empty else ""
    avg_score = round(ps.mean(), 1) if not ps.empty else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(
        f"<div class='metric-card'><div class='metric-value'>{total_tasks}</div><div class='metric-label'>任务数</div></div>",
        unsafe_allow_html=True)
    c2.markdown(
        f"<div class='metric-card'><div class='metric-value'>{total_people}</div><div class='metric-label'>人数</div></div>",
        unsafe_allow_html=True)
    c3.markdown(
        f"<div class='metric-card'><div class='metric-value'>{top_person}</div><div class='metric-label'>覆盖率最高</div></div>",
        unsafe_allow_html=True)
    c4.markdown(
        f"<div class='metric-card'><div class='metric-value'>{avg_score}</div><div class='metric-label'>平均数</div></div>",
        unsafe_allow_html=True)
    st.markdown("<hr/>", unsafe_allow_html=True)


# -------------------- 定义鲜艳的颜色列表（用于能力分析） --------------------
BRIGHT_COLORS = [
    "#FF0000",  # 红色
    "#00FF00",  # 绿色
    "#0000FF",  # 蓝色
    "#FFA500",  # 橙色
    "#800080",  # 紫色
    "#00FFFF",  # 青色
    "#FFC0CB",  # 粉色
    "#FFFF00",  # 黄色
    "#008080",  # 蓝绿色
    "#FF00FF"  # 洋红
]

# -------------------- 主页面 --------------------
st.title("📊 技能覆盖分析大屏")

if view == "编辑数据":
    if not time_choice:
        st.warning("⚠️ 请在左侧选择时间点（月或季）后再编辑数据")
    elif len(time_choice) > 1:
        st.warning("⚠️ 编辑数据时仅支持选择单个时间点，请重新选择！")
    else:
        # 卡片
        show_cards(df)
        st.info("你可以直接编辑下面的表格，修改完成后点击【保存】按钮。")

        # 读取原始完整数据（解决问题5：保留其他分组数据）
        sheet_name = time_choice[0]
        original_df = pd.read_excel(SAVE_FILE, sheet_name=sheet_name)

        # 显示筛选后的编辑表格
        edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)

        if st.button("💾 保存修改到库里"):
            try:
                # 核心修复：只更新筛选分组的数据，保留原始数据中其他分组
                if selected_groups and "分组" in original_df.columns:
                    # 1. 删除原始数据中选中分组的行
                    mask = original_df["分组"].isin(selected_groups)
                    original_df = original_df[~mask].reset_index(drop=True)
                    # 2. 合并编辑后的选中分组数据
                    final_df = pd.concat([original_df, edited_df], ignore_index=True)
                else:
                    final_df = edited_df.copy()

                # ---------- 自动计算数量总和 ----------
                if "明细" in final_df.columns and "值" in final_df.columns:
                    sum_df = (
                        final_df.groupby("明细", as_index=False)["值"].sum()
                        .rename(columns={"值": "数量总和"})
                    )
                    final_df = final_df.drop(columns=["数量总和"], errors="ignore")
                    final_df = final_df.merge(sum_df, on="明细", how="left")

                # ---------- 保存 ----------
                with pd.ExcelWriter(SAVE_FILE, mode="a", if_sheet_exists="replace", engine="openpyxl") as writer:
                    final_df.to_excel(writer, sheet_name=sheet_name, index=False)

                st.cache_data.clear()
                # 重新加载数据
                sheets, sheet_frames = load_sheets(SAVE_FILE)
                st.success(f"✅ 修改已保存到 {SAVE_FILE} ({sheet_name})，仅更新选中分组数据")
            except Exception as e:
                st.error(f"保存失败：{e}")

elif view == "大屏轮播":
    if not time_choice:
        st.warning("⚠️ 请在左侧选择时间点（月或季）后查看大屏轮播")
    else:
        st_autorefresh(interval=10000, key="aut")
        show_cards(df)
        # 移除热门任务，只保留3个图表轮播
        secs = [("完成排名", chart_total(df)),
                ("任务对比", chart_stack(df)),
                ("热力图", chart_heat(df))]
        t, op = secs[int(time.time() / 10) % len(secs)]
        st.subheader(t)
        if isinstance(op, go.Figure):
            st.plotly_chart(op, use_container_width=True)
        else:
            # 热力图添加滚动容器
            st.markdown('<div class="heatmap-container">', unsafe_allow_html=True)
            st_echarts(op, height=f"{max(600, len(df['明细'].unique()) * 25)}px", theme="dark")
            st.markdown('</div>', unsafe_allow_html=True)

elif view == "单页模式":
    if not time_choice:
        st.warning("⚠️ 请在左侧选择时间点（月或季）后查看单页模式")
    else:
        show_cards(df)
        choice = st.sidebar.selectbox("单页查看", sections_names, index=0)
        mapping = {
            "人员完成任务数量排名": chart_total(df),
            "任务对比（堆叠柱状图）": chart_stack(df),
            "任务-人员热力图": chart_heat(df)
        }
        chart_func = mapping.get(choice, chart_total(df))
        if isinstance(chart_func, go.Figure):
            st.plotly_chart(chart_func, use_container_width=True)
        else:
            # 热力图添加滚动容器
            st.markdown('<div class="heatmap-container">', unsafe_allow_html=True)
            st_echarts(chart_func, height=f"{max(600, len(df['明细'].unique()) * 25)}px", theme="dark")
            st.markdown('</div>', unsafe_allow_html=True)

elif view == "显示所有视图":
    if not time_choice:
        st.warning("⚠️ 请在左侧选择时间点（月或季）后查看所有视图")
    else:
        show_cards(df)
        # 移除热门任务，只保留3个图表
        charts = [("完成排名", chart_total(df)),
                  ("任务对比（堆叠柱状图）", chart_stack(df)),
                  ("热图", chart_heat(df))]
        for label, f in charts:
            st.subheader(label)
            if isinstance(f, go.Figure):
                st.plotly_chart(f, use_container_width=True)
            else:
                # 热力图添加滚动容器
                st.markdown('<div class="heatmap-container">', unsafe_allow_html=True)
                st_echarts(f, height=f"{max(600, len(df['明细'].unique()) * 25)}px", theme="dark")
                st.markdown('</div>', unsafe_allow_html=True)

elif view == "能力分析":
    if not time_choice:
        st.warning("⚠️ 请在左侧选择时间点（月或季）后查看能力分析")
    else:
        st.subheader("📈 能力分析")
        employees = df["员工"].unique().tolist()
        selected_emps = st.sidebar.multiselect("选择员工（图1显示）", employees, default=employees)
        tasks = df["明细"].unique().tolist()

        fig1, fig2, fig3 = go.Figure(), go.Figure(), go.Figure()

        # 核心优化：为每个时间点分配固定颜色，确保fig2和fig3颜色一致
        sheet_color_map = {}
        for idx, sheet in enumerate(time_choice):
            sheet_color_map[sheet] = BRIGHT_COLORS[idx % len(BRIGHT_COLORS)]

        # 遍历每个时间点处理数据
        emp_color_idx = 0
        for sheet in time_choice:
            df_sheet = get_merged_df([sheet], selected_groups)
            df_sheet = df_sheet[df_sheet["明细"] != "分数总和"]
            df_pivot = df_sheet.pivot(index="明细", columns="员工", values="值").fillna(0)

            # 1. 员工任务完成情况 - 折线图
            for emp in selected_emps:
                fig1.add_trace(go.Scatter(
                    x=tasks,
                    y=df_pivot[emp].reindex(tasks, fill_value=0),
                    mode="lines+markers",
                    name=f"{sheet}-{emp}",
                    line=dict(color=BRIGHT_COLORS[emp_color_idx % len(BRIGHT_COLORS)], width=3),
                    marker=dict(size=8)
                ))
                emp_color_idx += 1

            # 2. 任务整体完成度趋势 - 折线图（固定颜色映射）
            fig2.add_trace(go.Scatter(
                x=tasks,
                y=df_pivot.sum(axis=1).reindex(tasks, fill_value=0),
                mode="lines+markers",
                name=sheet,
                line=dict(color=sheet_color_map[sheet], width=3),
                marker=dict(size=8)
            ))

            # 3. 员工整体完成度对比 - 分组柱状图（彻底解决重叠问题）
            fig3.add_trace(go.Bar(
                x=df_pivot.columns,
                y=df_pivot.sum(axis=0),
                name=sheet,
                marker=dict(color=sheet_color_map[sheet]),
                width=0.3,  # 极致缩小宽度，避免重叠
            ))

        # 优化图表样式 - 重点修复柱状图布局
        fig1.update_layout(
            title="员工任务完成情况",
            template="plotly_dark",
            font=dict(size=12),
            legend=dict(orientation="h", yanchor="bottom", y=-0.3, xanchor="center", x=0.5),
            height=500
        )

        fig2.update_layout(
            title="任务整体完成度趋势",
            template="plotly_dark",
            font=dict(size=12),
            legend=dict(orientation="h", yanchor="bottom", y=-0.3, xanchor="center", x=0.5),
            height=500
        )

        # 柱状图核心优化配置
        fig3.update_layout(
            title="员工整体完成度对比",
            template="plotly_dark",
            font=dict(size=12),
            barmode="group",  # 分组模式（核心）
            bargap=0.25,  # 员工组之间的间距（增大）
            bargroupgap=0.005,  # 同一员工不同时间点柱子的间距（减小）
            legend=dict(orientation="h", yanchor="bottom", y=-0.3, xanchor="center", x=0.5),
            height=600,  # 增加图表高度，提升展示效果
            xaxis=dict(
                tickangle=45,  # X轴标签旋转45度，避免拥挤
                tickfont=dict(size=10)
            ),
            yaxis=dict(
                tickfont=dict(size=10)
            )
        )

        st.plotly_chart(fig1, use_container_width=True)
        st.plotly_chart(fig2, use_container_width=True)
        st.plotly_chart(fig3, use_container_width=True)
