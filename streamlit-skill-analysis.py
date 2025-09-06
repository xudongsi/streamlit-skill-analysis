# app.py

import os
import time
import io
from typing import List, Tuple

import pandas as pd
import streamlit as st
from streamlit_autorefresh import st_autorefresh
from streamlit_echarts import st_echarts
import plotly.graph_objects as go

# -------------------- 页面配置 --------------------
st.set_page_config(page_title="技能覆盖分析大屏", layout="wide")

# -------------------- 页面样式 --------------------
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

/* ----------- 分辨率自适应 ----------- */
@media screen and (max-width: 1600px) {
    .metric-value { font-size:28px; }
    .metric-card { padding:16px; }
}
@media screen and (max-width: 1200px) {
    .metric-value { font-size:22px; }
    .metric-label { font-size:12px; }
    div.stButton>button { height:36px; font-size:13px; }
}
@media screen and (max-width: 900px) {
    .metric-card { padding:12px; }
    .metric-value { font-size:18px; }
    .metric-label { font-size:11px; }
    div.stButton>button { height:32px; font-size:12px; }
    .block-container { padding-left:0.5rem; padding-right:0.5rem; }
}
@media screen and (max-width: 600px) {
    .metric-card { padding:8px; }
    .metric-value { font-size:16px; }
    .metric-label { font-size:10px; }
    div.stButton>button { width:100%; font-size:11px; }
}
</style>
"""
st.markdown(PAGE_CSS, unsafe_allow_html=True)


# -------------------- 左侧栏 --------------------
st.sidebar.header("📂 数据控制区")
upload = st.sidebar.file_uploader("上传 Excel（Sheet 名称＝月份或季度）", type=["xlsx", "xls"])

# -------------------- 数据导入 --------------------
@st.cache_data
def load_sheets(file) -> Tuple[List[str], dict]:
    xpd = pd.ExcelFile(file)
    frames = {}
    for s in xpd.sheet_names:
        df0 = pd.read_excel(xpd, sheet_name=s)
        if df0.iloc[0, 0] == "分组":  # 第一行是分组信息
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
    return xpd.sheet_names, frames

sheets, sheet_frames = [], {}
if upload:
    try:
        sheets, sheet_frames = load_sheets(upload)
    except Exception as e:
        st.sidebar.error(f"读取失败：{e}")

if not sheets:
    st.sidebar.info("未上传, 使用示例")
    sheet_frames = {
        "示例": pd.DataFrame({
            "明细": ["任务A", "任务B", "任务C", "分数总和"],
            "数量总和": [3, 2, 5, 10],
            "员工": ["张三", "李四", "王五", ""],
            "值": [1, 1, 1, 0],
            "分组": ["A8", "B7", "VN", ""]
        })
    }
    sheets = ["示例"]

# -------------------- 时间和分组选择 --------------------
time_choice = st.sidebar.multiselect("选择时间点（月或季）", sheets, default=sheets[:3] if len(sheets) >= 3 else sheets)
all_groups = pd.concat(sheet_frames.values())["分组"].dropna().unique().tolist()
selected_groups = st.sidebar.multiselect("选择分组", all_groups, default=all_groups)

sections_names = [
    "人员完成任务数量排名",
    "任务对比（堆叠柱状图）",
    "人员对比（气泡图）",
    "任务掌握情况（热门任务）",
    "任务-人员热力图"
]
view = st.sidebar.radio("切换视图", ["编辑数据", "大屏轮播", "单页模式", "显示所有视图", "能力分析"])

# -------------------- 数据合并 --------------------
def get_merged_df(keys: List[str], groups: List[str]) -> pd.DataFrame:
    dfs = []
    for k in keys:
        df0 = sheet_frames.get(k)
        if df0 is not None:
            if groups and "分组" in df0.columns:
                df0 = df0[df0["分组"].isin(groups)]
            dfs.append(df0)
    if not dfs:
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
        fig.add_trace(go.Bar(
            x=df_pivot.index,
            y=df_pivot[emp],
            name=emp,
            hovertemplate="任务: %{x}<br>员工: " + emp + "<br>完成值: %{y}<extra></extra>"
        ))
    fig.update_layout(barmode="stack", template="plotly_dark", xaxis_title="任务", yaxis_title="完成值")
    return fig

def chart_bubble(df0):
    df0 = df0[df0["明细"] != "分数总和"]
    emp_stats = df0.groupby("员工").agg(
        任务数=("明细","nunique"),
        总值=("值","sum")
    ).reset_index()
    emp_stats["覆盖率"] = emp_stats["任务数"] / df0["明细"].nunique()
    sizes = emp_stats["总值"].astype(float).tolist()
    fig = go.Figure(data=[go.Scatter(
        x=emp_stats["任务数"],
        y=emp_stats["覆盖率"],
        mode="markers+text",
        text=emp_stats["员工"],
        textposition="top center",
        hovertemplate="员工: %{text}<br>任务数: %{x}<br>覆盖率: %{y:.2%}<br>总值: %{marker.size}",
        marker=dict(
            size=sizes,
            sizemode="area",
            sizeref=2.*max(sizes)/(40.**2),
            sizemin=8,
            color=emp_stats["总值"],
            colorscale="Viridis",
            showscale=True
        )
    )])
    fig.update_layout(template="plotly_dark", xaxis_title="任务数", yaxis_title="覆盖率")
    return fig

def chart_hot(df0):
    df0 = df0[df0["明细"] != "分数总和"]
    ts = df0.groupby("明细")["员工"].nunique()
    return {
        "backgroundColor":"transparent",
        "yAxis":{"type":"category","data":ts.index.tolist(),"axisLabel":{"color":"#fff"}},
        "xAxis":{"type":"value","axisLabel":{"color":"#fff"}},
        "series":[{"data":ts.tolist(),"type":"bar","itemStyle":{"color":"#ffb703"}}]
    }

def chart_heat(df0):
    df0 = df0[df0["明细"] != "分数总和"]
    tasks = df0["明细"].unique().tolist()
    emps = df0["员工"].unique().tolist()
    data=[]
    for i,t in enumerate(tasks):
        for j,e in enumerate(emps):
            v=int(df0[(df0["明细"]==t)&(df0["员工"]==e)]["值"].sum())
            data.append([j,i,v])
    return {
        "backgroundColor":"transparent",
        "tooltip":{"position":"top"},
        "xAxis":{"type":"category","data":emps,"axisLabel":{"color":"#fff"}},
        "yAxis":{"type":"category","data":tasks,"axisLabel":{"color":"#fff"}},
        "visualMap":{"min":0,"max":1,"show":False,"inRange":{"color":["#ff4d4d","#4caf50"]}},
        "series":[{"type":"heatmap","data":data}]
    }

# -------------------- 卡片显示 --------------------
def show_cards(df0):
    df0 = df0[df0["明细"] != "分数总和"]
    total_tasks = df0["明细"].nunique()
    total_people = df0["员工"].nunique()
    ps = df0.groupby("员工")["值"].sum()
    top_person = ps.idxmax() if not ps.empty else ""
    avg_score = round(ps.mean(),1) if not ps.empty else 0

    c1,c2,c3,c4 = st.columns(4)
    c1.markdown(f"<div class='metric-card'><div class='metric-value'>{total_tasks}</div><div class='metric-label'>任务数</div></div>", unsafe_allow_html=True)
    c2.markdown(f"<div class='metric-card'><div class='metric-value'>{total_people}</div><div class='metric-label'>人数</div></div>", unsafe_allow_html=True)
    c3.markdown(f"<div class='metric-card'><div class='metric-value'>{top_person}</div><div class='metric-label'>覆盖率最高</div></div>", unsafe_allow_html=True)
    c4.markdown(f"<div class='metric-card'><div class='metric-value'>{avg_score}</div><div class='metric-label'>平均数</div></div>", unsafe_allow_html=True)
    st.markdown("<hr/>", unsafe_allow_html=True)

# -------------------- 主页面 --------------------
st.title("📊 技能覆盖分析大屏")

if view == "编辑数据":
    show_cards(df)
    st.dataframe(df)

elif view == "大屏轮播":
    st_autorefresh(interval=10000, key="aut")
    show_cards(df)
    secs = [
        ("完成排名", chart_total(df)),
        ("任务对比", chart_stack(df)),
        ("人员对比", chart_bubble(df)),
        ("热门任务", chart_hot(df)),
        ("热力图", chart_heat(df))
    ]
    t, op = secs[int(time.time()/10) % len(secs)]
    st.subheader(t)
    if isinstance(op, go.Figure):
        st.plotly_chart(op, use_container_width=True)
    else:
        st_echarts(op, height="600px", theme="dark")

elif view == "单页模式":
    show_cards(df)
    choice = st.sidebar.selectbox("单页查看", sections_names, index=0)
    st.subheader(choice)
    mapping = {
        "人员完成任务数量排名": chart_total(df),
        "任务对比（堆叠柱状图）": chart_stack(df),
        "人员对比（气泡图）": chart_bubble(df),
        "任务掌握情况（热门任务）": chart_hot(df),
        "任务-人员热力图": chart_heat(df)
    }
    chart_func = mapping.get(choice, chart_total(df))
    if isinstance(chart_func, go.Figure):
        st.plotly_chart(chart_func, use_container_width=True)
    else:
        st_echarts(chart_func, height="600px", theme="dark")

elif view == "显示所有视图":
    show_cards(df)
    charts = [
        ("完成排名", chart_total(df)),
        ("任务对比（堆叠柱状图）", chart_stack(df)),
        ("人员对比（气泡图）", chart_bubble(df)),
        ("热门任务", chart_hot(df)),
        ("热图", chart_heat(df))
    ]
    for label, f in charts:
        st.subheader(label)
        if isinstance(f, go.Figure):
            st.plotly_chart(f, use_container_width=True)
        else:
            st_echarts(f, height="520px", theme="dark")

elif view == "能力分析":
    st.subheader("📊 能力分析")
    employees = df["员工"].unique().tolist()
    selected_emps = st.sidebar.multiselect("选择员工（图1显示）", employees, default=employees)
    tasks = df["明细"].unique().tolist()

    fig1, fig2, fig3 = go.Figure(), go.Figure(), go.Figure()
    for sheet in time_choice:
        df_sheet = get_merged_df([sheet], selected_groups)
        df_sheet = df_sheet[df_sheet["明细"] != "分数总和"]
        df_pivot = df_sheet.pivot(index="明细", columns="员工", values="值").fillna(0)

        # 图1: 员工在任务上的表现
        for emp in selected_emps:
            fig1.add_trace(go.Scatter(
                x=tasks,
                y=df_pivot[emp].reindex(tasks, fill_value=0),
                mode="lines+markers",
                name=f"{sheet}-{emp}"
            ))

        # 图2: 各任务整体完成度趋势
        fig2.add_trace(go.Scatter(
            x=tasks,
            y=df_pivot.sum(axis=1).reindex(tasks, fill_value=0),
            mode="lines+markers",
            name=sheet
        ))

        # 图3: 各员工整体完成度
        fig3.add_trace(go.Scatter(
            x=df_pivot.columns,
            y=df_pivot.sum(axis=0),
            mode="lines+markers",
            name=sheet
        ))

    fig1.update_layout(title="员工任务完成情况", xaxis_title="任务", yaxis_title="完成值", template="plotly_dark")
    fig2.update_layout(title="任务整体完成度趋势", xaxis_title="任务", yaxis_title="总完成值", template="plotly_dark")
    fig3.update_layout(title="员工整体完成度对比", xaxis_title="员工", yaxis_title="总完成值", template="plotly_dark")

    st.plotly_chart(fig1, use_container_width=True)
    st.plotly_chart(fig2, use_container_width=True)
    st.plotly_chart(fig3, use_container_width=True)
