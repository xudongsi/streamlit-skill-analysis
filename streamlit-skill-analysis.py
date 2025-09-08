# app.py

import os, time, io
import pandas as pd
import streamlit as st
from typing import List, Tuple
from streamlit_autorefresh import st_autorefresh
from streamlit_echarts import st_echarts
import plotly.graph_objects as go

# -------------------- 页面配置 --------------------
st.set_page_config(page_title="技能覆盖分析大屏", layout="wide")

# -------------------- 数据读取 --------------------
def load_excel_files(folder: str) -> dict:
    data = {}
    if not os.path.exists(folder):
        return data
    for file in os.listdir(folder):
        if file.endswith(".xlsx"):
            filepath = os.path.join(folder, file)
            try:
                xls = pd.ExcelFile(filepath)
                for sheet in xls.sheet_names:
                    df = xls.parse(sheet)
                    data[(file, sheet)] = df
            except Exception as e:
                st.error(f"读取 {file} 出错: {e}")
    return data

# -------------------- 数据处理 --------------------
def get_merged_df(time_choice: List[str], groups: List[str]) -> pd.DataFrame:
    if not time_choice or not groups:
        return pd.DataFrame()

    dfs = []
    for (file, sheet), df in DATA.items():
        if sheet in time_choice:
            if "分组" in df.columns:
                df = df[df["分组"].isin(groups)]
            dfs.append(df)

    if not dfs:
        return pd.DataFrame()

    merged = pd.concat(dfs, ignore_index=True)
    return merged

# -------------------- 视图函数 --------------------
def show_cards(df: pd.DataFrame):
    if df.empty:
        return
    total_tasks = df["明细"].nunique()
    total_emps = df["员工"].nunique()
    total_value = df["值"].sum()

    cols = st.columns(3)
    cols[0].metric("任务数", total_tasks)
    cols[1].metric("员工数", total_emps)
    cols[2].metric("总完成值", total_value)

def chart_total(df: pd.DataFrame):
    if df.empty:
        return go.Figure()
    df_sum = df.groupby("员工")["值"].sum().sort_values(ascending=False)
    fig = go.Figure([go.Bar(x=df_sum.index, y=df_sum.values)])
    fig.update_layout(title="人员完成任务数量排名", template="plotly_dark")
    return fig

def chart_stack(df: pd.DataFrame):
    if df.empty:
        return go.Figure()
    df_pivot = df.pivot_table(index="明细", columns="员工", values="值", aggfunc="sum", fill_value=0)
    fig = go.Figure()
    for col in df_pivot.columns:
        fig.add_bar(name=col, x=df_pivot.index, y=df_pivot[col])
    fig.update_layout(barmode="stack", title="任务对比（堆叠柱状图）", template="plotly_dark")
    return fig

def chart_bubble(df: pd.DataFrame):
    if df.empty:
        return go.Figure()
    df_sum = df.groupby(["员工", "明细"])["值"].sum().reset_index()
    fig = go.Figure()
    for emp in df_sum["员工"].unique():
        d = df_sum[df_sum["员工"] == emp]
        fig.add_trace(go.Scatter(x=d["明细"], y=d["值"], mode="markers", name=emp,
                                 marker=dict(size=d["值"], sizemode="area", sizeref=2.*max(d["值"])/(40.**2))))
    fig.update_layout(title="人员对比（气泡图）", template="plotly_dark")
    return fig

def chart_hot(df: pd.DataFrame):
    if df.empty:
        return go.Figure()
    df_sum = df.groupby("明细")["值"].sum().sort_values(ascending=False).head(10)
    fig = go.Figure([go.Bar(x=df_sum.index, y=df_sum.values)])
    fig.update_layout(title="任务掌握情况（热门任务）", template="plotly_dark")
    return fig

def chart_heat(df: pd.DataFrame):
    if df.empty:
        return {}
    df_pivot = df.pivot_table(index="明细", columns="员工", values="值", aggfunc="sum", fill_value=0)
    data = []
    for i, task in enumerate(df_pivot.index):
        for j, emp in enumerate(df_pivot.columns):
            data.append([j, i, df_pivot.loc[task, emp]])
    option = {
        "tooltip": {"position": "top"},
        "xAxis": {"type": "category", "data": list(df_pivot.columns)},
        "yAxis": {"type": "category", "data": list(df_pivot.index)},
        "visualMap": {"min": 0, "max": int(df_pivot.values.max()), "calculable": True, "orient": "horizontal"},
        "series": [{
            "type": "heatmap",
            "data": data,
            "label": {"show": True}
        }]
    }
    return option

# -------------------- 主逻辑 --------------------
DATA = load_excel_files("data")

st.sidebar.title("📂 参数选择")
time_choice = st.sidebar.multiselect("选择时间点", sorted({sheet for _, sheet in DATA.keys()}))
groups = []
if DATA:
    sample_df = list(DATA.values())[0]
    if "分组" in sample_df.columns:
        groups = sample_df["分组"].unique().tolist()
selected_groups = st.sidebar.multiselect("选择分组", groups)

df = get_merged_df(time_choice, selected_groups)

view = st.sidebar.radio("选择视图模式", ["编辑数据", "大屏轮播", "单页模式", "显示所有视图", "能力分析"])

st.title("📊 技能覆盖分析大屏")

# -------- 统一拦截空数据 --------
if df.empty:
    st.warning("⚠️ 请选择对应窗口（时间点 / 分组）")
else:
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
        choice = st.sidebar.selectbox("单页查看", [
            "人员完成任务数量排名",
            "任务对比（堆叠柱状图）",
            "人员对比（气泡图）",
            "任务掌握情况（热门任务）",
            "任务-人员热力图"
        ], index=0)
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
