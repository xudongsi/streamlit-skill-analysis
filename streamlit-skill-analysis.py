# app.py
import os
import time
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
</style>
"""
st.markdown(PAGE_CSS, unsafe_allow_html=True)

SAVE_FILE = "jixiao.xlsx"   # 固定保存的文件

# -------------------- 数据导入 --------------------
@st.cache_data
def load_sheets(file) -> Tuple[List[str], dict]:
    xpd = pd.ExcelFile(file, engine="openpyxl")
    frames = {}
    for s in xpd.sheet_names:
        df0 = pd.read_excel(xpd, sheet_name=s, engine="openpyxl")
        if not df0.empty and df0.iloc[0, 0] == "分组":  # 第一行是分组信息
            col_names = list(df0.columns)
            groups = df0.iloc[0, 1:].tolist()
            df0 = df0.drop(0).reset_index(drop=True)

            emp_cols = [c for c in col_names if c not in ["明细", "数量总和", "编号"]]

            group_map = {}
            for emp in emp_cols:
                try:
                    pos = col_names.index(emp)
                    group_map[emp] = groups[pos - 1] if (pos - 1) < len(groups) and pos - 1 >= 0 else None
                except ValueError:
                    group_map[emp] = None

            id_vars = ["明细"] + (["数量总和"] if "数量总和" in df0.columns else [])
            df_long = df0.melt(
                id_vars=id_vars,
                value_vars=emp_cols,
                var_name="员工",
                value_name="值"
            )
            df_long["分组"] = df_long["员工"].map(group_map)
            frames[s] = df_long
        else:
            frames[s] = df0
    return xpd.sheet_names, frames


# -------------------- 文件读取逻辑 --------------------
sheets, sheet_frames = [], {}
try:
    sheets, sheet_frames = load_sheets(SAVE_FILE)
    st.sidebar.success(f"已加载库文件 {SAVE_FILE}")
except Exception as e:
    st.sidebar.warning(f"读取库文件失败：{e}")
    sheet_frames = {
        "示例": pd.DataFrame({
            "明细": ["任务A", "任务B", "任务C"],
            "数量总和": [3, 2, 5],
            "员工": ["张三", "李四", "王五"],
            "值": [1, 1, 1],
            "分组": ["A8", "B7", "VN"]
        })
    }
    sheets = ["示例"]

# -------------------- ✅ 新增月份/季度 --------------------
new_sheet_name = st.sidebar.text_input("➕ 新增时间点（月或季）")

if st.sidebar.button("创建新的时间点"):
    if new_sheet_name:
        try:
            if os.path.exists(SAVE_FILE):
                with pd.ExcelWriter(SAVE_FILE, mode="a", engine="openpyxl") as writer:
                    pd.DataFrame(columns=["明细", "数量总和", "员工", "值", "分组"]).to_excel(
                        writer, sheet_name=new_sheet_name, index=False
                    )
            else:
                with pd.ExcelWriter(SAVE_FILE, engine="openpyxl") as writer:
                    pd.DataFrame(columns=["明细", "数量总和", "员工", "值", "分组"]).to_excel(
                        writer, sheet_name=new_sheet_name, index=False
                    )
            st.cache_data.clear()
            st.sidebar.success(f"✅ 已在 {SAVE_FILE} 创建新时间点: {new_sheet_name}")
        except Exception as e:
            st.sidebar.error(f"创建失败：{e}")
    else:
        st.sidebar.warning("请输入时间点名称后再点击创建")


# -------------------- 🗑️ 删除月份/季度 --------------------
del_sheet_name = st.sidebar.selectbox("🗑️ 选择要删除的时间点", sheets)

if st.sidebar.button("删除所选时间点"):
    if del_sheet_name:
        try:
            xls = pd.ExcelFile(SAVE_FILE, engine="openpyxl")
            keep_sheets = [s for s in xls.sheet_names if s != del_sheet_name]

            if not keep_sheets:
                st.sidebar.warning("⚠️ 至少要保留一个时间点，不能全部删除")
            else:
                # 先读入保留的 sheet
                keep_dfs = {}
                for s in keep_sheets:
                    keep_dfs[s] = pd.read_excel(SAVE_FILE, sheet_name=s, engine="openpyxl")

                # 再写入（覆盖原文件）
                with pd.ExcelWriter(SAVE_FILE, engine="openpyxl", mode="w") as writer:
                    for s, df_tmp in keep_dfs.items():
                        df_tmp.to_excel(writer, sheet_name=s, index=False)

                st.cache_data.clear()
                st.sidebar.success(f"✅ 已删除时间点: {del_sheet_name}")
        except Exception as e:
            st.sidebar.error(f"删除失败：{e}")


# -------------------- 时间和分组选择 --------------------
time_choice = st.sidebar.multiselect("选择时间点（月或季）", sheets, default=sheets[:1])

def collect_all_groups(sheet_frames):
    dfs_with_group = [df for df in sheet_frames.values() if isinstance(df, pd.DataFrame) and "分组" in df.columns]
    if not dfs_with_group:
        return []
    return pd.concat(dfs_with_group, axis=0, ignore_index=True)["分组"].dropna().unique().tolist()

all_groups = collect_all_groups(sheet_frames)
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
    if not {"员工", "值"}.issubset(df0.columns):
        return go.Figure()
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
    if not {"员工", "明细", "值"}.issubset(df0.columns):
        return go.Figure()
    df0 = df0[df0["明细"] != "分数总和"]
    df_pivot = df0.pivot_table(index="明细", columns="员工", values="值", aggfunc="sum", fill_value=0)
    fig = go.Figure()
    for emp in df_pivot.columns:
        fig.add_trace(go.Bar(x=df_pivot.index, y=df_pivot[emp], name=emp))
    fig.update_layout(barmode="stack", template="plotly_dark", xaxis_title="任务", yaxis_title="完成值")
    return fig

def chart_bubble(df0):
    if not {"员工", "明细", "值"}.issubset(df0.columns):
        return go.Figure()
    df0 = df0[df0["明细"] != "分数总和"]
    emp_stats = df0.groupby("员工").agg(
        任务数=("明细","nunique"),
        总值=("值","sum")
    ).reset_index()
    emp_stats["覆盖率"] = emp_stats["任务数"] / df0["明细"].nunique()
    sizes = emp_stats["总值"].astype(float).tolist()
    if not sizes:
        return go.Figure()
    fig = go.Figure(data=[go.Scatter(
        x=emp_stats["任务数"],
        y=emp_stats["覆盖率"],
        mode="markers+text",
        text=emp_stats["员工"],
        textposition="top center",
        marker=dict(size=sizes, sizemode="area",
                    sizeref=2.*max(sizes)/(40.**2),
                    sizemin=8, color=emp_stats["总值"],
                    colorscale="Viridis", showscale=True)
    )])
    fig.update_layout(template="plotly_dark", xaxis_title="任务数", yaxis_title="覆盖率")
    return fig

def chart_hot(df0):
    if not {"员工", "明细"}.issubset(df0.columns):
        return {}
    ts = df0[df0["明细"] != "分数总和"].groupby("明细")["员工"].nunique()
    return {
        "backgroundColor":"transparent",
        "yAxis":{"type":"category","data":ts.index.tolist(),"axisLabel":{"color":"#fff"}},
        "xAxis":{"type":"value","axisLabel":{"color":"#fff"}},
        "series":[{"data":ts.tolist(),"type":"bar","itemStyle":{"color":"#ffb703"}}]
    }

def chart_heat(df0):
    if not {"员工", "明细", "值"}.issubset(df0.columns):
        return {}
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
    if not {"员工", "明细", "值"}.issubset(df0.columns):
        return
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
    if not time_choice:
        st.warning("⚠️ 请在左侧选择时间点（月或季）后再编辑数据")
    else:
        show_cards(df)
        st.info("你可以直接编辑下面的表格，修改完成后点击【保存】按钮。")

        edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)

        if st.button("💾 保存修改到库里"):
            try:
                sheet_name = time_choice[0]
                tmp_file = SAVE_FILE + ".tmp.xlsx"
                existing = {}
                if os.path.exists(SAVE_FILE):
                    xls = pd.ExcelFile(SAVE_FILE, engine="openpyxl")
                    for s in xls.sheet_names:
                        if s != sheet_name:
                            existing[s] = pd.read_excel(SAVE_FILE, sheet_name=s, engine="openpyxl")
                with pd.ExcelWriter(tmp_file, engine="openpyxl", mode="w") as writer:
                    for s, df_s in existing.items():
                        df_s.to_excel(writer, sheet_name=s, index=False)
                    edited_df.to_excel(writer, sheet_name=sheet_name, index=False)
                os.replace(tmp_file, SAVE_FILE)
                st.cache_data.clear()
                st.success(f"✅ 修改已保存到 {SAVE_FILE} ({sheet_name})")
            except Exception as e:
                st.error(f"保存失败：{e}")
        st.dataframe(edited_df)

elif view == "大屏轮播":
    if not time_choice:
        st.warning("⚠️ 请在左侧选择时间点（月或季）后查看大屏轮播")
    else:
        st_autorefresh(interval=10000, key="aut")
        show_cards(df)
        secs = [("完成排名", chart_total(df)),
                ("任务对比", chart_stack(df)),
                ("人员对比", chart_bubble(df)),
                ("热门任务", chart_hot(df)),
                ("热力图", chart_heat(df))]
        t, op = secs[int(time.time()/10) % len(secs)]
        st.subheader(t)
        if isinstance(op, go.Figure):
            st.plotly_chart(op, use_container_width=True)
        else:
            st_echarts(op, height="600px", theme="dark")

elif view == "单页模式":
    if not time_choice:
        st.warning("⚠️ 请在左侧选择时间点（月或季）后查看单页模式")
    else:
        show_cards(df)
        choice = st.sidebar.selectbox("单页查看", sections_names, index=0)
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
    if not time_choice:
        st.warning("⚠️ 请在左侧选择时间点（月或季）后查看所有视图")
    else:
        show_cards(df)
        charts = [("完成排名", chart_total(df)),
                  ("任务对比（堆叠柱状图）", chart_stack(df)),
                  ("人员对比（气泡图）", chart_bubble(df)),
                  ("热门任务", chart_hot(df)),
                  ("热图", chart_heat(df))]
        for label, f in charts:
            st.subheader(label)
            if isinstance(f, go.Figure):
                st.plotly_chart(f, use_container_width=True)
            else:
                st_echarts(f, height="520px", theme="dark")

elif view == "能力分析":
    if not time_choice:
        st.warning("⚠️ 请在左侧选择时间点（月或季）后查看能力分析")
else:
    show_cards(df)
    if "值" in df.columns and "员工" in df.columns:
        st.subheader("📈 分数分布情况")
        score = df.groupby("员工")["值"].sum()
        fig = go.Figure()
        fig.add_trace(go.Histogram(x=score, nbinsx=10, name="分数分布"))
        fig.update_layout(template="plotly_dark", xaxis_title="分数", yaxis_title="人数")
        st.plotly_chart(fig, use_container_width=True)

    if "分组" in df.columns and "员工" in df.columns:
        st.subheader("📊 分组人数")
        grp = df.groupby("分组")["员工"].nunique()
        fig = go.Figure(go.Bar(x=grp.index, y=grp.values, text=grp.values, textposition="outside"))
        fig.update_layout(template="plotly_dark", xaxis_title="分组", yaxis_title="人数")
        st.plotly_chart(fig, use_container_width=True)
