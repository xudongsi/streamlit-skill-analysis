import os, io, time, requests
import pandas as pd
import streamlit as st
from typing import List, Tuple
from streamlit_autorefresh import st_autorefresh
from streamlit_echarts import st_echarts
import plotly.graph_objects as go
from github import Github   # pip install PyGithub

st.set_page_config(page_title="技能覆盖分析大屏", layout="wide")

# ========== GitHub 配置 ==========
GITHUB_REPO = "xudongsi/streamlit-skill-analysis"  # 你的 GitHub 用户名/仓库名
GITHUB_FILE_PATH = "data/jixiao.xlsx"              # 仓库里的 Excel 文件路径
BRANCH = "main"
TOKEN = st.secrets["GITHUB_TOKEN"]                 # secrets.toml 里配置你的 GitHub token

# ========== 确保 GitHub 文件存在 ==========
def ensure_github_file():
    g = Github(TOKEN)
    repo = g.get_repo(GITHUB_REPO)
    try:
        # 检查文件是否存在
        repo.get_contents(GITHUB_FILE_PATH, ref=BRANCH)
    except:
        # 文件不存在，创建 data 文件夹 + 空 Excel 文件
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            pd.DataFrame({"明细": [], "值": []}).to_excel(writer, sheet_name="Sheet1", index=False)
        content = output.getvalue()
        repo.create_file(GITHUB_FILE_PATH, "create initial excel file", content, branch=BRANCH)
        st.info("📁 GitHub 文件不存在，已自动创建 jixiao.xlsx 初始模板。")

# 确保文件存在
ensure_github_file()

# ========== 从 GitHub 读取 Excel ==========
@st.cache_data
def load_excel_from_github() -> Tuple[List[str], dict]:
    url = f"https://raw.githubusercontent.com/{GITHUB_REPO}/{BRANCH}/{GITHUB_FILE_PATH}"
    r = requests.get(url)
    r.raise_for_status()
    xls = io.BytesIO(r.content)
    xpd = pd.ExcelFile(xls)
    frames = {}
    for s in xpd.sheet_names:
        df0 = pd.read_excel(xpd, sheet_name=s)
        if not df0.empty and df0.iloc[0, 0] == "分组":
            groups = df0.iloc[0, 1:].tolist()
            df0 = df0.drop(0).reset_index(drop=True)
            emp_cols = [c for c in df0.columns if c not in ["明细", "数量总和", "编号"]]
            group_map = {emp: groups[i] if i < len(groups) else None for i, emp in enumerate(emp_cols)}
            df_long = df0.melt(
                id_vars=["明细", "数量总和"] if "数量总和" in df0.columns else ["明细"],
                value_vars=emp_cols,
                var_name="员工", value_name="值"
            )
            df_long["分组"] = df_long["员工"].map(group_map)
            frames[s] = df_long
        else:
            frames[s] = df0
    return xpd.sheet_names, frames

# ========== 保存 Excel 到 GitHub ==========
def save_excel_to_github(frames: dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for name, df in frames.items():
            df.to_excel(writer, sheet_name=name, index=False)
    content = output.getvalue()

    g = Github(TOKEN)
    repo = g.get_repo(GITHUB_REPO)
    try:
        file = repo.get_contents(GITHUB_FILE_PATH, ref=BRANCH)
        repo.update_file(file.path, "update excel from Streamlit", content, file.sha, branch=BRANCH)
    except:
        repo.create_file(GITHUB_FILE_PATH, "create excel from Streamlit", content, branch=BRANCH)
    st.success("✅ 数据已保存到 GitHub！")

# ========== 读取数据 ==========
sheets, sheet_frames = load_excel_from_github()

# ========== 左侧栏配置 ==========
st.sidebar.header("📂 控制区")
time_choice = st.sidebar.multiselect("选择时间点（月或季）", sheets, default=sheets[:3] if len(sheets)>=3 else sheets)
all_groups = pd.concat(sheet_frames.values())["分组"].dropna().unique().tolist()
selected_groups = st.sidebar.multiselect("选择分组", all_groups, default=all_groups)
view = st.sidebar.radio("切换视图", ["编辑数据", "大屏轮播", "单页模式", "显示所有视图", "能力分析"])

# ========== 合并数据 ==========
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

# ========== 图表函数 ==========
def chart_total(df0):
    df0 = df0[df0["明细"] != "分数总和"]
    emp_stats = df0.groupby("员工")["值"].sum().sort_values(ascending=False).reset_index()
    fig = go.Figure(go.Bar(
        x=emp_stats["员工"], y=emp_stats["值"], text=emp_stats["值"], textposition="outside",
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
            x=df_pivot.index, y=df_pivot[emp], name=emp,
            hovertemplate="任务: %{x}<br>员工: " + emp + "<br>完成值: %{y}<extra></extra>"
        ))
    fig.update_layout(barmode="stack", template="plotly_dark", xaxis_title="任务", yaxis_title="完成值")
    return fig

def chart_bubble(df0):
    df0 = df0[df0["明细"] != "分数总和"]
    emp_stats = df0.groupby("员工").agg(任务数=("明细","nunique"), 总值=("值","sum")).reset_index()
    emp_stats["覆盖率"] = emp_stats["任务数"] / df0["明细"].nunique()
    sizes = emp_stats["总值"].astype(float).tolist()
    fig = go.Figure(data=[go.Scatter(
        x=emp_stats["任务数"], y=emp_stats["覆盖率"], mode="markers+text",
        text=emp_stats["员工"], textposition="top center",
        hovertemplate="员工: %{text}<br>任务数: %{x}<br>覆盖率: %{y:.2%}<br>总值: %{marker.size}",
        marker=dict(size=sizes, sizemode="area", sizeref=2.*max(sizes)/(40.**2), sizemin=8,
                    color=emp_stats["总值"], colorscale="Viridis", showscale=True)
    )])
    fig.update_layout(template="plotly_dark", xaxis_title="任务数", yaxis_title="覆盖率")
    return fig

def chart_hot(df0):
    df0 = df0[df0["明细"] != "分数总和"]
    ts = df0.groupby("明细")["员工"].nunique()
    return {"backgroundColor":"transparent","yAxis":{"type":"category","data":ts.index.tolist(),"axisLabel":{"color":"#fff"}},"xAxis":{"type":"value","axisLabel":{"color":"#fff"}},"series":[{"data":ts.tolist(),"type":"bar","itemStyle":{"color":"#ffb703"}}]}

def chart_heat(df0):
    df0 = df0[df0["明细"] != "分数总和"]
    tasks = df0["明细"].unique().tolist()
    emps = df0["员工"].unique().tolist()
    data=[]
    for i,t in enumerate(tasks):
        for j,e in enumerate(emps):
            v=int(df0[(df0["明细"]==t)&(df0["员工"]==e)]["值"].sum())
            data.append([j,i,v])
    return {"backgroundColor":"transparent","tooltip":{"position":"top"},
            "xAxis":{"type":"category","data":emps,"axisLabel":{"color":"#fff"}},
            "yAxis":{"type":"category","data":tasks,"axisLabel":{"color":"#fff"}},
            "visualMap":{"min":0,"max":1,"show":False,"inRange":{"color":["#ff4d4d","#4caf50"]}},
            "series":[{"type":"heatmap","data":data}]}

# ========== 主页面逻辑 ==========
st.title("📊 技能覆盖分析大屏")

if view == "编辑数据":
    st.subheader("✏️ 编辑数据")
    edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)
    if time_choice:
        sheet_frames[time_choice[0]] = edited_df
    if st.sidebar.button("💾 保存到 GitHub"):
        save_excel_to_github(sheet_frames)
    st.dataframe(edited_df)

elif view == "大屏轮播":
    st_autorefresh(interval=10000, key="aut")
    secs=[("完成排名",chart_total(df)),("任务对比",chart_stack(df)),
          ("人员对比",chart_bubble(df)),("热门任务",chart_hot(df)),("热力图",chart_heat(df))]
    t,op=secs[int(time.time()/10)%len(secs)]
    st.subheader(t)
    if isinstance(op, go.Figure):
        st.plotly_chart(op, use_container_width=True)
    else:
        st_echarts(op,height="600px",theme="dark")

elif view == "单页模式":
    choice = st.sidebar.selectbox("单页查看", ["人员完成任务数量排名","任务对比","人员对比","热门任务","热力图"], index=0)
    mapping = {"人员完成任务数量排名": chart_total(df),"任务对比": chart_stack(df),
               "人员对比": chart_bubble(df),"热门任务": chart_hot(df),"热力图": chart_heat(df)}
    st.subheader(choice)
    op = mapping[choice]
    if isinstance(op, go.Figure):
        st.plotly_chart(op, use_container_width=True)
    else:
        st_echarts(op, height="600px", theme="dark")

elif view == "显示所有视图":
    charts = [("完成排名", chart_total(df)),("任务对比", chart_stack(df)),
              ("人员对比", chart_bubble(df)),("热门任务", chart_hot(df)),("热图", chart_heat(df))]
    for label, f in charts:
        st.subheader(label)
        if isinstance(f, go.Figure):
            st.plotly_chart(f, use_container_width=True)
        else:
            st_echarts(f, height="520px", theme="dark")

elif view == "能力分析":
    st.subheader("📊 能力分析")
    employees = df["员工"].unique().tolist()
    selected_emps = st.sidebar.multiselect("选择员工", employees, default=employees)
    tasks = df["明细"].unique().tolist()
    fig1, fig2, fig3 = go.Figure(), go.Figure(), go.Figure()
    for sheet in time_choice:
        df_sheet = get_merged_df([sheet], selected_groups)
        df_sheet = df_sheet[df_sheet["明细"] != "分数总和"]
        df_pivot = df_sheet.pivot(index="明细", columns="员工", values="值").fillna(0)
        for emp in selected_emps:
            fig1.add_trace(go.Scatter(x=tasks,y=df_pivot[emp].reindex(tasks, fill_value=0),
                                      mode="lines+markers",name=f"{sheet}-{emp}"))
        fig2.add_trace(go.Scatter(x=tasks,y=df_pivot.sum(axis=1).reindex(tasks, fill_value=0),
                                  mode="lines+markers",name=sheet))
        fig3.add_trace(go.Scatter(x=df_pivot.columns,y=df_pivot.sum(axis=0),
                                  mode="lines+markers",name=sheet))
    fig1.update_layout(title="员工任务完成情况", template="plotly_dark")
    fig2.update_layout(title="任务整体完成度趋势", template="plotly_dark")
    fig3.update_layout(title="员工整体完成度对比", template="plotly_dark")
    st.plotly_chart(fig1, use_container_width=True)
    st.plotly_chart(fig2, use_container_width=True)
    st.plotly_chart(fig3, use_container_width=True)
