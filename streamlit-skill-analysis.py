# app.py
import os, time, io, re
import pandas as pd
import streamlit as st
from typing import List, Tuple
from streamlit_autorefresh import st_autorefresh
from streamlit_echarts import st_echarts
import plotly.graph_objects as go

st.set_page_config(page_title="技能覆盖分析大屏", layout="wide")

# ----------- 页面样式 ------------
PAGE_CSS = """
<style>
:root {--primary:#4cc9f0; --bg:#0d1b2a; --bg2:#1b263b; --text:#fff; --muted:#ccc;}
[data-testid="stAppViewContainer"]{background:var(--bg); color:var(--text);}
[data-testid="stSidebar"]{background:var(--bg2); color:var(--text);}
div.stButton>button{background:var(--primary); color:#000; border-radius:10px; height:40px; font-weight:700; margin:5px 0; width:100%;}
div.stButton>button:hover{background:#4895ef;color:#fff;}
.metric-card{background:var(--bg2); padding:20px; border-radius:16px; text-align:center; box-shadow:0 0 15px rgba(0,0,0,0.4);}
.metric-value{font-size:36px; font-weight:800; color:var(--primary);}
.metric-label{font-size:14px; color:var(--muted);}
hr{border:none;border-top:1px solid rgba(255,255,255,.12);margin:16px 0;}
</style>
"""
st.markdown(PAGE_CSS, unsafe_allow_html=True)

# ---------- 左侧栏 ----------
st.sidebar.header("📂 数据控制区")
upload = st.sidebar.file_uploader("上传 Excel（Sheet 名称＝月份或季度）", type=["xlsx","xls"])

@st.cache_data
def load_sheets(file) -> Tuple[List[str], dict]:
    xpd = pd.ExcelFile(file)
    return xpd.sheet_names, {s: pd.read_excel(xpd, sheet_name=s) for s in xpd.sheet_names}

sheets, sheet_frames = ([] , {})
if upload:
    try:
        sheets, sheet_frames = load_sheets(upload)
    except Exception as e:
        st.sidebar.error(f"读取失败：{e}")

if sheets:
    time_choice = st.sidebar.multiselect("选择时间点（月或季）", sheets, default=sheets[:3] if len(sheets)>=3 else sheets)
else:
    st.sidebar.info("未上传, 使用示例")
    sheet_frames = {"示例": pd.DataFrame({
        "编号":[1,2,3],"明细":["任务A","任务B","任务C"],
        "数量总和":[3,2,5],
        "张三":[1,0,1],"李四":[0,1,1],"王五":[1,1,0]
    })}
    time_choice = ["示例"]

# 单页模式下的选项
sections_names = ["人员完成任务数量排名", "任务覆盖率分布", "任务掌握情况（热门任务）", "任务-人员热力图", "人员对比（雷达图)"]

view = st.sidebar.radio("切换视图", ["编辑数据", "大屏轮播", "单页模式", "能力对比", "显示所有视图", "Plotly 三图"])
save_path = st.sidebar.text_input("💾 保存路径", os.path.join(os.getcwd(),"结果.xlsx"))
save_click = st.sidebar.button("保存数据")
download_tpl = st.sidebar.button("下载模板")

if download_tpl:
    tpl = pd.DataFrame({"编号":[1,2,3],"明细":["任务A","任务B","任务C"],"数量总和":[1,2,3],"张三":[1,0,0],"李四":[0,1,0],"王五":[0,0,1]})
    buf = io.BytesIO()
    tpl.to_excel(buf, index=False, engine="xlsxwriter")
    st.sidebar.download_button("📥 点击下载模板", data=buf.getvalue(), file_name="模板.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ----------- 数据处理 ------------
def get_merged_df(keys: List[str]) -> pd.DataFrame:
    dfs = []
    for k in keys:
        df0 = sheet_frames.get(k)
        if df0 is not None:
            dfs.append(df0)
    if not dfs:
        return pd.DataFrame()
    merged = pd.concat(dfs, axis=0, ignore_index=True)
    agg = merged.groupby(merged["明细"]).max().reset_index()
    return agg

df_raw = get_merged_df(time_choice)

def clean(df0):
    if df0 is None or df0.empty:
        return pd.DataFrame(columns=["编号","明细"])
    df1 = df0[df0["明细"]!="总分"] if "明细" in df0.columns else df0.copy()
    df1 = df1.drop(columns=["总分"], errors="ignore")
    return df1

df = clean(df_raw)

task_col="明细"
person_cols = [c for c in df.columns if c not in ["编号","数量总和",task_col]]
total_tasks=len(df); total_people=len(person_cols)
ps = df[person_cols].sum() if total_people>0 else pd.Series(dtype=float)
top_person = ps.idxmax() if not ps.empty else ""
avg_score = round(ps.mean(),1) if not ps.empty else 0

# ----------- echarts 图表函数 ------------
def chart_total():
    return {"backgroundColor":"transparent","tooltip":{"trigger":"axis"},
        "xAxis":{"type":"category","data":ps.sort_values(ascending=False).index.tolist(),"axisLabel":{"color":"#fff"}},
        "yAxis":{"type":"value","axisLabel":{"color":"#fff"}},
        "series":[{"data":ps.sort_values(ascending=False).tolist(),"type":"bar","itemStyle":{"color":"#4cc9f0"}}]}
def chart_cover():
    cov = df[person_cols].sum(axis=1).value_counts()
    dat=[{"name":f"{int(k)}人掌握","value":int(v)} for k,v in cov.items()]
    return {"backgroundColor":"transparent","tooltip":{"trigger":"item"},
            "series":[{"type":"pie","radius":"70%","data":dat,"label":{"color":"#fff"}}]}
def chart_hot():
    ts=df[person_cols].sum(axis=1) if total_people>0 else pd.Series(dtype=float)
    return {"backgroundColor":"transparent","yAxis":{"type":"category","data":df[task_col].tolist(),"axisLabel":{"color":"#fff"}},
            "xAxis":{"type":"value","axisLabel":{"color":"#fff"}},
            "series":[{"data":ts.tolist(),"type":"bar","itemStyle":{"color":"#ffb703"}}]}
def chart_radar_sel(sel):
    data = [{"name":p,"value":df[p].fillna(0).astype(int).tolist()} for p in sel if p in df.columns]
    return {"backgroundColor":"transparent","legend":{"data":sel,"textStyle":{"color":"#fff"}},
            "radar":{"indicator":[{"name":t,"max":1} for t in df[task_col].tolist()]},
            "series":[{"type":"radar","data":data}]}
def chart_heat():
    data = []
    for i, rt in enumerate(df.index, 1):
        for j, p in enumerate(person_cols):
            val = int(df.at[rt, p]) if p in df.columns else 0
            data.append([j, i, val])
    return {
        "backgroundColor": "transparent",
        "tooltip": {"position": "top"},
        "xAxis": {"type": "category", "data": person_cols, "axisLabel": {"color": "#fff"}},
        "yAxis": {"type": "category", "data": df[task_col].tolist(), "axisLabel": {"color": "#fff"}},
        "visualMap": {
            "min": 0,
            "max": 1,
            "calculable": False,
            "orient": "horizontal",
            "show": False,  # 不显示图例
            "inRange": {"color": ["#ff4d4d", "#4caf50"]}  # 0=红色，不会；1=绿色，会
        },
        "series": [{
            "type": "heatmap",
            "data": data,
            "label": {"show": False}
        }]
    }

# ----------- Plotly 三图函数 ------------
def make_plotly_figs(data: dict, sheet_names_sorted: List[str], selected_employees: List[str]):
    first_sheet = sheet_names_sorted[0]
    df_first = data[first_sheet]
    employees = [c for c in df_first.columns if c not in ["明细","数量总和","编号"]]
    tasks = df_first[df_first["明细"] != "分数总和"]["明细"].tolist()

    # 图1（可选员工）
    fig1 = go.Figure()
    for sheet in sheet_names_sorted:
        df_sheet = data[sheet]
        df_tasks = df_sheet[df_sheet["明细"]!="分数总和"].set_index("明细")
        for emp in employees:
            if emp not in selected_employees:
                continue
            y = [df_tasks.at[t,emp] if t in df_tasks.index and emp in df_tasks.columns else 0 for t in tasks]
            fig1.add_trace(go.Scatter(x=tasks,y=y,mode="lines+markers",name=f"{sheet}-{emp}"))
    fig1.update_layout(title="图1：员工每月得分对比 (明细项目)")

    # 图2（数量总和）
    fig2 = go.Figure()
    for sheet in sheet_names_sorted:
        df_sheet = data[sheet]
        df_tasks = df_sheet[df_sheet["明细"]!="分数总和"].set_index("明细")
        y=[df_tasks.at[t,"数量总和"] if t in df_tasks.index and "数量总和" in df_tasks.columns else 0 for t in tasks]
        fig2.add_trace(go.Scatter(x=tasks,y=y,mode="lines+markers",name=sheet))
    fig2.update_layout(title="图2：各月明细项目完成数量总和")

    # 图3（员工分数总和）
    fig3 = go.Figure()
    for sheet in sheet_names_sorted:
        df_sheet = data[sheet]
        df_tasks = df_sheet[df_sheet["明细"]!="分数总和"]
        totals=df_tasks[employees].sum()
        fig3.add_trace(go.Scatter(x=employees,y=[totals.get(emp,0) for emp in employees],
                                  mode="lines+markers",name=sheet))
    fig3.update_layout(title="图3：各月员工分数总和")

    return fig1, fig2, fig3
# ----------- 指标卡片 ------------
def show_cards():
    c1,c2,c3,c4=st.columns(4)
    c1.markdown(f"<div class='metric-card'><div class='metric-value'>{total_tasks}</div><div class='metric-label'>任务数</div></div>",unsafe_allow_html=True)
    c2.markdown(f"<div class='metric-card'><div class='metric-value'>{total_people}</div><div class='metric-label'>人数</div></div>",unsafe_allow_html=True)
    c3.markdown(f"<div class='metric-card'><div class='metric-value'>{top_person}</div><div class='metric-label'>覆盖率最高</div></div>",unsafe_allow_html=True)
    c4.markdown(f"<div class='metric-card'><div class='metric-value'>{avg_score}</div><div class='metric-label'>平均数</div></div>",unsafe_allow_html=True)
    st.markdown("<hr/>",unsafe_allow_html=True)

# ----------- 主页面 ------------
st.title("📊 技能覆盖分析大屏")

if view=="编辑数据":
    show_cards()
    st.subheader("当前数据表（行合并后）")
    st.dataframe(df, use_container_width=True)
    edt = st.data_editor(df, num_rows="dynamic", use_container_width=True)
    if save_click:
        edt.to_excel(save_path, index=False, engine="xlsxwriter")
        st.sidebar.success("已保存到 "+save_path)

elif view=="大屏轮播":
    st.info("⏱ 轮播中...")
    show_cards()
    st_autorefresh(interval=10000, key="aut")
    secs=[("完成排名",chart_total()),("覆盖率",chart_cover()),("热门任务",chart_hot()),("热力图",chart_heat())]
    t,op=secs[int(time.time()/10)%len(secs)]
    st.subheader(t); st_echarts(op,height="600px",theme="dark")

elif view=="单页模式":
    show_cards()
    choice = st.sidebar.selectbox("单页查看", sections_names, index=0)
    st.subheader(choice)
    if choice=="人员对比（雷达图)":
        sel = st.sidebar.multiselect("选择 2-5 人", person_cols, default=person_cols[:2])
        st_echarts(chart_radar_sel(sel),height="600px",theme="dark")
    else:
        mapping={"人员完成任务数量排名": chart_total(),
                 "任务覆盖率分布": chart_cover(),
                 "任务掌握情况（热门任务）": chart_hot(),
                 "任务-人员热力图": chart_heat()}
        st_echarts(mapping.get(choice,chart_total()),height="600px",theme="dark")

elif view=="能力对比":
    show_cards()
    st.subheader("📈 自由人员能力对比")
    sel = st.sidebar.multiselect("选择 2 人进行对比", person_cols, default=person_cols[:2])
    if len(sel)==2:
        st_echarts(chart_radar_sel(sel),height="600px",theme="dark")
    else:
        st.warning("请选择其中 2 人")

elif view=="显示所有视图":
    show_cards()
    for label,f in [("排名",chart_total),("覆盖",chart_cover),("热门任务",chart_hot),
                    ("雷达",lambda:chart_radar_sel(person_cols[:3])),("热图",chart_heat)]:
        st.subheader(label); st_echarts(f(),height="520px",theme="dark")

elif view=="Plotly 三图":
    st.subheader("📈 Plot"
                 "ly 交互图表")
    # 动态获取所有员工
    if time_choice:
        first_sheet = time_choice[0]
        employees = [c for c in sheet_frames[first_sheet].columns if c not in ["明细","数量总和","编号"]]
    else:
        employees = []

    selected_emps = st.sidebar.multiselect("选择员工（图1显示）", employees, default=employees)

    fig1, fig2, fig3 = make_plotly_figs(sheet_frames, time_choice, selected_emps)
    st.plotly_chart(fig1, use_container_width=True)
    st.plotly_chart(fig2, use_container_width=True)
    st.plotly_chart(fig3, use_container_width=True)