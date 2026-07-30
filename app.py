from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import streamlit as st


PROJECT_ROOT = Path(__file__).resolve().parent
SRC_DIR = PROJECT_ROOT / "src"
RAG_DIR = PROJECT_ROOT / "rag"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))
if str(RAG_DIR) not in sys.path:
    sys.path.insert(0, str(RAG_DIR))

from config import (  # noqa: E402
    INACTIVE_DAYS,
    MONTHLY_TARGET,
    NEAR_THRESHOLD_MIN,
    OUTPUT_FOLDER,
    PVI_THRESHOLD,
    REPORT_DATE,
    ROSTER_LOOKBACK_DAYS,
    SHEET_NAME,
    TEAM_KEYWORD,
)
from dashboard import draw_dashboard  # noqa: E402
from data_cleaning import clean_data, latest_source_file  # noqa: E402
from data_query import active_agents, format_active_agents  # noqa: E402
from insight import generate_insight, save_insight  # noqa: E402
from metrics import calculate_reports, determine_report_date, money  # noqa: E402
from qa import answer as rag_answer  # noqa: E402
from report_export import export_excel  # noqa: E402
from summary import generate_summary, save_summary  # noqa: E402


st.set_page_config(
    page_title="PVI 经营智能助手",
    page_icon="📊",
    layout="wide",
)


def save_uploaded_file(uploaded_file) -> Path:
    raw_dir = PROJECT_ROOT / "data" / "raw"
    raw_dir.mkdir(parents=True, exist_ok=True)
    output_path = raw_dir / uploaded_file.name
    output_path.write_bytes(uploaded_file.getbuffer())
    return output_path


def run_analysis() -> dict[str, object]:
    input_path = latest_source_file().resolve()
    output_dir = OUTPUT_FOLDER.expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    df = clean_data(input_path, SHEET_NAME, TEAM_KEYWORD)
    report_date = determine_report_date(df, REPORT_DATE)
    reports = calculate_reports(
        df=df,
        report_date=report_date,
        threshold=PVI_THRESHOLD,
        near_min=NEAR_THRESHOLD_MIN,
        inactive_days=INACTIVE_DAYS,
        roster_days=ROSTER_LOOKBACK_DAYS,
        monthly_target=MONTHLY_TARGET,
    )

    date_code = report_date.strftime("%Y%m%d")
    dashboard_path = output_dir / f"PVI经营看板_{date_code}.png"
    excel_path = output_dir / f"PVI经营报告_{date_code}.xlsx"
    summary_path = output_dir / f"经营摘要_{date_code}.txt"
    insight_path = output_dir / f"智能解读_{date_code}.txt"

    draw_dashboard(reports, dashboard_path, report_date, PVI_THRESHOLD)
    export_excel(reports, excel_path, dashboard_path, report_date)

    daily_report = reports["daily"]
    assert isinstance(daily_report, pd.DataFrame)
    today_pvi = float(
        daily_report.loc[
            daily_report["日期"].eq(report_date),
            "每日净PVI",
        ].sum()
    )

    summary_text = generate_summary(reports, report_date, today_pvi, MONTHLY_TARGET)
    insight_text = generate_insight(reports, report_date, today_pvi, MONTHLY_TARGET)
    save_summary(summary_text, summary_path)
    save_insight(insight_text, insight_path)

    return {
        "input_path": input_path,
        "report_date": report_date,
        "df": df,
        "reports": reports,
        "today_pvi": today_pvi,
        "summary_text": summary_text,
        "insight_text": insight_text,
        "dashboard_path": dashboard_path,
        "excel_path": excel_path,
        "summary_path": summary_path,
        "insight_path": insight_path,
    }


def file_download(label: str, path: Path, mime: str) -> None:
    if path.exists():
        st.download_button(
            label=label,
            data=path.read_bytes(),
            file_name=path.name,
            mime=mime,
            use_container_width=True,
        )


st.title("保险经销业务 PVI 经营智能助手")
st.caption("上传或放入最新保单汇总表后，可在网页中生成经营报告、智能解读并进行 RAG 问答。")
st.info("线上 Demo 默认使用 data/sample/ 中的脱敏样例数据；请勿上传真实业务数据。")

with st.sidebar:
    st.header("数据入口")
    uploaded = st.file_uploader(
        "上传新的保单汇总表",
        type=["xlsx"],
        help="文件名建议以“保单汇总列表”开头。上传后会保存到 data/raw/。",
    )
    if uploaded is not None:
        saved_path = save_uploaded_file(uploaded)
        st.success(f"已保存：{saved_path.name}")

    st.divider()
    st.header("操作")
    run_button = st.button("生成/刷新报告", type="primary", use_container_width=True)
    st.caption("也可以直接把新 Excel 放进 data/raw/，再点击此按钮。")

if run_button or "analysis_result" not in st.session_state:
    try:
        with st.spinner("正在读取最新数据并生成报告..."):
            st.session_state.analysis_result = run_analysis()
    except Exception as exc:
        st.error(f"生成失败：{exc}")
        st.stop()

result = st.session_state.analysis_result
reports = result["reports"]
metrics = reports["metrics"]
assert isinstance(metrics, dict)
report_date = result["report_date"]
today_pvi = float(result["today_pvi"])

st.subheader("经营概览")
col1, col2, col3, col4 = st.columns(4)
col1.metric("今日新增 PVI", f"{money(today_pvi)} 元")
col2.metric("本月累计净 PVI", f"{money(float(metrics['net_pvi']))} 元")
col3.metric("目标达成率", f"{float(metrics['target_completion']):.2%}")
col4.metric("目标缺口", f"{money(float(metrics['target_gap']))} 元")

data_type = "脱敏样例数据" if "data/sample" in str(result["input_path"]) else "用户上传/本地 raw 数据"
st.caption(
    f"当前分析文件：{result['input_path'].name}｜数据类型：{data_type}｜报告日期：{report_date:%Y-%m-%d}"
)

tab_report, tab_insight, tab_query, tab_files = st.tabs(
    ["经营结果", "智能解读", "问答助手", "输出文件"]
)

with tab_report:
    left, right = st.columns([1.1, 1])
    with left:
        st.markdown("#### 经营摘要")
        st.info(result["summary_text"])
        st.markdown("#### 今日出单名单")
        today_active = active_agents(result["df"], report_date, report_date)
        if today_active.empty:
            st.write("暂无今日出单明细。")
        else:
            st.dataframe(today_active, use_container_width=True, hide_index=True)
    with right:
        st.markdown("#### 关键指标")
        st.write(f"活跃代理人：{int(float(metrics['active_agents']))} 人")
        st.write(f"达标代理人：{int(float(metrics['qualified_agents']))} 人")
        st.write(f"Top5 贡献占比：{float(metrics['top5_share']):.2%}")
        st.write(f"正常业务预测：{money(float(metrics['normal_business_forecast']))} 元")
        st.write(f"剩余工作日所需日均：{money(float(metrics['required_daily']))} 元")

    st.markdown("#### 代理人 PVI Top")
    agent_ranking = reports["agent_ranking"]
    assert isinstance(agent_ranking, pd.DataFrame)
    st.dataframe(agent_ranking.head(10), use_container_width=True, hide_index=True)

with tab_insight:
    st.markdown("#### 智能经营解读")
    st.text(result["insight_text"])

with tab_query:
    st.markdown("#### RAG + Pandas 问答")
    question = st.text_input(
        "输入问题",
        value="今天都谁出单了，本月PVI是多少，PVI是什么",
    )
    if st.button("提问", use_container_width=True):
        with st.spinner("正在生成回答..."):
            st.text(rag_answer(question))

    st.markdown("常用问题：")
    examples = [
        "今天pvi多少",
        "今天都谁出单了",
        "本月哪个业务员最厉害",
        "PVI达成率咋算",
        "信泰方案规则是什么",
    ]
    for example in examples:
        if st.button(example, key=f"example-{example}"):
            st.text(rag_answer(example))

with tab_files:
    st.markdown("#### 下载输出文件")
    file_download(
        "下载 Excel 经营报告",
        result["excel_path"],
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    file_download("下载 PNG 经营看板", result["dashboard_path"], "image/png")
    file_download("下载经营摘要 TXT", result["summary_path"], "text/plain")
    file_download("下载智能解读 TXT", result["insight_path"], "text/plain")
