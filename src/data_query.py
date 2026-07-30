from __future__ import annotations

from functools import lru_cache
from pathlib import Path

import pandas as pd

from config import SHEET_NAME, TEAM_KEYWORD, find_source_file
from data_cleaning import clean_data
from metrics import determine_report_date, money


@lru_cache(maxsize=1)
def load_business_data() -> tuple[pd.DataFrame, pd.Timestamp, Path]:
    input_path = find_source_file().resolve()
    df = clean_data(input_path, SHEET_NAME, TEAM_KEYWORD)
    report_date = determine_report_date(df, None)
    return df, report_date, input_path


def agent_label(row: pd.Series) -> str:
    name = str(row.get("_agent_name", "") or "").strip()
    agent_id = str(row.get("_agent_id", "") or "").strip()
    if name and agent_id:
        return f"{name}（{agent_id}）"
    return name or agent_id or "未填写"


def top_agents(df: pd.DataFrame, start_date: pd.Timestamp, end_date: pd.Timestamp, top_n: int) -> pd.DataFrame:
    period = df.loc[
        df["_insured_date"].between(start_date, end_date, inclusive="both")
        & df["_is_valid"]
    ].copy()
    if period.empty:
        return pd.DataFrame(columns=["代理人", "PVI"])
    period["_agent_label"] = period.apply(agent_label, axis=1)
    ranking = (
        period.groupby("_agent_label", dropna=False)["_pvi"]
        .sum()
        .sort_values(ascending=False)
        .head(top_n)
        .reset_index()
        .rename(columns={"_agent_label": "代理人", "_pvi": "PVI"})
    )
    return ranking


def top_dimension(
    df: pd.DataFrame,
    column: str,
    label: str,
    start_date: pd.Timestamp,
    end_date: pd.Timestamp,
    top_n: int,
) -> pd.DataFrame:
    period = df.loc[
        df["_insured_date"].between(start_date, end_date, inclusive="both")
        & df["_is_valid"]
    ].copy()
    if period.empty:
        return pd.DataFrame(columns=[label, "PVI"])
    ranking = (
        period.groupby(column, dropna=False)["_pvi"]
        .sum()
        .sort_values(ascending=False)
        .head(top_n)
        .reset_index()
        .rename(columns={column: label, "_pvi": "PVI"})
    )
    return ranking


def format_ranking(title: str, ranking: pd.DataFrame) -> str:
    if ranking.empty:
        return f"{title}：暂无数据。"
    lines = [f"{title}："]
    label_column = ranking.columns[0]
    for index, (_, row) in enumerate(ranking.iterrows(), start=1):
        name = row[label_column]
        pvi = row["PVI"]
        lines.append(f"{index}. {name}：{money(float(pvi))} 元")
    return "\n".join(lines)


def active_agents(df: pd.DataFrame, start_date: pd.Timestamp, end_date: pd.Timestamp) -> pd.DataFrame:
    period = df.loc[
        df["_insured_date"].between(start_date, end_date, inclusive="both")
        & df["_is_valid"]
        & df["_pvi"].gt(0)
    ].copy()
    if period.empty:
        return pd.DataFrame(columns=["代理人", "PVI", "出单件数"])
    period["_agent_label"] = period.apply(agent_label, axis=1)
    summary = (
        period.groupby("_agent_label", dropna=False)
        .agg(PVI=("_pvi", "sum"), 出单件数=("保单号", "nunique"))
        .sort_values(["PVI", "出单件数"], ascending=[False, False])
        .reset_index()
        .rename(columns={"_agent_label": "代理人"})
    )
    return summary


def format_active_agents(period_label: str, summary: pd.DataFrame) -> str:
    if summary.empty:
        return f"{period_label}暂无正 PVI 出单代理人。"
    total_pvi = float(summary["PVI"].sum())
    total_policies = int(summary["出单件数"].sum())
    lines = [
        f"{period_label}出单代理人共 {len(summary)} 人，合计 PVI {money(total_pvi)} 元，合计出单 {total_policies} 件："
    ]
    for index, (_, row) in enumerate(summary.iterrows(), start=1):
        lines.append(
            f"{index}. {row['代理人']}：PVI {money(float(row['PVI']))} 元，出单 {int(row['出单件数'])} 件"
        )
    return "\n".join(lines)


def query_data(question: str) -> str | None:
    if not any(
        keyword in question
        for keyword in ["今日", "今天", "本月", "今月", "最高", "Top", "top", "产品", "保险公司", "代理人", "业务员", "出单", "谁", "几个人"]
    ):
        return None

    df, report_date, input_path = load_business_data()
    month_start = report_date.replace(day=1)

    asks_today = "今日" in question or "今天" in question
    asks_month = "本月" in question or "今月" in question
    start_date = report_date if asks_today else month_start
    end_date = report_date
    period_label = "今日" if asks_today else "本月"

    if "出单" in question and any(word in question for word in ["谁", "哪些", "名单", "人员", "人数", "代理人", "几个人"]):
        summary = active_agents(df, start_date, end_date)
        if any(word in question for word in ["人数", "几个人"]) and not any(word in question for word in ["谁", "哪些", "名单", "人员"]):
            return f"{period_label}出单代理人共 {len(summary)} 人。"
        return format_active_agents(period_label, summary)

    if "保险公司" in question:
        ranking = top_dimension(df, "保险公司", "保险公司", start_date, end_date, 5)
        return format_ranking(f"{period_label}保险公司 PVI Top 5", ranking)

    if "产品" in question or "险种" in question:
        ranking = top_dimension(df, "险种名称", "产品/险种", start_date, end_date, 5)
        return format_ranking(f"{period_label}产品 PVI Top 5", ranking)

    if "最高" in question or "Top" in question or "top" in question or "代理人" in question or "业务员" in question or "最厉害" in question:
        top_n = 5 if any(word in question for word in ["Top5", "top5", "前五", "5"]) else 1
        ranking = top_agents(df, start_date, end_date, top_n)
        return format_ranking(f"{period_label}代理人 PVI Top {top_n}", ranking)

    period = df.loc[
        df["_insured_date"].between(start_date, end_date, inclusive="both")
        & df["_is_valid"]
    ]
    total_pvi = float(period["_pvi"].sum())
    return (
        f"{period_label}PVI 为 {money(total_pvi)} 元。\n"
        f"报告日期：{report_date:%Y-%m-%d}\n"
        f"数据来源：{input_path.name}"
    )
