from __future__ import annotations

from pathlib import Path

import pandas as pd

from config import *

def latest_source_file() -> Path:
    selected = SELECTED_INPUT_FILE if "SELECTED_INPUT_FILE" in globals() else find_source_file()
    print(f"本次分析的表格：{selected.name}")
    print(f"表格完整路径：{selected}")
    return selected

def text_series(series: pd.Series) -> pd.Series:
    return series.astype("string").str.strip()

def parse_pvi(series: pd.Series) -> pd.Series:
    cleaned = text_series(series).str.replace(r"[,，￥¥\s]", "", regex=True)
    return pd.to_numeric(cleaned, errors="coerce").fillna(0.0)

def empty_frame(columns: list[str]) -> pd.DataFrame:
    return pd.DataFrame(columns=columns)

def normalize_region_name(value: object) -> str:
    text = "" if pd.isna(value) else str(value).strip()
    if not text:
        return "未填写"

    text = text.split("-", 1)[0].strip()

    aliases = {
        "Bees 旭日区": "旭日&旭日北辰区",
        "Bees 旭日北辰区": "旭日&旭日北辰区",
    }

    return aliases.get(text, text)

def clean_data(path: Path, sheet: str, team_keyword: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet, dtype=object)
    df.columns = df.columns.astype(str).str.strip()

    missing = [column for column in REQUIRED_COLUMNS if column not in df.columns]
    if missing:
        raise ValueError(f"工作表缺少必要字段：{missing}")

    df = df.copy()
    df["_pvi"] = parse_pvi(df["PVI"])

    for source, target in [
        ("交单日期", "_submit_date"),
        ("承保日期", "_insured_date"),
        ("回执日期", "_receipt_date"),
        ("回访完成日期", "_visit_date"),
    ]:
        df[target] = pd.to_datetime(df[source], errors="coerce").dt.normalize()

    df["_team"] = text_series(df["出单团队"])
    df["_status"] = text_series(df["保单状态"])
    df["_progress"] = text_series(df["承保进度"])
    df["_agency"] = text_series(df["出单机构"])
    df["_region"] = text_series(df["出单人所属区域"])
    df["_organization"] = text_series(df["出单人所属组织"])
    df["_agent_id"] = text_series(df["出单代理人工号"])
    df["_agent_name"] = text_series(df["出单代理人姓名"])

    valid_status = (
        df["_status"].notna()
        & df["_status"].ne("")
        & ~df["_status"].isin(EXCLUDED_POLICY_STATUSES)
        & ~df["_status"].str.contains("撤销|犹豫期退保", na=False)
    )
    df = df.loc[valid_status].copy()

    df["_is_team"] = df["_team"].eq(team_keyword)
    df = df.loc[df["_is_team"]].copy()

    # 注意：这里不要因为“承保进度=撤单”就剔除。
    # MDRT只按保单状态排除：撤销/犹豫期退保/待承保/空白。
    df["_is_cancelled"] = (
        df["_status"].str.contains("犹豫期退保|撤销", na=False)
    )

    df["_is_valid"] = ~df["_is_cancelled"]

    return df
