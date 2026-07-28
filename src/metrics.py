from __future__ import annotations

import math

import pandas as pd

from config import *
from data_cleaning import empty_frame, normalize_region_name, text_series

def money(value: float) -> str:
    return f"{value:,.2f}".rstrip("0").rstrip(".")

def short_money(value: float, _position: int | None = None) -> str:
    absolute = abs(value)
    if absolute >= 1_000_000:
        return f"{value / 1_000_000:.1f}M"
    if absolute >= 1_000:
        return f"{value / 1_000:.0f}K"
    return f"{value:.0f}"

def determine_report_date(
    df: pd.DataFrame,
    requested: str | None
) -> pd.Timestamp:

    if requested:
        parsed = pd.to_datetime(requested, errors="raise")
        parsed = pd.Timestamp(parsed).normalize()

        if parsed.year != ANALYSIS_YEAR:
            raise ValueError(
                f"报告日期必须属于{ANALYSIS_YEAR}年，"
                f"当前填写的是：{parsed:%Y-%m-%d}"
            )

        return parsed

    latest = df.loc[
        df["_is_team"]
        & df["_insured_date"].dt.year.eq(ANALYSIS_YEAR),
        "_insured_date"
    ].max()

    if pd.isna(latest):
        raise ValueError(
            f"{ANALYSIS_YEAR}年大湾区数据中没有可用的承保日期"
        )

    return pd.Timestamp(latest).normalize()

def calculate_reports(
    df: pd.DataFrame,
    report_date: pd.Timestamp,
    threshold: float,
    near_min: float,
    inactive_days: int,
    roster_days: int,
    monthly_target: float,
) -> dict[str, pd.DataFrame | dict[str, float] | pd.Series]:

    report_date = pd.Timestamp(report_date).normalize()

    # 防止误算其他年份
    if report_date.year != ANALYSIS_YEAR:
        raise ValueError(
            f"报告日期必须属于{ANALYSIS_YEAR}年，"
            f"当前日期是：{report_date:%Y-%m-%d}"
        )

    # 已承保业务按承保年份判断
    insured_in_year = (
        df["_insured_date"].dt.year.eq(ANALYSIS_YEAR)
    )

    # 未承保业务按交单年份判断
    pending_submitted_in_year = (
        df["_insured_date"].isna()
        & df["_submit_date"].dt.year.eq(ANALYSIS_YEAR)
    )

    # 本函数再次过滤，确保所有结果只来自2026年
    df_year = df.loc[
        insured_in_year | pending_submitted_in_year
    ].copy()

    month_start = report_date.replace(day=1)
    month_end = month_start + pd.offsets.MonthEnd(0)

    # 本月已承保数据
    in_period = df_year["_insured_date"].between(
        month_start,
        report_date,
        inclusive="both",
    )

    mtd = df_year.loc[
        df_year["_is_valid"] & in_period
    ].copy()

    # 2026年大湾区全部相关记录
    team_rows = df_year.loc[
        df_year["_is_team"]
    ].copy()

    # ============================================================
    # 每日PVI及本月累计
    # ============================================================
    daily = (
        mtd.groupby("_insured_date", dropna=True)["_pvi"]
        .sum()
        .reindex(
            pd.date_range(
                month_start,
                report_date,
                freq="D",
            ),
            fill_value=0.0,
        )
    )

    daily.index.name = "日期"

    daily_df = pd.DataFrame(
        {
            "日期": daily.index,
            "每日净PVI": daily.values,
        }
    )

    daily_df["本月累计PVI"] = (
        daily_df["每日净PVI"].cumsum()
    )

    # ============================================================
    # 代理人排行榜
    # ============================================================
    agent_ranking = (
        mtd.groupby(
            ["_agent_id", "_agent_name"],
            dropna=False,
        )
        .agg(
            本月PVI=("_pvi", "sum"),
            保单记录数=("保单号", "count"),
        )
        .reset_index()
        .rename(
            columns={
                "_agent_id": "出单代理人工号",
                "_agent_name": "出单代理人姓名",
            }
        )
        .sort_values(
            "本月PVI",
            ascending=False,
        )
    )

    agent_ranking["排名"] = range(
        1,
        len(agent_ranking) + 1,
    )

    agent_ranking["是否达到门槛"] = (
        agent_ranking["本月PVI"]
        .ge(threshold)
        .map(
            {
                True: "是",
                False: "否",
            }
        )
    )

    agent_ranking = agent_ranking[
        [
            "排名",
            "出单代理人工号",
            "出单代理人姓名",
            "本月PVI",
            "保单记录数",
            "是否达到门槛",
        ]
    ]

    # ============================================================
    # 产品排行榜
    # ============================================================
    product_ranking = (
        mtd.groupby(
            "险种名称",
            dropna=False,
        )["_pvi"]
        .sum()
        .sort_values(ascending=False)
        .rename("本月PVI")
        .reset_index()
    )

    # ============================================================
    # 保险公司排行榜
    # ============================================================
    company_ranking = (
        mtd.groupby(
            "保险公司",
            dropna=False,
        )["_pvi"]
        .sum()
        .sort_values(ascending=False)
        .rename("本月PVI")
        .reset_index()
    )

    # ============================================================
    # 代理人参考名单
    # 只使用2026年承保记录
    # ============================================================
    roster_start = max(
        report_date - pd.Timedelta(days=roster_days),
        pd.Timestamp(f"{ANALYSIS_YEAR}-01-01"),
    )

    roster_rows = team_rows.loc[
        team_rows["_insured_date"].between(
            roster_start,
            report_date,
            inclusive="both",
        )
        & team_rows["_insured_date"]
        .dt.year.eq(ANALYSIS_YEAR)
        & team_rows["_agent_id"].notna()
    ]

    roster = (
        roster_rows
        .sort_values("_insured_date")
        .drop_duplicates(
            "_agent_id",
            keep="last",
        )[
            [
                "_agent_id",
                "_agent_name",
            ]
        ]
        .rename(
            columns={
                "_agent_id": "出单代理人工号",
                "_agent_name": "出单代理人姓名",
            }
        )
    )

    # ============================================================
    # 本月尚未出单代理
    # ============================================================
    issued_this_month = set(
        mtd.loc[
            mtd["_pvi"].gt(0),
            "_agent_id",
        ]
        .dropna()
        .astype(str)
    )

    not_issued = roster.loc[
        ~roster["出单代理人工号"]
        .astype(str)
        .isin(issued_this_month)
    ].copy()

    not_issued["说明"] = (
        f"{ANALYSIS_YEAR}年最近{roster_days}天出现过，"
        "但本月尚无正PVI记录"
    )

    # ============================================================
    # 接近6000元门槛
    # ============================================================
    near_threshold = agent_ranking.loc[
        agent_ranking["本月PVI"].ge(near_min)
        & agent_ranking["本月PVI"].lt(threshold)
    ].copy()

    near_threshold["距离6000"] = (
        threshold
        - near_threshold["本月PVI"]
    )

    # ============================================================
    # 连续多日没有新增PVI
    # ============================================================
    positive_valid = team_rows.loc[
        ~team_rows["_is_cancelled"]
        & team_rows["_pvi"].gt(0)
        & team_rows["_insured_date"].notna()
        & team_rows["_insured_date"]
        .dt.year.eq(ANALYSIS_YEAR)
        & team_rows["_insured_date"].le(report_date)
    ]

    last_issue = (
        positive_valid.groupby(
            ["_agent_id", "_agent_name"],
            dropna=False,
        )["_insured_date"]
        .max()
        .reset_index()
        .rename(
            columns={
                "_agent_id": "出单代理人工号",
                "_agent_name": "出单代理人姓名",
                "_insured_date": "最近承保日期",
            }
        )
    )

    inactive = roster.merge(
        last_issue,
        how="left",
        on=[
            "出单代理人工号",
            "出单代理人姓名",
        ],
    )

    inactive["连续未出单天数"] = (
        report_date
        - inactive["最近承保日期"]
    ).dt.days

    inactive = inactive.loc[
        inactive["最近承保日期"].isna()
        | inactive["连续未出单天数"].ge(inactive_days)
    ].sort_values(
        "连续未出单天数",
        ascending=False,
        na_position="first",
    )

    # ============================================================
    # 交单超过3天仍未承保
    # 只统计2026年交单记录
    # ============================================================
    pending_underwrite = team_rows.loc[
        team_rows["_submit_date"].notna()
        & team_rows["_submit_date"]
        .dt.year.eq(ANALYSIS_YEAR)
        & team_rows["_submit_date"].le(
            report_date - pd.Timedelta(days=3)
        )
        & team_rows["_insured_date"].isna()
        & ~team_rows["_is_cancelled"]
    ].copy()

    pending_underwrite["等待天数"] = (
        report_date
        - pending_underwrite["_submit_date"]
    ).dt.days

    pending_underwrite = pending_underwrite[
        [
            "投保单号",
            "保单号",
            "出单代理人工号",
            "出单代理人姓名",
            "交单日期",
            "承保进度",
            "保单状态",
            "PVI",
            "等待天数",
        ]
    ].sort_values(
        "等待天数",
        ascending=False,
    )

    # ============================================================
    # 犹豫期退保、撤销及负PVI批改
    # ============================================================
    event_date = (
        team_rows["_insured_date"]
        .fillna(team_rows["_submit_date"])
    )

    event_in_month = event_date.between(
        month_start,
        report_date,
        inclusive="both",
    )

    event_in_year = event_date.dt.year.eq(
        ANALYSIS_YEAR
    )

    exception_mask = (
        team_rows["_status"].str.contains(
            "犹豫期退保|撤销",
            na=False,
        )
        | team_rows["_progress"].str.contains(
            "撤单",
            na=False,
        )
        | team_rows["_pvi"].lt(0)
    )

    exceptions = team_rows.loc[
        event_in_year
        & event_in_month
        & exception_mask
    ].copy()

    def exception_reason(
        row: pd.Series,
    ) -> str:
        reasons: list[str] = []

        if "犹豫期退保" in str(
            row["保单状态"]
        ):
            reasons.append("犹豫期退保")

        if (
            "撤销" in str(row["保单状态"])
            or "撤单" in str(row["承保进度"])
        ):
            reasons.append("撤销/撤单")

        if row["_pvi"] < 0:
            reasons.append("负PVI批改")

        return "、".join(reasons)

    if exceptions.empty:
        exceptions["异常类型"] = pd.Series(dtype="string")
    else:
        exceptions["异常类型"] = exceptions.apply(
            exception_reason,
            axis=1,
        )

    exceptions = exceptions[
        [
            "异常类型",
            "承保日期",
            "投保单号",
            "保单号",
            "出单代理人工号",
            "出单代理人姓名",
            "PVI",
            "保单状态",
            "承保进度",
        ]
    ].sort_values("PVI")

    # ============================================================
    # 高PVI待回执、待回访
    # ============================================================
    high_pending_mask = (
        team_rows["_is_valid"]
        & event_in_year
        & event_in_month
        & team_rows["_pvi"].ge(threshold)
        & (
            team_rows["_receipt_date"].isna()
            | team_rows["_visit_date"].isna()
        )
    )

    high_pending = team_rows.loc[
        high_pending_mask
    ].copy()

    high_pending["待办事项"] = ""

    high_pending.loc[
        high_pending["_receipt_date"].isna(),
        "待办事项",
    ] += "待回执"

    both_missing = (
        high_pending["_receipt_date"].isna()
        & high_pending["_visit_date"].isna()
    )

    high_pending.loc[
        both_missing,
        "待办事项",
    ] += "、"

    high_pending.loc[
        high_pending["_visit_date"].isna(),
        "待办事项",
    ] += "待回访"

    high_pending = high_pending[
        [
            "待办事项",
            "承保日期",
            "投保单号",
            "保单号",
            "出单代理人工号",
            "出单代理人姓名",
            "PVI",
            "承保进度",
            "回执日期",
            "回访完成日期",
        ]
    ].sort_values(
        "PVI",
        ascending=False,
    )

    # ============================================================
    # 汇总指标
    # ============================================================
    gross_pvi = float(
        mtd.loc[
            mtd["_pvi"].gt(0),
            "_pvi",
        ].sum()
    )

    negative_adjustment = float(
        mtd.loc[
            mtd["_pvi"].lt(0),
            "_pvi",
        ].sum()
    )

    net_pvi = float(
        mtd["_pvi"].sum()
    )

    positive_agents = agent_ranking.loc[
        agent_ranking["本月PVI"].gt(0),
        "本月PVI",
    ]

    positive_total = float(
        positive_agents.sum()
    )

    top5_share = (
        float(
            positive_agents.head(5).sum()
            / positive_total
        )
        if positive_total
        else 0.0
    )

    top10_share = (
        float(
            positive_agents.head(10).sum()
            / positive_total
        )
        if positive_total
        else 0.0
    )

    # ============================================================
    # 月末预测
    # ============================================================
    business_days = pd.bdate_range(
        month_start,
        month_end,
    )

    elapsed_business_days = business_days[
        business_days <= report_date
    ]

    remaining_business_days = business_days[
        business_days > report_date
    ]

    positive_mtd = mtd.loc[
        mtd["_pvi"].gt(0)
    ]

    if positive_mtd.empty:
        largest_single_pvi = 0.0
        normal_mtd_excluding_max = net_pvi
        normal_daily = daily.copy()

    else:
        largest_index = (
            positive_mtd["_pvi"].idxmax()
        )

        largest_single_pvi = float(
            mtd.loc[
                largest_index,
                "_pvi",
            ]
        )

        normal_mtd_excluding_max = (
            net_pvi
            - largest_single_pvi
        )

        normal_daily = (
            mtd.drop(index=largest_index)
            .groupby(
                "_insured_date",
                dropna=True,
            )["_pvi"]
            .sum()
            .reindex(
                pd.date_range(
                    month_start,
                    report_date,
                    freq="D",
                ),
                fill_value=0.0,
            )
        )

    normal_business_daily = normal_daily.reindex(
        elapsed_business_days,
        fill_value=0.0,
    )

    normal_daily_average = (
        float(normal_business_daily.mean())
        if len(normal_business_daily)
        else 0.0
    )

    normal_business_forecast = (
        normal_mtd_excluding_max
        + normal_daily_average
        * len(remaining_business_days)
    )

    target_gap = (
        max(
            monthly_target - net_pvi,
            0.0,
        )
        if monthly_target > 0
        else math.nan
    )

    required_daily = (
        target_gap
        / len(remaining_business_days)
        if (
            monthly_target > 0
            and len(remaining_business_days)
        )
        else math.nan
    )

    target_completion = (
        net_pvi / monthly_target
        if monthly_target > 0
        else math.nan
    )

    metrics = {
        "analysis_year": ANALYSIS_YEAR,
        "gross_pvi": gross_pvi,
        "negative_adjustment": negative_adjustment,
        "net_pvi": net_pvi,
        "active_agents": float(
            (
                agent_ranking["本月PVI"]
                > 0
            ).sum()
        ),
        "qualified_agents": float(
            (
                agent_ranking["本月PVI"]
                >= threshold
            ).sum()
        ),
        "top5_share": top5_share,
        "top10_share": top10_share,
        "elapsed_business_days": float(
            len(elapsed_business_days)
        ),
        "remaining_business_days": float(
            len(remaining_business_days)
        ),
        "largest_single_pvi": (
            largest_single_pvi
        ),
        "normal_mtd_excluding_max": (
            normal_mtd_excluding_max
        ),
        "normal_daily_average": (
            normal_daily_average
        ),
        "normal_business_forecast": (
            normal_business_forecast
        ),
        "target_gap": target_gap,
        "required_daily": required_daily,
        "target_completion": (
            target_completion
        ),
    }

    # ============================================================
    # 经营摘要
    # ============================================================
    summary_rows = [
        [
            "分析年份",
            ANALYSIS_YEAR,
            "所有经营结果限定为该年份",
        ],
        [
            "报告日期",
            report_date.strftime("%Y-%m-%d"),
            "必须属于分析年份",
        ],
        [
            "本月新增正PVI",
            gross_pvi,
            "只统计正PVI",
        ],
        [
            "本月批改冲减",
            negative_adjustment,
            "负数表示冲减",
        ],
        [
            "本月净PVI",
            net_pvi,
            "正PVI与负PVI合计，排除犹豫期退保和撤销",
        ],
        [
            "本月活跃代理数",
            int(metrics["active_agents"]),
            "本月PVI大于0",
        ],
        [
            f"达到{threshold:,.0f}代理数",
            int(metrics["qualified_agents"]),
            "按代理本月累计PVI",
        ],
        [
            "前5代理贡献占比",
            top5_share,
            "越高表示集中度越高",
        ],
        [
            "前10代理贡献占比",
            top10_share,
            "越高表示集中度越高",
        ],
        [
            "本月最大一笔PVI",
            largest_single_pvi,
            "用于识别大单影响",
        ],
        [
            "剔除最大一笔后本月PVI",
            normal_mtd_excluding_max,
            "正常业务基础",
        ],
        [
            "剔除最大一笔后的正常业务预测",
            normal_business_forecast,
            "按剔除最大单后的工作日日均推算",
        ],
    ]

    if monthly_target > 0:
        summary_rows.extend(
            [
                [
                    "月度目标",
                    monthly_target,
                    "在代码顶部修改",
                ],
                [
                    "目标完成率",
                    target_completion,
                    "本月净PVI/月度目标",
                ],
                [
                    "目标缺口",
                    target_gap,
                    "不足部分",
                ],
                [
                    "剩余工作日所需日均",
                    required_daily,
                    "未计中国法定节假日调休",
                ],
            ]
        )

    summary = pd.DataFrame(
        summary_rows,
        columns=[
            "指标",
            "数值",
            "说明",
        ],
    )


    # ============================================================
    # 安盛活力星达成追踪
    # ============================================================
    competition_tracking = (
        mtd.groupby(
            [
                "_agency",
                "_region",
                "_organization",
                "_agent_name",
                "_agent_id",
            ],
            dropna=False,
        )["_pvi"]
        .sum()
        .reset_index()
        .rename(
            columns={
                "_agency": "所属机构",
                "_region": "所属区域",
                "_organization": "所属组织",
                "_agent_name": "代理人姓名",
                "_agent_id": "工号",
                "_pvi": "竞赛业绩",
            }
        )
        .sort_values("竞赛业绩", ascending=False)
    )

    # 只保留本月已经出单的人：按代理人汇总后，竞赛业绩必须大于0
    competition_tracking = competition_tracking.loc[
        competition_tracking["竞赛业绩"].gt(0)
    ].copy()

    competition_tracking["_达标差距数值"] = (
        competition_tracking["竞赛业绩"] - threshold
    )
    competition_tracking["达标差距"] = [
        "达标" if value >= 0 else value
        for value in competition_tracking["_达标差距数值"].fillna(-threshold)
    ]
    competition_tracking = competition_tracking.drop(columns=["_达标差距数值"])
    competition_tracking.insert(0, "序号", range(1, len(competition_tracking) + 1))
    competition_tracking = competition_tracking[
        [
            "序号",
            "所属机构",
            "所属区域",
            "所属组织",
            "代理人姓名",
            "工号",
            "竞赛业绩",
            "达标差距",
        ]
    ]

    qualified_region_summary = (
        competition_tracking.loc[
            competition_tracking["达标差距"].eq("达标")
        ]
        .groupby("所属区域", dropna=False)
        .size()
        .rename("达标数")
        .reset_index()
        .rename(columns={"所属区域": "区域"})
        .sort_values("达标数", ascending=False)
    )
    total_qualified = int(qualified_region_summary["达标数"].sum()) if not qualified_region_summary.empty else 0
    qualified_region_summary = pd.concat(
        [
            qualified_region_summary,
            pd.DataFrame([{"区域": "总计", "达标数": total_qualified}]),
        ],
        ignore_index=True,
    )

    # ============================================================
    # (阳光)七八联动七月追踪：阳光人寿深圳分公司
    # ============================================================
    sunshine_rows = mtd.loc[
        text_series(mtd["保险公司"]).str.contains(SUNSHINE_COMPANY_KEYWORD, na=False)
    ].copy()

    sunshine_tracking = (
        sunshine_rows.groupby(
            [
                "_agency",
                "_region",
                "_organization",
                "_agent_name",
                "_agent_id",
            ],
            dropna=False,
        )
        .agg(
            件数=("保单号", "count"),
            **{"7月寿险PVI": ("_pvi", "sum")},
        )
        .reset_index()
        .rename(
            columns={
                "_agency": "所属机构",
                "_region": "所属区域",
                "_organization": "所属组织",
                "_agent_name": "代理人姓名",
                "_agent_id": "工号",
            }
        )
        .sort_values("7月寿险PVI", ascending=False)
    )

    sunshine_tracking = sunshine_tracking.loc[
        sunshine_tracking["7月寿险PVI"].gt(0)
    ].copy()
    sunshine_tracking["_差距数值"] = sunshine_tracking["7月寿险PVI"] - SUNSHINE_THRESHOLD
    sunshine_tracking["差距"] = [
        "达标" if value >= 0 else value
        for value in sunshine_tracking["_差距数值"].fillna(-SUNSHINE_THRESHOLD)
    ]
    sunshine_tracking = sunshine_tracking.drop(columns=["_差距数值"])
    sunshine_tracking.insert(0, "序号", range(1, len(sunshine_tracking) + 1))
    sunshine_tracking = sunshine_tracking[
        [
            "序号",
            "所属机构",
            "所属区域",
            "所属组织",
            "代理人姓名",
            "工号",
            "件数",
            "7月寿险PVI",
            "差距",
        ]
    ]

    sunshine_region_summary = (
        sunshine_tracking.groupby("所属区域", dropna=False)
        .agg(
            符合参赛人数=("工号", "count"),
            达标人数=("差距", lambda values: int((values == "达标").sum())),
        )
        .reset_index()
        .rename(columns={"所属区域": "区域"})
    )
    sunshine_region_summary["未达标人数"] = (
        sunshine_region_summary["符合参赛人数"] - sunshine_region_summary["达标人数"]
    )
    sunshine_region_summary = sunshine_region_summary.sort_values(
        ["达标人数", "符合参赛人数"],
        ascending=[False, False],
    )
    sunshine_region_summary.insert(0, "序号", range(1, len(sunshine_region_summary) + 1))
    sunshine_totals = pd.DataFrame(
        [
            {
                "序号": "总计",
                "区域": "",
                "符合参赛人数": int(sunshine_region_summary["符合参赛人数"].sum()) if not sunshine_region_summary.empty else 0,
                "达标人数": int(sunshine_region_summary["达标人数"].sum()) if not sunshine_region_summary.empty else 0,
                "未达标人数": int(sunshine_region_summary["未达标人数"].sum()) if not sunshine_region_summary.empty else 0,
            }
        ]
    )
    sunshine_region_summary = pd.concat(
        [sunshine_region_summary, sunshine_totals],
        ignore_index=True,
    )

    # ============================================================
    # （信泰）7月专属方案
    # ============================================================
    xintai_rows = mtd.loc[
        text_series(mtd["保险公司"]).str.contains(XINTAI_COMPANY_KEYWORD, na=False)
    ].copy()

    xintai_tracking = (
        xintai_rows.groupby(
            [
                "_agency",
                "_region",
                "_organization",
                "_agent_name",
                "_agent_id",
            ],
            dropna=False,
        )
        .agg(
            件数=("保单号", "count"),
            **{"7月寿险PVI": ("_pvi", "sum")},
        )
        .reset_index()
        .rename(
            columns={
                "_agency": "所属机构",
                "_region": "所属区域",
                "_organization": "所属组织",
                "_agent_name": "代理人姓名",
                "_agent_id": "工号",
            }
        )
        .sort_values("7月寿险PVI", ascending=False)
    )

    xintai_tracking = xintai_tracking.loc[
        xintai_tracking["7月寿险PVI"].gt(0)
    ].copy()
    xintai_tracking["_达标差距数值"] = xintai_tracking["7月寿险PVI"] - XINTAI_THRESHOLD
    xintai_tracking["达标情况"] = [
        "达标" if value >= 0 else value
        for value in xintai_tracking["_达标差距数值"].fillna(-XINTAI_THRESHOLD)
    ]
    xintai_tracking = xintai_tracking.drop(columns=["_达标差距数值"])
    xintai_tracking.insert(0, "序号", range(1, len(xintai_tracking) + 1))
    xintai_tracking = xintai_tracking[
        [
            "序号",
            "所属机构",
            "所属区域",
            "所属组织",
            "代理人姓名",
            "工号",
            "件数",
            "7月寿险PVI",
            "达标情况",
        ]
    ]

    xintai_region_summary = (
        xintai_tracking.groupby("所属区域", dropna=False)
        .agg(
            符合参赛人数=("工号", "count"),
            达标人数=("达标情况", lambda values: int((values == "达标").sum())),
        )
        .reset_index()
        .rename(columns={"所属区域": "区域"})
    )
    xintai_region_summary["未达标人数"] = (
        xintai_region_summary["符合参赛人数"] - xintai_region_summary["达标人数"]
    )
    xintai_region_summary = xintai_region_summary.sort_values(
        ["达标人数", "符合参赛人数"],
        ascending=[False, False],
    )
    xintai_region_summary.insert(0, "序号", range(1, len(xintai_region_summary) + 1))
    xintai_totals = pd.DataFrame(
        [
            {
                "序号": "总计",
                "区域": "",
                "符合参赛人数": int(xintai_region_summary["符合参赛人数"].sum()) if not xintai_region_summary.empty else 0,
                "达标人数": int(xintai_region_summary["达标人数"].sum()) if not xintai_region_summary.empty else 0,
                "未达标人数": int(xintai_region_summary["未达标人数"].sum()) if not xintai_region_summary.empty else 0,
            }
        ]
    )
    xintai_region_summary = pd.concat(
        [xintai_region_summary, xintai_totals],
        ignore_index=True,
    )

    # ============================================================
    # MDRT：全年个人PVI汇总
    # ============================================================

    # MDRT只做这四步筛选：
    # 1. 保留出单团队为大湾区计划
    # 2. 保单状态去掉撤销/犹豫期退保/待承保/空白
    # 3. 出单日期只保留2026年
    # 4. 按出单代理人工号合并PVI，只保留全年PVI大于0
    mdrt_valid = df.loc[
        df["_team"].eq(TEAM_KEYWORD)
        & df["_insured_date"].dt.year.eq(ANALYSIS_YEAR)
        & df["_status"].notna()
        & df["_status"].ne("")
        & ~df["_status"].isin(["撤销", "犹豫期退保", "待承保"])
        & ~df["_status"].str.contains("撤销|犹豫期退保", na=False)
    ].copy()

    annual_valid = mdrt_valid.copy()

    mdrt_pvi = (
        mdrt_valid.groupby("_agent_id", dropna=False)["_pvi"]
        .sum()
        .rename("全年业绩（PVI）")
        .reset_index()
    )

    mdrt_attrs = (
        mdrt_valid.sort_values("_insured_date")
        .drop_duplicates("_agent_id", keep="last")
        [
            [
                "_agent_id",
                "_team",
                "_agency",
                "_region",
                "_organization",
                "_agent_name",
            ]
        ]
    )

    mdrt_tracking = (
        mdrt_pvi.merge(mdrt_attrs, how="left", on="_agent_id")
        .rename(
            columns={
                "_team": "所属团队",
                "_agency": "所属机构",
                "_region": "所属区域",
                "_organization": "所属组织",
                "_agent_name": "代理人姓名",
                "_agent_id": "工号",
            }
        )
    )

    mdrt_tracking.loc[
        mdrt_tracking["所属区域"].eq("SUPA 威信区"),
        "所属机构",
    ] = "总部直辖"

    mdrt_tracking.loc[
        mdrt_tracking["所属区域"].eq("Bees 旭日区"),
        "所属区域",
    ] = "旭日&旭日北辰区"

    mdrt_tracking.loc[
        mdrt_tracking["所属区域"].eq("RH 喜悦区"),
        "所属机构",
    ] = "广东分公司"

    mdrt_tracking = mdrt_tracking.sort_values(
        "全年业绩（PVI）",
        ascending=False,
    )

    mdrt_tracking = mdrt_tracking.loc[
        mdrt_tracking["全年业绩（PVI）"].gt(0)
    ].copy()

    mdrt_tracking.insert(0, "序号", range(1, len(mdrt_tracking) + 1))

    mdrt_tracking = mdrt_tracking[
        [
            "序号",
            "所属团队",
            "所属机构",
            "所属区域",
            "所属组织",
            "代理人姓名",
            "工号",
            "全年业绩（PVI）",
        ]
    ]

    # ============================================================
    # MDRT-区域：全年区域PVI汇总
    # ============================================================

    mdrt_region_rows = annual_valid.copy()
    mdrt_region_rows["_region_normalized"] = (
        mdrt_region_rows["_region"].map(normalize_region_name)
    )

    mdrt_region_summary = (
        mdrt_region_rows.groupby("_region_normalized", dropna=False)["_pvi"]
        .sum()
        .rename("年度PVI")
        .reset_index()
        .rename(columns={"_region_normalized": "所属区域"})
        .sort_values("年度PVI", ascending=False)
    )

    mdrt_region_summary = mdrt_region_summary.loc[
        mdrt_region_summary["年度PVI"].gt(0)
    ].copy()

    mdrt_region_summary.insert(
        0,
        "序号",
        range(1, len(mdrt_region_summary) + 1),
    )

    mdrt_region_summary = mdrt_region_summary[
        ["序号", "所属区域", "年度PVI"]
    ]

    # ============================================================
    # 大湾区属地代理人达成情况：固定名单
    # ============================================================

    # ============================================================
    # 大湾区属地代理人达成情况：固定名单
    # ============================================================
    local_agent_tracking = pd.DataFrame(LOCAL_AGENT_ROSTER).copy()
    local_agent_tracking["_start_date"] = pd.to_datetime(
        local_agent_tracking["入职日期"],
        errors="coerce",
    ).dt.normalize()
    local_agent_tracking["_end_date"] = pd.to_datetime(
        local_agent_tracking["考核截止时间"].astype(str).str.replace("年", "-").str.replace("月", "-").str.replace("日", ""),
        errors="coerce",
    ).dt.normalize()

    # 大湾区属地代理人按各自考核周期累计：入职日期 <= 承保日期 <= 考核截止时间。
    # 这里不用 mtd，否则只会统计本月已承保，容易漏掉固定名单里已有保单的人。
    local_agent_pvi_rows = []
    for _, roster_row in local_agent_tracking.iterrows():
        agent_id = roster_row["工号"]
        start_date = roster_row["_start_date"]
        end_date = roster_row["_end_date"]
        agent_rows = df.loc[
            df["_agent_id"].eq(agent_id)
            & df["_insured_date"].between(start_date, end_date, inclusive="both")
        ]
        local_agent_pvi_rows.append(
            {
                "工号": agent_id,
                "PVI（元）": agent_rows["_pvi"].sum(),
            }
        )

    local_agent_pvi = pd.DataFrame(local_agent_pvi_rows)
    local_agent_tracking = local_agent_tracking.merge(
        local_agent_pvi,
        how="left",
        on="工号",
    )
    local_agent_tracking["PVI（元）"] = local_agent_tracking["PVI（元）"].fillna(0.0)
    local_agent_tracking["考核差距（元）"] = local_agent_tracking["PVI（元）"] - LOCAL_AGENT_THRESHOLD
    local_agent_tracking.insert(0, "序号", range(1, len(local_agent_tracking) + 1))
    local_agent_tracking = local_agent_tracking[
        [
            "序号",
            "所属区域",
            "所属组织",
            "工号",
            "代理人姓名",
            "PVI（元）",
            "考核差距（元）",
            "考核截止时间",
            "入职日期",
        ]
    ]

    local_agent_region_summary = (
        local_agent_tracking.groupby("所属区域", dropna=False)
        .agg(
            考核人数=("工号", "count"),
            已达标人数=("考核差距（元）", lambda values: int((values >= 0).sum())),
        )
        .reset_index()
        .rename(columns={"所属区域": "区域"})
    )
    local_agent_region_summary["未达标人数"] = (
        local_agent_region_summary["考核人数"] - local_agent_region_summary["已达标人数"]
    )
    local_agent_region_summary = local_agent_region_summary.sort_values(
        ["考核人数", "区域"],
        ascending=[False, True],
    )
    local_agent_region_summary.insert(0, "序号", range(1, len(local_agent_region_summary) + 1))
    local_agent_total = pd.DataFrame(
        [
            {
                "序号": "合计",
                "区域": "",
                "考核人数": int(local_agent_region_summary["考核人数"].sum()) if not local_agent_region_summary.empty else 0,
                "已达标人数": int(local_agent_region_summary["已达标人数"].sum()) if not local_agent_region_summary.empty else 0,
                "未达标人数": int(local_agent_region_summary["未达标人数"].sum()) if not local_agent_region_summary.empty else 0,
            }
        ]
    )
    local_agent_region_summary = pd.concat(
        [local_agent_region_summary, local_agent_total],
        ignore_index=True,
    )

    return {
        "summary": summary,
        "daily": daily_df,
        "agent_ranking": agent_ranking,
        "product_ranking": product_ranking,
        "company_ranking": company_ranking,
        "not_issued": not_issued,
        "near_threshold": near_threshold,
        "inactive": inactive,
        "pending_underwrite": pending_underwrite,
        "exceptions": exceptions,
        "high_pending": high_pending,
        "competition_tracking": competition_tracking,
        "qualified_region_summary": qualified_region_summary,
        "sunshine_tracking": sunshine_tracking,
        "sunshine_region_summary": sunshine_region_summary,
        "xintai_tracking": xintai_tracking,
        "xintai_region_summary": xintai_region_summary,
        "mdrt_tracking": mdrt_tracking,
        "mdrt_region_summary": mdrt_region_summary,
        "local_agent_tracking": local_agent_tracking,
        "local_agent_region_summary": local_agent_region_summary,
        "metrics": metrics,
    }
