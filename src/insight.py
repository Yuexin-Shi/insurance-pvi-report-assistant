from __future__ import annotations

from pathlib import Path

import pandas as pd

from metrics import money


def _first_value(frame: pd.DataFrame, column: str, default: str = "暂无") -> str:
    if frame.empty or column not in frame.columns:
        return default
    value = frame.iloc[0][column]
    if pd.isna(value):
        return default
    return str(value)


def _first_amount(frame: pd.DataFrame, column: str) -> float:
    if frame.empty or column not in frame.columns:
        return 0.0
    value = frame.iloc[0][column]
    if pd.isna(value):
        return 0.0
    return float(value)


def _completion_comment(completion: float) -> str:
    if completion >= 1:
        return "月度目标已完成，后续重点可转向稳固高贡献代理人和控制退保、撤单等负向影响。"
    if completion >= 0.8:
        return "当前距离月度目标较近，应优先推动临近达标代理人和高潜力产品，争取在剩余工作日内补足缺口。"
    if completion >= 0.6:
        return "当前达成率处于追赶区间，需要同时关注头部贡献延续和中腰部代理人激活。"
    return "当前达成率偏低，建议尽快拆解缺口到代理人、产品和区域层面，形成更明确的追踪动作。"


def _concentration_comment(top5_share: float) -> str:
    if top5_share >= 0.7:
        return "PVI 贡献高度集中在头部代理人，短期业绩弹性较强，但也需要关注大单依赖风险。"
    if top5_share >= 0.5:
        return "PVI 贡献存在一定头部集中度，建议在维持头部代理人产能的同时扩大有效出单面。"
    return "PVI 贡献分布相对分散，整体出单结构更均衡，可继续提升中腰部代理人的稳定产能。"


def generate_insight(
    reports: dict[str, pd.DataFrame | dict[str, float] | pd.Series],
    report_date: pd.Timestamp,
    today_pvi: float,
    monthly_target: float,
) -> str:
    metrics = reports["metrics"]
    agent_ranking = reports["agent_ranking"]
    product_ranking = reports["product_ranking"]
    company_ranking = reports["company_ranking"]
    near_threshold = reports["near_threshold"]
    inactive = reports["inactive"]

    assert isinstance(metrics, dict)
    assert isinstance(agent_ranking, pd.DataFrame)
    assert isinstance(product_ranking, pd.DataFrame)
    assert isinstance(company_ranking, pd.DataFrame)
    assert isinstance(near_threshold, pd.DataFrame)
    assert isinstance(inactive, pd.DataFrame)

    net_pvi = float(metrics["net_pvi"])
    target_completion = float(metrics["target_completion"])
    target_gap = float(metrics["target_gap"])
    required_daily = float(metrics["required_daily"])
    normal_forecast = float(metrics["normal_business_forecast"])
    top5_share = float(metrics["top5_share"])
    active_agents = int(float(metrics["active_agents"]))
    qualified_agents = int(float(metrics["qualified_agents"]))

    top_agent = _first_value(agent_ranking, "出单代理人姓名")
    top_agent_pvi = _first_amount(agent_ranking, "本月PVI")
    top_product = _first_value(product_ranking, "险种名称")
    top_product_pvi = _first_amount(product_ranking, "本月PVI")
    top_company = _first_value(company_ranking, "保险公司")
    top_company_pvi = _first_amount(company_ranking, "本月PVI")

    if target_gap <= 0:
        gap_sentence = "本月目标已达成，暂无目标缺口。"
    else:
        gap_sentence = (
            f"距离月度目标仍有 {money(target_gap)} 元缺口，"
            f"剩余工作日平均每日需新增约 {money(required_daily)} 元。"
        )

    if normal_forecast >= monthly_target:
        forecast_sentence = "剔除最大单后的正常业务预测显示，月末仍有机会达成目标。"
    else:
        forecast_sentence = (
            "剔除最大单后的正常业务预测低于月度目标，说明当前常规业务节奏仍需加速。"
        )

    near_count = len(near_threshold)
    inactive_count = len(inactive)

    lines = [
        f"{report_date:%Y-%m-%d} PVI 经营智能解读",
        "",
        "一、整体经营表现",
        (
            f"截至报告日，本月累计净 PVI 为 {money(net_pvi)} 元，"
            f"今日新增 PVI 为 {money(today_pvi)} 元，"
            f"月度目标达成率为 {target_completion:.2%}。{_completion_comment(target_completion)}"
        ),
        "",
        "二、目标达成与节奏判断",
        f"{gap_sentence}{forecast_sentence}",
        "",
        "三、代理人贡献结构",
        (
            f"本月活跃出单代理人共 {active_agents} 人，其中达到门槛代理人 {qualified_agents} 人。"
            f"当前排名第一的代理人为 {top_agent}，本月贡献 PVI {money(top_agent_pvi)} 元。"
            f"Top 5 代理人贡献占比为 {top5_share:.2%}，{_concentration_comment(top5_share)}"
        ),
        "",
        "四、产品与保险公司贡献",
        (
            f"当前贡献最高的产品/险种为 {top_product}，贡献 PVI {money(top_product_pvi)} 元；"
            f"贡献最高的保险公司为 {top_company}，贡献 PVI {money(top_company_pvi)} 元。"
            "建议持续跟踪高贡献产品的转化质量，同时关注贡献过度集中带来的结构风险。"
        ),
        "",
        "五、风险提示与跟进建议",
        (
            f"当前临近门槛代理人 {near_count} 人，连续未出单或低活跃代理人 {inactive_count} 人。"
            "建议优先跟进临近达标人员、低活跃区域和高潜力产品，"
            "并将目标缺口拆解到代理人和产品维度进行日度追踪。"
        ),
    ]
    return "\n".join(lines)


def save_insight(insight: str, output_path: Path) -> None:
    output_path.write_text(insight + "\n", encoding="utf-8")
