from __future__ import annotations

from pathlib import Path

import pandas as pd

from metrics import money


def generate_summary(
    reports: dict[str, pd.DataFrame | dict[str, float] | pd.Series],
    report_date: pd.Timestamp,
    today_pvi: float,
    monthly_target: float,
) -> str:
    metrics = reports["metrics"]
    assert isinstance(metrics, dict)

    net_pvi = float(metrics["net_pvi"])
    target_gap = float(metrics["target_gap"])
    target_completion = float(metrics["target_completion"])
    normal_forecast = float(metrics["normal_business_forecast"])

    if target_gap <= 0:
        target_sentence = "当前已完成月度目标。"
    else:
        target_sentence = f"距离月度目标仍有 {money(target_gap)} 元缺口。"

    if normal_forecast >= monthly_target:
        forecast_sentence = "按当前正常业务节奏预测，月末有望达成目标。"
    else:
        forecast_gap = monthly_target - normal_forecast
        forecast_sentence = (
            f"按剔除最大单后的正常业务节奏预测，月末预计仍有 "
            f"{money(forecast_gap)} 元缺口。"
        )

    return (
        f"{report_date:%Y-%m-%d} 经销业务 PVI 经营摘要："
        f"今日新增 PVI 为 {money(today_pvi)} 元，"
        f"本月累计净 PVI 为 {money(net_pvi)} 元，"
        f"月度目标达成率为 {target_completion:.2%}，"
        f"{target_sentence}"
        f"{forecast_sentence}"
        "建议重点关注未达标代理人、低活跃区域及临近达标人员。"
    )


def save_summary(summary: str, output_path: Path) -> None:
    output_path.write_text(summary + "\n", encoding="utf-8")
