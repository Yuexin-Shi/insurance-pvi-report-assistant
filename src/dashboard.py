from __future__ import annotations

from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
from matplotlib import font_manager
from matplotlib.ticker import FuncFormatter, PercentFormatter

from metrics import money, short_money

def setup_chinese_font() -> None:
    installed = {font.name for font in font_manager.fontManager.ttflist}
    candidates = [
        "PingFang SC",
        "Hiragino Sans GB",
        "Heiti SC",
        "Microsoft YaHei",
        "SimHei",
        "Arial Unicode MS",
        "Noto Sans CJK SC",
    ]
    selected = next((font for font in candidates if font in installed), "DejaVu Sans")
    plt.rcParams["font.sans-serif"] = [selected]
    plt.rcParams["axes.unicode_minus"] = False

def draw_dashboard(
    reports: dict[str, pd.DataFrame | dict[str, float] | pd.Series],
    output_path: Path,
    report_date: pd.Timestamp,
    threshold: float,
) -> None:
    setup_chinese_font()
    daily = reports["daily"]
    agents = reports["agent_ranking"]
    products = reports["product_ranking"]
    companies = reports["company_ranking"]
    metrics = reports["metrics"]
    assert isinstance(daily, pd.DataFrame)
    assert isinstance(agents, pd.DataFrame)
    assert isinstance(products, pd.DataFrame)
    assert isinstance(companies, pd.DataFrame)
    assert isinstance(metrics, dict)
    
    
    
    

    fig, axes = plt.subplots(2, 2, figsize=(18, 12), facecolor="#F5F7FA")
    for axis in axes.flat:
        axis.set_facecolor("white")
        axis.grid(axis="y", color="#E5E7EB", linewidth=0.8, alpha=0.8)
        for spine in axis.spines.values():
            spine.set_visible(False)

    ax = axes[0, 0]
    labels = [f"{date.month}/{date.day}" for date in daily["日期"]]
    ax.bar(labels, daily["每日净PVI"], color="#4C78A8", label="每日净PVI")
    ax.set_title("每日PVI与本月累计", loc="left", fontsize=15, fontweight="bold")
    ax.yaxis.set_major_formatter(FuncFormatter(short_money))
    ax.tick_params(axis="x", rotation=45)
    cumulative_axis = ax.twinx()
    cumulative_axis.plot(labels, daily["本月累计PVI"], color="#F28E2B", marker="o", linewidth=2.5, label="累计PVI")
    cumulative_axis.yaxis.set_major_formatter(FuncFormatter(short_money))
    cumulative_axis.grid(False)
    cumulative_axis.spines["top"].set_visible(False)
    cumulative_axis.spines["right"].set_visible(False)
    lines, line_labels = ax.get_legend_handles_labels()
    lines2, line_labels2 = cumulative_axis.get_legend_handles_labels()
    ax.legend(lines + lines2, line_labels + line_labels2, loc="upper left", frameon=False)

    ax = axes[0, 1]
    top_agents = agents.loc[agents["本月PVI"].gt(0)].head(15).copy().iloc[::-1]
    if top_agents.empty:
        ax.text(0.5, 0.5, "本月暂无代理PVI", ha="center", va="center")
    else:
        agent_labels = [
            f"{name}（{str(agent_id)[-4:]}）"
            for name, agent_id in zip(top_agents["出单代理人姓名"], top_agents["出单代理人工号"])
        ]
        colors = ["#59A14F" if value >= threshold else "#9C755F" for value in top_agents["本月PVI"]]
        ax.barh(agent_labels, top_agents["本月PVI"], color=colors)
        ax.xaxis.set_major_formatter(FuncFormatter(short_money))
    ax.set_title("代理人本月PVI Top 15", loc="left", fontsize=15, fontweight="bold")

    ax = axes[1, 0]
    concentration = agents.loc[agents["本月PVI"].gt(0), ["出单代理人姓名", "本月PVI"]].copy()
    concentration = concentration.head(20)
    if concentration.empty:
        ax.text(0.5, 0.5, "本月暂无贡献数据", ha="center", va="center")
    else:
        x = list(range(1, len(concentration) + 1))
        positive_total = agents.loc[agents["本月PVI"].gt(0), "本月PVI"].sum()
        cumulative_share = concentration["本月PVI"].cumsum() / positive_total
        ax.bar(x, concentration["本月PVI"], color="#76B7B2")
        ax.set_xlabel("代理排名")
        ax.yaxis.set_major_formatter(FuncFormatter(short_money))
        share_axis = ax.twinx()
        share_axis.plot(x, cumulative_share, color="#E15759", marker="o", linewidth=2)
        share_axis.yaxis.set_major_formatter(PercentFormatter(1.0))
        share_axis.set_ylim(0, 1.05)
        share_axis.grid(False)
        share_axis.spines["top"].set_visible(False)
        share_axis.spines["right"].set_visible(False)
        top5 = metrics["top5_share"]
        top10 = metrics["top10_share"]
        ax.text(
            0.02,
            0.94,
            f"前5占比 {top5:.1%} ｜ 前10占比 {top10:.1%}",
            transform=ax.transAxes,
            fontsize=11,
            va="top",
            bbox={"boxstyle": "round,pad=0.4", "facecolor": "#FFF4E6", "edgecolor": "none"},
        )
    ax.set_title("PVI贡献集中度（帕累托）", loc="left", fontsize=15, fontweight="bold")

    ax = axes[1, 1]
    top_products = products.head(5).copy()
    top_companies = companies.head(5).copy()
    mix_labels = [f"产品｜{str(value)[:18]}" for value in top_products["险种名称"]]
    mix_values = list(top_products["本月PVI"])
    mix_colors = ["#4C78A8"] * len(top_products)
    mix_labels += [f"公司｜{str(value)[:18]}" for value in top_companies["保险公司"]]
    mix_values += list(top_companies["本月PVI"])
    mix_colors += ["#F28E2B"] * len(top_companies)
    if mix_values:
        ax.barh(mix_labels[::-1], mix_values[::-1], color=mix_colors[::-1])
        ax.xaxis.set_major_formatter(FuncFormatter(short_money))
    else:
        ax.text(0.5, 0.5, "本月暂无产品/公司数据", ha="center", va="center")
    ax.set_title("产品与保险公司贡献 Top 5", loc="left", fontsize=15, fontweight="bold")

    title = f"大湾区PVI经营看板｜截至{report_date:%Y-%m-%d}"
    subtitle = (
        f"本月净PVI {money(metrics['net_pvi'])} 元　"
        f"活跃代理 {int(metrics['active_agents'])} 人　"
        f"最大一笔 {money(metrics['largest_single_pvi'])} 元　"
        f"正常业务预测 {money(metrics['normal_business_forecast'])} 元"
    )
    fig.suptitle(title, x=0.05, y=0.985, ha="left", fontsize=23, fontweight="bold", color="#1F2937")
    fig.text(0.05, 0.945, subtitle, fontsize=12, color="#4B5563")
    fig.tight_layout(rect=[0.03, 0.03, 0.98, 0.92], h_pad=3.2, w_pad=2.6)
    fig.savefig(output_path, dpi=180, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
