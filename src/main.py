from __future__ import annotations

import pandas as pd

from config import *
from dashboard import draw_dashboard
from data_cleaning import clean_data, latest_source_file
from metrics import calculate_reports, determine_report_date, money
from report_export import export_excel
from summary import generate_summary, save_summary

def main() -> None:
    input_path = latest_source_file().resolve()
    print(f"正在分析的Excel表格：{input_path.name}")
    print(f"完整路径：{input_path}")
    if not input_path.exists():
        raise FileNotFoundError(input_path)

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
    draw_dashboard(reports, dashboard_path, report_date, PVI_THRESHOLD)
    export_excel(reports, excel_path, dashboard_path, report_date)

    # =================【核心修复：在此处计算 today_pvi】=================
    daily_report = reports["daily"]
    assert isinstance(daily_report, pd.DataFrame)
    
    today_pvi = float(
        daily_report.loc[
            daily_report["日期"].eq(report_date),
            "每日净PVI"
        ].sum()
    )
    # ===================================================================

    metrics = reports["metrics"]
    assert isinstance(metrics, dict)
    summary_text = generate_summary(reports, report_date, today_pvi, MONTHLY_TARGET)
    save_summary(summary_text, summary_path)
    print(f"数据文件：{input_path}")
    print(f"报告日期：{report_date:%Y-%m-%d}")
    print(f"今日PVI：{money(today_pvi)} 元") # 现在这里就不会报错了
    print(f"本月净PVI：{money(metrics['net_pvi'])} 元")
    print(f"本月最大一笔PVI：{money(metrics['largest_single_pvi'])} 元")
    print(f"剔除最大一笔后的正常业务预测：{money(metrics['normal_business_forecast'])} 元")
    print(f"月度目标缺口：{money(metrics['target_gap'])} 元")
    print(f"经营摘要：{summary_path}")
    print(summary_text)
    print(f"图片看板：{dashboard_path}")
    print(f"Excel报告：{excel_path}")


if __name__ == "__main__":
    main()
