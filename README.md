# 保险经销业务 PVI 经营分析与自动报表生成系统

## 项目背景

本项目面向保险公司经销业务部日常经营管理场景，用于将每日保单汇总明细自动加工为 PVI 经营看板和多页签 Excel 追踪报告，减少人工汇总和重复制表工作。

## 核心功能

- 自动读取 `data/raw/` 或项目目录中的 `保单汇总列表*.xlsx` 文件
- 清洗保单状态、承保日期、代理人、区域、保险公司和 PVI 字段
- 按团队口径筛选经销业务数据
- 计算每日净 PVI、本月累计 PVI、月度目标达成率、目标缺口和月末预测
- 生成代理人、险种、保险公司、区域等维度分析
- 输出经营看板 PNG 和多页签 Excel 报告
- 支持安盛活力星、阳光七八联动、信泰专属方案、MDRT、属地代理人达成情况等专题追踪

## 项目结构

```text
pvi-report-assistant/
  data/          
    sample/       # 脱敏样例数据
  output/         # 自动生成的报告和看板
  src/
    config.py
    data_cleaning.py
    metrics.py
    dashboard.py
    report_export.py
    main.py
  README.md
  requirements.txt
```

## 技术栈

Python、Pandas、Matplotlib、OpenPyXL

## 运行方式

请先将输入文件放入 data/raw/，文件名需符合：保单汇总列表*.xlsx

然后在终端进入项目根目录，运行：

```bash
pip install -r requirements.txt
python src/main.py

```
如果本机默认使用 python3，可使用：
```bash
pip3 install -r requirements.txt
python3 src/main.py

```

## 示例输出

- `output/PVI经营报告_20260728.xlsx`
- `output/PVI经营看板_20260728.png`
- `output/经营摘要_20260728.txt`

## 数据说明

`data/raw/` 用于存放真实业务数据，不对外展示。  
`data/sample/` 存放脱敏样例数据，用于项目展示和复现。

## 数据安全说明

真实业务数据中可能包含代理人工号、代理人姓名、保单号、投保单号等敏感信息。对外展示时建议只使用 `data/sample/` 中的脱敏样例数据，并删除真实姓名、工号、保单号、投保单号等字段或替换为模拟编号。
