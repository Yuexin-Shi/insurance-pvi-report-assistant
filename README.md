# 保险经销业务 PVI 智能分析与 RAG 问答系统

## 项目背景

本项目面向保险公司经销业务部日常经营管理场景，用于将每日保单汇总明细自动加工为 PVI 经营看板、多页签 Excel 追踪报告和文字经营摘要，减少人工汇总、重复制表和报告撰写工作。

在自动报表基础上，项目进一步加入轻量级 RAG 问答与 Pandas 数据查询能力，支持对业务规则、指标口径和经营数据进行问答式查询。

## 核心功能

- 自动读取 `data/raw/` 中的 `保单汇总列表*.xlsx` 文件
- 清洗保单状态、承保日期、代理人、区域、保险公司和 PVI 字段
- 按团队口径筛选经销业务数据
- 计算每日净 PVI、本月累计 PVI、月度目标达成率、目标缺口和月末预测
- 生成代理人、险种、保险公司、区域等维度分析
- 输出经营看板 PNG、多页签 Excel 报告和文字经营摘要
- 支持安盛活力星、阳光七八联动、信泰专属方案、MDRT、属地代理人达成情况等专题追踪
- 支持轻量级 RAG 问答，查询指标口径、业务规则、报告生成规则和竞赛方案说明
- 支持 Pandas 结构化数据查询，回答今日 PVI、本月 PVI、出单名单、代理人排名、产品和保险公司贡献等问题

## 项目结构

```text
pvi-report-assistant/
  data/
    raw/              # 原始数据，不对外展示
    sample/           # 脱敏样例数据
  output/             # 自动生成的报告、看板和摘要
  src/
    config.py
    data_cleaning.py
    metrics.py
    dashboard.py
    report_export.py
    summary.py
    data_query.py
    main.py
  rag/
    knowledge_base/
      指标口径说明.md
      业务规则说明.md
      报告生成规则.md
      竞赛方案说明.md
      项目说明.md
    build_index.py
    qa.py
    index.json
  README.md
  requirements.txt

## 技术栈

Python、Pandas、Matplotlib、OpenPyXL、轻量级文本检索

## 运行方式

请先将输入文件放入 `data/raw/`，文件名需符合：

```text
保单汇总列表*.xlsx
```

然后在终端进入项目根目录，安装依赖并运行：

```bash
pip install -r requirements.txt
python src/main.py
```

如果本机默认使用 `python3`，可使用：

```bash
pip3 install -r requirements.txt
python3 src/main.py
```

## 示例输出

程序运行后会在 `output/` 文件夹生成：

- `output/PVI经营报告_20260728.xlsx`
- `output/PVI经营看板_20260728.png`
- `output/经营摘要_20260728.txt`

## 智能问答模块

项目包含轻量级 RAG + Pandas 数据查询模块：

- 指标口径、业务规则、竞赛方案等解释类问题：通过 `rag/knowledge_base/` 知识库检索回答；
- 今日 PVI、本月 PVI、代理人排名、出单名单等经营数据问题：通过 Pandas 读取 Excel 数据精确计算回答；
- 支持一次输入多个问题，系统会自动拆分并逐个回答。

### 构建知识库索引

如果修改了 `rag/knowledge_base/` 中的文档，需要重新构建索引：

```bash
python rag/build_index.py
```

### 问答示例

```bash
python rag/qa.py "PVI 是什么"
python rag/qa.py "目标达成率怎么定义"
python rag/qa.py "今日PVI是多少"
python rag/qa.py "今天都谁出单了"
python rag/qa.py "今天出单人数是多少"
python rag/qa.py "本月PVI最高的是谁"
python rag/qa.py "本月产品PVI最高的是哪个"
python rag/qa.py "本月保险公司PVI最高的是哪家"
python rag/qa.py "信泰方案规则是什么"
python rag/qa.py "今天都谁出单了，本月PVI是多少，PVI是什么"
```

## 数据说明

`data/raw/` 用于存放真实业务数据，不对外展示。  
`data/sample/` 存放脱敏样例数据，用于项目展示和复现。  
`rag/knowledge_base/` 存放脱敏后的业务规则和指标口径说明，用于 RAG 检索。
