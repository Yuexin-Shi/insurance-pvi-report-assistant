# 保险经销业务 PVI 经营分析与智能报表生成系统

在线 Demo：[Streamlit App](https://insurance-pvi-report-assistant.streamlit.app/)  
项目代码：[GitHub Repository](https://github.com/Yuexin-Shi/insurance-pvi-report-assistant)

## 项目背景

本项目面向保险公司经销业务部日常经营管理场景，用于将每日保单汇总明细自动加工为 PVI 经营看板、多页签 Excel 追踪报告和文字经营摘要，减少人工汇总、重复制表和报告撰写工作。

在自动报表基础上，项目进一步加入轻量级 RAG 问答与 Pandas 数据查询能力，支持对业务规则、指标口径和经营数据进行问答式查询。

## 核心功能

- 自动读取 `data/raw/` 中的 `保单汇总列表*.xlsx` 文件
- 清洗保单状态、承保日期、代理人、区域、保险公司和 PVI 字段
- 按团队口径筛选经销业务数据
- 计算每日净 PVI、本月累计 PVI、月度目标达成率、目标缺口和月末预测
- 生成代理人、险种、保险公司、区域等维度分析
- 输出经营看板 PNG、多页签 Excel 报告、文字经营摘要和智能经营解读
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
    insight.py
    data_query.py
    main.py
  app.py             # Streamlit 本地交互式应用
  rag/
    knowledge_base/
      指标口径说明.md
      业务规则说明.md
      报告生成规则.md
      竞赛方案说明.md
      项目说明.md
    build_index.py
    intent_classifier.py
    answer_generator.py
    logger.py
    qa.py
    index.json
    vector_index.py
    vector_qa.py
    vector_index.pkl
  eval/
    questions.json
    run_eval.py
  demo/
    index.html
    demo_script.md
  logs/              # 本地问答日志，不建议上传
  README.md
  requirements.txt
```

## 技术栈

Python、Pandas、Matplotlib、OpenPyXL、Scikit-learn、轻量级文本检索、TF-IDF 向量检索

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

## 快速 Demo

项目提供脱敏样例数据，可用于本地快速复现完整流程。首次运行时，请将样例数据复制到 `data/raw/`：

```bash
mkdir -p data/raw
cp data/sample/保单汇总列表_demo.xlsx data/raw/
```

然后依次运行：

```bash
pip install -r requirements.txt
python src/main.py
python rag/build_index.py
python rag/vector_index.py
python rag/qa.py "今天pvi多少"
python rag/qa.py "今天都谁出单了，本月PVI是多少，PVI是什么"
python rag/vector_qa.py "PVI达成率咋算"
python -u eval/run_eval.py
```

如果本机默认使用 `python3`，可将以上命令中的 `python` 替换为 `python3`。

Demo 会展示从脱敏保单明细到 Excel 报告、PNG 看板、经营摘要、智能解读、RAG 问答和标准评估集的完整链路。

## Streamlit 交互式应用

项目新增 `app.py`，用于将自动报表和智能问答封装成本地网页应用。每天使用时，可以将新的 `保单汇总列表*.xlsx` 放入 `data/raw/`，然后在网页中点击“生成/刷新报告”查看最新结果。

启动方式：

```bash
streamlit run app.py
```

页面功能包括：

- 上传或读取最新保单汇总表；
- 生成 Excel 经营报告、PNG 看板、经营摘要和智能解读；
- 查看今日 PVI、本月累计 PVI、目标达成率和目标缺口；
- 查看今日出单名单和代理人排名；
- 在网页中进行 RAG + Pandas 问答；
- 下载自动生成的输出文件。

## Streamlit Cloud 线上 Demo

本项目支持部署到 Streamlit Community Cloud。线上 Demo 默认读取 `data/sample/` 中的脱敏样例数据；当 `data/raw/` 不存在或没有原始文件时，系统会自动使用：

```text
data/sample/保单汇总列表_demo.xlsx
```

部署步骤：

1. 将项目上传到 GitHub；
2. 确认仓库中包含 `app.py`、`requirements.txt`、`packages.txt`、`runtime.txt`、`src/`、`rag/` 和 `data/sample/`；
3. 登录 Streamlit Community Cloud；
4. 新建应用，选择 GitHub 仓库；
5. Main file path 填写：

```text
app.py
```

6. 点击 Deploy。

部署后会得到一个公开链接，例如：

```text
https://your-app-name.streamlit.app
```

注意：线上部署只能使用脱敏样例数据，不能上传真实 `data/raw/`、真实代理人姓名、工号、保单号或投保单号。

## 示例输出

程序运行后会在 `output/` 文件夹生成：

- `output/PVI经营报告_20260728.xlsx`
- `output/PVI经营看板_20260728.png`
- `output/经营摘要_20260728.txt`
- `output/智能解读_20260728.txt`

## Demo 展示建议

可参考 `demo/demo_script.md` 进行项目展示，建议顺序为：

1. 展示 `data/sample/保单汇总列表_demo.xlsx` 脱敏样例数据；
2. 运行 `python src/main.py` 生成经营报告、看板、摘要和智能解读；
3. 展示 `output/PVI经营看板_20260728.png` 和 `output/PVI经营报告_20260728.xlsx`；
4. 展示 `output/经营摘要_20260728.txt` 和 `output/智能解读_20260728.txt`；
5. 运行 RAG 问答和 TF-IDF 向量检索问答；
6. 运行 `python -u eval/run_eval.py` 展示标准问题评估结果。

项目也提供一个静态产品界面 Demo：

```text
demo/index.html
```

可直接双击打开，或在浏览器中打开该文件。该页面模拟“保险经销业务 PVI 智能助手”的正式系统界面，展示经营问答、报表看板、指标口径、专题追踪、日志评估等入口。

## 报告智能解读

项目新增报告智能解读模块，由 `src/insight.py` 根据已计算的 PVI 指标和分析结果自动生成经营解读，内容包括：

- 整体经营表现
- 目标达成与节奏判断
- 代理人贡献结构
- 产品与保险公司贡献
- 风险提示与跟进建议

该模块会结合本月累计净 PVI、今日新增 PVI、目标达成率、目标缺口、月末预测、代理人 Top 贡献、产品/保险公司贡献和低活跃代理人情况，输出适合日报、周报或月报引用的文字分析。

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

### TF-IDF 向量检索 RAG

项目新增轻量级向量检索 RAG，通过 Scikit-learn 的 `TfidfVectorizer` 将业务规则、指标口径和报告生成文档切分并向量化，再基于余弦相似度检索最相关的文档片段。

构建向量索引：

```bash
python rag/vector_index.py
```

使用向量检索问答：

```bash
python rag/vector_qa.py "PVI 是什么"
python rag/vector_qa.py "PVI达成率咋算"
python rag/vector_qa.py "经营看板包含哪些内容"
python rag/vector_qa.py "信泰方案规则是什么"
```

该模块主要用于展示向量化检索能力，适合与 `rag/qa.py` 的意图识别 + Pandas 数据查询模块配合使用。

### 问答日志与评估

每次运行 `rag/qa.py` 后，系统会自动将问题、识别意图、查询路径和回答结果记录到：

```text
logs/qa_log.jsonl
```

该日志用于后续分析用户常问问题和优化意图识别规则。由于日志可能包含业务查询结果，`logs/` 已加入 `.gitignore`，不建议上传到 GitHub。

项目内置了一组标准评估问题，可用于检查意图识别和回答效果：

```bash
python eval/run_eval.py
```

评估脚本会检查每个问题的识别意图是否正确，以及回答中是否包含预期关键词。

## 数据说明

`data/raw/` 用于存放真实业务数据，不对外展示。  
`data/sample/` 存放脱敏样例数据，用于项目展示和复现。  
`rag/knowledge_base/` 存放脱敏后的业务规则和指标口径说明，用于 RAG 检索。
