# Demo 展示脚本

## 1. 项目一句话介绍

这是一个面向保险经销业务部的数据分析与智能报表生成项目。项目可以将每日保单汇总明细自动清洗、计算 PVI 经营指标，并生成 Excel 经营报告、PNG 可视化看板、文字经营摘要和智能经营解读。同时，项目加入了轻量级 RAG 问答、Pandas 数据查询、TF-IDF 向量检索和标准问答评估集。

## 2. Demo 数据说明

本项目使用脱敏样例数据进行展示：

```text
data/sample/保单汇总列表_demo.xlsx
```

真实业务数据不应上传或公开。展示前可将 demo 数据复制到 `data/raw/`：

```bash
mkdir -p data/raw
cp data/sample/保单汇总列表_demo.xlsx data/raw/
```

## 3. 运行自动报表流程

在项目根目录运行：

```bash
python src/main.py
```

运行后会自动生成：

```text
output/PVI经营报告_20260728.xlsx
output/PVI经营看板_20260728.png
output/经营摘要_20260728.txt
output/智能解读_20260728.txt
```

展示重点：

- Excel 报告包含多页签专题追踪和维度分析；
- PNG 看板展示每日 PVI、本月累计、代理人排名、贡献集中度、产品和保险公司 Top 贡献；
- 经营摘要适合直接放入日报、周报或月报；
- 智能解读进一步输出经营表现、节奏判断、结构风险和跟进建议。

## 4. 展示智能问答

构建知识库索引：

```bash
python rag/build_index.py
```

示例问题：

```bash
python rag/qa.py "今天pvi多少"
python rag/qa.py "本月哪个业务员最厉害"
python rag/qa.py "今天有几个人出单"
python rag/qa.py "PVI达成率咋算"
python rag/qa.py "今天都谁出单了，本月PVI是多少，PVI是什么"
```

展示重点：

- 指标解释类问题走 RAG 知识库；
- 今日 PVI、本月 PVI、出单名单、排名类问题走 Pandas 精确查询；
- 多问题输入会自动拆分并逐个回答；
- 不支持的问题会进入兜底回答，避免乱答。

## 5. 展示 TF-IDF 向量检索 RAG

构建向量索引：

```bash
python rag/vector_index.py
```

示例问题：

```bash
python rag/vector_qa.py "PVI 是什么"
python rag/vector_qa.py "PVI达成率咋算"
python rag/vector_qa.py "经营看板包含哪些内容"
python rag/vector_qa.py "信泰方案规则是什么"
```

展示重点：

- 将知识库文档切分为文本片段；
- 使用 TF-IDF 将问题和文档片段向量化；
- 通过余弦相似度检索最相关内容；
- 返回回答、来源文件和相似度。

## 6. 展示评估结果

运行标准评估集：

```bash
python -u eval/run_eval.py
```

当前评估覆盖：

- 指标解释；
- 业务规则；
- 经营数据查询；
- 出单名单；
- 代理人、产品、保险公司排名；
- 多问题组合；
- 口语化表达；
- 不支持问题兜底。

预期结果：

```text
评估结果：41/41 通过，通过率 100%
```

## 7. 面试讲法

可以这样总结：

```text
我把原始保单明细处理流程包装成一个可复现的数据分析项目：先用 Pandas 完成数据清洗和指标计算，再用 OpenPyXL 和 Matplotlib 生成 Excel 报告与经营看板。后续我加入了 RAG 问答、Pandas 精确数据查询、TF-IDF 向量检索、问答日志和标准评估集，并基于指标结果自动生成经营摘要和智能解读，形成从数据处理、报表生成到智能问答和评估验证的完整闭环。
```
