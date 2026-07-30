from __future__ import annotations


DATA_QUERY_INTENTS = {
    "pvi_total_query",
    "agent_ranking_query",
    "active_agent_query",
    "product_ranking_query",
    "company_ranking_query",
    "summary_query",
}


INTENT_LABELS = {
    "metric_definition": "指标解释",
    "business_rule": "业务规则/竞赛方案",
    "summary_query": "经营摘要/经营结果",
    "pvi_total_query": "PVI金额查询",
    "agent_ranking_query": "代理人排名查询",
    "active_agent_query": "出单名单查询",
    "product_ranking_query": "产品/险种排名查询",
    "company_ranking_query": "保险公司排名查询",
    "unknown": "未知意图",
}


def contains_any(text: str, keywords: list[str]) -> bool:
    lowered = text.lower()
    return any(keyword.lower() in lowered for keyword in keywords)


def classify_intent(question: str) -> str:
    q = question.strip()

    if contains_any(
        q,
        [
            "出单名单",
            "都谁出单",
            "谁出单",
            "哪些人出单",
            "出单人员",
            "出单人数",
            "出单代理人",
            "几个人出单",
        ],
    ):
        return "active_agent_query"

    if contains_any(q, ["产品", "险种"]):
        return "product_ranking_query"

    if contains_any(q, ["保险公司", "公司贡献", "哪家公司", "哪家保险"]):
        return "company_ranking_query"

    if contains_any(q, ["最高", "排名", "top", "前五", "top5", "最厉害"]) and contains_any(q, ["代理人", "业务员", "PVI", "pvi"]):
        return "agent_ranking_query"

    if contains_any(q, ["今日PVI", "今天PVI", "本月PVI", "今月PVI", "PVI是多少", "pvi是多少", "pvi多少"]):
        return "pvi_total_query"

    if contains_any(q, ["看板", "报告", "输出", "文字经营摘要"]) and contains_any(q, ["包含", "内容", "哪些", "生成", "有什么"]):
        return "business_rule"

    if contains_any(q, ["经营摘要", "经营情况", "报告结果", "当前结果"]):
        return "summary_query"

    if contains_any(q, ["规则", "方案", "标准", "奖励", "达标"]):
        return "business_rule"

    if contains_any(q, ["是什么", "什么意思", "怎么定义", "含义", "解释", "怎么算", "怎么计算", "怎么统计", "咋算"]):
        return "metric_definition"

    return "unknown"


def intent_label(intent: str) -> str:
    return INTENT_LABELS.get(intent, intent)


def should_use_data_query(intent: str) -> bool:
    return intent in DATA_QUERY_INTENTS
