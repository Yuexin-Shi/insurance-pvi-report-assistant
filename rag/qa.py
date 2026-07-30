from __future__ import annotations

import json
import math
import re
import sys
from collections import Counter
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
INDEX_PATH = PROJECT_ROOT / "rag" / "index.json"
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from data_query import query_data
from answer_generator import generate_data_answer, generate_not_found_answer, generate_rag_answer
from intent_classifier import classify_intent, should_use_data_query
from logger import log_qa


def tokenize(text: str) -> list[str]:
    tokens = re.findall(r"[A-Za-z0-9_]+", text.lower())
    chinese_blocks = re.findall(r"[\u4e00-\u9fff]+", text)
    for block in chinese_blocks:
        if len(block) <= 2:
            tokens.append(block)
        else:
            tokens.extend(block[i : i + 2] for i in range(len(block) - 1))
    return tokens


def vectorize(text: str) -> Counter[str]:
    return Counter(tokenize(text))


def cosine_similarity(left: Counter[str], right: Counter[str]) -> float:
    common = set(left) & set(right)
    dot = sum(left[token] * right[token] for token in common)
    left_norm = math.sqrt(sum(value * value for value in left.values()))
    right_norm = math.sqrt(sum(value * value for value in right.values()))
    if left_norm == 0 or right_norm == 0:
        return 0.0
    return dot / (left_norm * right_norm)


def load_index() -> list[dict[str, str]]:
    if not INDEX_PATH.exists():
        raise FileNotFoundError(
            f"没有找到知识库索引：{INDEX_PATH}\n"
            "请先运行：python rag/build_index.py"
        )
    return json.loads(INDEX_PATH.read_text(encoding="utf-8"))


def first_heading(text: str) -> str:
    for line in text.splitlines():
        stripped = line.strip()
        if stripped.startswith("#"):
            return stripped.lstrip("#").strip().lower()
    return ""


def query_keywords(question: str) -> set[str]:
    keywords = set(re.findall(r"[A-Za-z0-9_]+", question.lower()))
    for word in ["PVI", "MDRT"]:
        if word.lower() in question.lower():
            keywords.add(word.lower())
    chinese_terms = [
        "净PVI",
        "净 PVI",
        "每日净PVI",
        "每日净 PVI",
        "本月累计PVI",
        "本月累计 PVI",
        "最大一笔PVI",
        "最大一笔 PVI",
        "今日PVI",
        "今日新增",
        "目标达成率",
        "达成率",
        "目标缺口",
        "每日净",
        "本月累计",
        "月末预测",
        "正常业务预测",
        "阳光",
        "信泰",
        "竞赛方案",
        "经营摘要",
    ]
    for term in chinese_terms:
        if term in question:
            keywords.add(term.lower())
            keywords.add(term.replace(" ", "").lower())
    return keywords


def is_definition_question(question: str) -> bool:
    if is_result_question(question):
        return False
    return any(pattern in question for pattern in ["是什么", "什么意思", "含义", "定义", "解释", "怎么算", "怎么计算", "咋算"])


def is_result_question(question: str) -> bool:
    return any(
        pattern in question
        for pattern in [
            "今日",
            "今天",
            "本月",
            "具体",
            "当前",
            "结果",
        ]
    )


def rerank_score(question: str, item: dict[str, str], base_score: float) -> float:
    score = base_score
    heading = first_heading(item["text"])
    source = item["source"]
    keywords = query_keywords(question)

    for keyword in keywords:
        compact_heading = heading.replace(" ", "")
        compact_keyword = keyword.replace(" ", "")
        if heading == keyword or compact_heading == compact_keyword:
            score += 1.0
        elif keyword in heading or compact_keyword in compact_heading:
            score += 0.35

    if is_definition_question(question):
        if "指标口径说明" in source:
            score += 0.35
        if heading in keywords:
            score += 0.5
        if any(word in item["text"] for word in ["是本项目用于", "指", "=", "定义"]):
            score += 0.15
        if "净pvi" in question.replace(" ", "").lower():
            if heading.replace(" ", "") == "净pvi":
                score += 1.0
            elif heading == "pvi":
                score -= 0.4
        compact_question = question.replace(" ", "").lower()
        if "达成率" in compact_question:
            if heading.replace(" ", "") == "目标达成率":
                score += 1.5
            elif heading == "pvi":
                score -= 0.5
        for term in ["每日净pvi", "本月累计pvi", "最大一笔pvi"]:
            if term in compact_question:
                if heading.replace(" ", "") == term:
                    score += 1.0
                elif heading == "pvi":
                    score -= 0.4

    if "规则" in question and "竞赛方案说明" in source:
        score += 0.25

    if "文字经营摘要" in question and "报告生成规则" in source:
        score += 1.5

    if is_result_question(question) and "output/经营摘要" in source:
        score += 2.0

    return score


def search(question: str, top_k: int = 1) -> list[dict[str, str | float]]:
    query_vector = vectorize(question)
    results = []
    for item in load_index():
        base_score = cosine_similarity(query_vector, vectorize(item["text"]))
        score = rerank_score(question, item, base_score)
        if score > 0:
            results.append({**item, "score": score})
    results.sort(key=lambda item: item["score"], reverse=True)
    return results[:top_k]


def answer(question: str) -> str:
    intent = classify_intent(question)

    if should_use_data_query(intent):
        data_answer = query_data(question)
        if data_answer:
            response = generate_data_answer(question, intent, data_answer)
            log_qa(question, intent, "pandas", response)
            return response

    if intent == "unknown":
        data_answer = query_data(question)
        if data_answer:
            response = generate_data_answer(question, intent, data_answer)
            log_qa(question, intent, "pandas", response)
            return response

    results = search(question)
    if not results:
        response = generate_not_found_answer(question, intent)
        log_qa(question, intent, "unknown", response)
        return response

    response = generate_rag_answer(question, intent, results[0])
    log_qa(question, intent, "rag", response)
    return response


def split_questions(text: str) -> list[str]:
    parts = re.split(r"[；;？?，,]\s*|(?:以及|还有|并且)", text)
    questions = [part.strip() for part in parts if part.strip()]
    return questions or [text]


def main() -> None:
    if len(sys.argv) < 2:
        print('用法：python rag/qa.py "PVI 是什么"')
        raise SystemExit(1)
    question = " ".join(sys.argv[1:])
    questions = split_questions(question)
    if len(questions) == 1:
        print(answer(questions[0]))
        return
    for index, item in enumerate(questions, start=1):
        if index > 1:
            print("\n" + "-" * 60 + "\n")
        print(answer(item))


if __name__ == "__main__":
    main()
