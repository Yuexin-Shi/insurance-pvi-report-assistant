from __future__ import annotations

import pickle
import re
import sys
from pathlib import Path

from sklearn.metrics.pairwise import cosine_similarity


PROJECT_ROOT = Path(__file__).resolve().parents[1]
VECTOR_INDEX_PATH = PROJECT_ROOT / "rag" / "vector_index.pkl"


def load_vector_index() -> dict[str, object]:
    if not VECTOR_INDEX_PATH.exists():
        raise FileNotFoundError(
            f"没有找到向量索引：{VECTOR_INDEX_PATH}\n"
            "请先运行：python rag/vector_index.py"
        )
    return pickle.loads(VECTOR_INDEX_PATH.read_bytes())


def clean_retrieved_text(text: str) -> str:
    lines = []
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        if stripped.startswith("#"):
            continue
        lines.append(stripped)
    return "\n".join(lines)


def first_sentence(text: str) -> str:
    cleaned = clean_retrieved_text(text)
    parts = re.split(r"(?<=[。！？])", cleaned, maxsplit=1)
    return parts[0].strip() if parts and parts[0].strip() else cleaned


def first_heading(text: str) -> str:
    for line in text.splitlines():
        stripped = line.strip()
        if stripped.startswith("#"):
            return stripped.lstrip("#").strip().lower()
    return ""


def rerank_score(question: str, text: str, base_score: float) -> float:
    heading = first_heading(text)
    compact_question = question.replace(" ", "").lower()
    compact_heading = heading.replace(" ", "")
    score = base_score

    heading_terms = [
        "pvi",
        "净pvi",
        "每日净pvi",
        "本月累计pvi",
        "目标达成率",
        "目标缺口",
        "月末预测",
        "正常业务预测",
        "最大一笔pvi",
    ]
    for term in heading_terms:
        if term in compact_question and compact_heading == term:
            score += 0.8

    if "是什么" in question or "什么意思" in question:
        if compact_heading == compact_question.replace("是什么", "").replace("什么意思", ""):
            score += 1.0

    if "达成率" in compact_question and compact_heading == "目标达成率":
        score += 1.0

    return score


def search(question: str, top_k: int = 3) -> list[dict[str, object]]:
    index = load_vector_index()
    documents = index["documents"]
    vectorizer = index["vectorizer"]
    matrix = index["matrix"]

    query_vector = vectorizer.transform([question])
    scores = cosine_similarity(query_vector, matrix).ravel()
    reranked = []
    for item_index, base_score in enumerate(scores):
        document = documents[item_index]
        score = rerank_score(question, str(document["text"]), float(base_score))
        reranked.append((item_index, score))
    reranked.sort(key=lambda item: item[1], reverse=True)

    results = []
    for item_index, score in reranked[:top_k]:
        if score <= 0:
            continue
        document = documents[item_index]
        results.append({**document, "score": score})
    return results


def answer(question: str) -> str:
    results = search(question, top_k=1)
    if not results:
        return "\n".join(
            [
                f"问题：{question}",
                "检索方式：TF-IDF 向量检索",
                "",
                "回答：",
                "没有在向量知识库中找到明显相关内容。",
            ]
        )

    result = results[0]
    text = str(result["text"])
    detail = clean_retrieved_text(text)
    conclusion = first_sentence(text)

    lines = [
        f"问题：{question}",
        "检索方式：TF-IDF 向量检索",
        f"来源：{result['source']}｜相似度：{float(result['score']):.2f}",
        "",
        "回答：",
        conclusion,
    ]
    if detail and detail != conclusion:
        lines.extend(["", "补充说明：", detail])
    return "\n".join(lines)


def main() -> None:
    if len(sys.argv) < 2:
        print('用法：python rag/vector_qa.py "PVI 是什么"')
        raise SystemExit(1)
    question = " ".join(sys.argv[1:])
    print(answer(question))


if __name__ == "__main__":
    main()
