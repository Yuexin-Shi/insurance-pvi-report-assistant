from __future__ import annotations

import re

from intent_classifier import intent_label


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


def generate_data_answer(question: str, intent: str, data_answer: str) -> str:
    return "\n".join(
        [
            f"问题：{question}",
            f"识别意图：{intent_label(intent)}（{intent}）",
            "",
            "回答：",
            data_answer,
        ]
    )


def generate_rag_answer(
    question: str,
    intent: str,
    result: dict[str, str | float],
) -> str:
    source = str(result["source"])
    score = float(result["score"])
    text = str(result["text"])
    conclusion = first_sentence(text)
    detail = clean_retrieved_text(text)

    lines = [
        f"问题：{question}",
        f"识别意图：{intent_label(intent)}（{intent}）",
        "",
        "回答：",
        conclusion,
    ]
    if detail and detail != conclusion:
        lines.extend(["", "补充说明：", detail])
    lines.extend(["", f"来源：{source}｜相关度：{score:.2f}"])
    return "\n".join(lines)


def generate_not_found_answer(question: str, intent: str) -> str:
    return "\n".join(
        [
            f"问题：{question}",
            f"识别意图：{intent_label(intent)}（{intent}）",
            "",
            "回答：",
            "没有在知识库或数据表中找到明显相关内容。",
        ]
    )
