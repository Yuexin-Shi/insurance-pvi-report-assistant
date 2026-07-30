from __future__ import annotations

import json
import os
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
RAG_DIR = PROJECT_ROOT / "rag"
if str(RAG_DIR) not in sys.path:
    sys.path.insert(0, str(RAG_DIR))

os.environ.setdefault("PVI_USE_SAMPLE", "1")

from intent_classifier import classify_intent
from qa import answer, split_questions


QUESTIONS_PATH = PROJECT_ROOT / "eval" / "questions.json"


def load_questions() -> list[dict[str, object]]:
    return json.loads(QUESTIONS_PATH.read_text(encoding="utf-8"))


def evaluate_case(case: dict[str, object]) -> dict[str, object]:
    question = str(case["question"])
    expected_intent = str(case.get("expected_intent", ""))
    expected_intents = [str(item) for item in case.get("expected_intents", [])]
    expected_keywords = [str(item) for item in case.get("expected_keywords", [])]

    if expected_intents:
        split_items = split_questions(question)
        actual_intents = [classify_intent(item) for item in split_items]
        actual_intent = ",".join(actual_intents)
        intent_ok = actual_intents == expected_intents
        response = "\n".join(answer(item) for item in split_items)
    else:
        actual_intent = classify_intent(question)
        intent_ok = actual_intent == expected_intent
        response = answer(question)

    missing_keywords = [keyword for keyword in expected_keywords if keyword not in response]
    keyword_ok = not missing_keywords

    return {
        "question": question,
        "expected_intent": expected_intent,
        "expected_intents": expected_intents,
        "actual_intent": actual_intent,
        "intent_ok": intent_ok,
        "keyword_ok": keyword_ok,
        "missing_keywords": missing_keywords,
    }


def main() -> None:
    cases = load_questions()
    results = [evaluate_case(case) for case in cases]

    passed = 0
    for result in results:
        ok = result["intent_ok"] and result["keyword_ok"]
        passed += int(ok)
        status = "PASS" if ok else "FAIL"
        print(f"[{status}] {result['question']}")
        expected = result["expected_intents"] or result["expected_intent"]
        print(f"  意图：{result['actual_intent']} / 预期：{expected}")
        if result["missing_keywords"]:
            print(f"  缺少关键词：{', '.join(result['missing_keywords'])}")

    total = len(results)
    rate = passed / total if total else 0
    print(f"\n评估结果：{passed}/{total} 通过，通过率 {rate:.0%}")


if __name__ == "__main__":
    main()
