from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
LOG_DIR = PROJECT_ROOT / "logs"
LOG_PATH = LOG_DIR / "qa_log.jsonl"


def log_qa(question: str, intent: str, route: str, answer: str) -> None:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    record = {
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "question": question,
        "intent": intent,
        "route": route,
        "answer": answer,
    }
    with LOG_PATH.open("a", encoding="utf-8") as file:
        file.write(json.dumps(record, ensure_ascii=False) + "\n")
