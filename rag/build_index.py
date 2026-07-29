from __future__ import annotations

import json
import re
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
KNOWLEDGE_DIR = PROJECT_ROOT / "rag" / "knowledge_base"
OUTPUT_DIR = PROJECT_ROOT / "output"
INDEX_PATH = PROJECT_ROOT / "rag" / "index.json"


def split_markdown(text: str) -> list[str]:
    parts = re.split(r"\n(?=##?\s+)", text)
    chunks = []
    for part in parts:
        cleaned = part.strip()
        if cleaned:
            chunks.append(cleaned)
    return chunks


def build_index() -> list[dict[str, str]]:
    documents = []
    for path in sorted(KNOWLEDGE_DIR.glob("*.md")):
        text = path.read_text(encoding="utf-8")
        for i, chunk in enumerate(split_markdown(text), start=1):
            documents.append(
                {
                    "id": f"{path.stem}-{i}",
                    "source": str(path.relative_to(PROJECT_ROOT)),
                    "text": chunk,
                }
            )
    for path in sorted(OUTPUT_DIR.glob("经营摘要_*.txt")):
        text = path.read_text(encoding="utf-8").strip()
        if text:
            documents.append(
                {
                    "id": path.stem,
                    "source": str(path.relative_to(PROJECT_ROOT)),
                    "text": f"# 经营摘要\n\n{text}",
                }
            )
    INDEX_PATH.write_text(
        json.dumps(documents, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return documents


def main() -> None:
    documents = build_index()
    print(f"已构建知识库索引：{INDEX_PATH}")
    print(f"共写入 {len(documents)} 个文本片段")


if __name__ == "__main__":
    main()
