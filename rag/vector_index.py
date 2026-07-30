from __future__ import annotations

import pickle
import re
from pathlib import Path

from sklearn.feature_extraction.text import TfidfVectorizer


PROJECT_ROOT = Path(__file__).resolve().parents[1]
KNOWLEDGE_DIR = PROJECT_ROOT / "rag" / "knowledge_base"
OUTPUT_DIR = PROJECT_ROOT / "output"
VECTOR_INDEX_PATH = PROJECT_ROOT / "rag" / "vector_index.pkl"


def split_markdown(text: str) -> list[str]:
    parts = re.split(r"\n(?=##?\s+)", text)
    return [part.strip() for part in parts if part.strip()]


def load_documents() -> list[dict[str, str]]:
    documents: list[dict[str, str]] = []
    for path in sorted(KNOWLEDGE_DIR.glob("*.md")):
        text = path.read_text(encoding="utf-8")
        for index, chunk in enumerate(split_markdown(text), start=1):
            documents.append(
                {
                    "id": f"{path.stem}-{index}",
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
    return documents


def build_vector_index() -> dict[str, object]:
    documents = load_documents()
    if not documents:
        raise ValueError("没有找到可用于构建向量索引的知识库文档。")

    texts = [document["text"] for document in documents]
    vectorizer = TfidfVectorizer(
        analyzer="char_wb",
        ngram_range=(2, 4),
        lowercase=True,
    )
    matrix = vectorizer.fit_transform(texts)

    index = {
        "documents": documents,
        "vectorizer": vectorizer,
        "matrix": matrix,
    }
    VECTOR_INDEX_PATH.write_bytes(pickle.dumps(index))
    return index


def main() -> None:
    index = build_vector_index()
    documents = index["documents"]
    print(f"已构建 TF-IDF 向量索引：{VECTOR_INDEX_PATH}")
    print(f"共写入 {len(documents)} 个文本片段")


if __name__ == "__main__":
    main()
