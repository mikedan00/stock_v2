"""
rag/rag_engine.py
경량 RAG 엔진 — 외부 벡터 DB 없이 TF-IDF 기반 유사도 검색
(ChromaDB/sentence-transformers 설치 실패 환경 대비)
"""
from __future__ import annotations

import math
import re
from collections import defaultdict
from typing import Optional


# ── 텍스트 전처리 ─────────────────────────────────────────────────────────────

def _tokenize(text: str) -> list[str]:
    text = text.lower()
    # 한글 음절 + 영문 단어
    tokens = re.findall(r"[가-힣]+|[a-z0-9]+", text)
    return tokens


def _chunk_text(text: str, chunk_size: int = 400, overlap: int = 80) -> list[str]:
    """텍스트를 chunk_size 글자 단위로 분할 (overlap 포함)."""
    chunks = []
    start = 0
    while start < len(text):
        end = start + chunk_size
        chunks.append(text[start:end])
        start += chunk_size - overlap
    return [c.strip() for c in chunks if c.strip()]


# ── TF-IDF 인덱스 ─────────────────────────────────────────────────────────────

class RAGEngine:
    """
    인메모리 TF-IDF RAG 엔진.
    add_document()로 문서 추가 → query()로 관련 청크 검색.
    """

    def __init__(self):
        self.chunks: list[str] = []          # 청크 텍스트
        self.metadata: list[dict] = []       # 청크별 메타데이터
        self._tf: list[dict[str, float]] = []
        self._df: dict[str, int] = defaultdict(int)
        self._built = False

    def clear(self):
        self.__init__()

    def add_document(self, text: str, source: str = "", doc_type: str = ""):
        """문서를 청크로 분할해 인덱스에 추가."""
        chunks = _chunk_text(text)
        for i, chunk in enumerate(chunks):
            self.chunks.append(chunk)
            self.metadata.append({"source": source, "doc_type": doc_type, "chunk_idx": i})
        self._built = False   # 재빌드 필요 표시

    def _build_index(self):
        """TF-IDF 인덱스 구축."""
        self._tf = []
        self._df = defaultdict(int)
        token_sets = []
        for chunk in self.chunks:
            tokens = _tokenize(chunk)
            tf: dict[str, float] = defaultdict(float)
            for t in tokens:
                tf[t] += 1
            total = max(len(tokens), 1)
            self._tf.append({t: v / total for t, v in tf.items()})
            unique = set(tokens)
            token_sets.append(unique)
            for t in unique:
                self._df[t] += 1
        self._built = True
        self._n = len(self.chunks)

    def _tfidf_vec(self, tf_dict: dict[str, float]) -> dict[str, float]:
        n = self._n or 1
        return {
            t: v * math.log((n + 1) / (self._df.get(t, 0) + 1) + 1)
            for t, v in tf_dict.items()
        }

    def _cosine(self, a: dict[str, float], b: dict[str, float]) -> float:
        keys = set(a) & set(b)
        if not keys:
            return 0.0
        dot = sum(a[k] * b[k] for k in keys)
        na = math.sqrt(sum(v * v for v in a.values()))
        nb = math.sqrt(sum(v * v for v in b.values()))
        return dot / (na * nb + 1e-9)

    def query(self, question: str, top_k: int = 5) -> list[dict]:
        """질문과 가장 유사한 청크 top_k개 반환."""
        if not self.chunks:
            return []
        if not self._built:
            self._build_index()

        q_tokens = _tokenize(question)
        q_tf: dict[str, float] = defaultdict(float)
        for t in q_tokens:
            q_tf[t] += 1
        total = max(len(q_tokens), 1)
        q_tf = {t: v / total for t, v in q_tf.items()}
        q_vec = self._tfidf_vec(q_tf)

        scores = []
        for i, tf in enumerate(self._tf):
            vec = self._tfidf_vec(tf)
            score = self._cosine(q_vec, vec)
            scores.append((score, i))

        scores.sort(reverse=True)
        results = []
        for score, idx in scores[:top_k]:
            if score > 0:
                results.append({
                    "text": self.chunks[idx],
                    "score": round(score, 4),
                    "source": self.metadata[idx].get("source", ""),
                    "doc_type": self.metadata[idx].get("doc_type", ""),
                })
        return results

    def get_context_for_stock(self, stock_name: str, ticker: str, top_k: int = 6) -> str:
        """특정 종목 관련 컨텍스트를 RAG로 검색해 텍스트로 반환."""
        query = f"{stock_name} {ticker} 주가 분석 투자 전망"
        results = self.query(query, top_k=top_k)
        if not results:
            return ""
        parts = []
        for r in results:
            src = f"[출처: {r['source']}] " if r["source"] else ""
            parts.append(f"{src}{r['text']}")
        return "\n\n---\n\n".join(parts)

    @property
    def doc_count(self) -> int:
        return len(self.chunks)


# 전역 싱글턴 (Streamlit 세션 내에서 재사용)
_global_rag = RAGEngine()


def get_rag() -> RAGEngine:
    return _global_rag


def reset_rag():
    _global_rag.clear()
