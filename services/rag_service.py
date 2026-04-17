from __future__ import annotations

from collections import Counter, defaultdict
from pathlib import Path
import logging
import pickle
import re
import uuid

import faiss


class LocalRAGService:
    """Hybrid local RAG with vector search, lexical fallback, and lightweight reranking."""

    SCHEMA_VERSION = "2"

    STOPWORDS = {
        "的", "了", "和", "是", "在", "与", "及", "对", "把", "并", "中", "上", "下", "按", "将", "为", "到",
        "the", "and", "for", "with", "from", "this", "that", "into", "using",
    }

    def __init__(self, vector_dir: Path):
        self.vector_dir = vector_dir
        self.index_path = vector_dir / "knowledge.faiss"
        self.meta_path = vector_dir / "metadata.pkl"
        self.fallback_meta_path = vector_dir / "metadata_fallback.pkl"
        self.schema_path = vector_dir / "schema_version.txt"
        self.embedding_model = None
        self.dimension = 384
        self._embedding_load_attempted = False
        self._logger = logging.getLogger(__name__)
        self.global_session_id = "__global_kb__"

        self._ensure_store_schema()

        self.index = self._load_or_create_index()
        self.metadata = self._load_metadata()
        self.fallback_metadata = self._load_fallback_metadata()

    def _ensure_store_schema(self):
        current = self.schema_path.read_text(encoding="utf-8").strip() if self.schema_path.exists() else ""
        if current == self.SCHEMA_VERSION:
            return
        for path in (self.index_path, self.meta_path, self.fallback_meta_path):
            if path.exists():
                path.unlink()
        self.schema_path.write_text(self.SCHEMA_VERSION, encoding="utf-8")

    def _ensure_embedding_model(self):
        if self.embedding_model is not None:
            return True
        if self._embedding_load_attempted:
            return False
        self._embedding_load_attempted = True
        try:
            from sentence_transformers import SentenceTransformer

            self.embedding_model = SentenceTransformer("sentence-transformers/all-MiniLM-L6-v2")
            self.dimension = self.embedding_model.get_sentence_embedding_dimension()
            return True
        except BaseException as error:  # pragma: no cover
            self._logger.warning("Embedding model unavailable, RAG will use lexical fallback temporarily: %s", error)
            self.embedding_model = None
            return False

    def _load_or_create_index(self):
        if self.index_path.exists():
            index = faiss.read_index(str(self.index_path))
            self.dimension = index.d
            return index
        return faiss.IndexFlatIP(self.dimension)

    def _load_metadata(self):
        if self.meta_path.exists():
            with open(self.meta_path, "rb") as file:
                return pickle.load(file)
        return []

    def _load_fallback_metadata(self):
        if self.fallback_meta_path.exists():
            with open(self.fallback_meta_path, "rb") as file:
                return pickle.load(file)
        return []

    def _persist(self):
        faiss.write_index(self.index, str(self.index_path))
        with open(self.meta_path, "wb") as file:
            pickle.dump(self.metadata, file)
        with open(self.fallback_meta_path, "wb") as file:
            pickle.dump(self.fallback_metadata, file)

    def _tokenize(self, text: str):
        raw_tokens = re.split(r"[^\w\u4e00-\u9fff]+", (text or "").lower())
        return [token for token in raw_tokens if token and token not in self.STOPWORDS and len(token) > 1]

    def _extract_keywords(self, text: str, top_k: int = 8):
        counter = Counter(self._tokenize(text))
        return [token for token, _ in counter.most_common(top_k)]

    def _is_heading(self, line: str):
        stripped = line.strip()
        if not stripped or len(stripped) > 32:
            return False
        if re.match(r"^(第[一二三四五六七八九十百0-9]+[章节部分讲]|[0-9一二三四五六七八九十]+[、.])", stripped):
            return True
        return not re.search(r"[。！？；;,.，]", stripped)

    def _split_sections(self, text: str):
        lines = [line.strip() for line in str(text or "").splitlines() if line.strip()]
        if not lines:
            return []
        sections = []
        current_title = "正文"
        buffer = []
        for line in lines:
            if self._is_heading(line):
                if buffer:
                    sections.append({"title": current_title, "text": "\n".join(buffer)})
                    buffer = []
                current_title = line[:30]
            else:
                buffer.append(line)
        if buffer:
            sections.append({"title": current_title, "text": "\n".join(buffer)})
        return sections or [{"title": "正文", "text": "\n".join(lines)}]

    def _slice_text(self, text: str, chunk_size: int = 360, overlap: int = 60):
        cleaned = " ".join(str(text or "").split())
        if not cleaned:
            return []
        chunks = []
        start = 0
        while start < len(cleaned):
            end = min(len(cleaned), start + chunk_size)
            chunks.append(cleaned[start:end])
            if end == len(cleaned):
                break
            start = max(0, end - overlap)
        return chunks

    def _infer_chunk_type(self, title: str, text: str):
        sample = f"{title} {text[:180]}"
        if any(keyword in sample for keyword in ["案例", "例题", "情境", "应用", "任务"]):
            return "case"
        if any(keyword in sample for keyword in ["风格", "版式", "结构", "排版"]):
            return "style"
        if any(keyword in sample for keyword in ["定义", "概念", "原理", "方法", "公式", "定理", "知识"]):
            return "knowledge"
        return "overview"

    def _prepare_chunks(self, text: str, source_name: str, source_type: str, extra_meta: dict | None = None):
        extra_meta = extra_meta or {}
        summary = str(extra_meta.get("summary", "")).strip()
        sections = self._split_sections(text)
        prepared = []
        for section in sections:
            chunk_type = self._infer_chunk_type(section["title"], section["text"])
            for piece in self._slice_text(section["text"]):
                content = f"标题：{section['title']}\n类型：{chunk_type}\n内容：{piece}"
                keywords = self._extract_keywords(f"{section['title']} {piece}")
                prepared.append(
                    {
                        "chunk_id": uuid.uuid4().hex,
                        "title": section["title"],
                        "chunk_type": chunk_type,
                        "keywords": keywords,
                        "summary": summary[:220],
                        "text": content[:900],
                        "source_name": source_name,
                        "source_type": source_type,
                    }
                )
        if not prepared:
            overview = " ".join(str(text or "").split())[:800]
            prepared.append(
                {
                    "chunk_id": uuid.uuid4().hex,
                    "title": extra_meta.get("title", source_name),
                    "chunk_type": "overview",
                    "keywords": self._extract_keywords(overview),
                    "summary": summary[:220],
                    "text": overview,
                    "source_name": source_name,
                    "source_type": source_type,
                }
            )
        return prepared

    def _embed(self, texts: list[str]):
        if not self._ensure_embedding_model():
            return None
        return self.embedding_model.encode(
            texts,
            convert_to_numpy=True,
            normalize_embeddings=True,
            show_progress_bar=False,
        ).astype("float32")

    def _chunk_key(self, item: dict):
        return (item.get("session_id"), item.get("source_name"), item.get("title"), item.get("text"))

    def _has_vector_chunk(self, session_id: str, source_name: str, title: str, text: str):
        target_key = (session_id, source_name, title, text)
        return any(self._chunk_key(item) == target_key for item in self.metadata)

    def _allowed_ids(self, session_id: str, include_global: bool):
        allowed = {session_id}
        if include_global:
            allowed.add(self.global_session_id)
        return allowed

    def add_document(self, session_id: str, source_name: str, text: str, source_type: str, extra_meta: dict | None = None):
        chunks = self._prepare_chunks(text, source_name, source_type, extra_meta=extra_meta)
        if not chunks:
            return 0
        embeddings = self._embed([chunk["text"] for chunk in chunks])
        if embeddings is None:
            return self.add_document_fallback(session_id, source_name, text, source_type, extra_meta=extra_meta)

        self.index.add(embeddings)
        for chunk in chunks:
            self.metadata.append({"session_id": session_id, **chunk})
        self._persist()
        return len(chunks)

    def add_document_fallback(self, session_id: str, source_name: str, text: str, source_type: str, extra_meta: dict | None = None):
        chunks = self._prepare_chunks(text, source_name, source_type, extra_meta=extra_meta)
        if not chunks:
            return 0
        for chunk in chunks:
            self.fallback_metadata.append({"session_id": session_id, **chunk})
        self._persist()
        return len(chunks)

    def add_global_document(self, source_name: str, text: str, source_type: str, extra_meta: dict | None = None):
        return self.add_document(self.global_session_id, source_name, text, source_type, extra_meta=extra_meta)

    def add_global_document_fallback(self, source_name: str, text: str, source_type: str, extra_meta: dict | None = None):
        return self.add_document_fallback(self.global_session_id, source_name, text, source_type, extra_meta=extra_meta)

    def upgrade_scope_to_vector(self, session_id: str):
        if not self._ensure_embedding_model():
            return 0
        candidates = []
        for item in self.fallback_metadata:
            if item.get("session_id") != session_id:
                continue
            if self._has_vector_chunk(session_id, item.get("source_name", ""), item.get("title", ""), item.get("text", "")):
                continue
            candidates.append(item)
        if not candidates:
            return 0

        embeddings = self._embed([item["text"] for item in candidates])
        if embeddings is None:
            return 0
        self.index.add(embeddings)
        self.metadata.extend(candidates)
        self._persist()
        return len(candidates)

    def _lexical_score(self, query_tokens: set[str], item: dict):
        bag = " ".join(
            [
                item.get("title", ""),
                " ".join(item.get("keywords", [])),
                item.get("summary", ""),
                item.get("text", ""),
                item.get("source_name", ""),
            ]
        )
        text_tokens = set(self._tokenize(bag))
        if not text_tokens:
            return 0.0
        overlap = len(query_tokens & text_tokens)
        if overlap <= 0:
            return 0.0
        return round(overlap / max(len(query_tokens), 1), 4)

    def _chunk_type_bonus(self, item: dict, query_hints: dict):
        chunk_type = item.get("chunk_type", "overview")
        bonus = 0.0
        desired = set(query_hints.get("desired_chunk_types", []))
        if chunk_type in desired:
            bonus += 0.08
        if item.get("scope") == "session":
            bonus += 0.06
        elif item.get("scope") == "global":
            bonus += 0.03
        return bonus

    def search(self, session_id: str, query: str, top_k: int = 5, include_global: bool = True, query_hints: dict | None = None):
        query = str(query or "").strip()
        query_hints = query_hints or {}
        if not query:
            return []

        allowed_ids = self._allowed_ids(session_id, include_global)
        query_text = " ".join([query] + query_hints.get("keywords", [])[:6]).strip()
        query_tokens = set(self._tokenize(query_text))

        ranked = {}
        vector_results = []

        if self.index.ntotal > 0:
            query_embedding = self._embed([query_text])
            if query_embedding is not None:
                scores, indices = self.index.search(query_embedding, min(max(top_k * 6, 20), self.index.ntotal))
                for score, index_id in zip(scores[0], indices[0]):
                    if index_id < 0:
                        continue
                    item = self.metadata[index_id]
                    if item.get("session_id") not in allowed_ids:
                        continue
                    key = self._chunk_key(item)
                    vector_results.append((key, float(score), item))

        for key, score, item in vector_results:
            ranked[key] = {
                **item,
                "scope": "global" if item.get("session_id") == self.global_session_id else "session",
                "vector_score": round(score, 4),
                "lexical_score": 0.0,
            }

        lexical_pool = []
        lexical_pool.extend([item for item in self.metadata if item.get("session_id") in allowed_ids])
        lexical_pool.extend([item for item in self.fallback_metadata if item.get("session_id") in allowed_ids])

        for item in lexical_pool:
            lexical_score = self._lexical_score(query_tokens, item)
            if lexical_score <= 0:
                continue
            key = self._chunk_key(item)
            if key not in ranked:
                ranked[key] = {
                    **item,
                    "scope": "global" if item.get("session_id") == self.global_session_id else "session",
                    "vector_score": 0.0,
                    "lexical_score": lexical_score,
                }
            else:
                ranked[key]["lexical_score"] = max(ranked[key]["lexical_score"], lexical_score)

        merged = []
        for item in ranked.values():
            final_score = (
                item["vector_score"] * 0.68
                + item["lexical_score"] * 0.32
                + self._chunk_type_bonus(item, query_hints)
            )
            merged.append(
                {
                    "score": round(final_score, 4),
                    "source_name": item["source_name"],
                    "source_type": item["source_type"],
                    "scope": item["scope"],
                    "title": item.get("title", ""),
                    "chunk_type": item.get("chunk_type", "overview"),
                    "keywords": item.get("keywords", [])[:6],
                    "summary": item.get("summary", ""),
                    "text": item["text"][:280],
                }
            )

        merged.sort(key=lambda item: item["score"], reverse=True)

        diversified = []
        per_source = defaultdict(int)
        for item in merged:
            if per_source[item["source_name"]] >= 2:
                continue
            diversified.append(item)
            per_source[item["source_name"]] += 1
            if len(diversified) >= top_k:
                break
        return diversified

    def get_stats(self, session_id: str):
        vector_chunks = sum(1 for item in self.metadata if item.get("session_id") == session_id)
        fallback_chunks = sum(1 for item in self.fallback_metadata if item.get("session_id") == session_id)
        indexed_files = {item["source_name"] for item in self.metadata if item.get("session_id") == session_id}
        indexed_files.update({item["source_name"] for item in self.fallback_metadata if item.get("session_id") == session_id})
        return {
            "indexed_chunks": vector_chunks + fallback_chunks,
            "vector_chunks": vector_chunks,
            "fallback_chunks": fallback_chunks,
            "indexed_files": sorted(indexed_files),
        }

    def get_global_stats(self):
        return self.get_stats(self.global_session_id)
