"""
Single-agent RAG baseline: retrieve from Chroma, ask llama3.1 to judge.

The prompt enforces a strict JSON output: {"is_hallucination": bool, "evidence": str}.
This is the naive pipeline the multi-agent system is compared against.
"""
import json
import os
import re

from langchain_community.chat_models import ChatOllama
from langchain_community.embeddings import HuggingFaceEmbeddings
from langchain_community.vectorstores import Chroma
from langchain_core.messages import SystemMessage

CHROMA_PERSIST_DIR = "./chroma_db"
OLLAMA_MODEL = os.getenv("OLLAMA_MODEL", "llama3.1")
OLLAMA_BASE_URL = os.getenv("OLLAMA_BASE_URL", "http://localhost:11434")

_embeddings = None
_vectorstore = None


def _get_store():
    global _embeddings, _vectorstore
    if _vectorstore is None:
        _embeddings = HuggingFaceEmbeddings(
            model_name="sentence-transformers/all-MiniLM-L6-v2"
        )
        _vectorstore = Chroma(
            persist_directory=CHROMA_PERSIST_DIR, embedding_function=_embeddings
        )
    return _vectorstore


def _extract_json(text: str) -> dict:
    # Strict JSON first, then fall back to a regex grab.
    try:
        return json.loads(text)
    except Exception:
        m = re.search(r"\{.*\}", text, re.DOTALL)
        if m:
            try:
                return json.loads(m.group(0))
            except Exception:
                pass
    return {"is_hallucination": True, "evidence": f"parse_fail: {text[:200]}"}


def verify_claim(claim: str, k: int = 3) -> dict:
    store = _get_store()
    docs = store.similarity_search(claim, k=k)
    evidence = "\n\n".join(d.page_content for d in docs)

    llm = ChatOllama(
        model=OLLAMA_MODEL,
        base_url=OLLAMA_BASE_URL,
        temperature=0.0,
        format="json",
    )
    prompt = f"""You are a fact-checking assistant. Decide whether the CLAIM below is supported by the EVIDENCE.
Rules:
- Rely ONLY on the EVIDENCE. Do not use outside knowledge.
- A claim is a hallucination if any part of it (including cited cases, quotations, dates, penalties) is not supported by the EVIDENCE.
- Output strictly JSON: {{"is_hallucination": true|false, "evidence": "<short justification with a quote from EVIDENCE or a reason>"}}.

EVIDENCE:
{evidence}

CLAIM:
{claim}
"""
    resp = llm.invoke([SystemMessage(content=prompt)])
    parsed = _extract_json(resp.content)
    return {
        "is_hallucination": bool(parsed.get("is_hallucination", True)),
        "evidence": str(parsed.get("evidence", ""))[:800],
        "retrieved_docs": [d.metadata.get("case_id") for d in docs],
    }
