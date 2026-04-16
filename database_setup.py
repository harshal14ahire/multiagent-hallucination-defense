"""
Load the Charlotin hallucination cases dataset into ChromaDB + Neo4j.

Replaces the previous dummy medical seeder. Each row becomes:
  - Chroma document: the `Details` text, with case metadata
  - Neo4j nodes:   (Case)-[:IN_COURT]->(Court)
                   (Case)-[:INVOLVED]->(Party)
                   (Case)-[:USED]->(AITool)
                   (Case)-[:EXHIBITED]->(HallucinationType)
"""
import os
import re
import pandas as pd
from neo4j import GraphDatabase
from langchain_community.vectorstores import Chroma
from langchain_community.embeddings import HuggingFaceEmbeddings

CSV_PATH = os.getenv(
    "HALLU_CSV", "./data/Charlotin-hallucination_cases.csv"
)
CHROMA_PERSIST_DIR = "./chroma_db"
NEO4J_URI = os.getenv("NEO4J_URI", "bolt://localhost:7687")
NEO4J_USER = os.getenv("NEO4J_USER", "neo4j")
NEO4J_PASSWORD = os.getenv("NEO4J_PASSWORD", "admin@123")


def _safe(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


def load_rows():
    df = pd.read_csv(CSV_PATH)
    # Keep rows with enough evidence in Details
    df = df[df["Details"].astype(str).str.len() > 100].copy()
    df = df[df["Hallucination"].astype(str).str.len() > 3].copy()
    df = df.reset_index(drop=True)
    df["case_id"] = df.index.map(lambda i: f"case_{i:04d}")
    return df


def setup_chroma(df: pd.DataFrame):
    print(f"Setting up ChromaDB with {len(df)} cases...")
    embeddings = HuggingFaceEmbeddings(
        model_name="sentence-transformers/all-MiniLM-L6-v2"
    )

    texts = []
    metadatas = []
    for _, r in df.iterrows():
        # Compose a retrieval-friendly document that keeps case identity close to the facts
        doc = (
            f"Case: {_safe(r['Case Name'])}\n"
            f"Court: {_safe(r['Court'])} ({_safe(r['State(s)'])})\n"
            f"Date: {_safe(r['Date'])}\n"
            f"Party: {_safe(r['Party(ies)'])}\n"
            f"AI Tool: {_safe(r['AI Tool'])}\n"
            f"Outcome: {_safe(r['Outcome'])}\n"
            f"Details: {_safe(r['Details'])}"
        )
        texts.append(doc)
        metadatas.append(
            {
                "case_id": r["case_id"],
                "case_name": _safe(r["Case Name"])[:200],
                "court": _safe(r["Court"])[:100],
                "date": _safe(r["Date"])[:30],
                "source": "Charlotin-hallucination_cases",
            }
        )

    Chroma.from_texts(
        texts=texts,
        embedding=embeddings,
        metadatas=metadatas,
        persist_directory=CHROMA_PERSIST_DIR,
    )
    print(f"  -> loaded {len(texts)} docs into {CHROMA_PERSIST_DIR}")


def setup_neo4j(df: pd.DataFrame):
    print("Setting up Neo4j...")
    try:
        driver = GraphDatabase.driver(NEO4J_URI, auth=(NEO4J_USER, NEO4J_PASSWORD))
        with driver.session() as session:
            # Clear the dataset namespace before reload
            session.run(
                "MATCH (n) WHERE n.source = 'Charlotin-hallucination_cases' DETACH DELETE n"
            )
            for _, r in df.iterrows():
                params = {
                    "case_id": r["case_id"],
                    "case_name": _safe(r["Case Name"]),
                    "court": _safe(r["Court"]),
                    "date": _safe(r["Date"]),
                    "party": _safe(r["Party(ies)"]),
                    "ai_tool": _safe(r["AI Tool"]),
                    "hallucination": _safe(r["Hallucination"]),
                    "outcome": _safe(r["Outcome"]),
                }
                session.run(
                    """
                    MERGE (c:Case {case_id: $case_id})
                    SET c.name = $case_name,
                        c.date = $date,
                        c.outcome = $outcome,
                        c.source = 'Charlotin-hallucination_cases'
                    MERGE (ct:Court {name: $court}) SET ct.source = 'Charlotin-hallucination_cases'
                    MERGE (c)-[:IN_COURT]->(ct)
                    MERGE (p:Party {name: $party}) SET p.source = 'Charlotin-hallucination_cases'
                    MERGE (c)-[:INVOLVED]->(p)
                    MERGE (a:AITool {name: $ai_tool}) SET a.source = 'Charlotin-hallucination_cases'
                    MERGE (c)-[:USED]->(a)
                    MERGE (h:HallucinationType {name: $hallucination}) SET h.source = 'Charlotin-hallucination_cases'
                    MERGE (c)-[:EXHIBITED]->(h)
                    """,
                    params,
                )
        driver.close()
        print(f"  -> Neo4j loaded with {len(df)} cases.")
    except Exception as e:
        print(f"Neo4j error: {e}")


if __name__ == "__main__":
    df = load_rows()
    print(f"Loaded {len(df)} usable rows from {CSV_PATH}")
    setup_chroma(df)
    setup_neo4j(df)
    print("Database setup complete.")
