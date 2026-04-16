# multiagent-hallucination-defense

Multi-Agent RAG with **Deliberate Information Asymmetry** for hallucination
detection, evaluated on a corpus of real legal cases in which AI tools produced
fabricated citations, false quotations, or misrepresented precedents.

## Architecture

Solver → Proposer → Checker (blinded) → Synthesizer, orchestrated with LangGraph.

| Component | Role |
|---|---|
| `graph_workflow.py` | LangGraph pipeline with the four agents |
| `database_setup.py` | Loads the hallucination-cases corpus into ChromaDB + Neo4j |
| `main.py` | FastAPI wrapper exposing `/ask` |
| `evaluation/generate_probes.py` | Builds TRUE/FALSE probe pairs from the corpus |
| `evaluation/baseline.py` | Single-agent RAG baseline |
| `evaluation/run_benchmark.py` | Runs both pipelines over all probes |
| `evaluation/metrics.py` | Precision / recall / F1 / accuracy |

## Stack

Python 3.10+ · FastAPI · LangGraph · LangChain · Ollama (`llama3.1`) ·
HuggingFace `all-MiniLM-L6-v2` embeddings · ChromaDB · Neo4j.

## Dataset

Charlotin AI Hallucination Cases (Kaggle). 426 rows; 243 usable after
filtering for non-empty `Details` and `Hallucination` fields. The CSV itself
is not committed — download it into `data/` before running.

## Run

```bash
# prerequisites
ollama serve &
ollama pull llama3.1
neo4j start

# setup
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
python database_setup.py

# benchmark
python -m evaluation.generate_probes --n 30 --out evaluation/probes.jsonl
python -m evaluation.run_benchmark --probes evaluation/probes.jsonl --out results/benchmark_results.csv
python -m evaluation.metrics --results results/benchmark_results.csv
```
