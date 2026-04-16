# System Documentation: Multi-Agent RAG with Deliberate Information Asymmetry

## 1. Project Genesis & Objective
**Objective**: Build an end-to-end Python FastAPI backend implementing a Multi-Agent Retrieval-Augmented Generation (RAG) system to strictly mitigate Large Language Model (LLM) hallucinations.
**Core Constraint**: Implement a specific workflow named "Deliberate Information Asymmetry." In this architecture, not every agent has access to all the data. Fact-checking agents must be intentionally blinded to the drafting agent's prose to prevent confirmation bias.

## 2. Technology Stack Overview
- **Language**: Python 3.10+
- **API Framework**: FastAPI & Uvicorn
- **Orchestration**: LangGraph and LangChain
- **LLM**: Local model execution using Ollama (`llama3.1`)
- **Embeddings**: Local HuggingFace embeddings (`sentence-transformers/all-MiniLM-L6-v2`)
- **Vector Database**: ChromaDB (for dense semantic retrieval of unstructured text)
- **Graph Database**: Neo4j (for exact extraction of strict node-edge relationships)

---

## 3. Directory & File Structure
We created five main scripts in the workspace to construct this pipeline:

1. **`requirements.txt`**: Standard dependencies list.
2. **`database_setup.py`**: A bootstrap script to seed the local ChromaDB vector store and the local Neo4j graph database with 5-10 interconnected "dummy" medical facts (about a fictional patient named John Doe and a fictional disease called Syndromia).
3. **`graph_workflow.py`**: The heart of the system. It contains the data models (`TypedDict`, `Pydantic` schemas) and the exact graph nodes and edges for the LangGraph execution string.
4. **`main.py`**: A FastAPI application that wraps the LangGraph workflow. It exposes a single POST endpoint (`/ask`) which accepts a user query and returns a JSON payload detailing the facts extracted, the hallucination score, and the final safe response.
5. **`test_client.py`**: An End-to-End caller script built to explicitly test if the pipeline accurately captures induced hallucinations.

---

## 4. LangGraph State Schema (`WorkflowState`)
To pass data continuously through the connected agents, we defined a shared state object in LangGraph containing:
- `query`: The user's original question.
- `chroma_context`: Raw unstructured text retrieved from the vector store.
- `neo4j_context`: Raw extracted entity paths retrieved via Cypher queries.
- `solver_response`: The initial, unverified draft written by the generative agent.
- `qa_pairs`: The atomic claims extracted from the draft formatted as `[{"question": "...", "answer": "..."}]`.
- `checker_results`: Blindly validated results for each QA pair (`is_hallucination`, `evidence`).
- `hallucinations_found`: Boolean flag stating whether the Checker caught lies.
- `confidence_score`: A float (0.0 to 1.0) mathematically calculated based on the percentage of factual claims that survived strict verification.
- `final_response`: The sanitized output delivered to the user.

---

## 5. Agent Pipelines (Graph Nodes)

The core innovation is the sequential restriction of state reading. Here is what each Node (Agent) does:

### A. `Retrieve_Node`
* **Inputs**: `query`
* **Action**: Connects to the local ChromaDB instance to pull the top-K relevant documents. Simultaneously fires Cypher queries against Neo4j to pull exact subgraphs.
* **Outputs to State**: `chroma_context`, `neo4j_context`

### B. `Solver_Node` (The Drafter)
* **Inputs**: `query`, `chroma_context`, `neo4j_context`
* **Action**: Prompted via LangChain to act as a friendly assistant. It reads the raw database snippets and composes a fluent, narrative draft answer. Because LLMs are probabilistic, it is structurally prone to generating logical leaps or hallucinations in this node.
* **Outputs to State**: `solver_response`

### C. `Proposer_Node` (The Atomizer)
* **Inputs**: Only the `solver_response`
* **Action**: Forced to output strict JSON. It acts as a scalpel, ignoring all the fluff in the text and isolating every discrete factual claim made by the solver into a list of Question/Answer pairs.
* **Outputs to State**: `qa_pairs`

### D. `Checker_Node` (The Auditor)
* **Inputs**: `query`, `qa_pairs`, `chroma_context`, `neo4j_context`
* *CRITICAL CONSTRAINT*: **This agent does NOT have access to the `solver_response`.**
* **Action**: It takes the extracted atomic claims from the Drafter, and looks purely at the raw database facts. Because it cannot read the Drafter's writing, it has no psychological pressure to "agree" with it. It returns a verdict on whether each exact claim exists in the raw contexts.
* **Outputs to State**: `checker_results`, `hallucinations_found`, `confidence_score`

### E. `Synthesizer_Node` (The Editor)
* **Inputs**: `query`, `solver_response`, `checker_results`, `hallucinations_found`
* **Action**: If hallucinations are found, it receives a detailed list of which claims failed validation and *why*. It executes an LLM rewrite to cut those facts entirely out of the final text. If no hallucinations are found, it bypasses edits.
* **Outputs to State**: `final_response`

---

## 6. How the "Deliberate Information Asymmetry" Solves Hallucination
By giving the Drafter (Solver) all the information but letting it write freely, we get a highly readable response. By taking that response away from the Auditor (Checker) and only giving it isolated facts + raw data, we enforce a strict Boolean logic check against the database. Any discrepancy guarantees an LLM cannot mask a lie using semantic prose.
