import os
import json
from typing import List, Dict, Any
from typing_extensions import TypedDict
from pydantic import BaseModel, Field

from langchain_community.vectorstores import Chroma
from langchain_community.embeddings import HuggingFaceEmbeddings
from neo4j import GraphDatabase
from langchain_community.chat_models import ChatOllama
from langchain_core.messages import SystemMessage, HumanMessage
from langgraph.graph import StateGraph, START, END

# Config
CHROMA_PERSIST_DIR = "./chroma_db"
NEO4J_URI = os.getenv("NEO4J_URI", "bolt://localhost:7687")
NEO4J_USER = os.getenv("NEO4J_USER", "neo4j")
NEO4J_PASSWORD = os.getenv("NEO4J_PASSWORD", "admin@123")

OLLAMA_MODEL = os.getenv("OLLAMA_MODEL", "llama3.1")
OLLAMA_BASE_URL = os.getenv("OLLAMA_BASE_URL", "http://localhost:11434")

# -- Models and State Structure --

class WorkflowState(TypedDict):
    """The State of the LangGraph workflow."""
    query: str
    chroma_context: str
    neo4j_context: str
    solver_response: str
    qa_pairs: list[dict]       # [{"question": "...", "answer": "..."}]
    checker_results: list[dict] # [{"question": "...", "proposer_answer": "...", "is_hallucination": bool, "evidence": "..."}]
    hallucinations_found: bool
    confidence_score: float
    final_response: str

class QAPair(BaseModel):
    question: str = Field(description="The extracted question for a factual claim")
    answer: str = Field(description="The extracted answer for a factual claim from the text")

class ProposerOutput(BaseModel):
    claims: List[QAPair] = Field(description="List of factual claims extracted from the narrative")

class CheckerResult(BaseModel):
    is_hallucination: bool = Field(description="True if the claim is contradicted by or absent from the raw context")
    evidence: str = Field(description="Short quote or explanation based ONLY on raw context")

class CheckerOutput(BaseModel):
    evaluations: List[CheckerResult] = Field(description="List of evaluations corresponding structurally to the input claims")

# -- Nodes --

def retrieve_node(state: WorkflowState) -> WorkflowState:
    """Retrieve data from ChromaDB and Neo4j."""
    query = state["query"]
    
    # 1. Fetch from VectorStore
    embeddings = HuggingFaceEmbeddings(model_name="sentence-transformers/all-MiniLM-L6-v2")
    vectorstore = Chroma(persist_directory=CHROMA_PERSIST_DIR, embedding_function=embeddings)
    docs = vectorstore.similarity_search(query, k=3)
    chroma_context = "\n".join([d.page_content for d in docs])
    
    # 2. Fetch from Neo4j: pull structured facts for any Case whose name appears in the query,
    #    plus structured facts for the top-retrieved Chroma cases (by case_id metadata).
    neo4j_context = ""
    try:
        case_ids = []
        for d in docs:
            cid = (d.metadata or {}).get("case_id")
            if cid:
                case_ids.append(cid)
        driver = GraphDatabase.driver(NEO4J_URI, auth=(NEO4J_USER, NEO4J_PASSWORD))
        with driver.session() as session:
            if case_ids:
                result = session.run(
                    """
                    MATCH (c:Case)-[r]->(n)
                    WHERE c.case_id IN $ids
                    RETURN c.case_id AS case_id, c.name AS case, type(r) AS rel, labels(n)[0] AS kind, n.name AS value
                    """,
                    {"ids": case_ids},
                )
                for record in result:
                    neo4j_context += (
                        f"{record['case']} --{record['rel']}--> ({record['kind']}) {record['value']}\n"
                    )
    except Exception as e:
        print(f"Neo4j retrieval error: {e}")
        neo4j_context = "No graph data available."
    
    return {"chroma_context": chroma_context, "neo4j_context": neo4j_context}

def solver_node(state: WorkflowState) -> WorkflowState:
    """Generate the initial response using the Solver Agent."""
    query = state["query"]
    chroma_context = state["chroma_context"]
    neo4j_context = state["neo4j_context"]
    
    llm = ChatOllama(model=OLLAMA_MODEL, base_url=OLLAMA_BASE_URL, temperature=0.7)
    
    prompt = f"""You are a helpful assistant. Use the following contexts to answer the user's question.
    Vector Context: {chroma_context}
    Graph Context: {neo4j_context}
    
    User Question: {query}
    
    Write a clear, comprehensive narrative response. Do not explicitly cite the contexts, just provide the answer natively.
    """
    
    response = llm.invoke([HumanMessage(content=prompt)])
    return {"solver_response": response.content}

def proposer_node(state: WorkflowState) -> WorkflowState:
    """Break the Solver's response into QA facts. Since ChatOllama doesn't natively support tool calling with pydantic well in older versions, we format explicitly."""
    solver_response = state["solver_response"]
    
    llm = ChatOllama(model=OLLAMA_MODEL, base_url=OLLAMA_BASE_URL, temperature=0.0, format="json")
    
    prompt = f"""You are an Atomizer Agent. Your task is to extract every factual claim from the given text and represent it strictly as a list of Question-Answer pairs.
    
    Constraint: Your output must be valid JSON in this exact format:
    {{
        "claims": [
            {{"question": "...", "answer": "..."}}
        ]
    }}
    
    Text: {solver_response}
    """
    
    response = llm.invoke([SystemMessage(content=prompt)])
    
    try:
        content = json.loads(response.content)
        qa_pairs = content.get("claims", [])
    except json.JSONDecodeError:
        qa_pairs = [] # Fallback
        
    return {"qa_pairs": qa_pairs}

def checker_node(state: WorkflowState) -> WorkflowState:
    """Blind evaluation of QA pairs using the Checker Agent against raw context."""
    qa_pairs = state["qa_pairs"]
    chroma_context = state["chroma_context"]
    neo4j_context = state["neo4j_context"]
    
    if not qa_pairs:
        return {"checker_results": [], "hallucinations_found": False, "confidence_score": 1.0}
    
    llm = ChatOllama(model=OLLAMA_MODEL, base_url=OLLAMA_BASE_URL, temperature=0.0, format="json")
    
    prompt = f"""You are a Checker Agent. You verify claims purely based on provided RAW EVIDENCE.
    You must NOT rely on outside knowledge. If the RAW EVIDENCE does not explicitly support the claim, it is a hallucination.
    
    RAW EVIDENCE:
    Vector: {chroma_context}
    Graph: {neo4j_context}
    
    CLAIMS TO VERIFY:
    {json.dumps(qa_pairs, indent=2)}
    
    Constraint: Your output must be valid JSON in this exact format:
    {{
        "evaluations": [
            {{"is_hallucination": true/false, "evidence": "quote or reason from raw evidence"}}
        ]
    }}
    Ensure the array size matches the number of claims.
    """
    
    response = llm.invoke([SystemMessage(content=prompt)])
    
    try:
        content = json.loads(response.content)
        evals = content.get("evaluations", [])
    except json.JSONDecodeError:
        evals = [{"is_hallucination": True, "evidence": "Failed to parse checker response."} for _ in qa_pairs]
    
    checker_results = []
    hallu_count = 0
    for i in range(len(qa_pairs)):
        ev = evals[i] if i < len(evals) else {"is_hallucination": True, "evidence": "Missing evaluation."}
        checker_results.append({
            "question": qa_pairs[i].get("question", ""),
            "proposer_answer": qa_pairs[i].get("answer", ""),
            "is_hallucination": ev.get("is_hallucination", True),
            "evidence": ev.get("evidence", "")
        })
        if ev.get("is_hallucination", True):
            hallu_count += 1
            
    confidence_score = 1.0 - (hallu_count / len(qa_pairs)) if qa_pairs else 1.0
    return {
        "checker_results": checker_results, 
        "hallucinations_found": hallu_count > 0,
        "confidence_score": confidence_score
    }

def synthesizer_node(state: WorkflowState) -> WorkflowState:
    """Rewrite response if hallucinations are found."""
    if not state["hallucinations_found"]:
        return {"final_response": state["solver_response"]}
        
    query = state["query"]
    solver_response = state["solver_response"]
    checker_results = state["checker_results"]
    
    llm = ChatOllama(model=OLLAMA_MODEL, base_url=OLLAMA_BASE_URL, temperature=0.7)
    
    hallucinations = [r for r in checker_results if r["is_hallucination"]]
    hallucination_text = "\n".join([f"- The claim '{r['proposer_answer']}' (Question: {r['question']}) is unsupported because: {r['evidence']}" for r in hallucinations])
    
    prompt = f"""You are a Revision Agent.
    User's Original Query: {query}
    
    Original Attempt: {solver_response}
    
    The Checker Agent found the following hallucinations in the Original Attempt:
    {hallucination_text}
    
    Rewrite the response so it completely excludes the hallucinated information. Only state what is definitively supported. Ensure it flows naturally.
    """
    
    response = llm.invoke([HumanMessage(content=prompt)])
    return {"final_response": response.content}

# -- Build Graph --

def build_workflow() -> StateGraph:
    workflow = StateGraph(WorkflowState)
    
    # Add nodes
    workflow.add_node("Retrieve_Node", retrieve_node)
    workflow.add_node("Solver_Node", solver_node)
    workflow.add_node("Proposer_Node", proposer_node)
    workflow.add_node("Checker_Node", checker_node)
    workflow.add_node("Synthesizer_Node", synthesizer_node)
    
    # Add edges
    workflow.add_edge(START, "Retrieve_Node")
    workflow.add_edge("Retrieve_Node", "Solver_Node")
    workflow.add_edge("Solver_Node", "Proposer_Node")
    workflow.add_edge("Proposer_Node", "Checker_Node")
    workflow.add_edge("Checker_Node", "Synthesizer_Node")
    workflow.add_edge("Synthesizer_Node", END)
    
    return workflow.compile()
