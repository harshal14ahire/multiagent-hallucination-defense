from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from typing import List, Dict, Any

from graph_workflow import build_workflow, WorkflowState

from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse

app = FastAPI(
    title="Multi-Agent Hallucination Mitigation API",
    description="An API that uses LangGraph and Deliberate Information Asymmetry to mitigate LLM hallucinations.",
    version="1.0"
)

app.mount("/static", StaticFiles(directory="static"), name="static")

# Compile LangGraph Workflow
workflow_app = build_workflow()

class AskRequest(BaseModel):
    query: str

class CheckerDetailedResult(BaseModel):
    question: str
    proposer_answer: str
    is_hallucination: bool
    evidence: str

class AskResponse(BaseModel):
    query: str
    initial_solver_response: str
    final_response: str
    hallucinations_found: bool
    confidence_score: float
    checker_results: List[CheckerDetailedResult]

@app.get("/", response_class=HTMLResponse)
def read_root():
    with open("static/index.html") as f:
        return f.read()

@app.post("/ask", response_model=AskResponse)
def process_query(request: AskRequest):
    """
    Process a user query through the LangGraph Multi-Agent RAG workflow.
    """
    try:
        # Initialize the state with the user query
        initial_state = {
            "query": request.query,
            "chroma_context": "",
            "neo4j_context": "",
            "solver_response": "",
            "qa_pairs": [],
            "checker_results": [],
            "hallucinations_found": False,
            "confidence_score": 0.0,
            "final_response": ""
        }
        
        # Run the workflow
        result_state = workflow_app.invoke(initial_state)
        
        return AskResponse(
            query=result_state["query"],
            initial_solver_response=result_state["solver_response"],
            final_response=result_state["final_response"],
            hallucinations_found=result_state["hallucinations_found"],
            confidence_score=result_state["confidence_score"],
            checker_results=result_state["checker_results"],
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
