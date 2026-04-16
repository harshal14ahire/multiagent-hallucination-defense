import requests
import json

def test_e2e():
    url = "http://localhost:8000/ask"
    
    # Query designed to test both factual retrieval and hallucination mitigation
    payload = {
        "query": "Who is John Doe's doctor, what disease does he have, and does he take Penicillin?"
    }
    
    print(f"Sending Query: {payload['query']}\n")
    
    try:
        response = requests.post(url, json=payload)
        response.raise_for_status()
        
        data = response.json()
        print("=== E2E Test Results ===\n")
        
        print("1. Original Narrative:")
        print(data.get("initial_solver_response"), "\n")
        
        print(f"2. Hallucinations Detected: {data.get('hallucinations_found')}")
        print(f"   Confidence Score: {data.get('confidence_score') * 100}%\n")
        
        if data.get("hallucinations_found"):
            print("3. Checker Agent Findings:")
            for result in data.get("checker_results", []):
                if result.get("is_hallucination"):
                    print(f"   * FALSE CLAIM: {result.get('proposer_answer')}")
                    print(f"   * EVIDENCE: {result.get('evidence')}\n")
                    
            print("4. Rewritten Narrative (Synthesizer output):")
            print(data.get("final_response"), "\n")
            
        else:
            print("3. Perfect Confidence. Verifications passed.")
            
    except requests.exceptions.ConnectionError:
        print("Error: Could not connect to the API. Make sure you run `uvicorn main:app` first.")
    except Exception as e:
        print(f"Error testing API: {e}")

if __name__ == "__main__":
    test_e2e()
