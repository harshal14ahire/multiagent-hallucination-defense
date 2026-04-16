"""
Benchmark runner: iterate every probe through BOTH pipelines and log per-probe results.

Outputs a tidy CSV suitable for pandas analysis and direct inclusion in the paper.
Columns:
  probe_id, case_id, probe_type, expected_is_hallucination,
  pipeline, predicted_is_hallucination, confidence, latency_s, evidence, error
"""
import argparse
import csv
import json
import time
import traceback
from pathlib import Path

from evaluation import baseline
from graph_workflow import build_workflow


def load_probes(path: str):
    with open(path) as f:
        return [json.loads(line) for line in f if line.strip()]


def run_baseline(probe: dict) -> dict:
    t0 = time.time()
    try:
        r = baseline.verify_claim(probe["claim"], k=3)
        return {
            "predicted_is_hallucination": r["is_hallucination"],
            "confidence": None,
            "evidence": r["evidence"],
            "latency_s": round(time.time() - t0, 2),
            "error": "",
        }
    except Exception as e:
        return {
            "predicted_is_hallucination": None,
            "confidence": None,
            "evidence": "",
            "latency_s": round(time.time() - t0, 2),
            "error": f"{type(e).__name__}: {e}",
        }


def run_multiagent(probe: dict, workflow) -> dict:
    """
    Drive the existing LangGraph pipeline. We repurpose it for claim verification:
      - query := the claim
      - the pipeline retrieves, drafts (Solver), atomizes (Proposer), checks (Checker)
      - predicted hallucination = `hallucinations_found` after the Checker
      - confidence = `confidence_score` already produced by the pipeline
    """
    t0 = time.time()
    try:
        result = workflow.invoke({"query": probe["claim"]})
        return {
            "predicted_is_hallucination": bool(result.get("hallucinations_found", True)),
            "confidence": result.get("confidence_score"),
            "evidence": json.dumps(result.get("checker_results", []))[:2000],
            "latency_s": round(time.time() - t0, 2),
            "error": "",
        }
    except Exception as e:
        traceback.print_exc()
        return {
            "predicted_is_hallucination": None,
            "confidence": None,
            "evidence": "",
            "latency_s": round(time.time() - t0, 2),
            "error": f"{type(e).__name__}: {e}",
        }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--probes", type=str, default="evaluation/probes.jsonl")
    ap.add_argument("--out", type=str, default="results/benchmark_results.csv")
    ap.add_argument("--limit", type=int, default=0, help="0 = all probes")
    ap.add_argument(
        "--pipelines",
        type=str,
        default="baseline,multiagent",
        help="comma-separated subset of {baseline, multiagent}",
    )
    args = ap.parse_args()

    probes = load_probes(args.probes)
    if args.limit > 0:
        probes = probes[: args.limit]

    pipelines = [p.strip() for p in args.pipelines.split(",") if p.strip()]
    workflow = build_workflow() if "multiagent" in pipelines else None

    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    fields = [
        "probe_id", "case_id", "probe_type", "expected_is_hallucination",
        "pipeline", "predicted_is_hallucination", "confidence",
        "latency_s", "evidence", "error",
    ]
    with open(args.out, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for i, probe in enumerate(probes):
            print(f"[{i+1}/{len(probes)}] {probe['probe_id']} ({probe['probe_type']})")
            for pipe in pipelines:
                runner = run_baseline if pipe == "baseline" else run_multiagent
                kwargs = {} if pipe == "baseline" else {"workflow": workflow}
                res = runner(probe, **kwargs) if kwargs else runner(probe)
                row = {
                    "probe_id": probe["probe_id"],
                    "case_id": probe["case_id"],
                    "probe_type": probe["probe_type"],
                    "expected_is_hallucination": probe["expected_is_hallucination"],
                    "pipeline": pipe,
                    **res,
                }
                w.writerow(row)
                f.flush()
                print(
                    f"    {pipe:10s} expected={probe['expected_is_hallucination']!s:<5} "
                    f"predicted={res['predicted_is_hallucination']!s:<5} "
                    f"latency={res['latency_s']}s"
                )
    print(f"\nResults written to {args.out}")


if __name__ == "__main__":
    main()
