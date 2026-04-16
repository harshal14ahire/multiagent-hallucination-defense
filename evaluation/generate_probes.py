"""
Generate the probe set for the hallucination-detection benchmark.

Per sampled case we emit two probes:
  - TRUE probe:  a claim grounded in `Details` (expected: NOT a hallucination)
  - FALSE probe: the TRUE claim mutated with a fabricated citation (expected: hallucination)

The FALSE mutation appends a plausible-looking but non-existent case citation.
Because the citation does not appear in the retrieved evidence, a correctly-behaving
pipeline must flag the FALSE probe as a hallucination. This mirrors the dominant
failure mode catalogued in the Charlotin dataset: fabricated citations / false
quotations / non-existent precedents (see the `Hallucination` column).
"""
import argparse
import json
import os
import random
import re
from pathlib import Path

import pandas as pd

CSV_PATH = os.getenv("HALLU_CSV", "./data/Charlotin-hallucination_cases.csv")


FABRICATED_CITATIONS = [
    "Whitmore v. Ardant, 612 F.3d 441 (9th Cir. 2019)",
    "State v. Delgado-Hoyt, 284 Conn. 112 (2021)",
    "In re Marigold Holdings, 2020 WL 8841197",
    "People v. Kavanaugh-Reyes, 77 N.Y.3d 509 (2022)",
    "Harper v. Sinclair Pharmaceuticals, 89 Cal. App. 5th 644 (2023)",
    "Johnson v. Maryland Civil Liberties Union, 601 U.S. 312 (2024)",
    "United States v. Brankowski, 998 F.3d 77 (4th Cir. 2018)",
    "Matter of Pennybaker, 233 A.D.3d 1102 (N.Y. App. Div. 2023)",
]


def _safe(s) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip()


def _first_sentence(text: str, max_len: int = 260) -> str:
    text = _safe(text)
    m = re.search(r"(.+?[.!?])(\s|$)", text)
    frag = m.group(1) if m else text
    return frag[:max_len]


def load_rows() -> pd.DataFrame:
    df = pd.read_csv(CSV_PATH)
    df = df[df["Details"].astype(str).str.len() > 100].copy()
    df = df[df["Hallucination"].astype(str).str.len() > 3].copy()
    df = df.reset_index(drop=True)
    df["case_id"] = df.index.map(lambda i: f"case_{i:04d}")
    return df


def build_probes(df: pd.DataFrame, n_cases: int, seed: int = 42) -> list[dict]:
    rng = random.Random(seed)
    sampled = df.sample(n=min(n_cases, len(df)), random_state=seed).reset_index(drop=True)
    probes = []
    for _, r in sampled.iterrows():
        case_name = _safe(r["Case Name"])
        true_fact = _first_sentence(r["Details"])
        true_claim = f"In {case_name}, {true_fact}"

        fake_cite = rng.choice(FABRICATED_CITATIONS)
        # Attach a fabricated supporting citation to the TRUE claim -> creates a FALSE probe.
        # The underlying proposition may still be correct, but the asserted supporting
        # citation is not present in the evidence; an honest fact-checker must flag it.
        false_claim = f"In {case_name}, {true_fact} The court relied on {fake_cite}."

        probes.append({
            "probe_id": f"{r['case_id']}_T",
            "case_id": r["case_id"],
            "case_name": case_name,
            "claim": true_claim,
            "expected_is_hallucination": False,
            "probe_type": "true",
        })
        probes.append({
            "probe_id": f"{r['case_id']}_F",
            "case_id": r["case_id"],
            "case_name": case_name,
            "claim": false_claim,
            "expected_is_hallucination": True,
            "probe_type": "false_citation",
            "injected_citation": fake_cite,
        })
    return probes


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--n", type=int, default=30, help="number of cases to sample (→ 2×N probes)")
    ap.add_argument("--seed", type=int, default=42)
    ap.add_argument("--out", type=str, default="evaluation/probes.jsonl")
    args = ap.parse_args()

    df = load_rows()
    probes = build_probes(df, n_cases=args.n, seed=args.seed)
    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    with open(args.out, "w") as f:
        for p in probes:
            f.write(json.dumps(p) + "\n")
    print(f"Wrote {len(probes)} probes ({len(probes)//2} cases × 2) to {args.out}")


if __name__ == "__main__":
    main()
