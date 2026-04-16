"""
Compute precision / recall / F1 / accuracy per pipeline from benchmark_results.csv.

Treats probes with `expected_is_hallucination=True` as the positive class.
"""
import argparse

import pandas as pd


def _confusion(df: pd.DataFrame) -> dict:
    valid = df.dropna(subset=["predicted_is_hallucination"])
    y_true = valid["expected_is_hallucination"].astype(bool)
    y_pred = valid["predicted_is_hallucination"].astype(bool)
    tp = int(((y_true == True) & (y_pred == True)).sum())
    tn = int(((y_true == False) & (y_pred == False)).sum())
    fp = int(((y_true == False) & (y_pred == True)).sum())
    fn = int(((y_true == True) & (y_pred == False)).sum())
    n = tp + tn + fp + fn
    acc = (tp + tn) / n if n else 0.0
    prec = tp / (tp + fp) if (tp + fp) else 0.0
    rec = tp / (tp + fn) if (tp + fn) else 0.0
    f1 = 2 * prec * rec / (prec + rec) if (prec + rec) else 0.0
    return {
        "n": n,
        "tp": tp, "tn": tn, "fp": fp, "fn": fn,
        "accuracy": round(acc, 4),
        "precision": round(prec, 4),
        "recall": round(rec, 4),
        "f1": round(f1, 4),
        "errors": int(df["predicted_is_hallucination"].isna().sum()),
        "mean_latency_s": round(valid["latency_s"].mean(), 2) if len(valid) else None,
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--results", default="results/benchmark_results.csv")
    ap.add_argument("--out", default="results/metrics.csv")
    args = ap.parse_args()

    df = pd.read_csv(args.results)
    rows = []
    for pipe, group in df.groupby("pipeline"):
        rows.append({"pipeline": pipe, **_confusion(group)})

    out = pd.DataFrame(rows).set_index("pipeline")
    print("\n=== Hallucination Detection Metrics (positive class = is_hallucination) ===\n")
    print(out.to_string())
    out.to_csv(args.out)
    print(f"\nWritten to {args.out}")


if __name__ == "__main__":
    main()
