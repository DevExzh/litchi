#!/usr/bin/env python3
"""Summarize the change-0601 baseline: per-case statistics and the A/A floor."""

from __future__ import annotations

import json
import sys
from pathlib import Path


def load(path: Path) -> dict[str, dict]:
    document = json.loads(path.read_text(encoding="utf-8"))
    return {result["case"]: result for result in document["results"]}


def main() -> int:
    out = Path(sys.argv[1])
    a = load(out / "baseline-a.json")
    b = load(out / "baseline-b.json")
    rows = []
    for case in sorted(a):
        first, second = a[case]["elapsed_ns"], b[case]["elapsed_ns"]
        corpus = a[case]["corpus"]
        drift = abs(second["p50"] - first["p50"]) / first["p50"] * 100
        rows.append(
            {
                "case": case,
                "corpus": corpus["name"],
                "archive_bytes": corpus["archive_bytes"],
                "worksheet_bytes": corpus["uncompressed_payload_bytes"],
                "a_p50_ns": first["p50"],
                "a_p95_ns": first["p95"],
                "a_p99_ns": first["p99"],
                "a_mean_ns": first["mean"],
                "b_p50_ns": second["p50"],
                "b_p95_ns": second["p95"],
                "b_p99_ns": second["p99"],
                "b_mean_ns": second["mean"],
                "aa_p50_drift_percent": round(drift, 2),
            }
        )
    real_path = out / "real-file.json"
    real = load(real_path) if real_path.exists() else {}
    for case in sorted(real):
        stats = real[case]["elapsed_ns"]
        corpus = real[case]["corpus"]
        rows.append(
            {
                "case": case,
                "corpus": corpus["name"],
                "archive_bytes": corpus["archive_bytes"],
                "worksheet_bytes": corpus["uncompressed_payload_bytes"],
                "a_p50_ns": stats["p50"],
                "a_p95_ns": stats["p95"],
                "a_p99_ns": stats["p99"],
                "a_mean_ns": stats["mean"],
                "b_p50_ns": None,
                "b_p95_ns": None,
                "b_p99_ns": None,
                "b_mean_ns": None,
                "aa_p50_drift_percent": None,
            }
        )
    drifts = [row["aa_p50_drift_percent"] for row in rows if row["aa_p50_drift_percent"] is not None]
    summary = {
        "schema": "litchi.perf.change-0601.baseline-summary.v1",
        "warmup_iterations": 20,
        "samples": 100,
        "aa_floor_p50_percent_max": max(drifts) if drifts else None,
        "aa_floor_p50_percent_median": sorted(drifts)[len(drifts) // 2] if drifts else None,
        "rows": rows,
    }
    (out / "summary.json").write_text(
        json.dumps(summary, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )
    width = max(len(row["case"]) for row in rows)
    print(f"{'case'.ljust(width)}  {'p50 ns':>12}  {'p95 ns':>12}  {'p99 ns':>12}  {'A/A p50 %':>9}")
    for row in rows:
        drift = "-" if row["aa_p50_drift_percent"] is None else f"{row['aa_p50_drift_percent']:.2f}"
        print(
            f"{row['case'].ljust(width)}  {row['a_p50_ns']:>12}  {row['a_p95_ns']:>12}  "
            f"{row['a_p99_ns']:>12}  {drift:>9}"
        )
    print(f"\nA/A p50 floor: max {summary['aa_floor_p50_percent_max']}%, "
          f"median {summary['aa_floor_p50_percent_median']}%")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
