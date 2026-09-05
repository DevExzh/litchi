#!/usr/bin/env python3
"""Summarize the eight bounded ZIP guard rows from one ABBA capture."""
from __future__ import annotations
import argparse
import hashlib
import json
import re
from pathlib import Path

ROLES = ("A1", "B1", "B2", "A2")
MODES = ("borrowed", "owned")
PAYLOADS = ("zeros", "mixed")
SIZES = (16 * 1024, 1024 * 1024)
DEFAULT_SAMPLES = 300
DEFAULT_WARMUPS = 30
U64_MAX = (1 << 64) - 1
RSS_RE = re.compile(r"^\s*Maximum resident set size \(kbytes\):\s*(\d+)\s*$", re.M)

def fail(message: str) -> None:
    raise ValueError(message)

def integer(value: object, where: str, *, maximum: int = U64_MAX, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum or value > maximum:
        fail(f"{where} must be an unsigned integer in [{minimum}, {maximum}]")
    return value

def percentile(samples: list[int], number: int) -> int | float:
    ordered = sorted(samples)
    if number == 50:
        total = ordered[(len(ordered) - 1) // 2] + ordered[len(ordered) // 2]
        return total // 2 if total % 2 == 0 else total // 2 + 0.5
    return ordered[((number * len(ordered) + 99) // 100) - 1]

def stats(samples: list[int]) -> dict[str, int | float]:
    return {f"p{number}": percentile(samples, number) for number in (50, 95, 99)}

def read_row(path: Path, mode: str, payload: str, size: int, samples_expected: int, warmups_expected: int) -> dict:
    try:
        raw = path.read_bytes()
        row = json.loads(raw)
    except (OSError, json.JSONDecodeError) as error:
        fail(f"{path}: cannot load JSON: {error}")
    if not isinstance(row, dict):
        fail(f"{path}: report must be an object")
    for key, expected in (("mode", mode), ("payload", payload), ("input_bytes", size), ("warmup", warmups_expected)):
        if row.get(key) != expected:
            fail(f"{path}: {key}={row.get(key)!r}, expected {expected!r}")
    if "sample_count" in row and row["sample_count"] != samples_expected:
        fail(f"{path}: sample_count={row['sample_count']!r}, expected {samples_expected}")
    values = row.get("samples_ns")
    if not isinstance(values, list) or len(values) != samples_expected:
        fail(f"{path}: samples_ns must contain exactly {samples_expected} values")
    row["samples_ns"] = [integer(value, f"{path}.samples_ns[{index}]", minimum=1) for index, value in enumerate(values)]
    for key, maximum in (("output_bytes", U64_MAX), ("output_writes", U64_MAX), ("output_crc32", 0xFFFFFFFF)):
        integer(row.get(key), f"{path}.{key}", maximum=maximum)
    row["_path"], row["_sha256"], row["_stats"] = path, hashlib.sha256(raw).hexdigest(), stats(row["samples_ns"])
    return row

def rss_for(root: Path, stem: str) -> tuple[int, Path]:
    candidates = (
        root / "time-v" / f"{stem}.txt", root / "time-v" / f"{stem}.time.txt",
        root / "time-v" / f"{stem}.time-v.txt", root / "guards" / f"{stem}.time.txt",
        root / "guards" / f"{stem}.time-v.txt", root / f"{stem}.time.txt",
        root / f"{stem}.time-v.txt",
    )
    for path in candidates:
        if path.is_file():
            match = RSS_RE.search(path.read_text())
            if match:
                return int(match.group(1)), path
            fail(f"{path}: missing Maximum resident set size (kbytes)")
    fail(f"{stem}: required time-v RSS file is missing")

def delta(later: int | float, earlier: int | float) -> float:
    if earlier <= 0:
        fail(f"cannot compute a delta from non-positive statistic {earlier!r}")
    return (later - earlier) * 100.0 / earlier

def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--samples", type=int, default=DEFAULT_SAMPLES)
    parser.add_argument("--warmups", type=int, default=DEFAULT_WARMUPS)
    args = parser.parse_args()
    if args.samples <= 0 or args.warmups < 0:
        fail("--samples must be positive and --warmups must be non-negative")
    rows, triggers = [], []
    for mode in MODES:
        for payload in PAYLOADS:
            for size in SIZES:
                label = f"{mode}/{payload}/{size}"
                by_role = {}
                for role in ROLES:
                    stem = f"{role}-{mode}-{payload}-{size}"
                    by_role[role] = read_row(args.root / "guards" / f"{stem}.json", mode, payload, size, args.samples, args.warmups)
                identity = tuple(by_role["A1"][name] for name in ("output_bytes", "output_writes", "output_crc32"))
                for role, row in by_role.items():
                    if tuple(row[name] for name in ("output_bytes", "output_writes", "output_crc32")) != identity:
                        fail(f"{label}: output oracle differs in {role}")
                metric_stats = {role: row["_stats"] for role, row in by_role.items()}
                pairs = {
                    "B1_minus_A1": ("candidate_regression", "B1", "A1"),
                    "B2_minus_A2": ("candidate_regression", "B2", "A2"),
                    "A2_minus_A1": ("within_revision_drift", "A2", "A1"),
                    "B2_minus_B1": ("within_revision_drift", "B2", "B1"),
                }
                pair_deltas = {}
                for pair, (kind, later, earlier) in pairs.items():
                    values = {}
                    for metric in ("p50", "p95", "p99"):
                        value = delta(metric_stats[later][metric], metric_stats[earlier][metric])
                        values[metric] = value
                        if value > 5.0:
                            triggers.append({"row": label, "comparison_kind": kind, "pair": pair, "metric": metric, "regression_percent": value})
                    pair_deltas[pair] = {"comparison_kind": kind, "metrics_percent": values}
                rss = {role: rss_for(args.root, f"{role}-{mode}-{payload}-{size}") for role in ROLES}
                rss_values = {role: value[0] for role, value in rss.items()}
                rss_deltas = {}
                for pair, (kind, later, earlier) in pairs.items():
                    value = delta(rss_values[later], rss_values[earlier])
                    rss_deltas[pair] = value
                    if value > 5.0:
                        triggers.append({"row": label, "comparison_kind": kind, "pair": pair, "metric": "max_rss_kib", "regression_percent": value})
                rows.append({
                    "mode": mode, "payload": payload, "input_bytes": size,
                    "sample_count": args.samples, "warmups": args.warmups, "raw_sample_unit": "ns",
                    "legs": {role: {"path": row["_path"].relative_to(args.root).as_posix(), "sha256": row["_sha256"], "statistics_ns": row["_stats"]} for role, row in by_role.items()},
                    "output_oracle": {"bytes": identity[0], "writes": identity[1], "crc32": identity[2]},
                    "paired_delta_percent": pair_deltas,
                    "time_v_max_rss_kib": {role: value[0] for role, value in rss.items()},
                    "time_v_paths": {role: value[1].relative_to(args.root).as_posix() for role, value in rss.items()},
                    "time_v_rss_delta_percent": rss_deltas,
                })
    result = {
        "schema_version": 1, "tool": "litchi-goal-0415-guard-summary", "abba_order": list(ROLES),
        "statistics": {"p50": "exact median (integer or .5)", "p95": "integer nearest rank", "p99": "integer nearest rank"},
        "threshold": {"regression_percent": 5.0, "rule": "record every positive candidate regression, within-revision drift, or RSS delta above threshold"},
        "rows": rows, "regression_triggers": triggers,
        "verification": {"row_count": len(rows), "all_input_identity_verified": True, "all_output_oracles_equal_per_row": True, "all_time_v_rss_present": True, "raw_samples_retained_in_input_files": True},
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n")

if __name__ == "__main__":
    try:
        main()
    except ValueError as error:
        raise SystemExit(f"error: {error}") from error
