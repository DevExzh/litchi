#!/usr/bin/env python3
"""Run adversarial mutations against the independent 0427 report verifier.

Every mutation is written below a temporary directory and checked through the
verifier CLI.  The supplied report is opened read-only and is never replaced.
"""

from __future__ import annotations

import argparse
import copy
import json
import re
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any, Callable


Mutation = Callable[[dict[str, Any]], None]


def report_samples(report: dict[str, Any]) -> list[dict[str, Any]]:
    rows = report.get("samples_raw")
    if not isinstance(rows, list) or not rows:
        raise ValueError("report needs at least one retained row for mutation probes")
    return rows


def target_sample(report: dict[str, Any]) -> dict[str, Any]:
    rows = report_samples(report)
    return rows[1] if len(rows) > 1 else rows[0]


def checkpoint(row: dict[str, Any], name: str) -> dict[str, Any]:
    value = row.get(name)
    if not isinstance(value, dict):
        raise ValueError(f"missing checkpoint {name}")
    return value


def mutate_missing_boundary(report: dict[str, Any]) -> None:
    phases = report["phases"]
    del phases[1]


def mutate_reordered_boundary(report: dict[str, Any]) -> None:
    phases = report["phases"]
    phases[0], phases[1] = phases[1], phases[0]


def mutate_owner_label(report: dict[str, Any]) -> None:
    phases = report["phases"]
    phases[1]["live_owners"] = phases[1]["live_owners"] + " [mutated]"


def mutate_impossible_live_counter(report: dict[str, Any]) -> None:
    row = target_sample(report)
    name = "prepared_inputs_and_sink"
    checkpoint(row, name)["live_bytes"] = -1


def mutate_missing_numeric(report: dict[str, Any]) -> None:
    checkpoint(target_sample(report), "prepared_inputs_and_sink").pop("live_bytes")


def mutate_sample_order(report: dict[str, Any]) -> None:
    rows = report_samples(report)
    if len(rows) > 1:
        rows[1]["sample_index"] = rows[0]["sample_index"]
    else:
        rows[0]["sample_index"] = 1


def mutate_region_peak(report: dict[str, Any]) -> None:
    row = target_sample(report)
    checkpoints = [
        row[name]
        for name in (
            "baseline_before_inputs",
            "prepared_inputs_and_sink",
            "opened_documents",
            "planned",
            "published",
            "drop_result",
            "drop_plan",
            "drop_document_handles",
            "drop_caller_source_arcs",
            "drop_sink",
        )
        if name in row
    ]
    maximum = max(item["live_bytes"] for item in checkpoints)
    row["retention_probe"]["retention_probe_region_peak_live_bytes"] = maximum - 1 if maximum else -1


def mutate_absolute_baseline_offset(report: dict[str, Any]) -> None:
    for row in report_samples(report):
        for name in (
            "baseline_before_inputs",
            "prepared_inputs_and_sink",
            "opened_documents",
            "planned",
            "published",
            "drop_result",
            "drop_plan",
            "drop_document_handles",
            "drop_caller_source_arcs",
            "drop_sink",
        ):
            if name in row:
                row[name]["live_bytes"] += 1
                row[name]["peak_live_bytes"] += 1
        probe = row["retention_probe"]
        probe["live_bytes_before"] += 1
        probe["live_bytes_after"] += 1
        probe["peak_live_bytes_before"] += 1
        probe["peak_live_bytes_after"] += 1
        probe["retention_probe_region_peak_live_bytes"] += 1


def mutate_u64_overflow(report: dict[str, Any]) -> None:
    report_samples(report)[0]["baseline_before_inputs"]["allocated_bytes"] = 1 << 64


def mutate_missing_region_numeric(report: dict[str, Any]) -> None:
    report_samples(report)[0]["retention_probe"].pop("retention_probe_region_peak_live_bytes")


def mutate_missing_arc_count(report: dict[str, Any]) -> None:
    row = report_samples(report)[0]
    if "source_arc_counts_after_document_drop" in row:
        del row["source_arc_counts_after_document_drop"]
    else:
        row["source_arc_counts_after_document_drop"] = None


def mutate_output_flag(report: dict[str, Any]) -> None:
    report["all_iteration_output_bytes_verified"] = False


def mutate_output_hash(report: dict[str, Any]) -> None:
    report["expected_output_sha256"] = "0" * 63 + "z"


def mutate_gate(report: dict[str, Any]) -> None:
    report["corpus_gates_verified"] = False


def mutate_counter_revision(report: dict[str, Any]) -> None:
    report["allocator_counter_revision"] = "obsolete_counter_revision"


def mutate_counter_scope(report: dict[str, Any]) -> None:
    report["samples_raw"][0]["retention_probe"]["scope"] = "wrong_scope"


def mutate_callback_scope(report: dict[str, Any]) -> None:
    report["callback_scope"] = "wrong callback scope"


def mutate_ownership_scope(report: dict[str, Any]) -> None:
    report["ownership_scope"] = "wrong ownership scope"


def duplicate_json_key(original: bytes) -> bytes:
    text = original.decode("utf-8")
    opening = text.find("{")
    if opening < 0:
        raise ValueError("report JSON has no object root")
    return (text[: opening + 1] + '\n  "schema": "pptx_retention_v1",\n  "schema": "pptx_retention_v1",' + text[opening + 1 :]).encode("utf-8")


def nonfinite_json_constant(original: bytes) -> bytes:
    text = original.decode("utf-8")
    mutated, count = re.subn(r'("binary_bytes"\s*:\s*)[0-9]+', r"\1NaN", text, count=1)
    if count != 1:
        raise ValueError("report JSON has no binary_bytes number")
    return mutated.encode("utf-8")


MUTATIONS: tuple[tuple[str, Mutation], ...] = (
    ("missing_boundary", mutate_missing_boundary),
    ("reordered_boundary", mutate_reordered_boundary),
    ("owner_label", mutate_owner_label),
    ("impossible_live_counter", mutate_impossible_live_counter),
    ("missing_numeric", mutate_missing_numeric),
    ("sample_order", mutate_sample_order),
    ("region_peak_below_endpoint", mutate_region_peak),
    ("absolute_baseline_offset", mutate_absolute_baseline_offset),
    ("u64_overflow", mutate_u64_overflow),
    ("missing_region_numeric", mutate_missing_region_numeric),
    ("missing_arc_count", mutate_missing_arc_count),
    ("output_equality_flag", mutate_output_flag),
    ("output_hash_shape", mutate_output_hash),
    ("corpus_gate", mutate_gate),
    ("counter_revision", mutate_counter_revision),
    ("counter_scope", mutate_counter_scope),
    ("callback_scope", mutate_callback_scope),
    ("ownership_scope", mutate_ownership_scope),
)
RAW_MUTATIONS = ("duplicate_json_key", "nonfinite_json_constant")


def run_verifier(verifier: Path, report: Path, protocol: Path | None) -> subprocess.CompletedProcess[str]:
    command = [sys.executable, str(verifier), "--report", str(report)]
    if protocol is not None:
        command.extend(["--protocol", str(protocol)])
    return subprocess.run(command, capture_output=True, text=True, check=False)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("path", nargs="?", type=Path, help="retention report JSON")
    parser.add_argument("--report", dest="report_option", type=Path, help="retention report JSON")
    parser.add_argument("--verifier", type=Path, default=Path(__file__).with_name("verify-report.py"))
    parser.add_argument("--protocol", type=Path, default=None)
    args = parser.parse_args(argv)
    path = args.report_option or args.path
    if path is None:
        parser.error("a report path is required")
    try:
        original_bytes = path.read_bytes()
        original = json.loads(original_bytes.decode("utf-8"))
        if not isinstance(original, dict):
            raise ValueError("report root must be an object")
    except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as exc:
        print(f"INVALID INPUT: {exc}", file=sys.stderr)
        return 1

    baseline = run_verifier(args.verifier, path, args.protocol)
    if baseline.returncode != 0:
        print("BASELINE REJECTED", file=sys.stderr)
        if baseline.stderr:
            print(baseline.stderr.rstrip(), file=sys.stderr)
        return 1

    results: list[dict[str, Any]] = []
    with tempfile.TemporaryDirectory(prefix="litchi-0427-report-probes-") as temporary:
        directory = Path(temporary)
        for name, mutation in MUTATIONS:
            candidate = copy.deepcopy(original)
            try:
                mutation(candidate)
            except (KeyError, TypeError, ValueError, IndexError) as exc:
                print(f"MUTATION {name} COULD NOT BE APPLIED: {exc}", file=sys.stderr)
                return 1
            target = directory / f"{name}.json"
            target.write_text(json.dumps(candidate, indent=2) + "\n", encoding="utf-8")
            result = run_verifier(args.verifier, target, args.protocol)
            rejected = result.returncode != 0 and result.stderr.lstrip().startswith("INVALID:")
            results.append({"mutation": name, "rejected": rejected})
            if not rejected:
                print(f"MUTATION DID NOT PRODUCE STRUCTURED INVALID: {name}", file=sys.stderr)
                if result.stderr:
                    print(result.stderr.rstrip(), file=sys.stderr)
                return 1
        for name in RAW_MUTATIONS:
            target = directory / f"{name}.json"
            raw = duplicate_json_key(original_bytes) if name == "duplicate_json_key" else nonfinite_json_constant(original_bytes)
            target.write_bytes(raw)
            result = run_verifier(args.verifier, target, args.protocol)
            rejected = result.returncode != 0 and result.stderr.lstrip().startswith("INVALID:")
            results.append({"mutation": name, "rejected": rejected})
            if not rejected:
                print(f"MUTATION DID NOT PRODUCE STRUCTURED INVALID: {name}", file=sys.stderr)
                if result.stderr:
                    print(result.stderr.rstrip(), file=sys.stderr)
                return 1

    if path.read_bytes() != original_bytes:
        print("ORIGINAL REPORT CHANGED", file=sys.stderr)
        return 1
    print(json.dumps({"status": "pass", "mutations": results}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
