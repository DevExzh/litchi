#!/usr/bin/env python3
"""Rebuild the bounded 0945 summary from retained stage and reduction receipts.

This reads JSON and hashes referenced descriptors only; it does not launch a
workload, test, build, archive, or cleanup command.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
from typing import Any

SCHEMA = "litchi.performance.0945-summary.v1"
CAPTURE_SCHEMA = "litchi.paired-audit-capture-0945.v1"
CASES = (
    "primary-percent-medium",
    "primary-percent-dense-sparse",
    "control-xlsx-one-edit-medium",
    "control-xlsx-managed-one-edit-medium",
)
STAGES = {
    "qualification": {"children": 16, "samples": 2, "warmup": 0},
    "native": {"children": 48, "samples": 100, "warmup": 10},
    "allocator": {"children": 16, "samples": 3, "warmup": 2},
}
QUALITY_FILES = (
    "quality-v2/quality-audit-final-v1.json",
    "quality-v2/quality-summary-followup-v1.json",
    "quality-v2/quality-inputs-v2.json",
    "quality-v2/source-freeze-v2.json",
    "receipts/root-quality-admission-v1.json",
)
SAFETY_METRICS = (
    "allocation_calls",
    "region_peak_live_bytes",
    "live_bytes_after",
    "max_rss_kib",
)


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def descriptor(root: Path, rel: str) -> dict[str, Any]:
    p = root / rel
    if not p.is_file() or p.is_symlink():
        raise AssertionError(f"missing or non-regular descriptor: {rel}")
    st = p.stat()
    return {"path": rel, "bytes": st.st_size, "sha256": sha256(p)}


def recorded(root: Path, item: dict[str, Any], label: str) -> dict[str, Any]:
    rel = item.get("path")
    if not isinstance(rel, str):
        raise AssertionError(f"{label}: descriptor path missing")
    actual = descriptor(root, rel)
    for key in ("bytes", "sha256"):
        if item.get(key) != actual[key]:
            raise AssertionError(
                f"{label}: {key} differs for {rel}: {item.get(key)!r} != {actual[key]!r}"
            )
    return actual


def load(root: Path, rel: str) -> dict[str, Any]:
    with (root / rel).open() as f:
        value = json.load(f)
    if not isinstance(value, dict):
        raise AssertionError(f"{rel}: expected JSON object")
    return value


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def native_rows(analysis: dict[str, Any]) -> dict[str, Any]:
    reduction = analysis["reduction"]
    rows: dict[str, Any] = {}
    for case in CASES:
        raw = reduction["cases"][case]
        metrics: dict[str, Any] = {}
        for metric, ratio_key in (
            ("publication_p50", "publication_ratio"),
            ("whole_workflow_p50", "whole_workflow_ratio"),
        ):
            boot = raw[metric]
            metrics[metric] = {
                "pair_ratios": [pair[ratio_key] for pair in boot["pairs"]],
                "bootstrap": {
                    "estimate": boot["estimate"],
                    "lower_95": boot["lower_95"],
                    "upper_95": boot["upper_95"],
                    "replicates": boot["replicates"],
                    "seed": boot["seed"],
                    "statistic": boot["statistic"],
                    "accepted": boot["accepted"],
                },
            }
        rows[case] = metrics
    return rows


def allocation_rows(analysis: dict[str, Any]) -> dict[str, Any]:
    source = analysis["allocation_reduction"]
    rows: dict[str, Any] = {}
    for case in CASES:
        raw = source["cases"][case]
        pairs = []
        for pair in raw["pairs"]:
            pairs.append(
                {
                    "repeat": pair["repeat"],
                    "ratios": {
                        "allocated_bytes": pair["metrics"]["allocated_bytes"]["ratio"],
                        **{
                            metric: pair["metrics"][metric]["ratio"]
                            for metric in SAFETY_METRICS
                        },
                    },
                }
            )
        rows[case] = {
            "primary_every_pair_gate_pass": raw["primary_every_pair_gate_pass"],
            "publication_allocated_bytes_pair_ratios": raw[
                "publication_allocated_bytes_pair_ratios"
            ],
            "pairs": pairs,
            "safety": raw["safety"],
        }
    return rows


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--packet", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    root = args.packet.resolve()
    output = (args.output or root / "summary-v1.json").resolve()
    require(root.is_dir(), f"packet root does not exist: {root}")

    analysis = load(root, "analysis.json")
    plan = load(root, "plan-v2.json")
    captures = {stage: load(root, f"{stage}/capture.json") for stage in STAGES}
    require(analysis["schema"] == "litchi.paired-audit-reduction-0945.v1", "analysis schema")
    require(analysis["stage"] == "all" and analysis["status"] == "pass", "analysis status")
    require(plan["schema"] == CAPTURE_SCHEMA and plan["status"] == "frozen", "plan binding")
    require(plan["plan_id"] == "0945" and plan["seed"] == 945009, "plan identity")

    stage_summary: dict[str, Any] = {}
    total_children = 0
    total_samples = 0
    source_refs: dict[str, Any] | None = None
    binary_bindings: dict[str, Any] | None = None
    capture_descriptors: dict[str, Any] = {}
    for stage, expected in STAGES.items():
        capture = captures[stage]
        require(capture["schema"] == CAPTURE_SCHEMA, f"{stage}: capture schema")
        require(capture["stage"] == stage and capture["status"] == "pass", f"{stage}: status")
        children = capture["children"]
        jobs = capture["jobs"]
        receipts = capture["receipts"]
        require(children == expected["children"], f"{stage}: child count")
        require(len(jobs) == children and len(receipts) == children, f"{stage}: receipt count")
        expected_samples = children * expected["samples"]
        require(capture["samples_per_child"] == expected["samples"], f"{stage}: samples/child")
        require(capture["warmup_per_child"] == expected["warmup"], f"{stage}: warmup")
        require(capture["total_samples"] == expected_samples, f"{stage}: total samples")
        require(all(job["stage"] == stage for job in jobs), f"{stage}: job stage")
        source = capture["source_start"]
        require(source["full"] is True, f"{stage}: source closure is not full")
        require(source["checked_files"] == {"baseline": 10788, "candidate": 10788}, f"{stage}: source count")
        if source_refs is None:
            source_refs = source
            binary_bindings = capture["binary_bindings"]
        else:
            require(source == source_refs, f"{stage}: source binding differs")
            require(capture["binary_bindings"] == binary_bindings, f"{stage}: binary binding differs")
        capture_descriptors[stage] = descriptor(root, f"{stage}/capture.json")
        stage_summary[stage] = {
            "children": children,
            "jobs": len(jobs),
            "receipts": len(receipts),
            "samples_per_child": expected["samples"],
            "warmup_per_child": expected["warmup"],
            "total_samples": expected_samples,
        }
        total_children += children
        total_samples += expected_samples
    require(total_children == 80 and total_samples == 4880, "aggregate population")

    # Confirm source manifest descriptors and the plan's recorded references.
    plan_source = plan["source"]
    source_descriptors = {
        "baseline_manifest": recorded(root, plan_source["baseline_manifest"], "baseline manifest"),
        "candidate_manifest": recorded(root, plan_source["candidate_manifest"], "candidate manifest"),
    }
    require(source_refs["baseline_manifest"] == plan_source["baseline_manifest"], "baseline source binding")
    require(source_refs["candidate_manifest"] == plan_source["candidate_manifest"], "candidate source binding")

    # Check each retained binary/build descriptor against the file on disk.
    binaries: dict[str, Any] = {}
    for role in ("baseline", "candidate"):
        binaries[role] = {}
        for kind in ("native", "allocator"):
            bound = binary_bindings[role][kind]
            binaries[role][kind] = {
                "binary": recorded(root, bound["binary"], f"{role}/{kind} binary"),
                "build_receipt": recorded(root, bound["build_receipt"], f"{role}/{kind} build"),
            }

    quality: dict[str, Any] = {}
    for rel in QUALITY_FILES:
        doc = load(root, rel)
        quality[rel] = {"descriptor": descriptor(root, rel), "schema": doc.get("schema"), "status": doc.get("status")}
    adoption = load(root, "receipts/root-adoption-v1.json")
    quality["receipts/root-adoption-v1.json"] = {
        "descriptor": descriptor(root, "receipts/root-adoption-v1.json"),
        "schema": adoption.get("schema"),
        "decision": adoption.get("decision"),
        "frozen_gates_pass": adoption.get("frozen_gates_pass"),
    }

    result = {
        "schema": SCHEMA,
        "status": "pass",
        "scope": "bounded projection of retained 0945 analysis and stage receipts; no execution",
        "generator": descriptor(root, os.path.relpath(Path(__file__).resolve(), root)),
        "inputs": {
            "analysis": descriptor(root, "analysis.json"),
            "plan": descriptor(root, "plan-v2.json"),
            "captures": capture_descriptors,
        },
        "population": {
            "children": total_children,
            "samples": total_samples,
            "stages": stage_summary,
        },
        "protocol": {
            "seed": plan["seed"],
            "native": plan["native"],
            "allocator": plan["allocator"],
            "qualification": plan["qualification"],
            "bootstrap": plan["gate"]["bootstrap"],
            "latency_gate": plan["gate"]["latency"],
            "allocation_gate": plan["gate"]["allocation"],
        },
        "source": {
            "manifests": source_descriptors,
            "candidate_files": plan_source["candidate_files"],
            "checked_files": source_refs["checked_files"],
            "full": source_refs["full"],
        },
        "binaries": binaries,
        "quality": quality,
        "native": {
            "all_case_latency_gate_pass": analysis["reduction"]["all_case_latency_gate_pass"],
            "rows": native_rows(analysis),
            "review_flags": analysis["reduction"]["review_flags"],
        },
        "allocator": {
            "all_row_safety_gate_pass": analysis["allocation_reduction"]["all_row_safety_gate_pass"],
            "primary_every_pair_gate_pass": analysis["allocation_reduction"]["primary_every_pair_gate_pass"],
            "rows": allocation_rows(analysis),
            "review_flags": analysis["allocation_reduction"]["review_flags"],
        },
        "analysis_status": analysis["status"],
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n")
    print(json.dumps({"status": "pass", "output": str(output), "bytes": output.stat().st_size}, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (AssertionError, KeyError, OSError, json.JSONDecodeError) as exc:
        raise SystemExit(f"summary verification failed: {exc}")
