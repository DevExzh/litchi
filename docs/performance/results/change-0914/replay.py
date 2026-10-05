#!/usr/bin/env python3
"""Replay retained statistics and receipt links without external artifacts.

This does not rebuild or inspect the external executable/source archive.
Use capture.py verify for that additional custody check.
"""

from __future__ import annotations

import argparse
import datetime
import hashlib
import json
import math
from pathlib import Path
import statistics
import sys

import capture


def check(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def bound(path: Path, identity: dict) -> bytes:
    data = path.read_bytes()
    check(len(data) == identity["bytes"], f"size mismatch: {path}")
    check(hashlib.sha256(data).hexdigest() == identity["sha256"], f"hash mismatch: {path}")
    return data


def statistics_for(values: list[int]) -> dict:
    ordered = sorted(values)
    count = len(values)
    return {
        "samples": count,
        "min": ordered[0],
        "p50": (ordered[(count - 1) // 2] + ordered[count // 2]) // 2,
        "p95": ordered[math.ceil(0.95 * count) - 1],
        "p99": ordered[math.ceil(0.99 * count) - 1],
        "max": ordered[-1],
        "mean": statistics.fmean(values),
    }


def replay(packet: Path) -> dict:
    summary_bytes = (packet / "reduced.json").read_bytes()
    summary = json.loads(summary_bytes)
    receipts = packet / "receipts"
    source = json.loads(bound(receipts / "source-v1.json", summary["source_receipt"]))
    protocol = json.loads(bound(receipts / "protocol-v1.json", summary["protocol"]))
    build = json.loads(bound(receipts / "build-v1.json", summary["build"]))
    check(build["exit_code"] == 0, "build failed")
    check(build["source"] == source, "build/source identity mismatch")
    check(build["binary"] == summary["binary"], "binary identity mismatch")
    check(source["source_revision"] == protocol["source_revision"] ==
          build["source_revision"] == summary["source_revision"], "revision mismatch")
    check(summary["case"] == protocol["case"] == capture.CASE, "case mismatch")
    check(summary["performance_claim"] == protocol["performance_claim"] == "none", "claim mismatch")
    check(summary["corpus"] == capture.EXPECTED, "corpus mismatch")
    check(summary["children"] == 4 and summary["samples_per_child"] == capture.SAMPLES and
          summary["warmup_iterations"] == capture.WARMUP, "sampling mismatch")
    check(protocol["samples"] == capture.SAMPLES and
          protocol["warmup_iterations"] == capture.WARMUP, "protocol sampling mismatch")
    for field in ("timing_scope", "process_scope", "allocation_metrics", "physical_io_metrics", "cpu"):
        check(summary[field] == protocol[field], f"scope mismatch: {field}")
    labels = [f"child-{letter}" for letter in "abcd"]
    check(protocol["children"] == labels, "protocol children mismatch")
    check([child["label"] for child in summary["children_detail"]] == labels, "children mismatch")
    totals: list[int] = []
    phases = {key: [] for key in summary["pooled_phase_descriptive_only"]}
    medians = []
    for child in summary["children_detail"]:
        label = child["label"]
        receipt = json.loads(bound(receipts / f"{label}.receipt.json", child["receipt"]))
        report = json.loads(bound(packet / "reports" / f"{label}.report.json", child["report"]))
        check(receipt["exit_code"] == 0 and receipt["label"] == label, f"{label}: failed child")
        for field in ("protocol", "build", "binary", "cpu"):
            check(receipt[field] == summary[field], f"{label}: changed {field}")
        check(receipt["report"] == child["report"], f"{label}: report binding mismatch")
        start = datetime.datetime.fromisoformat(receipt["started_utc"])
        finish = datetime.datetime.fromisoformat(receipt["finished_utc"])
        check(start < finish, f"{label}: invalid interval")
        for recorded in (source["created_utc"], protocol["created_utc"], build["finished_utc"]):
            check(datetime.datetime.fromisoformat(recorded) <= start, f"{label}: metadata postdates capture")
        process_path = receipts / f"{label}.process.txt"
        bound(process_path, receipt["process"])
        process = capture.parse_time_v(process_path)
        process["path"] = receipt["process_metrics"]["path"]
        check(process == receipt["process_metrics"] == child["process"], f"{label}: process metrics mismatch")
        observed = capture.validate_report(report, label, summary["binary"])
        for key, value in observed.items():
            check(value == child[key], f"{label}: reduction mismatch: {key}")
        result = report["results"][0]
        samples = result["elapsed_ns"]["samples"]
        independent = statistics_for(samples)
        check(independent == child["elapsed_ns"], f"{label}: independent statistics mismatch")
        medians.append(independent["p50"])
        totals.extend(samples)
        for key in phases:
            phases[key].extend(result["source"]["docx_section_layout"][key])
    check(statistics_for(totals) == summary["pooled_elapsed_descriptive_only"], "pooled statistics mismatch")
    check(medians == summary["child_elapsed_p50_ns"], "child medians mismatch")
    check({"min": min(medians), "max": max(medians)} == summary["child_elapsed_p50_range_ns"], "median range mismatch")
    for key, values in phases.items():
        check(statistics_for(values) == summary["pooled_phase_descriptive_only"][key], f"phase statistics mismatch: {key}")
        check(sum(values) / sum(totals) == summary["pooled_phase_fraction_of_elapsed"][key], f"phase share mismatch: {key}")
    check(sum(map(sum, phases.values())) == sum(totals), "pooled phase conservation failed")
    guards = (
        "all_children_exit_zero", "all_reports_match_binary", "all_reports_match_source_corpus",
        "all_reports_match_warmup_and_samples", "all_semantic_and_preservation_gates",
        "all_source_and_sink_gates", "all_output_hashes_stable", "no_speedup_claim",
    )
    check(summary["guards"] == dict.fromkeys(guards, True), "summary guard failed")
    return {"status": "verified", "scope": "retained reports, receipts and statistics; external artifacts not inspected",
            "children": len(labels), "samples": len(totals), "reduced_sha256": hashlib.sha256(summary_bytes).hexdigest()}


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--packet", type=Path, default=Path(__file__).resolve().parent)
    args = parser.parse_args()
    try:
        print(json.dumps(replay(args.packet), sort_keys=True))
    except (OSError, ValueError, KeyError, TypeError, capture.EvidenceError) as error:
        print(f"replay failed: {error}", file=sys.stderr)
        sys.exit(1)
