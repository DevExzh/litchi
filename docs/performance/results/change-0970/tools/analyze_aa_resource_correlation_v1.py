#!/usr/bin/env python3
"""Offline p50/resource correlation for four retained 0969 A/A cells.

This addendum reads existing packet JSON and does not launch a workload,
binary, build, test, or reducer.  Correlations are descriptive associations;
the fixed six-child tails are rank quartiles, not value thresholds.
"""
from __future__ import annotations

import hashlib
import json
import math
import statistics
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


ROOT = Path("/home/zhuhe/litchi-performance-0969")
RAW_ROOT = ROOT / "capture/aa-v4/aa"
PRIOR_AUDIT = ROOT / "analysis/aa-independent-audit-v1.json"
OUT_JSON = Path("/home/zhuhe/litchi-performance-0970/analysis/aa-resource-correlation-v1.json")
OUT_MD = Path("/home/zhuhe/litchi-performance-0970/analysis/aa-resource-correlation-v1.md")

SELECTED_CELLS = (
    "docfloat/e1/default",
    "docfloat/e32/default",
    "docnohf/e1/default",
    "docnohf/e32/default",
)
EXPECTED_PRIOR_AUDIT_SHA256 = "f0d0d0b63a896be41dab84550ff39f59954c26538423be2a3d3d017df9f26775"
EXPECTED_RAW_DIGEST = "3eecb0db82fb14be38852685956dd6fa485fad2c609b23261fa5c75443befdd8"
EXPECTED_RAW_FILES = 2304
EXPECTED_PACKETS = 576
EXPECTED_SAMPLES_PER_CHILD = 80


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def raw_set_digest(paths: list[Path], root: Path) -> str:
    rows = sorted(f"{sha256_file(path)}  {path.relative_to(root).as_posix()}\n" for path in paths)
    return hashlib.sha256("".join(rows).encode()).hexdigest()


def average_rank(values: list[float]) -> list[float]:
    ordered = sorted((value, index) for index, value in enumerate(values))
    ranks = [0.0] * len(values)
    start = 0
    while start < len(ordered):
        end = start + 1
        while end < len(ordered) and ordered[end][0] == ordered[start][0]:
            end += 1
        rank = (start + end - 1) / 2.0 + 1.0
        for _, index in ordered[start:end]:
            ranks[index] = rank
        start = end
    return ranks


def pearson(xs: list[float], ys: list[float]) -> float | None:
    if len(xs) != len(ys) or len(xs) < 2:
        return None
    mean_x, mean_y = statistics.mean(xs), statistics.mean(ys)
    numerator = sum((x - mean_x) * (y - mean_y) for x, y in zip(xs, ys))
    denominator = math.sqrt(sum((x - mean_x) ** 2 for x in xs) * sum((y - mean_y) ** 2 for y in ys))
    return numerator / denominator if denominator else None


def correlations(xs: list[float], ys: list[float]) -> dict[str, Any]:
    result = {
        "pearson_r": pearson(xs, ys),
        "spearman_rho": pearson(average_rank(xs), average_rank(ys)),
        "n_children": len(xs),
    }
    if result["pearson_r"] is None:
        result["undefined_reason"] = "one or both variables are constant"
    return result


def range_summary(values: list[float | int]) -> dict[str, Any]:
    vals = [float(value) for value in values]
    return {
        "n": len(vals),
        "min": min(vals),
        "median": statistics.median(vals),
        "max": max(vals),
    }


def metric_ranges(rows: list[dict[str, Any]]) -> dict[str, Any]:
    return {
        "elapsed_p50_ns": range_summary([row["elapsed_p50_ns"] for row in rows]),
        "rss_kb": range_summary([row["rss_kb"] for row in rows]),
        "whole_child_wall_ns": range_summary([row["whole_child_wall_ns"] for row in rows]),
        "output_capacity_bytes": range_summary([row["output_capacity_bytes"] for row in rows]),
        "output_bytes": range_summary([row["output_bytes"] for row in rows]),
    }


def quartile_tail(rows: list[dict[str, Any]], label: str) -> dict[str, Any]:
    ordered = sorted(rows, key=lambda row: (row["elapsed_p50_ns"], row["packet_id"]))
    tail_count = len(ordered) // 4
    selected = ordered[:tail_count] if label == "fastest_p50_quartile" else ordered[-tail_count:]
    return {
        "label": label,
        "children": len(selected),
        "measured_sample_count": len(selected) * EXPECTED_SAMPLES_PER_CHILD,
        "ranges": metric_ranges(selected),
        "median_rss_kb": statistics.median(row["rss_kb"] for row in selected),
        "median_whole_child_wall_ns": statistics.median(row["whole_child_wall_ns"] for row in selected),
        "median_output_capacity_bytes": statistics.median(row["output_capacity_bytes"] for row in selected),
    }


def correlation_block(rows: list[dict[str, Any]]) -> dict[str, Any]:
    elapsed = [float(row["elapsed_p50_ns"]) for row in rows]
    return {
        "p50_vs_rss": correlations(elapsed, [float(row["rss_kb"]) for row in rows]),
        "p50_vs_whole_child_wall": correlations(elapsed, [float(row["whole_child_wall_ns"]) for row in rows]),
        "p50_vs_output_capacity": correlations(elapsed, [float(row["output_capacity_bytes"]) for row in rows]),
        "p50_vs_output_bytes": correlations(elapsed, [float(row["output_bytes"]) for row in rows]),
    }


def revalidate_inputs(errors: list[str], prior: dict[str, Any]) -> dict[str, Any]:
    audit_actual = sha256_file(PRIOR_AUDIT) if PRIOR_AUDIT.is_file() else None
    raw_files = sorted(path for path in RAW_ROOT.iterdir() if path.is_file()) if RAW_ROOT.is_dir() else []
    packet_paths = sorted(RAW_ROOT.glob("*.json")) if RAW_ROOT.is_dir() else []
    raw_digest = raw_set_digest(raw_files, RAW_ROOT) if RAW_ROOT.is_dir() else None
    result = {
        "prior_audit": {
            "path": str(PRIOR_AUDIT),
            "sha256": audit_actual,
            "expected_sha256": EXPECTED_PRIOR_AUDIT_SHA256,
            "match": audit_actual == EXPECTED_PRIOR_AUDIT_SHA256,
            "error_count": prior.get("error_count"),
            "complete_valid_packet_set": prior.get("complete_valid_packet_set"),
        },
        "raw_packet_set": {
            "root": str(RAW_ROOT),
            "file_count": len(raw_files),
            "expected_file_count": EXPECTED_RAW_FILES,
            "set_sha256": raw_digest,
            "expected_set_sha256": EXPECTED_RAW_DIGEST,
            "json_packet_count": len(packet_paths),
            "expected_json_packet_count": EXPECTED_PACKETS,
            "total_bytes": sum(path.stat().st_size for path in raw_files),
            "match": len(raw_files) == EXPECTED_RAW_FILES and raw_digest == EXPECTED_RAW_DIGEST and len(packet_paths) == EXPECTED_PACKETS,
        },
    }
    if not result["prior_audit"]["match"] or prior.get("error_count") != 0 or prior.get("complete_valid_packet_set") is not True:
        errors.append("prior audit binding is not complete and valid")
    if not result["raw_packet_set"]["match"]:
        errors.append("raw packet set binding changed")
    return result


def load_selected_rows(errors: list[str]) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    selected_counts = {cell: 0 for cell in SELECTED_CELLS}
    binary_hashes: set[str] = set()
    fixture_hashes: set[str] = set()
    source_manifest_hashes: set[str] = set()
    for path in sorted(RAW_ROOT.glob("*.json")):
        packet = load_json(path)
        context = packet.get("context") or {}
        cell = (context.get("cell") or {}).get("cell_id")
        if cell not in SELECTED_CELLS:
            continue
        packet_id = packet.get("packet_id")
        report = packet.get("report") or {}
        elapsed = report.get("elapsed_ns")
        rss = packet.get("rss") or {}
        wall = packet.get("wall_ns_excluded_from_metric")
        capacity = report.get("output_capacity")
        output_bytes = report.get("output_bytes")
        if path.stem != packet_id or context.get("packet_id") != packet_id:
            errors.append(f"packet identity mismatch: {path.name}")
        if packet.get("role") != "aa_baseline" or packet.get("schema") != "litchi.performance.0969.runner.v4":
            errors.append(f"packet role/schema mismatch: {packet_id}")
        if (packet.get("terminal") or {}).get("returncode") != 0:
            errors.append(f"packet terminal mismatch: {packet_id}")
        if not isinstance(elapsed, list) or len(elapsed) != EXPECTED_SAMPLES_PER_CHILD or any(type(value) is not int or value <= 0 for value in elapsed):
            errors.append(f"elapsed sample mismatch: {packet_id}")
            continue
        if type(rss.get("kb")) is not int or rss["kb"] <= 0:
            errors.append(f"RSS mismatch: {packet_id}")
            continue
        if type(wall) is not int or wall <= 0:
            errors.append(f"whole-child wall mismatch: {packet_id}")
            continue
        if type(capacity) is not int or capacity < 0 or type(output_bytes) is not int or output_bytes < 0:
            errors.append(f"output size mismatch: {packet_id}")
            continue
        binary_hashes.add(str(context.get("binary_sha256")))
        fixture_hashes.add(str(context.get("fixture_sha256")))
        source_manifest_hashes.add(str(context.get("source_manifest_sha256")))
        selected_counts[cell] += 1
        rows.append({
            "packet_id": packet_id,
            "cell": cell,
            "elapsed_p50_ns": statistics.median(elapsed),
            "elapsed_sample_count": len(elapsed),
            "rss_kb": rss["kb"],
            "whole_child_wall_ns": wall,
            "output_capacity_bytes": capacity,
            "output_bytes": output_bytes,
        })
    if any(count != 24 for count in selected_counts.values()):
        errors.append(f"selected cell coverage mismatch: {selected_counts}")
    if len(rows) != 96 or len({row["packet_id"] for row in rows}) != len(rows):
        errors.append("selected packet population is not 96 unique children")
    return rows, {
        "selected_cells": selected_counts,
        "binary_sha256_values": sorted(binary_hashes),
        "fixture_sha256_values": sorted(fixture_hashes),
        "source_manifest_sha256_values": sorted(source_manifest_hashes),
    }


def analyze(rows: list[dict[str, Any]]) -> dict[str, Any]:
    per_cell = []
    for cell in SELECTED_CELLS:
        group = [row for row in rows if row["cell"] == cell]
        per_cell.append({
            "cell": cell,
            "children": len(group),
            "measured_sample_count": sum(row["elapsed_sample_count"] for row in group),
            "ranges": metric_ranges(group),
            "correlations": correlation_block(group),
            "output_capacity_unique_values": sorted({row["output_capacity_bytes"] for row in group}),
            "output_bytes_unique_values": sorted({row["output_bytes"] for row in group}),
            "quartile_tails": [quartile_tail(group, "fastest_p50_quartile"), quartile_tail(group, "slowest_p50_quartile")],
        })
    pooled = {
        "children": len(rows),
        "measured_sample_count": sum(row["elapsed_sample_count"] for row in rows),
        "ranges": metric_ranges(rows),
        "correlations": correlation_block(rows),
        "interpretation": "Pooled correlations are descriptive and include between-cell composition; per-cell rows are the within-cell comparison.",
    }
    return {"pooled": pooled, "cells": per_cell}


def fmt(value: Any, digits: int = 3) -> str:
    if value is None:
        return "undefined"
    return f"{value:.{digits}f}"


def markdown(result: dict[str, Any]) -> str:
    revalidation = result["revalidation"]
    pooled = result["analysis"]["pooled"]
    cells = result["analysis"]["cells"]
    lines = [
        "# 0969 selected-cell resource correlation addendum",
        "",
        "This is a bounded descriptive analysis of retained 0969 A/A packet metadata. It launches no workload, binary, build, test, or reducer and makes no causal or host claim.",
        "",
        f"Input custody: **{'PASS' if revalidation['pass'] else 'FAIL'}**. The selected set contains {pooled['children']} children and {pooled['measured_sample_count']} elapsed samples across four cells; the raw packet set is bound by `{revalidation['raw_packet_set']['set_sha256']}`.",
        "",
        "Per-cell ranges and correlations (r is Pearson; rho is Spearman):",
        "",
        "| cell | children/samples | p50 range (ns) | RSS range (KiB) | whole-child wall range (ns) | output capacity | p50↔RSS r/rho | p50↔wall r/rho |",
        "|---|---:|---:|---:|---:|---:|---:|---:|",
    ]
    for row in cells:
        ranges = row["ranges"]
        rss = ranges["rss_kb"]
        wall = ranges["whole_child_wall_ns"]
        p50 = ranges["elapsed_p50_ns"]
        cap = row["output_capacity_unique_values"]
        cr = row["correlations"]
        lines.append(
            f"| `{row['cell']}` | {row['children']}/{row['measured_sample_count']} | {p50['min']:.0f}–{p50['max']:.0f} | {rss['min']:.0f}–{rss['max']:.0f} | {wall['min']:.0f}–{wall['max']:.0f} | {','.join(f'{x:.0f}' for x in cap)} | {fmt(cr['p50_vs_rss']['pearson_r'])}/{fmt(cr['p50_vs_rss']['spearman_rho'])} | {fmt(cr['p50_vs_whole_child_wall']['pearson_r'])}/{fmt(cr['p50_vs_whole_child_wall']['spearman_rho'])} |"
        )
    lines += [
        "",
        f"The pooled p50 correlations are RSS r/rho {fmt(pooled['correlations']['p50_vs_rss']['pearson_r'])}/{fmt(pooled['correlations']['p50_vs_rss']['spearman_rho'])}, whole-child wall {fmt(pooled['correlations']['p50_vs_whole_child_wall']['pearson_r'])}/{fmt(pooled['correlations']['p50_vs_whole_child_wall']['spearman_rho'])}, and output capacity {fmt(pooled['correlations']['p50_vs_output_capacity']['pearson_r'])}/{fmt(pooled['correlations']['p50_vs_output_capacity']['spearman_rho'])}. These pooled values include cell composition and are not process-level causal evidence.",
        "",
        "Within each cell, output capacity is constant, so its p50 correlation is undefined; the four cell capacities differ and explain the pooled association. The fixed six-child fastest/slowest p50 quartiles show whole-child wall medians for the slow tail above the fast tail in all four cells, while RSS medians are equal or lower in the slow tail. This is an association in retained metadata, without a selected value threshold or causal interpretation.",
        "",
        "Fixed-tail quartile medians (six children per tail; 480 elapsed samples per tail):",
        "",
        "| cell | fast p50 / RSS / wall | slow p50 / RSS / wall |",
        "|---|---:|---:|",
    ]
    for row in cells:
        tails = {tail["label"]: tail for tail in row["quartile_tails"]}
        fast, slow = tails["fastest_p50_quartile"], tails["slowest_p50_quartile"]
        lines.append(
            f"| `{row['cell']}` | {fast['ranges']['elapsed_p50_ns']['median']:.0f} ns / {fast['median_rss_kb']:.0f} KiB / {fast['median_whole_child_wall_ns']:.0f} ns | {slow['ranges']['elapsed_p50_ns']['median']:.0f} ns / {slow['median_rss_kb']:.0f} KiB / {slow['median_whole_child_wall_ns']:.0f} ns |"
        )
    lines += [
        "",
        f"Analyzer: `{result['analyzer']['path']}` (SHA-256 `{result['analyzer']['sha256']}`). JSON: `{OUT_JSON}` (SHA-256 `{result['receipt_sha256']}`).",
        "",
    ]
    return "\n".join(lines)


def main() -> int:
    errors: list[str] = []
    prior = load_json(PRIOR_AUDIT) if PRIOR_AUDIT.is_file() else {}
    revalidation_base = revalidate_inputs(errors, prior)
    rows, selected_binding = load_selected_rows(errors)
    result: dict[str, Any] = {
        "schema": "litchi.performance.0970.aa_resource_correlation.v1",
        "created_utc": datetime.now(timezone.utc).isoformat(),
        "analyzer": {"path": str(Path(__file__).resolve()), "sha256": sha256_file(Path(__file__).resolve())},
        "scope": {
            "offline_only": True,
            "no_workload": True,
            "no_build": True,
            "no_tests": True,
            "no_reducer": True,
            "no_causal_claim": True,
            "selected_cells": list(SELECTED_CELLS),
            "tail_method": "Within each 24-child cell, sort by child elapsed p50 and describe fixed six-child lower/upper quartile tails; no value threshold is selected.",
        },
        "revalidation": {
            **revalidation_base,
            "selected_packet_binding": selected_binding,
            "pass": not errors,
        },
        "errors": errors,
        "analysis": analyze(rows),
    }
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    receipt_sha = sha256_file(OUT_JSON)
    OUT_MD.write_text(markdown({**result, "receipt_sha256": receipt_sha}), encoding="utf-8")
    print(json.dumps({
        "json": str(OUT_JSON),
        "json_sha256": receipt_sha,
        "markdown": str(OUT_MD),
        "analyzer_sha256": result["analyzer"]["sha256"],
        "errors": len(errors),
    }, sort_keys=True))
    return 0 if not errors else 2


if __name__ == "__main__":
    sys.exit(main())
