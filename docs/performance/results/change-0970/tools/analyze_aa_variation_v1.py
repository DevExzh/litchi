#!/usr/bin/env python3
"""Offline descriptive analysis of the retained 0969 DOC A/A samples.

The script rechecks the sealed 0969 custody receipt and raw-file set before
reading elapsed samples.  It does not import or invoke the 0969 reducer and
does not launch a binary, workload, build, or test.
"""
from __future__ import annotations

import hashlib
import itertools
import json
import math
import random
import re
import statistics
import sys
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


ROOT = Path("/home/zhuhe/litchi-performance-0969")
RAW_ROOT = ROOT / "capture/aa-v4/aa"
PRIOR_AUDIT = ROOT / "analysis/aa-independent-audit-v1.json"
PRIOR_HELPER = ROOT / "tools/independent_audit_aa_v1.py"
PLAN = ROOT / "plans/plan-0969-aa-v6.json"
ROOT_GO = ROOT / "receipts/root-go-aa-v1.json"
TERMINAL = ROOT / "receipts/root-terminal-aa-v1.json"
READINESS = ROOT / "capture/aa-v4/aa-readiness-v4.json"
CAPTURE_STAGE = ROOT / "capture/aa-v4/capture-stage-v4.json"
OUT_JSON = Path("/home/zhuhe/litchi-performance-0970/analysis/aa-variation-v1.json")
OUT_MD = Path("/home/zhuhe/litchi-performance-0970/analysis/aa-variation-v1.md")
ADR_ROOT = ROOT / "source-baseline/docs/adr"

EXPECTED = {
    str(PLAN): "e853030180d7464c06cf1ed5d525957ce2d50514b8337a1b168392374d00b5df",
    str(ROOT_GO): "f3584215eade930b7748e1e3b7ea5a9b17bf331c037d2cda35f43afe8c93c7fa",
    str(TERMINAL): "187494d87f4c0e3d7786dffcfb266094cea0faf7a625ac66d7d4e0acbaf71d9e",
    str(READINESS): "6bb110145d0583be3ff12c3c6e1ada6611be6b5396f55f24698bb0358bb51efe",
    str(CAPTURE_STAGE): "59485d2c231f2c56e39300acfc5916c5a14ceea91918ebe0c2706fc7fe666bc3",
    str(PRIOR_HELPER): "09a8749910a1da49579a8ea782c3a6a11b34a303cd56818ac30d1b500069441a",
    str(PRIOR_AUDIT): "f0d0d0b63a896be41dab84550ff39f59954c26538423be2a3d3d017df9f26775",
}
EXPECTED_RAW_DIGEST = "3eecb0db82fb14be38852685956dd6fa485fad2c609b23261fa5c75443befdd8"
EXPECTED_RAW_FILES = 2304
EXPECTED_PACKETS = 576
EXPECTED_SAMPLES = 80
EXPECTED_ADR_COUNT = 32
EXPECTED_ADR_BYTES = 3147627
EXPECTED_ADR_SET_DIGEST = "bf84c9e75149dda400f9f93c4c72f07b26afa5969ff133c9f4b4427421a9bc40"
RETENTION_SHORT = {"zero": "zero", "default_8mib": "default", "release_8mib": "release"}
PACKET_RE = re.compile(r"^aa-(docnohf|docfloat)-e(1|2|8|32)-(zero|default|release)-p(\d{2})-(first|second)$")


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def raw_set_digest(paths: list[Path], root: Path) -> str:
    rows = sorted(f"{sha256_file(path)}  {path.relative_to(root).as_posix()}\n" for path in paths)
    return hashlib.sha256("".join(rows).encode()).hexdigest()


def percentile(values: list[float], q: float) -> float:
    ordered = sorted(float(value) for value in values)
    if not ordered:
        raise ValueError("percentile of empty values")
    pos = (len(ordered) - 1) * q
    low, high = math.floor(pos), math.ceil(pos)
    if low == high:
        return ordered[low]
    return ordered[low] + (ordered[high] - ordered[low]) * (pos - low)


def avg_rank(values: list[float]) -> list[float]:
    ordered = sorted((value, index) for index, value in enumerate(values))
    result = [0.0] * len(values)
    start = 0
    while start < len(ordered):
        end = start + 1
        while end < len(ordered) and ordered[end][0] == ordered[start][0]:
            end += 1
        rank = (start + end - 1) / 2.0 + 1.0
        for _, index in ordered[start:end]:
            result[index] = rank
        start = end
    return result


def spearman(xs: list[float], ys: list[float]) -> float:
    if len(xs) != len(ys) or len(xs) < 2:
        return 0.0
    rx, ry = avg_rank(xs), avg_rank(ys)
    mx, my = statistics.mean(rx), statistics.mean(ry)
    numerator = sum((a - mx) * (b - my) for a, b in zip(rx, ry))
    denominator = math.sqrt(sum((a - mx) ** 2 for a in rx) * sum((b - my) ** 2 for b in ry))
    return numerator / denominator if denominator else 0.0


def slope(xs: list[float], ys: list[float]) -> float:
    if len(xs) != len(ys) or len(xs) < 2:
        return 0.0
    mx, my = statistics.mean(xs), statistics.mean(ys)
    denominator = sum((x - mx) ** 2 for x in xs)
    return sum((x - mx) * (y - my) for x, y in zip(xs, ys)) / denominator if denominator else 0.0


def summary(values: list[float]) -> dict[str, Any]:
    if not values:
        return {"n": 0}
    vals = [float(value) for value in values]
    return {
        "n": len(vals),
        "min": min(vals),
        "p05": percentile(vals, 0.05),
        "p10": percentile(vals, 0.10),
        "p25": percentile(vals, 0.25),
        "median": percentile(vals, 0.50),
        "p75": percentile(vals, 0.75),
        "p90": percentile(vals, 0.90),
        "p95": percentile(vals, 0.95),
        "p99": percentile(vals, 0.99),
        "max": max(vals),
        "mean": statistics.mean(vals),
        "stdev_population": statistics.pstdev(vals),
    }


def relative_cv(values: list[float]) -> float:
    mean = statistics.mean(values)
    return statistics.pstdev(values) / mean if mean else math.inf


def parse_timestamp(value: str) -> float:
    return datetime.fromisoformat(value).timestamp()


def compressed_runs(labels: list[str]) -> list[str]:
    return [f"{label}{len(list(group))}" for label, group in itertools.groupby(labels)]


def sequence_modes(values: list[int]) -> dict[str, Any]:
    """Describe a conservative two-level sequence split, not a test of modality."""
    ordered = sorted(float(value) for value in values)
    best: tuple[float, dict[str, Any]] | None = None
    # Exclude singleton tails: a one-sample outlier is not a second mode.
    for split_count in range(10, 71):
        low, high = ordered[:split_count], ordered[split_count:]
        low_median, high_median = statistics.median(low), statistics.median(high)
        low_mad = statistics.median([abs(value - low_median) for value in low])
        high_mad = statistics.median([abs(value - high_median) for value in high])
        pooled_mad = (low_mad + high_mad) / 2.0
        if pooled_mad <= 0:
            continue
        gap = ordered[split_count] - ordered[split_count - 1]
        score = gap / pooled_mad
        center = (low_median + high_median) / 2.0
        labels = ["H" if value > center else "L" for value in values]
        transitions = sum(a != b for a, b in zip(labels, labels[1:]))
        runs = compressed_runs(labels)
        high_count = labels.count("H")
        candidate = {
            "split_count_low": split_count,
            "split_count_high": len(values) - split_count,
            "low_median_ns": low_median,
            "high_median_ns": high_median,
            "relative_center_gap": abs(high_median - low_median) / center if center else math.inf,
            "sorted_gap_ns": gap,
            "gap_over_pooled_mad": score,
            "transitions": transitions,
            "runs": len(runs),
            "longest_run": max((int(re.sub(r"^[LH]", "", run)) for run in runs), default=0),
            "high_count": high_count,
            "low_count": len(values) - high_count,
            "run_encoding": runs,
        }
        if best is None or score > best[0]:
            best = (score, candidate)
    if best is None:
        return {"candidate_two_level_sequence": False, "reason": "no non-degenerate split"}
    candidate = best[1]
    candidate["candidate_two_level_sequence"] = bool(
        candidate["gap_over_pooled_mad"] >= 10.0
        and candidate["relative_center_gap"] >= 0.05
        and min(candidate["low_count"], candidate["high_count"]) >= 10
        and candidate["runs"] >= 3
    )
    candidate["heuristic_thresholds"] = {
        "gap_over_pooled_mad_at_least": 10.0,
        "relative_center_gap_at_least": 0.05,
        "each_mode_count_at_least": 10,
        "chronological_runs_at_least": 3,
    }
    return candidate


def descriptor_hashes() -> dict[str, Any]:
    result = {}
    for raw_path, expected in EXPECTED.items():
        path = Path(raw_path)
        actual = sha256_file(path) if path.is_file() else None
        result[raw_path] = {"sha256": actual, "expected_sha256": expected, "match": actual == expected}
    return result


def check_binding(errors: list[str], path: Path, expected: str) -> None:
    if not path.is_file():
        errors.append(f"missing binding: {path}")
        return
    actual = sha256_file(path)
    if actual != expected:
        errors.append(f"SHA mismatch {path}: {actual} != {expected}")


def revalidate_bound_custody(errors: list[str], plan: dict[str, Any]) -> dict[str, Any]:
    """Recheck executable, fixtures, source manifest, and all baseline files."""
    result: dict[str, Any] = {"binaries": {}, "fixtures": {}, "source": {}}
    for role in ("native_baseline", "allocator_baseline"):
        row = (plan.get("binaries") or {}).get(role) or {}
        path = Path(row.get("path", ""))
        actual = sha256_file(path) if path.is_file() else None
        result["binaries"][role] = {"path": str(path), "bytes": path.stat().st_size if path.is_file() else None, "sha256": actual, "expected_sha256": row.get("sha256"), "match": actual == row.get("sha256")}
        if not result["binaries"][role]["match"]:
            errors.append(f"binary custody changed: {role}")
    for case, row in sorted((plan.get("fixtures") or {}).items()):
        path = Path(row.get("path", ""))
        actual = sha256_file(path) if path.is_file() else None
        result["fixtures"][case] = {"path": str(path), "bytes": path.stat().st_size if path.is_file() else None, "sha256": actual, "expected_bytes": row.get("bytes"), "expected_sha256": row.get("sha256"), "match": path.is_file() and path.stat().st_size == row.get("bytes") and actual == row.get("sha256")}
        if not result["fixtures"][case]["match"]:
            errors.append(f"fixture custody changed: {case}")
    source = (plan.get("source") or {}).get("baseline") or {}
    manifest = Path(source.get("manifest", ""))
    manifest_hash = sha256_file(manifest) if manifest.is_file() else None
    result["source"]["manifest"] = {"path": str(manifest), "sha256": manifest_hash, "expected_sha256": source.get("manifest_sha256"), "match": manifest_hash == source.get("manifest_sha256")}
    if not result["source"]["manifest"]["match"]:
        errors.append("baseline source manifest custody changed")
    file_manifest = manifest.parent / "source-file-manifest-v1.json"
    entries = load_json(file_manifest).get("entries", []) if file_manifest.is_file() else []
    source_root = Path(source.get("root", ""))
    checked = 0
    mismatches: list[str] = []
    total_bytes = 0
    for entry in entries:
        rel = entry.get("path")
        path = source_root / rel if isinstance(rel, str) else None
        if path is None or not path.is_file():
            mismatches.append(str(rel)); continue
        checked += 1
        total_bytes += path.stat().st_size
        if path.stat().st_size != entry.get("bytes") or sha256_file(path) != entry.get("baseline_sha256"):
            mismatches.append(str(rel))
    result["source"].update({"root": str(source_root), "manifest_entries": len(entries), "checked_entries": checked, "bytes": total_bytes, "expected_bytes": 18259541, "mismatches": mismatches, "match": len(entries) == 924 and checked == 924 and not mismatches and total_bytes == 18259541})
    if not result["source"]["match"]:
        errors.append("baseline source file custody changed")
    return result


def revalidate_accepted_adrs(errors: list[str]) -> dict[str, Any]:
    """Read and bind every numbered accepted ADR before interpreting samples."""
    paths = sorted(ADR_ROOT.glob("[0-9][0-9][0-9][0-9]-*.md"))
    total_bytes = sum(path.stat().st_size for path in paths if path.is_file())
    digest = raw_set_digest(paths, ADR_ROOT) if all(path.is_file() for path in paths) else None
    result = {
        "root": str(ADR_ROOT),
        "files": [path.name for path in paths],
        "count": len(paths),
        "expected_count": EXPECTED_ADR_COUNT,
        "bytes": total_bytes,
        "expected_bytes": EXPECTED_ADR_BYTES,
        "set_sha256": digest,
        "expected_set_sha256": EXPECTED_ADR_SET_DIGEST,
        "match": len(paths) == EXPECTED_ADR_COUNT and total_bytes == EXPECTED_ADR_BYTES and digest == EXPECTED_ADR_SET_DIGEST,
    }
    if not result["match"]:
        errors.append("accepted ADR binding changed")
    return result


def load_packets(errors: list[str], prior: dict[str, Any]) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    paths = sorted(RAW_ROOT.glob("*.json"))
    raw_files = sorted(path for path in RAW_ROOT.iterdir() if path.is_file())
    raw_digest = raw_set_digest(raw_files, RAW_ROOT)
    packet_custody = {
        "raw_root": str(RAW_ROOT),
        "raw_file_count": len(raw_files),
        "raw_file_count_expected": EXPECTED_RAW_FILES,
        "raw_file_set_sha256": raw_digest,
        "raw_file_set_sha256_expected": EXPECTED_RAW_DIGEST,
        "raw_total_bytes": sum(path.stat().st_size for path in raw_files),
        "json_count": len(paths),
        "packet_ids_unique": True,
        "stdout_report_equal": 0,
        "stderr_empty": 0,
        "sidecar_bytes_match": 0,
    }
    if len(raw_files) != EXPECTED_RAW_FILES or raw_digest != EXPECTED_RAW_DIGEST or len(paths) != EXPECTED_PACKETS:
        errors.append("raw packet file-set binding changed")
    prior_packets = (prior.get("observations") or {}).get("packets") or {}
    for key, value in (("raw_file_count", len(raw_files)), ("raw_file_set_digest", raw_digest), ("json_count", len(paths))):
        prior_key = "raw_file_set_digest" if key == "raw_file_set_digest" else key
        if prior_packets.get(prior_key) != value:
            errors.append(f"prior audit custody mismatch: {key}")

    packets: list[dict[str, Any]] = []
    seen: set[str] = set()
    for path in paths:
        try:
            packet = load_json(path)
        except Exception as exc:
            errors.append(f"invalid packet JSON {path.name}: {exc}")
            continue
        pid = packet.get("packet_id")
        match = PACKET_RE.fullmatch(str(pid)) if pid is not None else None
        if match is None or path.stem != pid or pid in seen:
            errors.append(f"packet identity mismatch: {path.name}")
            packet_custody["packet_ids_unique"] = False
        seen.add(str(pid))
        context = packet.get("context") or {}
        report = packet.get("report") or {}
        if packet.get("role") != "aa_baseline" or packet.get("schema") != "litchi.performance.0969.runner.v4":
            errors.append(f"packet role/schema mismatch: {pid}")
        terminal = packet.get("terminal") or {}
        if terminal.get("returncode") != 0 or terminal.get("timed_out") is not False or terminal.get("launch_error") is not None:
            errors.append(f"packet terminal mismatch: {pid}")
        elapsed = report.get("elapsed_ns")
        if not isinstance(elapsed, list) or len(elapsed) != EXPECTED_SAMPLES or any(type(value) is not int or value <= 0 for value in elapsed):
            errors.append(f"sample sequence mismatch: {pid}")
            continue
        for ext, key in (("stdout", "stdout"), ("stderr", "stderr"), ("rss", "rss")):
            descriptor = packet.get(key) or {}
            sidecar = Path(descriptor.get("path", ""))
            expected_sidecar = RAW_ROOT / f"{pid}.{ext}"
            if sidecar.resolve() != expected_sidecar.resolve() or not sidecar.is_file() or descriptor.get("bytes") != sidecar.stat().st_size:
                errors.append(f"sidecar custody mismatch: {pid}.{ext}")
            else:
                packet_custody["sidecar_bytes_match"] += 1
            if ext == "stderr" and sidecar.is_file() and sidecar.read_bytes() == b"":
                packet_custody["stderr_empty"] += 1
        stdout_path = Path((packet.get("stdout") or {}).get("path", ""))
        try:
            if load_json(stdout_path) == report:
                packet_custody["stdout_report_equal"] += 1
            else:
                errors.append(f"stdout/report mismatch: {pid}")
        except Exception as exc:
            errors.append(f"stdout JSON mismatch {pid}: {exc}")
        if context.get("packet_id") != pid or context.get("lane") != "aa" or context.get("variant") != "baseline":
            errors.append(f"packet context mismatch: {pid}")
        packets.append(packet)
    if len(packets) != EXPECTED_PACKETS or len(seen) != EXPECTED_PACKETS:
        errors.append("packet coverage is not 576 unique children")
    return packets, packet_custody


def make_child(packet: dict[str, Any]) -> dict[str, Any]:
    report = packet["report"]
    context = packet["context"]
    cell = context["cell"]
    values = [int(value) for value in report["elapsed_ns"]]
    mean = statistics.mean(values)
    median = percentile(values, 0.5)
    q1, q3 = percentile(values, 0.25), percentile(values, 0.75)
    first10, last10 = statistics.mean(values[:10]), statistics.mean(values[-10:])
    sequence = sequence_modes(values)
    return {
        "packet_id": packet["packet_id"],
        "cell": cell["cell_id"],
        "case": cell["case"],
        "edits": cell["edits"],
        "retention": cell["retention"],
        "order": context["aa_order"],
        "pair": context["pair"],
        "position": context["aa_position"],
        "created_utc": packet["created_utc"],
        "created_ts": parse_timestamp(packet["created_utc"]),
        "wall_ns_excluded_from_metric": packet.get("wall_ns_excluded_from_metric"),
        "rss_kb": (packet.get("rss") or {}).get("kb"),
        "median_ns": median,
        "mean_ns": mean,
        "cv": statistics.pstdev(values) / mean if mean else math.inf,
        "iqr_rel": (q3 - q1) / median if median else math.inf,
        "p05_ns": percentile(values, 0.05),
        "p95_ns": percentile(values, 0.95),
        "min_ns": min(values),
        "max_ns": max(values),
        "first10_mean_ns": first10,
        "last10_mean_ns": last10,
        "first10_to_last10_rel": (last10 - first10) / first10 if first10 else math.inf,
        "sequence_modes": sequence,
    }


def cell_dispersion(children: list[dict[str, Any]]) -> dict[str, Any]:
    medians = [child["median_ns"] for child in children]
    cvs = [child["cv"] for child in children]
    between = relative_cv(medians)
    within = statistics.median(cvs)
    return {
        "child_count": len(children),
        "child_median_ns": summary(medians),
        "within_child_sample_cv": summary(cvs),
        "within_child_cv_median": within,
        "between_child_median_cv": between,
        "between_to_within_cv_ratio": between / within if within else math.inf,
        "within_child_iqr_rel_median": statistics.median([child["iqr_rel"] for child in children]),
        "first10_to_last10_rel": summary([child["first10_to_last10_rel"] for child in children]),
        "two_level_candidate_count": sum(bool(child["sequence_modes"].get("candidate_two_level_sequence")) for child in children),
    }


def grouped_summary(children: list[dict[str, Any]], keys: tuple[str, ...]) -> list[dict[str, Any]]:
    groups: dict[tuple[Any, ...], list[dict[str, Any]]] = defaultdict(list)
    for child in children:
        groups[tuple(child[key] for key in keys)].append(child)
    result = []
    for key, group in sorted(groups.items(), key=lambda item: tuple(str(x) for x in item[0])):
        row = {name: value for name, value in zip(keys, key)}
        row.update(cell_dispersion(group))
        result.append(row)
    return result


def trends_by_cell(children: list[dict[str, Any]]) -> dict[str, Any]:
    groups: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for child in children:
        groups[child["cell"]].append(child)
    rows = []
    for cell, group in sorted(groups.items()):
        ab = [child["median_ns"] for child in group if child["order"] == "AB"]
        ba = [child["median_ns"] for child in group if child["order"] == "BA"]
        first = [child["median_ns"] for child in group if child["position"] == "first"]
        second = [child["median_ns"] for child in group if child["position"] == "second"]
        by_pair = []
        for pair in range(1, 13):
            values = [child["median_ns"] for child in group if child["pair"] == pair]
            by_pair.append(statistics.median(values))
        ordered = sorted(group, key=lambda child: child["created_ts"])
        order_pair_rows = []
        for order in ("AB", "BA"):
            pair_values = []
            pair_numbers = list(range(1, 7)) if order == "AB" else list(range(7, 13))
            for pair in pair_numbers:
                values = [child["median_ns"] for child in group if child["order"] == order and child["pair"] == pair]
                pair_values.append(statistics.median(values))
            order_pair_rows.append({
                "order": order,
                "pair_numbers": pair_numbers,
                "pair_medians_ns": pair_values,
                "spearman_pair_median": spearman([float(x) for x in range(1, 7)], pair_values),
                "slope_ns_per_pair": slope([float(x) for x in range(1, 7)], pair_values),
            })
        rows.append({
            "cell": cell,
            "child_count": len(group),
            "AB_median_ns": percentile(ab, 0.5),
            "BA_median_ns": percentile(ba, 0.5),
            "BA_over_AB": percentile(ba, 0.5) / percentile(ab, 0.5),
            "first_median_ns": percentile(first, 0.5),
            "second_median_ns": percentile(second, 0.5),
            "second_over_first": percentile(second, 0.5) / percentile(first, 0.5),
            "paired_second_over_first": summary([
                next(child["median_ns"] for child in group if child["pair"] == pair and child["position"] == "second")
                / next(child["median_ns"] for child in group if child["pair"] == pair and child["position"] == "first")
                for pair in range(1, 13)
            ]),
            "pair_medians_ns": by_pair,
            "pair_spearman_median": spearman([float(x) for x in range(1, 13)], by_pair),
            "chronological_child_spearman_median": spearman([float(x) for x in range(1, len(ordered) + 1)], [child["median_ns"] for child in ordered]),
            "chronological_child_slope_ns_per_child": slope([float(x) for x in range(1, len(ordered) + 1)], [child["median_ns"] for child in ordered]),
            "created_span_seconds": ordered[-1]["created_ts"] - ordered[0]["created_ts"],
            "order_specific_pair_trends": order_pair_rows,
        })
    return {
        "cells": rows,
        "global": {
            "BA_over_AB": summary([row["BA_over_AB"] for row in rows]),
            "second_over_first": summary([row["second_over_first"] for row in rows]),
            "paired_second_over_first": summary([row["paired_second_over_first"]["median"] for row in rows]),
            "pair_spearman_median": summary([row["pair_spearman_median"] for row in rows]),
            "chronological_child_spearman_median": summary([row["chronological_child_spearman_median"] for row in rows]),
            "chronological_child_slope_ns_per_child": summary([row["chronological_child_slope_ns_per_child"] for row in rows]),
        },
    }


def family_edit_retention(children: list[dict[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for name, keys in (("family", ("case",)), ("family_edit", ("case", "edits")), ("family_retention", ("case", "retention"))):
        result[name] = grouped_summary(children, keys)
    return result


def markdown(result: dict[str, Any]) -> str:
    pop = result["population"]
    overall = result["within_between"]["overall"]
    trends = result["trends"]["global"]
    warm = result["warmup_startup"]
    bim = result["bimodal_review"]
    family = result["family_edit_retention"]
    top = sorted(bim["heuristic_two_level_candidates"], key=lambda row: row["sequence_modes"].get("gap_over_pooled_mad", 0), reverse=True)[:5]
    family_drift = {row["case"]: row["first10_to_last10_rel"]["median"] for row in family["family"]}
    edit_drift = {}
    for row in family["family_edit"]:
        edit_drift.setdefault(row["case"], {})[row["edits"]] = row["first10_to_last10_rel"]["median"]
    retention_drift = {}
    for row in family["family_retention"]:
        retention_drift.setdefault(row["case"], {})[row["retention"]] = row["first10_to_last10_rel"]["median"]
    candidate_by_family = {row["case"]: row["two_level_candidate_count"] for row in family["family"]}
    lines = [
        "# 0969 DOC A/A variation analysis",
        "",
        "This is a descriptive offline analysis of the retained baseline-only A/A raw packets. It launches no workload, build, test, or reducer and makes no host-cause claim.",
        "",
        f"Custody revalidation: **{'PASS' if result['revalidation']['pass'] else 'FAIL'}**. The sealed 0969 audit, plan/GO/terminal/readiness/stage bindings, and raw packet set were rechecked before analysis. The packet set contains {pop['children']} children and {pop['samples']} measured samples.",
        "",
        f"Within-child sample CV has median **{overall['within_child_cv']['median']:.3%}** and p95 **{overall['within_child_cv']['p95']:.3%}**. Across the 24 cell-level groups, the between-child median CV has median **{overall['between_child_median_cv']['median']:.3%}**; the between/within CV ratio has median **{overall['between_to_within_cv_ratio']['median']:.3f}**. Between-child median dispersion exceeds within-child CV in {overall['between_gt_within_cell_count']}/24 cells and exceeds it 2× in {overall['between_gt_2x_within_cell_count']}/24; this is a descriptive scale comparison, and cell and order composition remain part of the measured grouping.",
        "",
        f"Across cells, the BA/AB median ratio is {trends['BA_over_AB']['median']:.4f} (p10–p90 {trends['BA_over_AB']['p10']:.4f}–{trends['BA_over_AB']['p90']:.4f}); the second/first position ratio is {trends['second_over_first']['median']:.4f} (p10–p90 {trends['second_over_first']['p10']:.4f}–{trends['second_over_first']['p90']:.4f}). Median chronological Spearman rho is {trends['chronological_child_spearman_median']['median']:.3f}; median pair-index rho is {trends['pair_spearman_median']['median']:.3f}. These are mixed cell-level trends, not evidence of a common direction or cause.",
        "",
        f"Family medians drift {family_drift['docnohf']:.3%} for docnohf and {family_drift['docfloat']:.3%} for docfloat. The edit-group medians range {min(edit_drift['docnohf'].values()):.3%}–{max(edit_drift['docnohf'].values()):.3%} for docnohf and {min(edit_drift['docfloat'].values()):.3%}–{max(edit_drift['docfloat'].values()):.3%} for docfloat; retention-group medians range {min(retention_drift['docnohf'].values()):.3%}–{max(retention_drift['docnohf'].values()):.3%} and {min(retention_drift['docfloat'].values()):.3%}–{max(retention_drift['docfloat'].values()):.3%}, respectively. The heuristic candidates are {candidate_by_family.get('docfloat', 0)} docfloat and {candidate_by_family.get('docnohf', 0)} docnohf.",
        "",
        f"Each child declares five warmups, but warmup samples are not retained in the report. The retained measured sequence compares its first and last ten samples: overall median drift is **{warm['first10_to_last10_rel']['median']:.3%}** (p05–p95 {warm['first10_to_last10_rel']['p05']:.3%}–{warm['first10_to_last10_rel']['p95']:.3%}). This describes measured-sequence drift after warmup; it cannot identify warmup behavior itself.",
        "",
        f"The conservative two-level sequence screen flags {bim['candidate_count']} of {pop['children']} children as heuristic candidates. It requires a 10× pooled-MAD sorted gap, at least a 5% center separation, at least ten observations per side, and at least three chronological runs. The screen is exploratory and does not establish a statistical bimodality test or a host explanation.",
        "",
        "Largest heuristic candidates:",
        "",
    ]
    for row in top:
        mode = row["sequence_modes"]
        lines.append(f"- `{row['packet_id']}`: low/high medians {mode.get('low_median_ns', 0):.0f}/{mode.get('high_median_ns', 0):.0f} ns, gap/MAD {mode.get('gap_over_pooled_mad', 0):.1f}, runs {mode.get('runs')}, transitions {mode.get('transitions')}.")
    lines += [
        "",
        "The 0969 readiness result remains the governing A/A outcome: complete evidence exceeded its frozen symmetric noise bound. This variation report does not convert that failure into a workload or host diagnosis.",
        "",
        f"Analyzer: `{result['analyzer']['path']}` (SHA-256 `{result['analyzer']['sha256']}`). JSON receipt: `{OUT_JSON}` (SHA-256 `{result['receipt_sha256']}`).",
        "",
    ]
    return "\n".join(lines)


def main() -> int:
    errors: list[str] = []
    prior = load_json(PRIOR_AUDIT) if PRIOR_AUDIT.is_file() else {}
    descriptor_results = descriptor_hashes()
    for path, row in descriptor_results.items():
        if not row["match"]:
            errors.append(f"sealed binding changed: {path}")
    if prior.get("error_count") != 0 or prior.get("complete_valid_packet_set") is not True:
        errors.append("prior 0969 independent audit is not complete and valid")
    plan = load_json(PLAN) if PLAN.is_file() else {}
    accepted_adrs = revalidate_accepted_adrs(errors)
    bound_custody = revalidate_bound_custody(errors, plan)
    packets, packet_custody = load_packets(errors, prior)
    children = []
    for packet in packets:
        try:
            children.append(make_child(packet))
        except Exception as exc:
            errors.append(f"child analysis failed for {packet.get('packet_id')}: {exc}")
    if len(children) != EXPECTED_PACKETS:
        errors.append("variation population is not 576 children")

    cells = grouped_summary(children, ("cell",))
    overall_cell_rows = cells
    within_cvs = [row["within_child_cv_median"] for row in overall_cell_rows]
    between_cvs = [row["between_child_median_cv"] for row in overall_cell_rows]
    ratios = [row["between_to_within_cv_ratio"] for row in overall_cell_rows]
    drift_values = [child["first10_to_last10_rel"] for child in children]
    candidate_rows = [child for child in children if child["sequence_modes"].get("candidate_two_level_sequence")]
    for row in candidate_rows:
        # Keep only compact mode evidence in the receipt; the full sequence is
        # reproducible from the bound raw packet.
        row["sequence_modes"] = dict(row["sequence_modes"])
    family = defaultdict(list)
    for child in children:
        family[child["case"]].append(child)
    family_drift = []
    for case, group in sorted(family.items()):
        family_drift.append({"case": case, "children": len(group), "first10_to_last10_rel": summary([x["first10_to_last10_rel"] for x in group]), "within_child_cv": summary([x["cv"] for x in group]), "median_ns": summary([x["median_ns"] for x in group]), "two_level_candidate_count": sum(bool(x["sequence_modes"].get("candidate_two_level_sequence")) for x in group)})

    revalidation = {
        "pass": not errors[:],
        "prior_audit": {"path": str(PRIOR_AUDIT), "sha256": sha256_file(PRIOR_AUDIT) if PRIOR_AUDIT.is_file() else None, "recorded_sha256": EXPECTED[str(PRIOR_AUDIT)], "error_count": prior.get("error_count"), "complete_valid_packet_set": prior.get("complete_valid_packet_set")},
        "sealed_bindings": descriptor_results,
        "bound_executable_fixture_source": bound_custody,
        "raw_packet_custody": packet_custody,
    }
    result: dict[str, Any] = {
        "schema": "litchi.performance.0970.aa_variation.v1",
        "created_utc": datetime.now(timezone.utc).isoformat(),
        "analyzer": {"path": str(Path(__file__).resolve()), "sha256": sha256_file(Path(__file__).resolve())},
        "scope": {"offline_only": True, "no_build": True, "no_workload": True, "no_tests": True, "no_reducer": True, "no_host_cause_claim": True},
        "revalidation": revalidation,
        "accepted_adr_binding": accepted_adrs,
        "errors": errors,
        "population": {"children": len(children), "samples": len(children) * EXPECTED_SAMPLES, "cells": len(overall_cell_rows), "pairs_per_cell": 12, "positions_per_pair": 2, "warmups_declared_per_child": 5, "measured_samples_per_child": EXPECTED_SAMPLES, "time_range_utc": [min((x["created_utc"] for x in children), default=None), max((x["created_utc"] for x in children), default=None)]},
        "within_between": {"overall": {"within_child_cv": summary([child["cv"] for child in children]), "within_child_iqr_rel": summary([child["iqr_rel"] for child in children]), "between_child_median_cv": summary(between_cvs), "between_to_within_cv_ratio": summary(ratios), "between_gt_within_cell_count": sum(ratio > 1.0 for ratio in ratios), "between_gt_2x_within_cell_count": sum(ratio > 2.0 for ratio in ratios), "cell_count": len(overall_cell_rows)}, "cells": overall_cell_rows},
        "trends": trends_by_cell(children),
        "family_edit_retention": family_edit_retention(children),
        "warmup_startup": {"warmups_declared": 5, "warmup_samples_retained": False, "first10_to_last10_rel": summary(drift_values), "by_family": family_drift, "child_wall_ns_including_warmup": summary([float(child["wall_ns_excluded_from_metric"]) for child in children if isinstance(child["wall_ns_excluded_from_metric"], int)])},
        "bimodal_review": {"method": "restricted sorted split with chronological labels; exploratory screen only", "candidate_count": len(candidate_rows), "candidate_ids": [row["packet_id"] for row in candidate_rows], "heuristic_two_level_candidates": candidate_rows},
        "governing_readiness": {"status": (load_json(READINESS).get("status") if READINESS.is_file() else None), "interpretation": "uninterpretable if readiness_fail; this report does not override it"},
    }
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    final_sha = sha256_file(OUT_JSON)
    OUT_MD.write_text(markdown({**result, "receipt_sha256": final_sha}), encoding="utf-8")
    print(json.dumps({"json": str(OUT_JSON), "json_sha256": final_sha, "markdown": str(OUT_MD), "analyzer_sha256": result["analyzer"]["sha256"], "errors": len(errors), "candidate_count": len(candidate_rows)}, sort_keys=True))
    return 0 if not errors else 2


if __name__ == "__main__":
    sys.exit(main())
