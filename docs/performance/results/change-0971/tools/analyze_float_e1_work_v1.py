#!/usr/bin/env python3
"""Bounded offline analysis of the retained 0970 DOC Float/e1 children.

This helper reads only the sealed 0970 capture, its packet sidecars, and the
sealed v6 reduction.  It selects every docfloat/e1/default child (12 pinned
and 12 unpinned), verifies custody, and compares whole-process counters with
the baseline harness timed-region samples.  It never launches a workload,
builds code, or invokes a reducer.

The report samples are for the timed region named by report.region_definition.
Perf and GNU time counters cover the wrapper/process tree.  The helper does
not apportion process-wide instructions to report stages.
"""
from __future__ import annotations

import hashlib
import json
import math
import statistics
import sys
from collections import Counter
from pathlib import Path
from typing import Any

ROOT = Path("/home/zhuhe/litchi-performance-0970")
CAPTURE = ROOT / "capture/diagnostic-v3"
PACKETS = CAPTURE / "packets"
MANIFEST = CAPTURE / "capture-manifest-v2.json"
STAGE = CAPTURE / "capture-stage-v2.json"
PLAN = ROOT / "plans/plan-0970-v6.json"
REDUCED = ROOT / "analysis/affinity-diagnostic-v6-final.json"
REDUCER = ROOT / "tools/reduce_0970_v6.py"
ADR_ROOT = Path("/home/zhuhe/litchi-performance-0969/source-baseline/docs/adr")
OUT_JSON = Path("/home/zhuhe/litchi-performance-0971/analysis/float-e1-work-v1.json")
OUT_MD = Path("/home/zhuhe/litchi-performance-0971/analysis/float-e1-work-v1.md")

CELL_ID = "docfloat/e1/default"
EXPECTED = {
    str(MANIFEST): "b0943c4572304dccc3a70dde17cce3894a35535e4c52fb7efb2526eb760b8685",
    str(STAGE): "244e033250f303328b7c4f3b1b729c4d3548886e9153b7ecfe5a996669871143",
    str(PLAN): "4043a33e413476961feb5c4712689be9bd392f341c30d7d5a8bd73406d033e7b",
    str(REDUCED): "560e4cae2cd338fd936c57d8f3aa9ecd663aa84736ce8ea5aa431a161066c4da",
    str(REDUCER): "1cc1245b9132c80ff57172e146e5628b148f2d108ed108c4c80b9de6777e2423",
}
EXPECTED_ADR_SET = "bf84c9e75149dda400f9f93c4c72f07b26afa5969ff133c9f4b4427421a9bc40"
EXPECTED_ADR_COUNT = 32
EXPECTED_ADR_BYTES = 3147627
EXPECTED_PACKET_COUNT = 96
EXPECTED_CHILDREN = 24
EXPECTED_PER_ARM = 12
EXPECTED_SAMPLES = 80
EXPECTED_WARMUPS = 5
EXPECTED_BINARY = "0db34357a1e725dc397a4fd895532433db1214c5862b7bc96bf23f21c0396ecb"
EXPECTED_FIXTURE = "41cc0cd8f2d7390266f844f5bddbcd29d08ac5338abc33832d9c9a6ed5d0f85d"
EXPECTED_SOURCE = "f1818af5c306fb6ea01b59ca8df29710e773386b246071202bd4aa896f4e8e2b"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def raw_set_digest(paths: list[Path], root: Path) -> str:
    rows = sorted(
        f"{sha256_file(path)}  {path.relative_to(root).as_posix()}\n" for path in paths
    )
    return hashlib.sha256("".join(rows).encode("utf-8")).hexdigest()


def file_inventory(paths: list[Path], root: Path) -> dict[str, Any]:
    unique = sorted(set(paths))
    missing = [str(path) for path in unique if not path.is_file()]
    existing = [path for path in unique if path.is_file()]
    return {
        "file_count": len(existing),
        "expected_file_count": len(unique),
        "missing": missing,
        "bytes": sum(path.stat().st_size for path in existing),
        "set_sha256": raw_set_digest(existing, root) if not missing else None,
    }


def linear_percentile(values: list[float], q: float) -> float:
    ordered = sorted(float(value) for value in values)
    if not ordered:
        raise ValueError("percentile of empty sequence")
    position = (len(ordered) - 1) * q
    low, high = math.floor(position), math.ceil(position)
    if low == high:
        return ordered[low]
    return ordered[low] + (ordered[high] - ordered[low]) * (position - low)


def average_ranks(values: list[float]) -> list[float]:
    ordered = sorted((value, index) for index, value in enumerate(values))
    ranks = [0.0] * len(values)
    start = 0
    while start < len(ordered):
        end = start + 1
        while end < len(ordered) and ordered[end][0] == ordered[start][0]:
            end += 1
        rank = (start + end + 1) / 2.0
        for _, index in ordered[start:end]:
            ranks[index] = rank
        start = end
    return ranks


def correlation(xs: list[float], ys: list[float]) -> tuple[float | None, str | None]:
    if len(xs) != len(ys):
        raise ValueError("correlation vectors have different lengths")
    if len(xs) < 2:
        return None, "fewer_than_two_children"
    mean_x, mean_y = statistics.mean(xs), statistics.mean(ys)
    dx = [x - mean_x for x in xs]
    dy = [y - mean_y for y in ys]
    denominator = math.sqrt(sum(x * x for x in dx) * sum(y * y for y in dy))
    if denominator == 0:
        return None, "constant_metric"
    return sum(x * y for x, y in zip(dx, dy)) / denominator, None


def correlation_pair(xs: list[float], ys: list[float]) -> dict[str, Any]:
    pearson, pearson_reason = correlation(xs, ys)
    spearman, spearman_reason = correlation(average_ranks(xs), average_ranks(ys))
    return {
        "n": len(xs),
        "pearson": pearson,
        "pearson_reason": pearson_reason,
        "spearman": spearman,
        "spearman_reason": spearman_reason,
    }


def numeric_summary(values: list[float]) -> dict[str, Any]:
    if not values:
        return {"n": 0}
    vals = [float(value) for value in values]
    mean = statistics.mean(vals)
    stdev = statistics.pstdev(vals)
    median = statistics.median(vals)
    return {
        "n": len(vals),
        "min": min(vals),
        "median": median,
        "max": max(vals),
        "mean": mean,
        "population_stdev": stdev,
        "cv_population": stdev / mean if mean else None,
        "relative_range_over_median": (max(vals) - min(vals)) / median if median else None,
    }


def parse_perf(path: Path) -> dict[str, dict[str, float | None]]:
    events: dict[str, dict[str, float | None]] = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        if not line or line.startswith("#"):
            continue
        fields = line.split(",")
        if len(fields) < 5:
            raise ValueError(f"short perf row in {path}: {line!r}")
        try:
            value = float(fields[0].strip())
        except ValueError as exc:
            raise ValueError(f"non-numeric perf value in {path}: {line!r}") from exc
        event = fields[2].strip()
        runtime = fields[3].strip()
        percent = fields[4].strip()
        events[event] = {
            "value": value,
            "runtime_ns": float(runtime) if runtime else None,
            "percent": float(percent) if percent else None,
            "unit": fields[1].strip(),
        }
    return events


def bools_all_true(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, dict):
        return all(bools_all_true(item) for item in value.values())
    if isinstance(value, list):
        return all(bools_all_true(item) for item in value)
    return True


def active_stage(stage: dict[str, Any]) -> bool:
    return any(stage.get(key) is not None for key in ("before", "after", "after_release"))


def stage_signature(stages: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [
        {
            "stage": stage.get("stage"),
            "requested_text": stage.get("requested_text"),
            "before": stage.get("before"),
            "after": stage.get("after"),
            "after_release": stage.get("after_release"),
        }
        for stage in stages
        if active_stage(stage)
    ]


def route_signature(routes: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [
        {
            "retention": route.get("retention"),
            "retention_ceiling_bytes": route.get("retention_ceiling_bytes"),
            "output_bytes": route.get("output_bytes"),
            "output_capacity": route.get("output_capacity"),
            "output_sha256": route.get("output_sha256"),
            "semantic_change_count": route.get("semantic_change_count"),
            "final_retained_before_commit": route.get("final_retained_before_commit"),
            "active_stages": stage_signature(route.get("stages", [])),
        }
        for route in routes
    ]


def extract_child(packet: dict[str, Any]) -> dict[str, Any]:
    context = packet["context"]
    report = packet["report"]
    packet_id = packet["packet_id"]
    if context["cell_id"] != CELL_ID:
        raise ValueError(f"unexpected cell for selected packet {packet_id}")
    samples = report["samples"]
    elapsed = [float(sample["elapsed_ns"]) for sample in samples]
    if len(elapsed) != EXPECTED_SAMPLES:
        raise ValueError(f"{packet_id}: expected {EXPECTED_SAMPLES} samples")
    if report.get("elapsed_ns") != [sample["elapsed_ns"] for sample in samples]:
        raise ValueError(f"{packet_id}: top-level elapsed_ns differs from sample elapsed_ns")
    stage_key_sets = {
        tuple(sorted(stage))
        for sample in samples
        for stage in sample["stages"]
    }
    if len(stage_key_sets) != 1:
        raise ValueError(f"{packet_id}: stage field keys vary within report")
    perf = parse_perf(Path(packet["sidecars"]["perf"]["path"]))
    time_values = load_json(Path(packet["sidecars"]["time"]["path"]))
    required_events = (
        "task-clock",
        "cycles",
        "instructions",
        "context-switches",
        "cpu-migrations",
        "page-faults",
    )
    if set(perf) != set(required_events):
        raise ValueError(f"{packet_id}: unexpected perf event set {sorted(perf)}")
    for name in required_events:
        if perf[name]["value"] is None or perf[name]["runtime_ns"] is None:
            raise ValueError(f"{packet_id}: incomplete perf event {name}")
        if perf[name]["percent"] != 100.0:
            raise ValueError(f"{packet_id}: multiplexed event {name}")
    terminal = packet["terminal"]
    if terminal.get("returncode") != 0 or terminal.get("timed_out") or terminal.get("launch_error"):
        raise ValueError(f"{packet_id}: terminal failure {terminal}")
    if packet.get("validation_error") is not None:
        raise ValueError(f"{packet_id}: validation error")
    launch = load_json(Path(packet["sidecars"]["launch"]["path"]))
    stage_rows = [sample["stages"] for sample in samples]
    all_sample_active = [stage for stages in stage_rows for stage in stages if active_stage(stage)]
    after_values = [float(stage["after"]) for stage in all_sample_active if stage.get("after") is not None]
    route_rows = report["route_outputs"]
    route_active = [stage for route in route_rows for stage in route.get("stages", []) if active_stage(stage)]
    report_signature = {
        "samples_requested": report["samples_requested"],
        "warmups": report["warmups"],
        "output_bytes": report["output_bytes"],
        "output_capacity": report["output_capacity"],
        "source_bytes": report["source_bytes"],
        "reference_output_bytes": report["reference_output_bytes"],
        "output_sha256": report["output_sha256"],
        "source_sha256": report["source_sha256"],
        "schema": report["schema"],
        "schema_version": report["schema_version"],
        "timing_claim": report["timing_claim"],
        "allocator_instrumented": report["allocator_instrumented"],
        "allocator_timing_claim": report["allocator_timing_claim"],
        "allocation": report["allocation"],
        "stage_keys": list(next(iter(stage_key_sets))),
        "sample_stage_slots": len(samples[0]["stages"]),
        "sample_active_stage_signature": stage_signature(samples[0]["stages"]),
        "route_signature": route_signature(route_rows),
        "correctness": report["correctness"],
        "alternating_texts": report["alternating_texts"],
        "change": report["change"],
        "region_definition": report["region_definition"],
    }
    perf_values = {
        "task_clock_msec": perf["task-clock"]["value"],
        "cycles": perf["cycles"]["value"],
        "instructions": perf["instructions"]["value"],
        "perf_context_switches": perf["context-switches"]["value"],
        "cpu_migrations": perf["cpu-migrations"]["value"],
        "perf_page_faults": perf["page-faults"]["value"],
        "task_clock_runtime_ns": perf["task-clock"]["runtime_ns"],
    }
    perf_values["cycles_per_instruction"] = (
        perf_values["cycles"] / perf_values["instructions"]
        if perf_values["instructions"]
        else None
    )
    perf_values["cycles_per_task_clock_msec"] = (
        perf_values["cycles"] / perf_values["task_clock_msec"]
        if perf_values["task_clock_msec"]
        else None
    )
    process_values = {
        **perf_values,
        "user_s": float(time_values["user_s"]),
        "system_s": float(time_values["system_s"]),
        "voluntary_context_switches": float(time_values["voluntary_context_switches"]),
        "involuntary_context_switches": float(time_values["involuntary_context_switches"]),
        "minor_faults": float(time_values["minor_faults"]),
        "major_faults": float(time_values["major_faults"]),
        "max_rss_kb": float(time_values["max_rss_kb"]),
    }
    return {
        "packet_id": packet_id,
        "pair": context["pair"],
        "condition": context["condition"],
        "order_position": context["order_position"],
        "requested_affinity": context["expected_affinity"],
        "observed_affinity": context["observed_affinity"],
        "p50_ns": linear_percentile(elapsed, 0.50),
        "p95_ns": linear_percentile(elapsed, 0.95),
        "mean_ns": statistics.mean(elapsed),
        "min_ns": min(elapsed),
        "max_ns": max(elapsed),
        "timed_sample_cv_population": statistics.pstdev(elapsed) / statistics.mean(elapsed),
        "whole_process": process_values,
        "report_work": {
            "sample_count": len(samples),
            "stage_slots_per_sample": len(samples[0]["stages"]),
            "active_stage_count_per_sample": sum(active_stage(stage) for stage in samples[0]["stages"]),
            "active_stage_count_total": len(all_sample_active),
            "active_stage_after_sum": sum(after_values),
            "active_stage_after_min": min(after_values) if after_values else None,
            "active_stage_after_max": max(after_values) if after_values else None,
            "active_stage_after_mean": statistics.mean(after_values) if after_values else None,
            "route_count": len(route_rows),
            "route_active_stage_count": len(route_active),
            "route_active_stage_counts_by_retention": {
                route["retention"]: sum(active_stage(stage) for stage in route.get("stages", []))
                for route in route_rows
            },
            "route_retained_before_commit_sum": sum(
                route.get("final_retained_before_commit") or 0 for route in route_rows
            ),
            "semantic_change_counts": sorted(
                {route.get("semantic_change_count") for route in route_rows}
            ),
            "output_bytes": report["output_bytes"],
            "output_capacity": report["output_capacity"],
            "route_signature": route_signature(route_rows),
            "report_signature": report_signature,
        },
        "correctness_all_true": bools_all_true(report["correctness"]),
        "terminal_returncode": terminal.get("returncode"),
        "terminal_timed_out": terminal.get("timed_out"),
        "terminal_wall_ns_excluded_from_metric": terminal.get("wall_ns_excluded_from_metric"),
        "launcher_actual_affinity": launch.get("actual_affinity"),
    }


def association_rows(rows: list[dict[str, Any]], target: str) -> list[dict[str, Any]]:
    target_values = [float(row[target]) for row in rows]
    features = [
        ("instructions", lambda row: row["whole_process"]["instructions"]),
        ("cycles", lambda row: row["whole_process"]["cycles"]),
        ("task_clock_msec", lambda row: row["whole_process"]["task_clock_msec"]),
        ("user_s", lambda row: row["whole_process"]["user_s"]),
        ("system_s", lambda row: row["whole_process"]["system_s"]),
        ("perf_context_switches", lambda row: row["whole_process"]["perf_context_switches"]),
        ("cpu_migrations", lambda row: row["whole_process"]["cpu_migrations"]),
        ("perf_page_faults", lambda row: row["whole_process"]["perf_page_faults"]),
        ("minor_faults", lambda row: row["whole_process"]["minor_faults"]),
        ("involuntary_context_switches", lambda row: row["whole_process"]["involuntary_context_switches"]),
        ("max_rss_kb", lambda row: row["whole_process"]["max_rss_kb"]),
    ]
    result = []
    for name, getter in features:
        values = [float(getter(row)) for row in rows]
        result.append(
            {
                "metric": name,
                **correlation_pair(target_values, values),
                "metric_summary": numeric_summary(values),
            }
        )
    result.sort(
        key=lambda item: (
            -(abs(item["spearman"]) if item["spearman"] is not None else -1.0),
            item["metric"],
        )
    )
    for index, item in enumerate(result, 1):
        item["rank_by_absolute_spearman"] = index if item["spearman"] is not None else None
    return result


def arm_summary(rows: list[dict[str, Any]]) -> dict[str, Any]:
    process_names = (
        "instructions",
        "cycles",
        "task_clock_msec",
        "user_s",
        "system_s",
        "perf_context_switches",
        "cpu_migrations",
        "perf_page_faults",
        "minor_faults",
        "involuntary_context_switches",
        "max_rss_kb",
    )
    p50_associations = association_rows(rows, "p50_ns")
    p95_associations = association_rows(rows, "p95_ns")
    instruction_p50 = next(item for item in p50_associations if item["metric"] == "instructions")
    instruction_p95 = next(item for item in p95_associations if item["metric"] == "instructions")
    return {
        "children": len(rows),
        "pairs": sorted(row["pair"] for row in rows),
        "orders": dict(Counter(row["order_position"] for row in rows)),
        "child_rows": [
            {
                "packet_id": row["packet_id"],
                "pair": row["pair"],
                "order_position": row["order_position"],
                "requested_affinity": row["requested_affinity"],
                "observed_affinity": row["observed_affinity"],
                "p50_ns": row["p50_ns"],
                "p95_ns": row["p95_ns"],
                "mean_ns": row["mean_ns"],
                "timed_sample_cv_population": row["timed_sample_cv_population"],
                "instructions": row["whole_process"]["instructions"],
                "cycles": row["whole_process"]["cycles"],
                "task_clock_msec": row["whole_process"]["task_clock_msec"],
                "cycles_per_instruction": row["whole_process"]["cycles_per_instruction"],
                "cycles_per_task_clock_msec": row["whole_process"]["cycles_per_task_clock_msec"],
                "stage_active_after_sum": row["report_work"]["active_stage_after_sum"],
                "stage_active_count_total": row["report_work"]["active_stage_count_total"],
                "route_count": row["report_work"]["route_count"],
                "route_active_stage_count": row["report_work"]["route_active_stage_count"],
                "output_bytes": row["report_work"]["output_bytes"],
                "output_capacity": row["report_work"]["output_capacity"],
            }
            for row in rows
        ],
        "timing": {
            name: numeric_summary([row[name] for row in rows])
            for name in ("p50_ns", "p95_ns", "mean_ns", "timed_sample_cv_population")
        },
        "whole_process": {
            name: numeric_summary([row["whole_process"][name] for row in rows])
            for name in process_names
        },
        "instruction_timing": {
            "p50": {
                "target": "report.samples[*].elapsed_ns linear p50 per child",
                "instruction": "whole-process perf instructions",
                **instruction_p50,
            },
            "p95": {
                "target": "report.samples[*].elapsed_ns linear p95 per child",
                "instruction": "whole-process perf instructions",
                **instruction_p95,
            },
        },
        "ranked_whole_process_associations": {
            "target_p50_ns": p50_associations,
            "target_p95_ns": p95_associations,
        },
    }


def self_test() -> dict[str, Any]:
    values = list(range(1, 13))
    assert linear_percentile(values, 0.50) == 6.5
    assert linear_percentile(values, 0.95) == 11.45
    assert correlation([1.0, 2.0, 3.0], [7.0, 8.0, 9.0])[0] == 1.0
    assert correlation([1.0, 2.0, 3.0], [4.0, 4.0, 4.0])[0] is None
    assert numeric_summary([9.0, 10.0, 11.0])["cv_population"] > 0
    return {
        "status": "pass",
        "linear_percentile_control": True,
        "correlation_constant_control": True,
        "population_cv_control": True,
        "workload_started": False,
    }


def revalidate_adr_binding() -> dict[str, Any]:
    files = sorted(ADR_ROOT.glob("[0-9][0-9][0-9][0-9]-*.md"))
    total = sum(path.stat().st_size for path in files)
    digest = raw_set_digest(files, ADR_ROOT)
    return {
        "root": str(ADR_ROOT),
        "count": len(files),
        "bytes": total,
        "set_sha256": digest,
        "expected_count": EXPECTED_ADR_COUNT,
        "expected_bytes": EXPECTED_ADR_BYTES,
        "expected_set_sha256": EXPECTED_ADR_SET,
        "match": len(files) == EXPECTED_ADR_COUNT
        and total == EXPECTED_ADR_BYTES
        and digest == EXPECTED_ADR_SET,
    }


def build_report() -> dict[str, Any]:
    errors: list[str] = []
    for path_string, expected_hash in EXPECTED.items():
        path = Path(path_string)
        actual = sha256_file(path) if path.is_file() else None
        if actual != expected_hash:
            errors.append(f"hash mismatch {path}: expected {expected_hash}, got {actual}")
    adr_binding = revalidate_adr_binding()
    if not adr_binding["match"]:
        errors.append("accepted ADR binding mismatch")
    manifest = load_json(MANIFEST)
    stage = load_json(STAGE)
    plan = load_json(PLAN)
    reduced = load_json(REDUCED)
    packet_paths = [Path(path) for path in manifest["packets"]]
    actual_packet_paths = sorted(PACKETS.glob("*.json"))
    if sorted(packet_paths) != actual_packet_paths:
        errors.append("manifest packet set differs from packet directory")
    packet_inventory = file_inventory(packet_paths, CAPTURE)
    if packet_inventory["missing"]:
        errors.append("manifest packet inventory has missing files")
    if not manifest.get("complete") or manifest.get("no_retry") is not True:
        errors.append("manifest is not complete/no-retry")
    if stage.get("status") != "complete" or stage.get("packet_count") != EXPECTED_PACKET_COUNT:
        errors.append("capture stage is not complete at expected count")
    if reduced.get("status") != "complete" or reduced.get("errors") != []:
        errors.append("sealed v6 reduction is not complete/error-free")
    if reduced.get("packet_count") != EXPECTED_PACKET_COUNT:
        errors.append("sealed v6 reduction packet count mismatch")
    packets = [load_json(path) for path in packet_paths]
    selected_packets = [packet for packet in packets if packet["context"]["cell_id"] == CELL_ID]
    if len(selected_packets) != EXPECTED_CHILDREN:
        errors.append(f"selected child count is {len(selected_packets)}, expected {EXPECTED_CHILDREN}")
    selected_ids = {packet["packet_id"] for packet in selected_packets}
    reduced_ids = {
        child["packet_id"]
        for child in reduced["cells"].get(CELL_ID, {}).get("children", [])
    }
    if selected_ids != reduced_ids:
        errors.append("selected packet ids differ from sealed v6 child ids")
    conditions = Counter(packet["context"]["condition"] for packet in selected_packets)
    if conditions != Counter({"pinned": EXPECTED_PER_ARM, "unpinned": EXPECTED_PER_ARM}):
        errors.append(f"arm counts mismatch: {conditions}")
    selected_paths: list[Path] = []
    for packet in selected_packets:
        packet_path = PACKETS / f"{packet['packet_id']}.json"
        selected_paths.append(packet_path)
        for descriptor in (*packet["sidecars"].values(), packet["stdout"], packet["stderr"]):
            path = Path(descriptor["path"])
            selected_paths.append(path)
            if not path.is_file():
                errors.append(f"missing selected input {path}")
            else:
                if path.stat().st_size != descriptor["bytes"]:
                    errors.append(f"size mismatch {path}")
                if sha256_file(path) != descriptor["sha256"]:
                    errors.append(f"hash mismatch sidecar/raw {path}")
    selected_inventory = file_inventory(selected_paths, CAPTURE)
    rows = [extract_child(packet) for packet in selected_packets]
    by_arm = {
        condition: sorted(
            (row for row in rows if row["condition"] == condition),
            key=lambda row: (row["pair"], row["packet_id"]),
        )
        for condition in ("pinned", "unpinned")
    }
    for condition, arm_rows in by_arm.items():
        if len(arm_rows) != EXPECTED_PER_ARM:
            errors.append(f"{condition} child count mismatch")
        if any(row["correctness_all_true"] is not True for row in arm_rows):
            errors.append(f"{condition} has failed report correctness flags")
        if any(row["observed_affinity"] != row["requested_affinity"] for row in arm_rows):
            errors.append(f"{condition} observed affinity differs from requested")
    identities = {
        "binary": sorted({packet["binary"]["sha256"] for packet in selected_packets}),
        "fixture": sorted({packet["fixture"]["sha256"] for packet in selected_packets}),
        "source": sorted({packet["context"]["source_manifest_sha256"] for packet in selected_packets}),
    }
    if identities["binary"] != [EXPECTED_BINARY]:
        errors.append("selected binary identity mismatch")
    if identities["fixture"] != [EXPECTED_FIXTURE]:
        errors.append("selected fixture identity mismatch")
    if identities["source"] != [EXPECTED_SOURCE]:
        errors.append("selected source manifest identity mismatch")
    signatures = [row["report_work"]["report_signature"] for row in rows]
    signature_counts = Counter(json.dumps(signature, sort_keys=True) for signature in signatures)
    if len(signature_counts) != 1:
        errors.append(f"report work/output signature varies across children: {len(signature_counts)}")
    signature = signatures[0] if signatures else {}
    work_keys = (
        "sample_count",
        "stage_slots_per_sample",
        "active_stage_count_per_sample",
        "active_stage_count_total",
        "active_stage_after_sum",
        "active_stage_after_min",
        "active_stage_after_max",
        "active_stage_after_mean",
        "route_count",
        "route_active_stage_count",
        "route_retained_before_commit_sum",
        "output_bytes",
        "output_capacity",
    )
    work_values = {
        key: sorted({row["report_work"][key] for row in rows}, key=lambda value: str(value))
        for key in work_keys
    }
    work_values["route_active_stage_counts_by_retention"] = sorted(
        {
            json.dumps(row["report_work"]["route_active_stage_counts_by_retention"], sort_keys=True)
            for row in rows
        }
    )
    work_values["semantic_change_counts"] = sorted(
        {json.dumps(row["report_work"]["semantic_change_counts"], sort_keys=True) for row in rows}
    )
    stage_keys = signature.get("stage_keys", [])
    stage_duration_candidates = [
        key for key in stage_keys
        if any(token in key.lower() for token in ("duration", "elapsed", "time", "ns"))
    ]
    report_structure = {
        "children": len(rows),
        "children_by_arm": dict(conditions),
        "samples_per_child": sorted({row["report_work"]["sample_count"] for row in rows}),
        "warmups": sorted({row["report_work"]["report_signature"]["warmups"] for row in rows}),
        "timed_region_fields": {
            "top_level": ["elapsed_ns"],
            "sample": ["samples[*].elapsed_ns"],
            "scope": "report.region_definition",
        },
        "stage_field_keys": stage_keys,
        "stage_duration_fields_present": bool(stage_duration_candidates),
        "stage_duration_field_candidates": stage_duration_candidates,
        "subphase_counter_fields_present": False,
        "subphase_counter_field_candidates": [],
        "allocator_timing_fields_present": False,
        "allocator_fields": {
            "allocation_all_null": all(
                row["report_work"]["report_signature"]["allocation"] is None for row in rows
            ),
            "allocator_instrumented_values": sorted(
                {row["report_work"]["report_signature"]["allocator_instrumented"] for row in rows}
            ),
            "allocator_timing_claim_values": sorted(
                {row["report_work"]["report_signature"]["allocator_timing_claim"] for row in rows}
            ),
        },
        "exact_work_and_output_route_values": work_values,
        "stable_across_all_children": len(signature_counts) == 1,
        "no_fast_slow_cutoff": True,
        "selection": "all docfloat/e1/default packet children; no timing or counter filter",
    }
    arms = {condition: arm_summary(arm_rows) for condition, arm_rows in by_arm.items()}
    return {
        "schema": "litchi.performance.0971.float_e1_work_analysis.v1",
        "status": "complete" if not errors else "blocked",
        "analysis_only": True,
        "workload_started": False,
        "build_started": False,
        "errors": errors,
        "input_binding": {
            "capture_root": str(CAPTURE),
            "manifest": {"path": str(MANIFEST), "sha256": EXPECTED[str(MANIFEST)]},
            "stage": {"path": str(STAGE), "sha256": EXPECTED[str(STAGE)]},
            "plan": {"path": str(PLAN), "sha256": EXPECTED[str(PLAN)]},
            "sealed_reducer": {"path": str(REDUCER), "sha256": EXPECTED[str(REDUCER)]},
            "sealed_reduction": {"path": str(REDUCED), "sha256": EXPECTED[str(REDUCED)]},
            "packet_inventory": packet_inventory,
            "selected_raw_input_inventory": selected_inventory,
            "selected_cell": CELL_ID,
            "selected_packet_ids": sorted(selected_ids),
            "expected_packet_count": EXPECTED_PACKET_COUNT,
            "selected_child_count": len(selected_packets),
            "identities": identities,
            "accepted_adr_binding": adr_binding,
        },
        "capture_revalidation": {
            "manifest_complete": manifest.get("complete"),
            "manifest_no_retry": manifest.get("no_retry"),
            "stage_status": stage.get("status"),
            "stage_packet_count": stage.get("packet_count"),
            "stage_expected_packet_count": stage.get("expected_packet_count"),
            "sealed_reduction_status": reduced.get("status"),
            "sealed_reduction_errors": reduced.get("errors"),
            "all_selected_terminal_returncodes_zero": all(
                row["terminal_returncode"] == 0 and not row["terminal_timed_out"] for row in rows
            ),
        },
        "scope": {
            "cell": CELL_ID,
            "arms": {"pinned": EXPECTED_PER_ARM, "unpinned": EXPECTED_PER_ARM},
            "samples_per_child": EXPECTED_SAMPLES,
            "warmups": EXPECTED_WARMUPS,
            "selection_policy": "all 24 retained Float/e1 children, analyzed separately by arm",
            "timed_region_scope": signature.get("region_definition"),
            "whole_process_perf_scope": plan["wrapper"]["scope"]["perf"],
            "whole_process_time_scope": plan["wrapper"]["scope"]["gnu_time"],
            "harness_scope": plan["wrapper"]["scope"]["harness"],
            "interpretation_limit": "descriptive association only; whole-process counters are not decomposed into report stages",
        },
        "report_structure": report_structure,
        "stage_duration_associations": {
            "status": "not_available",
            "timed_fields_available": ["report.samples[*].elapsed_ns"],
            "stage_duration_fields_available": [],
            "reason": "Retained stage records expose before/after/after_release capacities and requested text, but no per-stage duration or counter. All exposed stage/work/output-route values are constant across the 24 selected children.",
        },
        "arms": arms,
        "findings": [
            "The instruction/timing associations use all 12 children in each arm and no fast/slow cutoff.",
            "The report exposes one timed region per sample; it does not expose a decomposable per-stage duration or subphase counter.",
            "The exact report work and output-route signature is stable across all selected children, so the retained report cannot explain child-level timing spread through changing work/output counts.",
            "Perf instructions, cycles, and task-clock are whole-process wrapper counters and should be treated as correlated telemetry rather than an attribution of report-stage cost.",
        ],
        "self_test": self_test(),
    }


def markdown(report: dict[str, Any]) -> str:
    lines = [
        "# 0971 Float/e1 retained-work analysis",
        "",
        f"- Status: {report['status']}; analysis only; workload/build started: {report['workload_started']}/{report['build_started']}.",
        f"- Input: sealed 0970 {report['input_binding']['selected_cell']}; {report['input_binding']['selected_child_count']} children, 12 pinned and 12 unpinned, {report['scope']['samples_per_child']} timed samples each.",
        f"- Capture/reduction custody: manifest {report['input_binding']['manifest']['sha256'][:12]}, stage {report['input_binding']['stage']['sha256'][:12]}, v6 reduction {report['input_binding']['sealed_reduction']['sha256'][:12]}.",
        "",
        "## Timing and whole-process instruction association",
        "",
        "Each association below is descriptive across all 12 children in that arm. Pearson and Spearman are child-level correlations; no fast/slow threshold was selected.",
        "",
        "| arm | timing target | instruction range | timing range (ns) | Pearson | Spearman |",
        "|---|---|---:|---:|---:|---:|",
    ]
    for arm in ("pinned", "unpinned"):
        summary = report["arms"][arm]["instruction_timing"]
        timing = report["arms"][arm]["timing"]
        ins = report["arms"][arm]["whole_process"]["instructions"]
        for target in ("p50", "p95"):
            item = summary[target]
            timing_summary = timing[target + "_ns"]
            lines.append(
                f"| {arm} | {target} | {ins['min']:.0f}–{ins['max']:.0f} | "
                f"{timing_summary['min']:.0f}–{timing_summary['max']:.0f} | "
                f"{item['pearson'] if item['pearson'] is not None else 'null'} | "
                f"{item['spearman'] if item['spearman'] is not None else 'null'} |"
            )
    lines.extend(
        [
            "",
            "The strongest whole-process associations are ranked in the JSON under each arm. The retained report stage/work values are invariant: 80 samples, 32 stage slots per sample, one active stage per sample, three output routes, and stable output bytes/capacity and semantic-change counts. Stage records have capacity fields (before, after, after_release) but no duration or counter.",
            "",
            "## Scope and interpretation",
            "",
            f"- Timed region: {report['scope']['timed_region_scope']}",
            f"- Perf scope: {report['scope']['whole_process_perf_scope']}",
            f"- GNU time scope: {report['scope']['whole_process_time_scope']}",
            f"- Harness scope: {report['scope']['harness_scope']}",
            "- These associations do not attribute process-wide instructions to a stage. The report has no subphase timer or work counter from which such a decomposition could be computed.",
            "",
            "The machine-readable output is float-e1-work-v1.json; the reproducible helper is tools/analyze_float_e1_work_v1.py.",
        ]
    )
    return "\n".join(lines) + "\n"


def main() -> int:
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_MD.parent.mkdir(parents=True, exist_ok=True)
    report = build_report()
    OUT_JSON.write_text(json.dumps(report, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    OUT_MD.write_text(markdown(report), encoding="utf-8")
    print(
        json.dumps(
            {
                "status": report["status"],
                "errors": report["errors"],
                "json": str(OUT_JSON),
                "markdown": str(OUT_MD),
                "children": report["input_binding"]["selected_child_count"],
                "arms": report["input_binding"]["selected_child_count"] // 2,
            },
            sort_keys=True,
        )
    )
    return 0 if report["status"] == "complete" else 1


if __name__ == "__main__":
    sys.exit(main())
