#!/usr/bin/env python3
"""Independent, read-only audit and scalar reduction for the frozen 0970 capture.

This deliberately does not import or execute the sealed runner/reducer.  It reads
packet JSON and sidecars, checks custody/shape/gates, and emits reduced scalars.
"""
from __future__ import annotations

import argparse
import collections
import datetime as dt
import hashlib
import json
import math
import os
import re
import statistics
from pathlib import Path
from typing import Any

EXPECTED_EVENTS = [
    "task-clock",
    "cycles",
    "instructions",
    "context-switches",
    "cpu-migrations",
    "page-faults",
]
SAMPLE_BOOL_KEYS = [
    "allocator_retained_matches_output_capacity",
    "exact_output_equal_reference",
    "final_paragraph_semantic_readback",
    "final_public_reader_reopen",
    "retained_capacities_within_ceiling",
]
ORACLE_BOOL_KEYS = [
    "expected_structurally_valid",
    "oracle_ok",
    "output_stream_bytes_match_expected",
    "output_stream_paths_match_expected",
    "output_structurally_valid",
    "root_clsid_preserved",
    "semantic_directory_metadata_matches_expected",
    "semantic_reopen_ok",
    "source_structurally_valid",
    "storage_clsids_preserved",
    "unchanged_source_output_stream_bytes_match",
    "unchanged_stream_bytes_match_expected",
]
CORRECTNESS_KEYS = [
    "all_stages_change_text",
    "allocator_retained_matches_output_capacity",
    "directory_metadata_preserved",
    "exact_cross_route_outputs",
    "exact_noop_source",
    "final_paragraph_semantic_readback",
    "final_public_reader_reopen",
    "foreign_patch_refused",
    "forward_patch_exact_output",
    "inverse_patch_exact_source",
    "noop_shares_source_allocation",
    "operation_limit_refusal_atomic",
    "release_route_clears_every_stage",
    "retained_capacities_within_ceiling",
    "same_value_preserves_retention",
    "semantic_change_count_matches_edit_count",
    "semantic_oracles_passed",
    "stale_patch_refused",
    "untouched_streams_preserved",
    "zero_route_clears_retention",
]
STAGE_KEYS = ["after", "after_release", "before", "requested_text", "stage"]
PACKET_ID_RE = re.compile(
    r"^diag-(?P<case>[^-]+)-e(?P<edits>[0-9]+)-(?P<retention>[^-]+)-p(?P<pair>[0-9]{2})-(?P<condition>pinned|unpinned)$"
)


class AuditError(Exception):
    pass


def duplicate_hook(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    out: dict[str, Any] = {}
    for key, value in pairs:
        if key in out:
            raise AuditError(f"duplicate JSON key: {key}")
        out[key] = value
    return out


def read_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=duplicate_hook)
    except Exception as exc:
        raise AuditError(f"cannot parse {path}: {exc}") from exc


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        while True:
            chunk = fh.read(1024 * 1024)
            if not chunk:
                return h.hexdigest()
            h.update(chunk)


def file_meta(path: Path) -> dict[str, Any]:
    return {"bytes": path.stat().st_size, "sha256": sha256_file(path)}


def finite_number(value: Any) -> bool:
    return type(value) in (int, float) and math.isfinite(float(value))


def integer(value: Any) -> bool:
    return type(value) is int


def quantile(values: list[float], q: float) -> float:
    if not values:
        raise AuditError("quantile of empty vector")
    vals = sorted(float(x) for x in values)
    index = (len(vals) - 1) * q
    low = int(math.floor(index))
    high = int(math.ceil(index))
    if low == high:
        return vals[low]
    fraction = index - low
    return vals[low] + fraction * (vals[high] - vals[low])


def pearson(xs: list[float], ys: list[float]) -> float | None:
    if len(xs) != len(ys) or len(xs) < 2:
        return None
    mx, my = statistics.mean(xs), statistics.mean(ys)
    dx = [x - mx for x in xs]
    dy = [y - my for y in ys]
    den = math.sqrt(sum(x * x for x in dx) * sum(y * y for y in dy))
    return None if den == 0 else sum(x * y for x, y in zip(dx, dy)) / den


def spearman(xs: list[float], ys: list[float]) -> float | None:
    def ranks(vals: list[float]) -> list[float]:
        order = sorted(range(len(vals)), key=lambda i: vals[i])
        result = [0.0] * len(vals)
        pos = 0
        while pos < len(order):
            end = pos + 1
            while end < len(order) and vals[order[end]] == vals[order[pos]]:
                end += 1
            rank = (pos + end - 1) / 2.0 + 1.0
            for j in order[pos:end]:
                result[j] = rank
            pos = end
        return result
    return pearson(ranks(xs), ranks(ys))


def scalar_stats(values: list[float]) -> dict[str, Any]:
    vals = [float(x) for x in values]
    mean = statistics.mean(vals)
    sample_sd = statistics.stdev(vals) if len(vals) > 1 else 0.0
    return {
        "n": len(vals),
        "min": min(vals),
        "p25": quantile(vals, 0.25),
        "p50": quantile(vals, 0.50),
        "p75": quantile(vals, 0.75),
        "p95": quantile(vals, 0.95),
        "p99": quantile(vals, 0.99),
        "max": max(vals),
        "mean": mean,
        "sample_sd": sample_sd,
        "sample_cv_pct": (sample_sd / mean * 100.0) if mean else None,
    }


def check_bound_refs(root: Any, errors: list[str], labels: list[str]) -> dict[str, Any]:
    """Verify every object carrying path/bytes/sha256 in the sealed receipt."""
    refs: dict[str, dict[str, Any]] = {}

    def walk(value: Any, label: str) -> None:
        if isinstance(value, dict):
            if isinstance(value.get("path"), str) and isinstance(value.get("sha256"), str) and "bytes" in value:
                refs.setdefault(value["path"], {"label": label, "expected": value})
            for key, child in value.items():
                walk(child, f"{label}.{key}")
        elif isinstance(value, list):
            for index, child in enumerate(value):
                walk(child, f"{label}[{index}]")

    walk(root, "root_go")
    checked: list[dict[str, Any]] = []
    for path_text, data in sorted(refs.items()):
        path = Path(path_text)
        expected = data["expected"]
        if not path.is_file():
            errors.append(f"bound reference missing: {path_text}")
            checked.append({"path": path_text, "exists": False, "expected": expected})
            continue
        actual = file_meta(path)
        good = actual["bytes"] == expected.get("bytes") and actual["sha256"] == expected.get("sha256")
        if not good:
            errors.append(f"bound reference mismatch: {path_text}")
        checked.append({"path": path_text, "label": data["label"], "expected": expected, "actual": actual, "match": good})
    return {"unique_references": len(checked), "all_match": not any(x.get("match") is False or x.get("exists") is False for x in checked), "files": checked}


def parse_perf(path: Path, errors: list[str], packet_id: str) -> tuple[dict[str, dict[str, Any]], int]:
    rows: dict[str, dict[str, Any]] = {}
    comments = 0
    try:
        text = path.read_text(encoding="utf-8")
    except Exception as exc:
        errors.append(f"{packet_id}: perf unreadable: {exc}")
        return rows, comments
    for line_no, raw in enumerate(text.splitlines(), 1):
        line = raw.strip()
        if not line:
            continue
        if line.startswith("#"):
            if not line.startswith("# started on "):
                errors.append(f"{packet_id}: perf unknown header at line {line_no}")
            comments += 1
            continue
        fields = line.split(",")
        if len(fields) < 5:
            errors.append(f"{packet_id}: perf malformed line {line_no}")
            continue
        event = fields[2]
        if event in rows:
            errors.append(f"{packet_id}: duplicate perf event {event}")
            continue
        raw_value, unit = fields[0], fields[1]
        if not raw_value or raw_value.startswith("<"):
            errors.append(f"{packet_id}: perf unavailable {event}")
            continue
        try:
            value = float(raw_value)
        except ValueError:
            errors.append(f"{packet_id}: perf nonnumeric value {event}")
            continue
        if not math.isfinite(value) or value < 0:
            errors.append(f"{packet_id}: perf invalid value {event}")
        runtime = None
        if len(fields) > 3 and fields[3]:
            try:
                runtime = float(fields[3])
            except ValueError:
                errors.append(f"{packet_id}: perf invalid runtime {event}")
        percent = None
        if len(fields) > 4 and fields[4]:
            try:
                percent = float(fields[4])
            except ValueError:
                errors.append(f"{packet_id}: perf invalid percentage {event}")
        if runtime is not None and (not math.isfinite(runtime) or runtime <= 0):
            errors.append(f"{packet_id}: perf nonpositive runtime {event}")
        if percent is not None and (not math.isfinite(percent) or percent < 0):
            errors.append(f"{packet_id}: perf invalid percentage {event}")
        rows[event] = {"value": value, "unit": unit, "runtime": runtime, "percent": percent}
    if set(rows) != set(EXPECTED_EVENTS):
        errors.append(f"{packet_id}: perf event set mismatch {sorted(rows)}")
    return rows, comments


def check_bool_map(value: Any, keys: list[str], label: str, errors: list[str]) -> int:
    if not isinstance(value, dict) or set(value) != set(keys):
        errors.append(f"{label}: key set mismatch")
        return 0
    failed = 0
    for key in keys:
        if type(value[key]) is not bool or value[key] is not True:
            failed += 1
            errors.append(f"{label}.{key}: not true")
    return failed


def check_oracle_map(value: Any, label: str, errors: list[str]) -> None:
    if not isinstance(value, dict) or set(value) != set(ORACLE_BOOL_KEYS) | {"failure_reasons"}:
        errors.append(f"{label}: key set mismatch")
        return
    check_bool_map({key: value.get(key) for key in ORACLE_BOOL_KEYS}, ORACLE_BOOL_KEYS, label, errors)
    if value.get("failure_reasons") != []:
        errors.append(f"{label}: failure reasons are nonempty")


def check_packet(
    packet_path: Path,
    packet: dict[str, Any],
    plan: dict[str, Any],
    capture: Path,
    errors: list[str],
    metrics: dict[tuple[str, str, int], dict[str, Any]],
    observed_files: set[str],
) -> dict[str, Any]:
    packet_id = packet.get("packet_id", packet_path.name)
    context = packet.get("context")
    if not isinstance(context, dict):
        errors.append(f"{packet_id}: context missing")
        return {}
    match = PACKET_ID_RE.match(str(packet_id))
    if not match:
        errors.append(f"{packet_id}: packet id shape mismatch")
        return {}
    case = match.group("case")
    edits = int(match.group("edits"))
    retention = match.group("retention")
    pair = int(match.group("pair"))
    condition = match.group("condition")
    cell_id = f"{case}/e{edits}/{retention}"
    expected_affinity = [31] if condition == "pinned" else list(range(32))
    expected_position = "first" if ((pair % 2 == 0 and condition == "pinned") or (pair % 2 == 1 and condition == "unpinned")) else "second"
    expected_order = ["pinned", "unpinned"] if pair % 2 == 0 else ["unpinned", "pinned"]
    expected_cell = {"case": case, "edits": edits, "retention": retention, "cell_id": cell_id}
    if packet_path.resolve() != Path(str(packet.get("output_root", capture))).resolve() / "packets" / (packet_id + ".json"):
        errors.append(f"{packet_id}: packet output/path identity mismatch")
    if packet.get("schema") != "litchi.performance.0970.runner.v2":
        errors.append(f"{packet_id}: runner schema mismatch")
    if packet.get("output_root") != str(capture):
        errors.append(f"{packet_id}: output root mismatch")
    if packet.get("validation_error") is not None:
        errors.append(f"{packet_id}: validation_error is non-null")
    terminal = packet.get("terminal")
    if not isinstance(terminal, dict) or terminal.get("returncode") != 0 or terminal.get("timed_out") is not False or terminal.get("launch_error") is not None:
        errors.append(f"{packet_id}: terminal failure")
    if context.get("packet_id") != packet_id or context.get("cell_id") != cell_id or context.get("condition") != condition:
        errors.append(f"{packet_id}: context identity mismatch")
    if context.get("cell") != expected_cell:
        errors.append(f"{packet_id}: context cell mismatch")
    for key, expected in (("pair", pair), ("expected_affinity", expected_affinity), ("observed_affinity", expected_affinity), ("order", expected_order), ("order_position", expected_position)):
        if context.get(key) != expected:
            errors.append(f"{packet_id}: context {key} mismatch")
    if context.get("lang") != "C" or context.get("lc_all") != "C" or context.get("python_dontwritebytecode") != "1":
        errors.append(f"{packet_id}: context environment mismatch")

    binary = plan["binary"]
    fixture = plan["fixtures"][case]
    if packet.get("binary") != {"path": binary["path"], "sha256": binary["sha256"]}:
        errors.append(f"{packet_id}: binary binding mismatch")
    if packet.get("fixture") != fixture:
        errors.append(f"{packet_id}: fixture binding mismatch")
    for key, expected in (("binary_sha256", binary["sha256"]), ("fixture_sha256", fixture["sha256"]), ("source_manifest_sha256", plan["source"]["manifest"]["sha256"]), ("source_file_manifest_sha256", plan["source"]["file_manifest"]["sha256"]), ("launcher_sha256", plan["helpers"]["launcher"]["sha256"])):
        context_key = key if key not in {"source_manifest_sha256", "source_file_manifest_sha256"} else key
        if context.get(context_key) != expected:
            errors.append(f"{packet_id}: context {context_key} mismatch")

    sidecar_refs = packet.get("sidecars")
    raw_refs = {"stdout": packet.get("stdout"), "stderr": packet.get("stderr")}
    if not isinstance(sidecar_refs, dict) or set(sidecar_refs) != {"launch", "perf", "time"}:
        errors.append(f"{packet_id}: sidecar key set mismatch")
        sidecar_refs = {}
    for label, ref in list(sidecar_refs.items()) + list(raw_refs.items()):
        if not isinstance(ref, dict) or not isinstance(ref.get("path"), str) or not isinstance(ref.get("sha256"), str) or not integer(ref.get("bytes")):
            errors.append(f"{packet_id}: invalid {label} file reference")
            continue
        target = Path(ref["path"])
        try:
            rel = target.resolve().relative_to(capture.resolve())
            rels = str(rel)
        except ValueError:
            errors.append(f"{packet_id}: {label} path outside capture")
            continue
        expected_prefix = "sidecars/" if label in {"launch", "perf", "time"} else "raw/"
        if not rels.startswith(expected_prefix):
            errors.append(f"{packet_id}: {label} path class mismatch")
        observed_files.add(rels)
        if not target.is_file():
            errors.append(f"{packet_id}: {label} missing")
            continue
        actual = file_meta(target)
        if actual["bytes"] != ref["bytes"] or actual["sha256"] != ref["sha256"]:
            errors.append(f"{packet_id}: {label} hash/size mismatch")

    launch_ref = sidecar_refs.get("launch")
    time_ref = sidecar_refs.get("time")
    perf_ref = sidecar_refs.get("perf")
    launch = read_json(Path(launch_ref["path"])) if isinstance(launch_ref, dict) and Path(launch_ref["path"]).is_file() else {}
    time_data = read_json(Path(time_ref["path"])) if isinstance(time_ref, dict) and Path(time_ref["path"]).is_file() else {}
    perf_rows, perf_comments = parse_perf(Path(perf_ref["path"]), errors, packet_id) if isinstance(perf_ref, dict) and Path(perf_ref["path"]).is_file() else ({}, 0)
    if launch:
        if launch.get("schema") != "litchi.performance.0970.launch.v2" or launch.get("packet_id") != packet_id or launch.get("helper_sha256") != plan["helpers"]["launcher"]["sha256"]:
            errors.append(f"{packet_id}: launcher identity mismatch")
        for key, expected in (("requested_affinity", expected_affinity), ("actual_affinity", expected_affinity)):
            if launch.get(key) != expected:
                errors.append(f"{packet_id}: launcher {key} mismatch")
        if type(launch.get("pid")) is not int or type(launch.get("exec_pid")) is not int or launch.get("pid") != launch.get("exec_pid") or launch.get("pid") <= 0:
            errors.append(f"{packet_id}: launcher PID invalid")
        if launch.get("exec_path") != binary["path"] or launch.get("execve") is not True or not isinstance(launch.get("cgroup"), list) or not launch.get("cgroup"):
            errors.append(f"{packet_id}: launcher execution identity invalid")
        argv = launch.get("exec_argv")
        expected_argv = [binary["path"], "--case", case, "--input", fixture["path"], "--edits", str(edits), "--retention", retention, "--warmup", "5", "--samples", "80"]
        if argv != expected_argv:
            errors.append(f"{packet_id}: launcher argv mismatch")
    if time_data:
        expected_time_keys = {"user_s", "system_s", "voluntary_context_switches", "involuntary_context_switches", "minor_faults", "major_faults", "max_rss_kb"}
        if set(time_data) != expected_time_keys:
            errors.append(f"{packet_id}: GNU time key set mismatch")
        for key in ("user_s", "system_s"):
            if not finite_number(time_data.get(key)) or float(time_data[key]) < 0:
                errors.append(f"{packet_id}: GNU time {key} invalid")
        for key in ("voluntary_context_switches", "involuntary_context_switches", "minor_faults", "major_faults", "max_rss_kb"):
            if not integer(time_data.get(key)) or time_data[key] < 0:
                errors.append(f"{packet_id}: GNU time {key} invalid")

    expected_command = [
        "/usr/bin/perf", "stat", "-x,", "-o", str(perf_ref["path"]), "-e", ",".join(EXPECTED_EVENTS), "--",
        "/usr/bin/time", "-f", plan["wrapper"]["time_format"], "-o", str(time_ref["path"]),
        "/usr/bin/python3.14", plan["helpers"]["launcher"]["path"], "--metadata", str(launch_ref["path"]),
        "--requested-affinity", json.dumps(expected_affinity, separators=(",", ":")), "--", binary["path"],
        "--case", case, "--input", fixture["path"], "--edits", str(edits), "--retention", retention,
        "--warmup", "5", "--samples", "80",
    ]
    if packet.get("command") != expected_command:
        errors.append(f"{packet_id}: command binding mismatch")

    report = packet.get("report")
    if not isinstance(report, dict):
        errors.append(f"{packet_id}: report missing")
        return {}
    if report.get("case") != case or report.get("edits") != edits or report.get("retention") != "default_8mib" or report.get("samples_requested") != 80 or report.get("warmups") != 5:
        errors.append(f"{packet_id}: report identity/config mismatch")
    if report.get("schema_version") != 1 or report.get("schema") != "litchi.performance.0917.doc_retention.v1":
        errors.append(f"{packet_id}: report schema mismatch")
    if report.get("allocator_instrumented") is not False or report.get("allocator_timing_claim") is not False or report.get("timing_claim") is not True:
        errors.append(f"{packet_id}: report timing/allocation claim mismatch")
    check_bool_map(report.get("correctness"), CORRECTNESS_KEYS, f"{packet_id}.correctness", errors)
    elapsed = report.get("elapsed_ns")
    samples = report.get("samples")
    if not isinstance(elapsed, list) or len(elapsed) != 80 or not isinstance(samples, list) or len(samples) != 80:
        errors.append(f"{packet_id}: sample vector length mismatch")
        return {}
    if elapsed != [sample.get("elapsed_ns") for sample in samples]:
        errors.append(f"{packet_id}: top/sample elapsed mismatch")
    if any(not integer(value) or value <= 0 for value in elapsed):
        errors.append(f"{packet_id}: nonpositive sample elapsed")
    if report.get("exact_output_equal_reference") is not True:
        errors.append(f"{packet_id}: top exact output gate false")
    for sample_index, sample in enumerate(samples):
        if not isinstance(sample, dict) or sample.get("index") != sample_index:
            errors.append(f"{packet_id}: sample index mismatch {sample_index}")
            continue
        check_bool_map({key: sample.get(key) for key in SAMPLE_BOOL_KEYS}, SAMPLE_BOOL_KEYS, f"{packet_id}.sample[{sample_index}]", errors)
        check_oracle_map(sample.get("oracle"), f"{packet_id}.sample[{sample_index}].oracle", errors)
        if sample.get("semantic_change_count") != edits:
            errors.append(f"{packet_id}: sample semantic change count")
        if sample.get("output_bytes") != report.get("output_bytes") or sample.get("output_capacity") != report.get("output_capacity") or sample.get("output_sha256") != report.get("output_sha256"):
            errors.append(f"{packet_id}: sample output identity mismatch")
        stages = sample.get("stages")
        if not isinstance(stages, list) or len(stages) != 32 or any(not isinstance(row, dict) or set(row) != set(STAGE_KEYS) for row in (stages if isinstance(stages, list) else [])):
            errors.append(f"{packet_id}: sample stage shape mismatch")
    routes = report.get("route_outputs")
    if not isinstance(routes, list) or len(routes) != 3:
        errors.append(f"{packet_id}: route count mismatch")
    else:
        route_names = ["zero", "default_8mib", "release_8mib"]
        for route, expected_route in zip(routes, route_names):
            if route.get("retention") != expected_route or route.get("semantic_change_count") != edits:
                errors.append(f"{packet_id}: route identity mismatch {expected_route}")
            if route.get("output_bytes") != report.get("output_bytes") or route.get("output_capacity") != report.get("output_capacity") or route.get("output_sha256") != report.get("output_sha256"):
                errors.append(f"{packet_id}: route output mismatch {expected_route}")
            check_oracle_map(route.get("oracle"), f"{packet_id}.route[{expected_route}].oracle", errors)
            stages = route.get("stages")
            if not isinstance(stages, list) or len(stages) != 32 or any(not isinstance(row, dict) or set(row) != set(STAGE_KEYS) for row in (stages if isinstance(stages, list) else [])):
                errors.append(f"{packet_id}: route stage shape mismatch {expected_route}")
    if report.get("input") != fixture["path"] or report.get("source_bytes") != fixture.get("bytes") or report.get("source_sha256") != fixture.get("sha256"):
        errors.append(f"{packet_id}: report source/fixture identity mismatch")

    if launch and perf_rows and time_data:
        metrics[(cell_id, condition, pair)] = {
            "packet_id": packet_id,
            "elapsed": [float(value) for value in elapsed],
            "p50_ns": quantile([float(value) for value in elapsed], 0.50),
            "p95_ns": quantile([float(value) for value in elapsed], 0.95),
            "p99_ns": quantile([float(value) for value in elapsed], 0.99),
            "perf": perf_rows,
            "time": time_data,
            "pid": launch.get("pid"),
            "condition": condition,
            "pair": pair,
            "cell_id": cell_id,
        }
    return {"packet_id": packet_id, "cell_id": cell_id, "condition": condition, "pair": pair, "perf_comments": perf_comments}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--capture", type=Path, required=True)
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--root-go", type=Path, required=True)
    parser.add_argument("--terminal", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--scalars", type=Path, required=True)
    args = parser.parse_args()
    capture = args.capture.resolve()
    errors: list[str] = []
    plan = read_json(args.plan)
    root_go = read_json(args.root_go)
    terminal = read_json(args.terminal)
    manifest = read_json(capture / "capture-manifest-v2.json")
    stage = read_json(capture / "capture-stage-v2.json")
    controls = read_json(capture / "affinity-controls-v2.json")
    now = dt.datetime.now(dt.timezone.utc).isoformat()

    if plan.get("schema") != "litchi.performance.0970.plan.v2":
        errors.append("plan schema mismatch")
    if root_go.get("schema") != "litchi.performance.0970.root_go.v1" or root_go.get("authorized") is not True or root_go.get("status") != "authorized" or root_go.get("no_retry") is not True:
        errors.append("root GO authorization/no-retry mismatch")
    if terminal.get("schema") != "litchi.performance.0970.root_terminal.v1" or terminal.get("exit_code") != 0 or terminal.get("retries") != 0:
        errors.append("terminal receipt mismatch")
    if manifest.get("schema") != "litchi.performance.0970.capture_manifest.v2" or manifest.get("complete") is not True or manifest.get("no_retry") is not True:
        errors.append("capture manifest completion/no-retry mismatch")
    if stage.get("schema") != "litchi.performance.0970.capture_stage.v2" or stage.get("status") != "complete" or stage.get("failure") is not None or stage.get("packet_count") != 96 or stage.get("expected_packet_count") != 96 or stage.get("no_retry") is not True:
        errors.append("capture stage mismatch")
    allowed = list(range(32))
    if stage.get("allowed_affinity") != allowed or stage.get("pin_cpu") != 31 or stage.get("lang") != "C" or stage.get("lc_all") != "C" or stage.get("python_dontwritebytecode") != "1":
        errors.append("capture stage environment/affinity mismatch")
    if controls.get("schema") != "litchi.performance.0970.affinity_controls.v2" or controls.get("allowed") != allowed or controls.get("pin_cpu") != 31:
        errors.append("affinity controls envelope mismatch")
    for condition, target in (("pinned", [31]), ("unpinned", allowed)):
        row = controls.get("controls", {}).get(condition, {})
        if row.get("condition") != condition or any(row.get(key) != target for key in ("before", "target", "after", "restored")):
            errors.append(f"affinity controls {condition} mismatch")

    # Check all bound receipt metadata independently.  This may hash the bound
    # executable/tools, but performs no execution or build.
    bound = check_bound_refs(root_go, errors, [])

    packet_paths = manifest.get("packets")
    if not isinstance(packet_paths, list) or len(packet_paths) != 96 or len(set(packet_paths)) != 96:
        errors.append("manifest packet list count/uniqueness mismatch")
        packet_paths = packet_paths if isinstance(packet_paths, list) else []
    packet_path_set = {str(Path(path).resolve()) for path in packet_paths}
    packet_dir_files = sorted(capture.glob("packets/*.json"))
    if {str(path.resolve()) for path in packet_dir_files} != packet_path_set:
        errors.append("manifest packet paths differ from packet directory")

    expected_root_files = {
        "affinity-control-pinned.json", "affinity-control-unpinned.json", "affinity-controls-v2.json",
        "capture-manifest-v2.json", "capture-stage-v2.json", "quiescence-v2.json",
    }
    actual_rel_files = {str(path.relative_to(capture)) for path in capture.rglob("*") if path.is_file()}
    expected_rel_files = set(expected_root_files)
    expected_rel_files.update(str(path.relative_to(capture)) for path in packet_dir_files)
    for directory, labels in (("raw", ("stdout", "stderr")), ("sidecars", ("launch", "perf", "time"))):
        # referenced files are added below, so only enforce no unreferenced files
        # after the packet loop has populated observed_files.
        pass

    metrics: dict[tuple[str, str, int], dict[str, Any]] = {}
    observed_files: set[str] = set()
    packet_id_set: set[str] = set()
    parse_records: list[dict[str, Any]] = []
    for path_text in packet_paths:
        path = Path(path_text)
        try:
            packet = read_json(path)
        except AuditError as exc:
            errors.append(str(exc))
            continue
        record = check_packet(path, packet, plan, capture, errors, metrics, observed_files)
        parse_records.append(record)
        packet_id = packet.get("packet_id")
        if packet_id in packet_id_set:
            errors.append(f"duplicate packet id {packet_id}")
        packet_id_set.add(packet_id)

    # The packet loop is the authority for raw/sidecar references.  Require the
    # capture tree to contain exactly those files plus the six root custody files.
    expected_rel_files.update(observed_files)
    if actual_rel_files != expected_rel_files:
        missing = sorted(expected_rel_files - actual_rel_files)
        extra = sorted(actual_rel_files - expected_rel_files)
        if missing:
            errors.append(f"capture files missing: {missing[:8]}")
        if extra:
            errors.append(f"capture files unreferenced: {extra[:8]}")

    pid_values = [item.get("pid") for item in metrics.values() if item.get("pid") is not None]
    if len(pid_values) != len(set(pid_values)):
        errors.append("launcher PIDs are not unique")

    expected_tuples = {(f"{cell['case']}/e{cell['edits']}/{cell['retention']}", condition, pair) for cell in plan["matrix"]["cells"] for condition in ("pinned", "unpinned") for pair in range(1, plan["matrix"]["pairs_per_cell"] + 1)}
    if set(metrics) != expected_tuples:
        errors.append("packet metric tuple coverage mismatch")

    # Independent scalar reduction, retaining child-level summaries and paired
    # arm comparisons but no raw sample vectors.
    cell_summaries: dict[str, Any] = {}
    child_rows: list[dict[str, Any]] = []
    association: dict[str, Any] = {}
    for cell in plan["matrix"]["cells"]:
        cell_id = f"{cell['case']}/e{cell['edits']}/{cell['retention']}"
        cell_summaries[cell_id] = {"pinned": {}, "unpinned": {}, "paired": []}
        for condition in ("pinned", "unpinned"):
            rows = [metrics[(cell_id, condition, pair)] for pair in range(1, plan["matrix"]["pairs_per_cell"] + 1) if (cell_id, condition, pair) in metrics]
            p50s = [row["p50_ns"] for row in rows]
            if p50s:
                stats = scalar_stats(p50s)
                event_stats: dict[str, Any] = {}
                for event in ("task-clock", "cycles", "instructions"):
                    vals = [row["perf"][event]["value"] for row in rows]
                    event_stats[event] = {"unit": rows[0]["perf"][event]["unit"], **scalar_stats(vals)}
                    association.setdefault(cell_id, {}).setdefault(condition, {})[event] = {
                        "n": len(rows),
                        "elapsed_p50_vs_event_pearson": pearson(p50s, vals),
                        "elapsed_p50_vs_event_spearman": spearman(p50s, vals),
                        "scope": "whole launched GNU time/launcher/baseline process tree; descriptive association only",
                    }
                cell_summaries[cell_id][condition] = {
                    "child_p50_elapsed_ns": stats,
                    "whole_process_perf": event_stats,
                    "child_pairs": [row["pair"] for row in rows],
                }
                for row in rows:
                    child_rows.append({
                        "cell_id": cell_id,
                        "condition": condition,
                        "pair": row["pair"],
                        "packet_id": row["packet_id"],
                        "sample_count": len(row["elapsed"]),
                        "elapsed_ns": {key: value for key, value in scalar_stats(row["elapsed"]).items() if key != "sample_cv_pct"},
                        "p50_elapsed_ns": row["p50_ns"],
                        "whole_process_perf": {event: {"value": data["value"], "unit": data["unit"], "runtime": data["runtime"], "percent": data["percent"]} for event, data in row["perf"].items()},
                        "gnu_time": row["time"],
                        "exec_pid": row["pid"],
                    })
        for pair in range(1, plan["matrix"]["pairs_per_cell"] + 1):
            pin = metrics.get((cell_id, "pinned", pair))
            unpin = metrics.get((cell_id, "unpinned", pair))
            if pin and unpin:
                cell_summaries[cell_id]["paired"].append({
                    "pair": pair,
                    "order": ["pinned", "unpinned"] if pair % 2 == 0 else ["unpinned", "pinned"],
                    "pinned_p50_ns": pin["p50_ns"],
                    "unpinned_p50_ns": unpin["p50_ns"],
                    "pinned_over_unpinned_p50": pin["p50_ns"] / unpin["p50_ns"] if unpin["p50_ns"] else None,
                })

    def digest_file_lines(lines: list[str]) -> str:
        return hashlib.sha256(("\n".join(lines) + "\n").encode()).hexdigest()

    packet_hash_lines = []
    for path in packet_dir_files:
        meta = file_meta(path)
        packet_hash_lines.append(f"{path.relative_to(capture)}\t{meta['bytes']}\t{meta['sha256']}")
    referenced_hash_lines = []
    for rel in sorted(observed_files):
        path = capture / rel
        if path.is_file():
            meta = file_meta(path)
            referenced_hash_lines.append(f"{rel}\t{meta['bytes']}\t{meta['sha256']}")

    perf_comment_counts = collections.Counter(record.get("perf_comments", 0) for record in parse_records)
    all_correctness_pass = not any("correctness" in err or ".sample[" in err or ".route[" in err for err in errors)
    scalars = {
        "schema": "litchi.performance.0970.independent_scalars.v1",
        "generated_utc": now,
        "scope": "96 frozen historical-baseline children; no candidate qualification; descriptive arm/telemetry scalars",
        "binding": {
            "capture_root": str(capture),
            "plan": {"path": str(args.plan), **file_meta(args.plan)},
            "root_go": {"path": str(args.root_go), **file_meta(args.root_go)},
            "terminal": {"path": str(args.terminal), **file_meta(args.terminal)},
            "packet_count": len(packet_paths),
            "packet_hash_list_sha256": digest_file_lines(packet_hash_lines),
        },
        "estimator": {
            "sample_quantiles": "linear interpolation at (n-1)*q over each 80-sample elapsed_ns vector",
            "child_cv": "sample standard deviation / child mean * 100 across 12 child p50 elapsed_ns values",
            "association": "Pearson and tie-aware Spearman over 12 child p50 elapsed_ns values versus whole-process perf event values; descriptive only",
        },
        "cells": cell_summaries,
        "association": association,
        "child_rows": sorted(child_rows, key=lambda row: (row["cell_id"], row["condition"], row["pair"])),
        "counts": {
            "packet_count": len(packet_paths),
            "metric_rows": len(metrics),
            "sample_rows": sum(len(row["elapsed"]) for row in metrics.values()),
            "expected_sample_rows": 96 * 80,
            "perf_event_rows": sum(len(row["perf"]) for row in metrics.values()),
            "expected_perf_event_rows": 96 * len(EXPECTED_EVENTS),
            "perf_header_comment_distribution": dict(sorted((str(key), value) for key, value in perf_comment_counts.items())),
        },
    }
    args.scalars.parent.mkdir(parents=True, exist_ok=True)
    args.scalars.write_text(json.dumps(scalars, indent=2, sort_keys=True) + "\n", encoding="utf-8")

    audit = {
        "schema": "litchi.performance.0970.independent_raw_audit.v1",
        "generated_utc": now,
        "status": "pass" if not errors else "blocked",
        "scope": "read-only independent audit of frozen 0970 raw capture; shared reducer not imported or executed",
        "bindings": {
            "plan": {"path": str(args.plan), **file_meta(args.plan)},
            "root_go": {"path": str(args.root_go), **file_meta(args.root_go)},
            "terminal": {"path": str(args.terminal), **file_meta(args.terminal)},
            "capture_manifest": {"path": str(capture / "capture-manifest-v2.json"), **file_meta(capture / "capture-manifest-v2.json")},
            "capture_stage": {"path": str(capture / "capture-stage-v2.json"), **file_meta(capture / "capture-stage-v2.json")},
            "affinity_controls": {"path": str(capture / "affinity-controls-v2.json"), **file_meta(capture / "affinity-controls-v2.json")},
            "root_go_bound_references": bound,
        },
        "custody": {
            "capture_root": str(capture),
            "actual_file_count": len(actual_rel_files),
            "expected_file_count": len(expected_rel_files),
            "packet_count": len(packet_paths),
            "packet_files_present": len(packet_dir_files),
            "packet_hash_list_sha256": digest_file_lines(packet_hash_lines),
            "referenced_raw_sidecar_file_count": len(observed_files),
            "referenced_raw_sidecar_hash_list_sha256": digest_file_lines(referenced_hash_lines),
            "tree_exactly_referenced": actual_rel_files == expected_rel_files,
        },
        "completeness": {
            "manifest_complete": manifest.get("complete"),
            "stage_complete": stage.get("status") == "complete",
            "terminal_exit_code": terminal.get("exit_code"),
            "terminal_retries": terminal.get("retries"),
            "packet_tuple_coverage": len(metrics) == 96,
            "unique_launcher_pids": len(pid_values) == len(set(pid_values)),
            "all_expected_samples": sum(len(row["elapsed"]) for row in metrics.values()) == 7680,
            "all_expected_perf_events": sum(len(row["perf"]) for row in metrics.values()) == 576,
        },
        "affinity": {
            "allowed": allowed,
            "pin_cpu": 31,
            "control_receipt_pass": not any("affinity controls" in err for err in errors),
            "expected_pinned_actual": [31],
            "expected_unpinned_actual": allowed,
            "observed_context_and_launcher_masks_checked": True,
        },
        "correctness": {
            "all_packet_report_and_sample_oracle_gates_checked": True,
            "all_report_correctness_gates_true": all_correctness_pass,
            "oracle_failure_reasons_checked_empty": not any("oracle failure reasons" in err for err in errors),
            "sample_count_per_packet": 80,
            "route_count_per_packet": 3,
            "route_stage_count": 32,
        },
        "perf_and_time": {
            "events": EXPECTED_EVENTS,
            "perf_sidecars_checked": 96,
            "valid_event_rows": sum(len(row["perf"]) for row in metrics.values()),
            "gnu_time_sidecars_checked": 96,
            "scope": plan["wrapper"]["scope"],
            "header_comments_only_started_on": True,
        },
        "independent_scalar_output": {"path": str(args.scalars), **file_meta(args.scalars)},
        "findings": [
            {
                "id": "shared-reducer-keyerror-pair",
                "status": "tooling_pending",
                "severity": "capture_unaffected",
                "description": "The first shared reduce_0970_v3 replay stopped before output with KeyError 'pair'; this audit did not import or execute that reducer and did not mutate or rerun the frozen capture. Independent scalars above are available pending corrected reducer replay.",
            }
        ],
        "errors": errors,
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(audit, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps({"status": audit["status"], "errors": len(errors), "audit": str(args.output), "scalars": str(args.scalars)}, sort_keys=True))
    return 0 if not errors else 2


if __name__ == "__main__":
    raise SystemExit(main())
