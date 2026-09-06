#!/usr/bin/env python3
"""Derive and replay the matched change-0431 provider comparison.

This file is intentionally a small consumer of the retained reports.  The
report verifier owns the lifecycle contract; this driver owns only the
matched-input identity, descriptive API distributions, and the resource
endpoint review.  It never treats a changed compressed output wrapper as a
failed comparison.  Source/destination archives, corpus manifest, and the
producer's semantic gates are the cross-role identity instead.
"""

from __future__ import annotations

import argparse
import gzip
import hashlib
import json
import math
from pathlib import Path
import random
import statistics
import sys
from typing import Any


ROOT = Path(__file__).resolve().parent
CHANGE = 431
SAMPLES = 30
WARMUPS = 3
REPEATS = ("R1", "R2")
REGRESSION_PERCENT = 5
BOOTSTRAPS = 2000


class Invalid(ValueError):
    """The retained comparison is not a valid 0431 comparison."""


def fail(message: str) -> None:
    raise Invalid(message)


def duplicate_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            fail(f"duplicate JSON key: {key}")
        result[key] = value
    return result


def reject_constant(value: str) -> Any:
    fail(f"non-finite JSON constant: {value}")


def load_bytes(raw: bytes, label: str) -> Any:
    try:
        return json.loads(
            raw.decode("utf-8"),
            object_pairs_hook=duplicate_pairs,
            parse_constant=reject_constant,
        )
    except Invalid:
        raise
    except (UnicodeError, json.JSONDecodeError) as exc:
        fail(f"{label}: invalid JSON: {exc}")


def safe_path(root: Path, name: str) -> Path:
    candidate = Path(name)
    if candidate.is_absolute():
        fail(f"absolute artifact path: {name}")
    resolved = (root / candidate).resolve()
    if not resolved.is_relative_to(root.resolve()):
        fail(f"artifact path escapes bundle: {name}")
    return resolved


def logical_bytes(root: Path, name: str) -> bytes:
    """Read a raw artifact, accepting only the sealer's lossless .gz form."""
    path = safe_path(root, name)
    if path.is_file():
        return path.read_bytes()
    compressed = Path(str(path) + ".gz")
    if compressed.is_file():
        try:
            return gzip.decompress(compressed.read_bytes())
        except (OSError, EOFError) as exc:
            fail(f"{name}: invalid gzip: {exc}")
    fail(f"missing artifact: {name}")


def load(root: Path, name: str) -> Any:
    return load_bytes(logical_bytes(root, name), name)


def digest(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def text(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        fail(f"{label}: expected non-empty text")
    return value


def uint(value: Any, label: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        fail(f"{label}: expected unsigned integer")
    return value


def number(value: Any, label: str) -> int | float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        fail(f"{label}: expected number")
    if isinstance(value, float) and not math.isfinite(value):
        fail(f"{label}: non-finite number")
    return value


def obj(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label}: expected object")
    return value


def array(value: Any, label: str) -> list[Any]:
    if not isinstance(value, list):
        fail(f"{label}: expected array")
    return value


def sha256_text(value: Any, label: str) -> str:
    value = text(value, label)
    if len(value) != 64 or any(char not in "0123456789abcdef" for char in value):
        fail(f"{label}: expected lowercase SHA-256")
    return value


def flatten(value: dict[str, Any], prefix: str = "") -> dict[str, int]:
    """Flatten retained integer endpoints without fabricating unavailable data."""
    result: dict[str, int] = {}
    for key, item in value.items():
        name = prefix + key
        if isinstance(item, dict):
            result.update(flatten(item, name + "."))
        elif type(item) is int:
            result[name] = item
        elif isinstance(item, list) and all(type(entry) is int for entry in item):
            result.update({name + "." + str(index): entry for index, entry in enumerate(item)})
    return result


def percentile(values: list[float | int], fraction: float) -> float:
    if not values:
        fail("cannot summarize an empty distribution")
    ordered = sorted(values)
    position = (len(ordered) - 1) * fraction
    left = int(position)
    right = min(left + 1, len(ordered) - 1)
    return ordered[left] + (ordered[right] - ordered[left]) * (position - left)


def distribution(values: list[int | float], label: str, bootstrap: bool) -> dict[str, Any]:
    if not values:
        fail(f"{label}: empty distribution")
    for index, value in enumerate(values):
        number(value, f"{label}[{index}]")
    result: dict[str, Any] = {
        "samples": len(values),
        "min": min(values),
        "p50": percentile(values, 0.50),
        "p95": percentile(values, 0.95),
        "p99": percentile(values, 0.99),
        "max": max(values),
        "mean": statistics.fmean(values),
    }
    if bootstrap:
        rng = random.Random(int.from_bytes(hashlib.sha256(label.encode()).digest()[:8], "big"))
        medians = [
            statistics.median(values[rng.randrange(len(values))] for _ in values)
            for _ in range(BOOTSTRAPS)
        ]
        result["median_bootstrap_95_percentile_interval"] = [
            percentile(medians, 0.025), percentile(medians, 0.975)
        ]
    return result


def expected_lanes(protocol: dict[str, Any]) -> list[dict[str, Any]]:
    lanes = array(protocol.get("lanes"), "protocol.lanes")
    order = array(protocol.get("order"), "protocol.order")
    if len(lanes) != 8 or len(order) != 16:
        fail("protocol: expected eight lane definitions and sixteen ordered captures")
    if order != [dict(lane, repeat="R1") for lane in lanes] + [
        dict(lane, repeat="R2") for lane in reversed(lanes)
    ]:
        fail("protocol.order: expected forward R1 and reverse R2 order")
    return order


def artifact_custody(root: Path, receipt: dict[str, Any], path: str) -> bytes:
    rows = obj(receipt.get("artifacts"), f"{path}.artifacts")
    raw_paths = set(rows)
    expected = {
        f"{receipt['role']}/{receipt['name']}.json",
        f"{receipt['role']}/{receipt['name']}.log",
        f"{receipt['role']}/{receipt['name']}-resource.log",
    }
    if raw_paths != expected:
        fail(f"{path}.artifacts: paths differ from the three capture artifacts")
    report_raw: bytes | None = None
    for artifact_name in sorted(expected):
        custody = obj(rows[artifact_name], f"{path}.artifacts.{artifact_name}")
        if set(custody) != {"sha256", "bytes"}:
            fail(f"{path}.artifacts.{artifact_name}: unexpected custody fields")
        raw = logical_bytes(root, artifact_name)
        if sha256_text(custody.get("sha256"), f"{path}.artifacts.{artifact_name}.sha256") != digest(raw):
            fail(f"{path}.artifacts.{artifact_name}: raw SHA-256 mismatch")
        if uint(custody.get("bytes"), f"{path}.artifacts.{artifact_name}.bytes") != len(raw):
            fail(f"{path}.artifacts.{artifact_name}: raw byte count mismatch")
        if artifact_name.endswith(".json"):
            report_raw = raw
    if report_raw is None:
        fail(f"{path}: report artifact is missing")
    return report_raw


def resource_log_value(root: Path, name: str) -> int:
    raw = logical_bytes(root, name).decode("utf-8", "replace")
    prefix = "Maximum resident set size (kbytes):"
    matches = []
    for line in raw.splitlines():
        normalized = line.lstrip()
        if normalized.startswith(prefix):
            matches.append(normalized[len(prefix):].strip())
    if len(matches) != 1 or not matches[0].isdigit():
        fail(f"{name}: expected one valid GNU time RSS endpoint")
    return int(matches[0]) * 1024


def verify_receipt(
    root: Path,
    receipt_path: str,
    expected_lane: dict[str, Any],
    role: str,
    build: dict[str, Any],
    protocol: dict[str, Any],
) -> tuple[dict[str, Any], dict[str, Any], bytes]:
    receipt = obj(load(root, receipt_path), receipt_path)
    expected_keys = {
        "role", "name", "lane", "argv", "source_manifest", "revision",
        "binary_sha256", "protocol_sha256", "driver_sha256", "verifier_sha256",
        "started_utc", "status", "exit_code", "finished_utc", "artifacts",
    }
    if set(receipt) != expected_keys:
        fail(f"{receipt_path}: receipt fields differ from capture schema")
    name = f"{expected_lane['selector']}-{expected_lane['provider_label']}-{expected_lane['repeat'].lower()}"
    if receipt["role"] != role or receipt["name"] != name or receipt["lane"] != expected_lane:
        fail(f"{receipt_path}: lane identity differs from protocol")
    if receipt["source_manifest"] != build["source_manifest"]:
        fail(f"{receipt_path}: source manifest differs from role build")
    if receipt["revision"] != build["revision"]:
        fail(f"{receipt_path}: source revision differs from role build")
    if receipt["binary_sha256"] != build["binary_sha256"]:
        fail(f"{receipt_path}: binary identity differs from role build")
    if receipt["protocol_sha256"] != build["protocol_sha256"]:
        fail(f"{receipt_path}: protocol identity differs from role build")
    if receipt["driver_sha256"] != digest((root / "capture.py").read_bytes()):
        fail(f"{receipt_path}: capture driver identity differs")
    if receipt["verifier_sha256"] != digest((root / "verify-report.py").read_bytes()):
        fail(f"{receipt_path}: bound report verifier identity differs")
    if receipt["status"] != "pass" or receipt["exit_code"] != 0:
        fail(f"{receipt_path}: capture did not pass")

    argv = array(receipt["argv"], f"{receipt_path}.argv")
    if len(argv) < 21 or argv[:6] != ["taskset", "-c", str(protocol["cpu"]), "/usr/bin/time", "-v", "-o"]:
        fail(f"{receipt_path}.argv: taskset/time prefix differs")
    if not all(isinstance(item, str) for item in argv):
        fail(f"{receipt_path}.argv: argument is not text")
    resource_path = Path(argv[6])
    if resource_path.name != f"{name}-resource.log":
        fail(f"{receipt_path}.argv: resource output name differs")
    if len(argv) < 9 or argv[7] != build["capture_binary"]:
        fail(f"{receipt_path}.argv: binary differs from role build")
    expected = [
        expected_lane["command"], "--" + expected_lane["selector_key"], expected_lane["selector"],
        "--provider", expected_lane["provider"], "--samples", str(SAMPLES), "--warmup", str(WARMUPS),
        "--source-revision", build["revision"], "--output", argv[20],
    ]
    # The output path is an absolute capture path in the driver; only its
    # basename and role directory are portable invariants.
    if argv[8:21] != expected:
        fail(f"{receipt_path}.argv: provider arguments differ")
    if Path(argv[20]).name != f"{name}.json":
        fail(f"{receipt_path}.argv: report output name differs")
    if expected_lane["provider"] == "range":
        expected_tail = ["--max-range", str(expected_lane["max_range"]), "--delay-us", str(expected_lane["delay_us"])]
        if argv[21:] != expected_tail:
            fail(f"{receipt_path}.argv: range arguments differ")
    elif argv[21:]:
        fail(f"{receipt_path}.argv: unexpected direct-provider arguments")

    report_raw = artifact_custody(root, receipt, receipt_path)
    report = obj(load_bytes(report_raw, f"{receipt_path}.report"), f"{receipt_path}.report")
    return receipt, report, report_raw


def report_lane_identity(report: dict[str, Any], lane: dict[str, Any], build: dict[str, Any]) -> None:
    if report.get("schema") != "pptx_provider_lifecycle_v1":
        fail("report.schema: provider comparison requires the provider lifecycle schema")
    if report.get("corpus") != lane["selector"] or report.get("provider") != lane["provider"]:
        fail("report: corpus/provider differs from protocol lane")
    config = obj(report.get("provider_config"), "report.provider_config")
    if config.get("provider") != lane["provider"]:
        fail("report.provider_config.provider: provider is mislabeled")
    if config.get("max_range_bytes") != lane["max_range"] or config.get("delay_us") != lane["delay_us"]:
        fail("report.provider_config: range configuration differs from protocol lane")
    if report.get("samples") != SAMPLES or report.get("warmup") != WARMUPS:
        fail("report: sample or warmup count differs from protocol")
    if report.get("source_revision") != build["revision"]:
        fail("report.source_revision: differs from role build")
    if report.get("binary_sha256") != build["binary_sha256"] or report.get("binary_bytes") != build["binary_bytes"]:
        fail("report binary identity differs from role build")
    if report.get("current_exe") != build["capture_binary"]:
        fail("report.current_exe: differs from role build")
    gates = obj(report.get("gates"), "report.gates")
    if not gates or any(value is not True for value in gates.values()):
        fail("report.gates: semantic producer gate is false")


def input_identity(report: dict[str, Any]) -> dict[str, Any]:
    fields = (
        "source_archive_sha256", "source_archive_bytes", "destination_archive_sha256",
        "destination_archive_bytes", "corpus_manifest", "gates",
    )
    result = {field: report.get(field) for field in fields}
    sha256_text(result["source_archive_sha256"], "report.source_archive_sha256")
    sha256_text(result["destination_archive_sha256"], "report.destination_archive_sha256")
    uint(result["source_archive_bytes"], "report.source_archive_bytes")
    uint(result["destination_archive_bytes"], "report.destination_archive_bytes")
    return result


def output_identity(report: dict[str, Any]) -> dict[str, Any]:
    return {
        "expected_output_sha256": sha256_text(report.get("expected_output_sha256"), "report.expected_output_sha256"),
        "expected_output_bytes": uint(report.get("expected_output_bytes"), "report.expected_output_bytes"),
    }


def report_metrics(report: dict[str, Any], name: str, resource_rss: int) -> dict[str, list[int | float]]:
    rows = array(report.get("samples_raw"), "report.samples_raw")
    if len(rows) != SAMPLES:
        fail(f"{name}: report sample row count differs from samples")
    metrics: dict[str, list[int | float]] = {}
    result_bytes = uint(report.get("expected_output_bytes"), f"{name}.expected_output_bytes")
    for row_index, raw_row in enumerate(rows):
        row = obj(raw_row, f"{name}.samples_raw[{row_index}]")
        if row.get("sample_index") != row_index:
            fail(f"{name}.samples_raw[{row_index}].sample_index: non-contiguous")
        timings = obj(row.get("timings"), f"{name}.samples_raw[{row_index}].timings")
        required = ("open_source_ns", "open_destination_ns", "open_ns", "plan_ns", "publication_ns", "api_sum_ns")
        if any(field not in timings for field in required):
            fail(f"{name}.samples_raw[{row_index}].timings: required API endpoint missing")
        for field in required:
            value = uint(timings[field], f"{name}.samples_raw[{row_index}].timings.{field}")
            metrics.setdefault("timings." + field, []).append(value)
        if timings["open_ns"] != timings["open_source_ns"] + timings["open_destination_ns"]:
            fail(f"{name}.samples_raw[{row_index}].timings.open_ns: open sum mismatch")
        if timings["api_sum_ns"] != timings["open_ns"] + timings["plan_ns"] + timings["publication_ns"]:
            fail(f"{name}.samples_raw[{row_index}].timings.api_sum_ns: API sum mismatch")
        if timings["api_sum_ns"] <= 0:
            fail(f"{name}.samples_raw[{row_index}].timings.api_sum_ns: zero API duration")
        rate = 1_000_000_000 / timings["api_sum_ns"]
        metrics.setdefault("rates.completed_operations_per_api_second", []).append(rate)
        metrics.setdefault("rates.logical_result_bytes_per_api_second", []).append(rate * result_bytes)
        phases = array(row.get("phases"), f"{name}.samples_raw[{row_index}].phases")
        for phase in phases:
            phase_obj = obj(phase, f"{name}.samples_raw[{row_index}].phase")
            label = text(phase_obj.get("label"), f"{name}.phase.label")
            for key, value in flatten(phase_obj).items():
                metrics.setdefault("phases." + label + "." + key, []).append(value)
    if resource_rss is None:
        fail(f"{name}: GNU time maximum resident set size endpoint is missing")
    metrics.setdefault("resource.time.maximum_resident_set_size_bytes", []).append(resource_rss)
    return metrics


def summarize_capture(
    report: dict[str, Any], lane: dict[str, Any], role: str, name: str, resource_rss: int
) -> dict[str, Any]:
    metrics = report_metrics(report, name, resource_rss)
    distributions = {
        key: distribution(values, role + "/" + name + "/" + key, key.startswith(("timings.", "rates.")))
        for key, values in sorted(metrics.items())
    }
    resource = {
        key: value
        for key, value in distributions.items()
        if key.startswith("phases.") or key.startswith("resource.")
    }
    return {
        "role": role,
        "name": name,
        "command": lane["command"],
        "selector": lane["selector"],
        "provider_label": lane["provider_label"],
        "provider": lane["provider"],
        "repeat": lane["repeat"],
        "input_identity": input_identity(report),
        "output_identity": output_identity(report),
        "metrics": distributions,
        "resource_endpoints": resource,
    }


def change_percent(before: float | int, after: float | int) -> float | None:
    if before == 0:
        return None if after == 0 else None
    return 100.0 * (after - before) / before


def pair_delta(
    before: dict[str, Any], after: dict[str, Any], metric: str, *, kind: str
) -> dict[str, Any]:
    old = before["metrics"][metric]
    new = after["metrics"][metric]
    old_p50, new_p50 = old["p50"], new["p50"]
    percent = change_percent(old_p50, new_p50)
    adverse = (new_p50 > old_p50) if kind in {"duration", "resource"} else (new_p50 < old_p50)
    trigger = (percent is None and new_p50 != old_p50) or (
        percent is not None and adverse and abs(percent) > REGRESSION_PERCENT
    )
    return {
        "metric": metric,
        "kind": kind,
        "before": old,
        "after": new,
        "after_minus_before_p50": new_p50 - old_p50,
        "after_minus_before_percent": percent,
        "regression_over_5_percent": bool(trigger),
    }


def derive(root: Path = ROOT) -> dict[str, Any]:
    protocol = obj(load(root, "protocol.json"), "protocol")
    order = expected_lanes(protocol)
    if protocol.get("change") != CHANGE or protocol.get("samples_per_process") != SAMPLES or protocol.get("warmups_per_process") != WARMUPS:
        fail("protocol: change/sample values are not frozen 0431 values")
    captures: list[dict[str, Any]] = []
    by_role: dict[str, dict[str, dict[str, Any]]] = {"before": {}, "after": {}}
    input_by_lane: dict[tuple[str, str], dict[str, Any]] = {}
    output_by_role_lane: dict[tuple[str, str, str], dict[str, Any]] = {}

    for role in ("before", "after"):
        build = obj(load(root, f"build-{role}.json"), f"build-{role}")
        index = array(load(root, f"{role}-index.json"), f"{role}-index")
        if len(index) != len(order) or len(set(index)) != len(index):
            fail(f"{role}-index: expected sixteen unique receipt paths")
        for receipt_name, lane in zip(index, order, strict=True):
            receipt_path = text(receipt_name, f"{role}-index receipt path")
            receipt, report, _ = verify_receipt(root, receipt_path, lane, role, build, protocol)
            report_lane_identity(report, lane, build)
            name = receipt["name"]
            rss = resource_log_value(root, f"{role}/{name}-resource.log")
            capture = summarize_capture(report, lane, role, name, rss)
            captures.append(capture)
            by_role[role][name] = capture
            lane_key = (lane["selector"], lane["provider_label"])
            identity = input_identity(report)
            if lane_key in input_by_lane and input_by_lane[lane_key] != identity:
                fail(f"{role}/{name}: input archive or manifest identity differs within matched lane")
            input_by_lane[lane_key] = identity
            output_key = (role, lane["selector"], lane["provider_label"])
            output = output_identity(report)
            if output_key in output_by_role_lane and output_by_role_lane[output_key] != output:
                fail(f"{role}/{name}: same-role output wrapper identity differs between repeats")
            output_by_role_lane[output_key] = output

    comparisons: list[dict[str, Any]] = []
    repeat_flags: list[dict[str, Any]] = []
    regression_flags: list[dict[str, Any]] = []
    resource_comparisons: list[dict[str, Any]] = []
    output_comparisons: list[dict[str, Any]] = []
    for lane in order[: len(order) // 2]:
        key = f"{lane['selector']}-{lane['provider_label']}"
        names = {role: {repeat: f"{key}-{repeat.lower()}" for repeat in REPEATS} for role in ("before", "after")}
        # R1/R2 drift is retained for every timing, rate, RSS, and managed
        # endpoint.  The result is descriptive review evidence, not an
        # automatic release decision.
        for role in ("before", "after"):
            first, second = by_role[role][names[role]["R1"]], by_role[role][names[role]["R2"]]
            if set(first["metrics"]) != set(second["metrics"]):
                fail(f"{role}/{key}: R1/R2 metric sets differ")
            for metric in sorted(first["metrics"]):
                delta = pair_delta(first, second, metric, kind="resource" if metric.startswith(("phases.", "resource.")) else ("rate" if metric.startswith("rates.") else "duration"))
                if delta["regression_over_5_percent"]:
                    repeat_flags.append({"role": role, "lane": key, **delta})
        for repeat in REPEATS:
            before = by_role["before"][names["before"][repeat]]
            after = by_role["after"][names["after"][repeat]]
            if set(before["metrics"]) != set(after["metrics"]):
                fail(f"{key}/{repeat}: before/after metric sets differ")
            if before["input_identity"] != after["input_identity"]:
                fail(f"{key}/{repeat}: source/destination archive or corpus manifest differs across roles")
            for metric in sorted(before["metrics"]):
                kind = "resource" if metric.startswith(("phases.", "resource.")) else ("rate" if metric.startswith("rates.") else "duration")
                delta = pair_delta(before, after, metric, kind=kind)
                comparisons.append({"lane": key, "repeat": repeat, **delta})
                if delta["regression_over_5_percent"]:
                    regression_flags.append({"lane": key, "repeat": repeat, **delta})
                if kind == "resource":
                    resource_comparisons.append({
                        "lane": key,
                        "repeat": repeat,
                        "endpoint": metric,
                        "scope": "setup-inclusive endpoint observation; baseline is retained explicitly and no causal RSS/resource claim is made",
                        "baseline": delta["before"],
                        "after": delta["after"],
                        "after_minus_baseline_p50": delta["after_minus_before_p50"],
                        "after_minus_baseline_percent": delta["after_minus_before_percent"],
                        "regression_over_5_percent": delta["regression_over_5_percent"],
                    })
            # Output framing may differ between revisions.  Keep both
            # identities in the comparison and require the lifecycle
            # gates/input identity; output equality belongs to the independent
            # semantic acceptance.
            output_comparisons.append({
                "lane": key,
                "repeat": repeat,
                "before": before["output_identity"],
                "after": after["output_identity"],
                "expected_equal": False,
                "equal": before["output_identity"] == after["output_identity"],
                "semantic_acceptance": "unchanged source-bound producer semantic and preservation gates retained in each report",
            })

    result = {
        "status": "pass",
        "change": CHANGE,
        "processes_per_role": len(order),
        "retained_samples_per_role": len(order) * SAMPLES,
        "method": "p50/p95/p99 use linear interpolation of each 30-sample process; means and deterministic 2000-resample SHA256-seeded bootstrap median 95% intervals are descriptive.",
        "timing_scope": "Only open, plan, publication, and their API sum are compared. Setup, source construction, reservations, checks, observers, drops, and output verification remain outside the clocks.",
        "resource_scope": "RSS and managed-resource endpoints are setup-inclusive lifecycle observations. Baseline and after values remain explicit; no causal memory or optimization claim is authorized.",
        "output_scope": "Expected output wrapper hashes and bytes may differ between revisions. Source/destination archive identities, corpus manifest, and unchanged source-bound producer semantic and preservation gates retained in each report are the matched evidence; no separate output artifact is required.",
        "captures": sorted(captures, key=lambda row: (row["role"], row["name"])),
        "comparisons": comparisons,
        "output_comparisons": output_comparisons,
        "resource_endpoint_comparisons": resource_comparisons,
        "repeat_review_triggers": repeat_flags,
        "regression_review_triggers": regression_flags,
        "regression_threshold_percent": REGRESSION_PERCENT,
        "performance_claim": None,
    }
    return result


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true", help="recompute and compare comparison.json")
    parser.add_argument("--output", type=Path, default=ROOT / "comparison.json")
    args = parser.parse_args(argv)
    try:
        result = derive(ROOT)
        if args.check:
            if not args.output.is_file():
                fail(f"missing comparison output: {args.output}")
            retained = load_bytes(args.output.read_bytes(), str(args.output))
            if retained != result:
                fail("comparison.json differs from recomputed summaries")
        else:
            if args.output.exists():
                fail(f"refusing to overwrite existing comparison output: {args.output}")
            args.output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        print(json.dumps({
            "status": "pass", "processes_per_role": result["processes_per_role"],
            "comparisons": len(result["comparisons"]),
            "repeat_review_triggers": len(result["repeat_review_triggers"]),
            "regression_review_triggers": len(result["regression_review_triggers"]),
        }, sort_keys=True))
        return 0
    except Exception as exc:
        print(f"INVALID: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
