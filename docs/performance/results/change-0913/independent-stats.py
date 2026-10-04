#!/usr/bin/env python3
"""Recompute 0913 timing totals independently from retained raw receipts.

This is a small custody and arithmetic cross-check for ``analyze.py``.  It
does not import the analyzer, run a workload, invoke Cargo, or invoke any
external command.  It reads the frozen plan, the 48 child receipts, their
reports, and ``/usr/bin/time`` metrics, then writes a create-only JSON result.
The output intentionally contains raw route/variant summaries and paired
baseline-over-candidate bootstrap intervals so a reviewer can compare the
numbers without trusting the primary analyzer's implementation.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import random
import statistics
import sys
from pathlib import Path
from typing import Any


HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
ROUTES = ("materialized", "owned")
VARIANTS = ("baseline", "candidate")
PHASES = ("clone_ns", "ingress_ns", "snapshot_ns", "edit_commit_ns", "publish_ns")
TIMINGS = PHASES + ("total_ns",)
WARMUP = 30
SAMPLES = 20
PAIRS = 12
CHILDREN = 48
BOOTSTRAP_DRAWS = 10_000
BOOTSTRAP_SEED = 913
CPU = 12
SOURCE_BYTES = 17_017_139
SOURCE_SHA256 = "61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757"
OUTPUT_BYTES = 17_017_145
OUTPUT_SHA256 = "23d6d7b8dd433ff453307ee56485efadc34d6b18b455a6858d9f565c4b1b6cd9"
SINK_MAX_WRITE = 65_536


class EvidenceError(ValueError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def read_json(path: Path, label: str) -> Any:
    require(path.is_file() and not path.is_symlink(), f"{label}: missing regular file")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"{label}: invalid JSON: {error}") from error


def digest(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"cannot hash {path}")
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    path = path.resolve()
    require(path.is_file() and not path.is_symlink(), f"not a regular file: {path}")
    try:
        name = str(path.relative_to(ROOT))
    except ValueError as error:
        raise EvidenceError(f"path escapes repository root: {path}") from error
    return {"path": name, "bytes": path.stat().st_size, "sha256": digest(path)}


def check_descriptor(raw: Any, label: str) -> Path:
    require(isinstance(raw, dict), f"{label}: descriptor malformed")
    raw_path = raw.get("path")
    require(isinstance(raw_path, str) and raw_path, f"{label}: descriptor path missing")
    path = Path(raw_path)
    if not path.is_absolute():
        path = (ROOT / path).resolve()
    else:
        path = path.resolve()
    actual = descriptor(path)
    require(actual["bytes"] == raw.get("bytes") and actual["sha256"] == raw.get("sha256"), f"{label}: descriptor changed")
    return path


def integer(value: Any, label: str, *, positive: bool = False) -> int:
    require(type(value) is int, f"{label}: expected integer")
    require(value > 0 if positive else value >= 0, f"{label}: invalid integer")
    return value


def summary(values: list[int], label: str) -> dict[str, Any]:
    require(values, f"{label}: no values")
    ordered = sorted(integer(value, f"{label}[{index}]") for index, value in enumerate(values))
    count = len(ordered)
    return {
        "samples": count,
        "min": ordered[0],
        "p50": (ordered[(count - 1) // 2] + ordered[count // 2]) // 2,
        "p95": ordered[min((95 * count + 99) // 100 - 1, count - 1)],
        "p99": ordered[min((99 * count + 99) // 100 - 1, count - 1)],
        "max": ordered[-1],
        "mean": statistics.fmean(ordered),
    }


def percentile(values: list[float], fraction: float) -> float:
    ordered = sorted(values)
    require(ordered, "bootstrap: no draws")
    return ordered[min(max(math.ceil(fraction * len(ordered)) - 1, 0), len(ordered) - 1)]


def bootstrap(values: list[float]) -> dict[str, Any]:
    require(len(values) == PAIRS, "bootstrap: wrong paired ratio count")
    rng = random.Random(BOOTSTRAP_SEED)
    draws: list[float] = []
    for _ in range(BOOTSTRAP_DRAWS):
        draw = [values[rng.randrange(PAIRS)] for _ in range(PAIRS)]
        draws.append(float(statistics.median(draw)))
    return {
        "statistic": float(statistics.median(values)),
        "draws": BOOTSTRAP_DRAWS,
        "seed": BOOTSTRAP_SEED,
        "confidence": 0.95,
        "lower": percentile(draws, 0.025),
        "upper": percentile(draws, 0.975),
    }


def process_metrics(path: Path, label: str) -> dict[str, int | float]:
    result: dict[str, int | float] = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        require("=" in line, f"{label}: malformed process metric")
        key, value = line.split("=", 1)
        require(key not in result and key, f"{label}: duplicate/empty metric")
        if key in ("max_rss_kib", "major_faults", "minor_faults", "voluntary_context_switches", "involuntary_context_switches"):
            try:
                number = int(value)
            except ValueError as error:
                raise EvidenceError(f"{label}.{key}: integer expected") from error
            result[key] = integer(number, f"{label}.{key}", positive=key == "max_rss_kib")
        elif key in ("user_seconds", "system_seconds", "elapsed_seconds"):
            try:
                number = float(value)
            except ValueError as error:
                raise EvidenceError(f"{label}.{key}: float expected") from error
            require(math.isfinite(number) and number >= 0, f"{label}.{key}: invalid number")
            result[key] = number
        else:
            raise EvidenceError(f"{label}: unexpected metric {key}")
    expected = ("max_rss_kib", "user_seconds", "system_seconds", "elapsed_seconds", "major_faults", "minor_faults", "voluntary_context_switches", "involuntary_context_switches")
    require(set(result) == set(expected), f"{label}: metric keys differ")
    return result


def load_raw_legs() -> tuple[dict[str, Any], list[dict[str, Any]]]:
    plan_path = HERE / "capture-plan.json"
    plan = read_json(plan_path, "capture-plan")
    require(plan.get("status") == "frozen_before_capture", "capture-plan: not frozen")
    require(plan.get("warmup") == WARMUP and plan.get("samples") == SAMPLES and plan.get("cpu") == CPU, "capture-plan: schedule differs")
    require(plan.get("pairs_per_route") == PAIRS and plan.get("blocks") == 24, "capture-plan: pair/block count differs")
    require(plan.get("routes") == list(ROUTES) and plan.get("variants") == list(VARIANTS), "capture-plan: route/variant set differs")
    legs = plan.get("legs")
    require(isinstance(legs, list) and len(legs) == CHILDREN, "capture-plan: child count differs")
    result: list[dict[str, Any]] = []
    labels: set[str] = set()
    for raw in legs:
        require(isinstance(raw, dict), "capture-plan: malformed leg")
        label = raw.get("label")
        require(isinstance(label, str) and label not in labels, "capture-plan: duplicate label")
        require(raw.get("route") in ROUTES and raw.get("variant") in VARIANTS, "capture-plan: invalid route/variant")
        require(type(raw.get("pair")) is int and 0 <= raw["pair"] < PAIRS, "capture-plan: invalid pair")
        labels.add(label)
        result.append(raw)
    require(len({(leg["route"], leg["pair"], leg["variant"]) for leg in result}) == CHILDREN, "capture-plan: duplicate route/pair/variant")
    return plan, result


def load_leg(leg: dict[str, Any]) -> dict[str, Any]:
    label = leg["label"]
    folder = HERE / "capture" / label
    command = read_json(folder / "command.json", f"{label}.command")
    require(command.get("exit_code") == 0, f"{label}: child failed")
    require(command.get("leg") == leg and command.get("variant") == leg["variant"], f"{label}: receipt identity differs")
    tools = HERE / "process-tools.json"
    require(command.get("process_tools") is not None, f"{label}: process-tools binding missing")
    check_descriptor(command["process_tools"], f"{label}.process_tools")
    host_environment = HERE / "environment-after-recovery.json"
    require(command.get("host_environment") is not None, f"{label}: host environment binding missing")
    check_descriptor(command["host_environment"], f"{label}.host_environment")
    artifact_dir = folder / "artifacts"
    report = read_json(artifact_dir / "report.json", f"{label}.report")
    require(report.get("route") == leg["route"] and report.get("status") == "completed", f"{label}: report status/route differs")
    require(report.get("warmup_iterations") == WARMUP and report.get("measured_samples") == SAMPLES, f"{label}: report sample schedule differs")
    require(report.get("source") == {"file": "source.pptx", "bytes": SOURCE_BYTES, "sha256": SOURCE_SHA256}, f"{label}: source guard differs")
    require(report.get("output") == {"file": "output.pptx", "bytes": OUTPUT_BYTES, "sha256": OUTPUT_SHA256}, f"{label}: output guard differs")
    require(report.get("sink_max_write_bytes") == SINK_MAX_WRITE and report.get("sink_capacity_bytes") == 2 * SOURCE_BYTES + SINK_MAX_WRITE, f"{label}: sink guard differs")
    require((artifact_dir / "source.pptx").stat().st_size == SOURCE_BYTES and digest(artifact_dir / "source.pptx") == SOURCE_SHA256, f"{label}: source artifact differs")
    require((artifact_dir / "output.pptx").stat().st_size == OUTPUT_BYTES and digest(artifact_dir / "output.pptx") == OUTPUT_SHA256, f"{label}: output artifact differs")
    samples = report.get("samples")
    require(isinstance(samples, list) and len(samples) == SAMPLES, f"{label}: raw sample count differs")
    normalized: list[dict[str, int]] = []
    for index, sample in enumerate(samples):
        require(isinstance(sample, dict) and sample.get("index") == index, f"{label}: sample order differs")
        row: dict[str, int] = {}
        for field in TIMINGS:
            row[field] = integer(sample.get(field), f"{label}.samples[{index}].{field}", positive=True)
        require(row["total_ns"] >= sum(row[field] for field in PHASES), f"{label}: total below phase sum")
        require(sample.get("output_bytes") == OUTPUT_BYTES and sample.get("output_sha256") == OUTPUT_SHA256, f"{label}: sample output differs")
        require(sample.get("sink_write_calls") == (OUTPUT_BYTES + SINK_MAX_WRITE - 1) // SINK_MAX_WRITE and sample.get("sink_largest_write") == SINK_MAX_WRITE, f"{label}: sample sink differs")
        normalized.append(row)
    rss = process_metrics(folder / "process.txt", f"{label}.process")
    return {"label": label, "route": leg["route"], "variant": leg["variant"], "pair": leg["pair"], "samples": normalized, "rss": rss}


def pair_ratios(legs: list[dict[str, Any]], route: str) -> dict[str, Any]:
    selected = {leg["pair"]: {} for leg in legs if leg["route"] == route}
    for leg in legs:
        if leg["route"] == route:
            selected[leg["pair"]][leg["variant"]] = leg
    rows: list[dict[str, Any]] = []
    ratios = {field: [] for field in TIMINGS}
    for pair in range(PAIRS):
        row = selected[pair]
        require(set(row) == set(VARIANTS), f"{route}/pair-{pair}: missing variant")
        baseline = row["baseline"]
        candidate = row["candidate"]
        output: dict[str, Any] = {"pair": pair}
        for field in TIMINGS:
            baseline_median = float(statistics.median([sample[field] for sample in baseline["samples"]]))
            candidate_median = float(statistics.median([sample[field] for sample in candidate["samples"]]))
            value = baseline_median / candidate_median
            require(math.isfinite(value) and value > 0, f"{route}/pair-{pair}.{field}: invalid ratio")
            output[field] = {"baseline_median": baseline_median, "candidate_median": candidate_median, "baseline_over_candidate": value}
            ratios[field].append(value)
        rows.append(output)
    return {"pairs": rows, "median_ratio_bootstrap": {field: bootstrap(values) for field, values in ratios.items()}}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    output = args.output.resolve()
    try:
        require(not output.exists() and not output.is_symlink(), f"refusing to overwrite {output}")
        plan, plan_legs = load_raw_legs()
        raw = [load_leg(leg) for leg in plan_legs]
        routes: dict[str, Any] = {}
        paired: dict[str, Any] = {}
        for route in ROUTES:
            routes[route] = {}
            for variant in VARIANTS:
                selected = [leg for leg in raw if leg["route"] == route and leg["variant"] == variant]
                require(len(selected) == PAIRS, f"{route}/{variant}: process count differs")
                routes[route][variant] = {
                    "processes": len(selected),
                    "samples": len(selected) * SAMPLES,
                    **{field: summary([sample[field] for leg in selected for sample in leg["samples"]], f"{route}/{variant}.{field}") for field in TIMINGS},
                    "max_rss_kib": summary([int(leg["rss"]["max_rss_kib"]) for leg in selected], f"{route}/{variant}.max_rss_kib"),
                }
            paired[route] = pair_ratios(raw, route)
        result = {
            "schema": "litchi.performance.0913.independent-stats.v1",
            "status": "pass",
            "plan": descriptor(HERE / "capture-plan.json"),
            "child_processes": CHILDREN,
            "warmup": WARMUP,
            "samples": SAMPLES,
            "pairs_per_route": PAIRS,
            "bootstrap": {"draws": BOOTSTRAP_DRAWS, "seed": BOOTSTRAP_SEED, "confidence": 0.95},
            "raw": {
                "total_ns_by_child": {leg["label"]: [sample["total_ns"] for sample in leg["samples"]] for leg in raw},
                "max_rss_kib_by_child": {leg["label"]: leg["rss"]["max_rss_kib"] for leg in raw},
            },
            "routes": routes,
            "paired_by_route": paired,
            "checks": {"raw_receipts_read": True, "source_output_sink_guards": True, "phase_totals_recomputed": True, "route_specific_pairing": True, "bootstrap_recomputed_independently": True},
        }
        output.parent.mkdir(parents=True, exist_ok=True)
        with output.open("x", encoding="utf-8") as stream:
            json.dump(result, stream, indent=2, sort_keys=True)
            stream.write("\n")
    except (EvidenceError, OSError) as error:
        print(f"independent-stats.py: error: {error}", file=sys.stderr)
        return 2
    print("independent 0913 raw totals and bootstrap verification passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
