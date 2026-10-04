#!/usr/bin/env python3
"""Fail-closed offline analysis for the 0913 four-variant PPTX capture.

The packet contains twelve baseline/candidate process pairs for each route.
This program only reads frozen receipts and benchmark reports.  It never
builds, launches the benchmark, invokes a profiler, or rewrites an evidence
file.  A result JSON is created only after every receipt, source/output guard,
sample vector, and route-specific pairing check succeeds.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import random
import statistics
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Mapping, Sequence


HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]

SCHEMA = "litchi.performance.0913.pptx-binary-variant-analysis.v1"
REPORT_SCHEMA = "pptx-owned-input-benchmark-v1"
PARENT = "66b944b07fcf4c6fa37f10e7ed83a02417455f57"
BASELINE_PARENT = "468fe2931d6532c32f4c6608590816b8eaa88e31"
BASELINE_SOURCE_MANIFEST = ROOT / "docs/performance/results/change-0910/source-v1.json"
BASELINE_BINARY = ROOT / "docs/performance/results/change-0910/binaries/native"

WARMUP = 30
SAMPLES = 20
PAIRS_PER_ROUTE = 12
BLOCKS = PAIRS_PER_ROUTE * 2
LEGS = PAIRS_PER_ROUTE * len(("materialized", "owned")) * len(("baseline", "candidate"))
BOOTSTRAP_DRAWS = 10_000
BOOTSTRAP_SEED = 913
CPU = 12
ROUTES = ("materialized", "owned")
VARIANTS = ("baseline", "candidate")
PHASES = ("clone_ns", "ingress_ns", "snapshot_ns", "edit_commit_ns", "publish_ns")
ALL_TIMINGS = PHASES + ("total_ns",)
OVERLAYS = (
    "crates/soapberry-zip/src/office.rs",
    "crates/litchi-opc/src/payload.rs",
    "crates/soapberry-zip/tests/owned_direct_read.rs",
    "crates/litchi-opc/tests/owned_direct_concurrency.rs",
)

SOURCE_BYTES = 17_017_139
SOURCE_SHA256 = "61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757"
OUTPUT_BYTES = 17_017_145
OUTPUT_SHA256 = "23d6d7b8dd433ff453307ee56485efadc34d6b18b455a6858d9f565c4b1b6cd9"
ENTRY_COUNT = 229
MEMBER_COUNT = 445
UNCOMPRESSED_BYTES = 17_568_429
POSITIONS = (0, 28, 57, 85, 114, 142, 171, 199)
REPLACEMENT_COUNT = 64
SINK_MAX_WRITE = 65_536


class EvidenceError(ValueError):
    """A retained result is missing, malformed, or contradictory."""


def fail(message: str) -> None:
    raise EvidenceError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read_json(path: Path, label: str) -> Any:
    require(path.is_file() and not path.is_symlink(), f"{label}: missing regular JSON file")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"{label}: invalid JSON: {error}")


def sha256(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"cannot hash missing regular file: {path}")
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1 << 20), b""):
                digest.update(chunk)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def relative(path: Path, base: Path = ROOT) -> str:
    path = path.resolve()
    try:
        return str(path.relative_to(base.resolve()))
    except ValueError as error:
        raise EvidenceError(f"path escapes allowed root: {path}") from error


def descriptor(path: Path, base: Path = ROOT) -> dict[str, Any]:
    path = path.resolve()
    require(path.is_file() and not path.is_symlink(), f"descriptor target is not a regular file: {path}")
    return {"path": relative(path, base), "bytes": path.stat().st_size, "sha256": sha256(path)}


def descriptor_shape(raw: Any, label: str) -> dict[str, Any]:
    require(isinstance(raw, dict), f"{label}: descriptor is malformed")
    path = raw.get("path")
    require(isinstance(path, str) and path, f"{label}: descriptor path is missing")
    return {
        "path": path,
        "bytes": integer(raw.get("bytes"), f"{label}.bytes"),
        "sha256": valid_sha(raw.get("sha256"), f"{label}.sha256"),
    }


def resolve_path(raw_path: str) -> Path:
    candidate = Path(raw_path)
    if candidate.is_absolute():
        return candidate.resolve()
    # run.py records workspace-relative paths for external binaries and
    # packet-relative-looking paths for packet files.  Prefer the existing
    # workspace path when both forms are possible.
    workspace = (ROOT / candidate).resolve()
    packet = (HERE / candidate).resolve()
    if workspace.is_file() or workspace.is_symlink():
        return workspace
    return packet


def resolve_descriptor(raw: Any, label: str, *, allow_missing: bool = False) -> tuple[Path, dict[str, Any]]:
    shaped = descriptor_shape(raw, label)
    path = resolve_path(shaped["path"])
    if allow_missing and not path.exists():
        return path, shaped
    require(path.is_file() and not path.is_symlink(), f"{label}: described file is missing: {path}")
    actual = descriptor(path)
    require(actual["bytes"] == shaped["bytes"], f"{label}: byte count changed")
    require(actual["sha256"] == shaped["sha256"], f"{label}: SHA-256 changed")
    return path, actual


def optional_descriptor(raw: Any, label: str) -> None:
    shaped = descriptor_shape(raw, label)
    path = resolve_path(shaped["path"])
    if not path.exists():
        return
    require(path.is_file() and not path.is_symlink(), f"{label}: retained path is not regular")
    actual = descriptor(path)
    require(actual["bytes"] == shaped["bytes"] and actual["sha256"] == shaped["sha256"], f"{label}: retained bytes changed")


def valid_sha(value: Any, label: str) -> str:
    require(isinstance(value, str) and len(value) == 64, f"{label}: expected SHA-256")
    require(all(char in "0123456789abcdef" for char in value), f"{label}: expected lowercase SHA-256")
    return value


def integer(value: Any, label: str, *, positive: bool = False) -> int:
    require(type(value) is int, f"{label}: expected integer")
    require(value > 0 if positive else value >= 0, f"{label}: expected {'positive' if positive else 'non-negative'} integer")
    return value


def finite_number(value: Any, label: str, *, positive: bool = False) -> float:
    require(isinstance(value, (int, float)) and not isinstance(value, bool), f"{label}: expected number")
    result = float(value)
    require(math.isfinite(result), f"{label}: expected finite number")
    require(result > 0 if positive else result >= 0, f"{label}: expected {'positive' if positive else 'non-negative'} number")
    return result


def parse_timestamp(value: Any, label: str) -> datetime:
    require(isinstance(value, str) and value, f"{label}: missing timestamp")
    try:
        result = datetime.fromisoformat(value)
    except ValueError as error:
        raise EvidenceError(f"{label}: invalid timestamp") from error
    require(result.tzinfo is not None, f"{label}: timestamp has no timezone")
    return result


def canonical(value: Any) -> str:
    return json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=False)


def summary(values: Sequence[int], label: str) -> dict[str, Any]:
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


def median(values: Sequence[float | int], label: str) -> float:
    require(values, f"{label}: no values")
    return float(statistics.median(values))


def ratio(numerator: float, denominator: float, label: str) -> float:
    require(math.isfinite(numerator) and math.isfinite(denominator) and numerator > 0 and denominator > 0, f"{label}: invalid ratio operands")
    value = numerator / denominator
    require(math.isfinite(value) and value > 0, f"{label}: invalid ratio")
    return value


def percentile(values: Sequence[float], percent: float, label: str) -> float:
    require(values, f"{label}: no values")
    ordered = sorted(values)
    index = min(max(math.ceil(percent * len(ordered)) - 1, 0), len(ordered) - 1)
    return ordered[index]


def bootstrap_median(values: Sequence[float], label: str) -> dict[str, Any]:
    require(len(values) == PAIRS_PER_ROUTE, f"{label}: expected {PAIRS_PER_ROUTE} paired ratios")
    generator = random.Random(BOOTSTRAP_SEED)
    draws: list[float] = []
    for _ in range(BOOTSTRAP_DRAWS):
        resample = [values[generator.randrange(PAIRS_PER_ROUTE)] for _ in range(PAIRS_PER_ROUTE)]
        draws.append(float(statistics.median(resample)))
    return {
        "statistic": float(statistics.median(values)),
        "draws": BOOTSTRAP_DRAWS,
        "seed": BOOTSTRAP_SEED,
        "resampling_unit": "paired baseline/candidate process block within one route",
        "confidence": 0.95,
        "method": "nearest-rank percentile interval over bootstrap medians",
        "lower": percentile(draws, 0.025, f"{label}.lower"),
        "upper": percentile(draws, 0.975, f"{label}.upper"),
    }


def expected_corpus(corpus: Any, label: str) -> dict[str, Any]:
    require(isinstance(corpus, dict), f"{label}: corpus manifest missing")
    require(corpus.get("generator") == "litchi-pptx-source-edit-media-v1", f"{label}: generator differs")
    require(corpus.get("package_format") == "PPTX/OPC/ZIP", f"{label}: package format differs")
    require(corpus.get("entry_count") == ENTRY_COUNT, f"{label}: entry count differs")
    require(corpus.get("archive_member_count") == MEMBER_COUNT, f"{label}: member count differs")
    require(corpus.get("uncompressed_payload_bytes") == UNCOMPRESSED_BYTES, f"{label}: logical bytes differ")
    require(corpus.get("shape") == "media-rich", f"{label}: corpus shape differs")
    require(corpus.get("compression") == "deflate", f"{label}: compression differs")
    require(corpus.get("archive_sha256") == SOURCE_SHA256, f"{label}: source digest differs")
    valid_sha(corpus.get("archive_sha256"), f"{label}.archive_sha256")
    return corpus


def check_report_scope(report: Mapping[str, Any], label: str, route: str) -> None:
    require(report.get("schema") == REPORT_SCHEMA, f"{label}: benchmark schema differs")
    require(report.get("status") == "completed", f"{label}: benchmark did not complete")
    require(report.get("route") == route, f"{label}: report route differs")
    require(report.get("warmup_iterations") == WARMUP, f"{label}: warmup count differs")
    require(report.get("measured_samples") == SAMPLES, f"{label}: sample count differs")
    require(report.get("selected_slide_positions") == list(POSITIONS), f"{label}: selected positions differ")
    require(report.get("replacement_count") == REPLACEMENT_COUNT, f"{label}: replacement count differs")
    for field in ("timing_scope", "phase_scope", "teardown_scope", "verification_scope", "performance_claim"):
        require(isinstance(report.get(field), str) and report[field], f"{label}: {field} missing")
    timing = " ".join(str(report[field]).lower() for field in ("timing_scope", "phase_scope", "teardown_scope", "verification_scope"))
    for marker in ("clone", "ingress", "snapshot", "commit", "publish", "to_bytes", "drop"):
        require(marker in timing, f"{label}: timing scope omits {marker}")
    require("none" in report["performance_claim"].lower(), f"{label}: performance claim is not bounded")
    gates = report.get("gates")
    require(isinstance(gates, dict), f"{label}: gates missing")
    for key in ("source_identity_verified", "source_shape_verified", "semantic_output_verified", "stable_output_verified", "sink_verified", "source_immutability_verified"):
        require(gates.get(key) is True, f"{label}: gate {key} did not pass")


def parse_process_metrics(path: Path, label: str) -> dict[str, int | float]:
    values: dict[str, int | float] = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        require("=" in line, f"{label}: malformed process metric")
        key, raw = line.split("=", 1)
        require(key and key not in values, f"{label}: duplicate process metric {key}")
        if key == "max_rss_kib" or key.endswith("faults") or key.endswith("context_switches"):
            try:
                parsed = int(raw)
            except ValueError as error:
                raise EvidenceError(f"{label}.{key}: expected integer") from error
            values[key] = integer(parsed, f"{label}.{key}", positive=key == "max_rss_kib")
        elif key in ("user_seconds", "system_seconds", "elapsed_seconds"):
            try:
                parsed = float(raw)
            except ValueError as error:
                raise EvidenceError(f"{label}.{key}: expected number") from error
            values[key] = finite_number(parsed, f"{label}.{key}")
        else:
            fail(f"{label}: unexpected process metric {key}")
    for key in ("max_rss_kib", "user_seconds", "system_seconds", "elapsed_seconds", "major_faults", "minor_faults", "voluntary_context_switches", "involuntary_context_switches"):
        require(key in values, f"{label}: missing process metric {key}")
    return values


def check_sample(sample: Any, label: str) -> dict[str, Any]:
    require(isinstance(sample, dict), f"{label}: sample malformed")
    index = integer(sample.get("index"), f"{label}.index")
    require(0 <= index < SAMPLES, f"{label}.index outside measured range")
    result: dict[str, Any] = {"index": index}
    for field in ALL_TIMINGS:
        result[field] = integer(sample.get(field), f"{label}.{field}", positive=True)
    require(result["total_ns"] >= sum(result[field] for field in PHASES), f"{label}: total below phase sum")
    result["output_sha256"] = valid_sha(sample.get("output_sha256"), f"{label}.output_sha256")
    result["output_bytes"] = integer(sample.get("output_bytes"), f"{label}.output_bytes", positive=True)
    require(result["output_sha256"] == OUTPUT_SHA256 and result["output_bytes"] == OUTPUT_BYTES, f"{label}: sample output guard differs")
    return result


def validate_plan(packet: Path) -> tuple[dict[str, Any], list[dict[str, Any]], str]:
    plan = read_json(packet / "capture-plan.json", "capture-plan")
    require(isinstance(plan, dict), "capture-plan: expected object")
    require(plan.get("schema") == "litchi.performance.0913.pptx-binary-variant-plan.v1", "capture-plan: schema differs")
    require(plan.get("status") == "frozen_before_capture", "capture-plan: not frozen before capture")
    require(plan.get("parent") == PARENT, "capture-plan: parent differs")
    require(plan.get("warmup") == WARMUP and plan.get("samples") == SAMPLES, "capture-plan: schedule differs")
    require(plan.get("pairs_per_route") == PAIRS_PER_ROUTE and plan.get("blocks") == BLOCKS and plan.get("cpu") == CPU, "capture-plan: pair/block/CPU scope differs")
    require(plan.get("routes") == list(ROUTES) and plan.get("variants") == list(VARIANTS), "capture-plan: route or variant set differs")
    source = plan.get("source_manifest")
    require(isinstance(source, dict), "capture-plan: source manifest descriptor missing")
    _, source_checked = resolve_descriptor(source, "capture-plan.source_manifest")
    attempt = plan.get("attempt")
    require(isinstance(attempt, str) and attempt.replace("-", "").isalnum(), "capture-plan: invalid attempt")
    legs = plan.get("legs")
    require(isinstance(legs, list) and len(legs) == LEGS, f"capture-plan: expected {LEGS} legs")
    labels: set[str] = set()
    by_block: dict[int, list[dict[str, Any]]] = {index: [] for index in range(BLOCKS)}
    for index, raw in enumerate(legs):
        require(isinstance(raw, dict), f"capture-plan.legs[{index}]: malformed")
        label = raw.get("label")
        require(isinstance(label, str) and label and label not in labels, f"capture-plan.legs[{index}]: duplicate label")
        require(raw.get("route") in ROUTES and raw.get("variant") in VARIANTS, f"capture-plan.legs[{index}]: route/variant differs")
        require(type(raw.get("block")) is int and 0 <= raw["block"] < BLOCKS, f"capture-plan.legs[{index}]: block differs")
        require(type(raw.get("pair")) is int and 0 <= raw["pair"] < PAIRS_PER_ROUTE, f"capture-plan.legs[{index}]: pair differs")
        require(raw.get("position") in (0, 1), f"capture-plan.legs[{index}]: position differs")
        labels.add(label)
        by_block[raw["block"]].append(raw)
    for block, pair in by_block.items():
        require(len(pair) == 2, f"capture-plan block {block}: expected two legs")
        require(sorted(item["position"] for item in pair) == [0, 1], f"capture-plan block {block}: positions differ")
        require(len({item["route"] for item in pair}) == 1 and len({item["pair"] for item in pair}) == 1, f"capture-plan block {block}: pair identity differs")
        route = pair[0]["route"]
        route_pair = pair[0]["pair"]
        require({item["variant"] for item in pair} == set(VARIANTS), f"capture-plan block {block}: variants differ")
        expected = list(VARIANTS) if route_pair % 2 == 0 else list(reversed(VARIANTS))
        ordered = [item["variant"] for item in sorted(pair, key=lambda item: item["position"])]
        require(ordered == expected, f"capture-plan block {block}: AB/BA order differs")
    for route in ROUTES:
        selected = [leg for leg in legs if leg["route"] == route]
        require(len(selected) == PAIRS_PER_ROUTE * 2, f"capture-plan {route}: expected 24 legs")
        require({leg["pair"] for leg in selected} == set(range(PAIRS_PER_ROUTE)), f"capture-plan {route}: pair set differs")
    require(source_checked == descriptor(packet / f"source-{attempt}.json"), "capture-plan: source descriptor differs from attempt manifest")
    return plan, legs, attempt


def baseline_binary_descriptor() -> dict[str, Any]:
    require(BASELINE_BINARY.is_file(), f"missing retained 0910 binary: {BASELINE_BINARY}")
    actual = descriptor(BASELINE_BINARY)
    require(actual["bytes"] == 9_301_456 and actual["sha256"] == "f4ad7ee158061cc80354c3c8674c9fe6a8100ed76a0c8f743653efcb06bb2e8d", "retained 0910 binary identity differs")
    return actual


def source_manifest_checks(packet: Path, attempt: str) -> dict[str, Any]:
    path = packet / f"source-{attempt}.json"
    raw = read_json(path, f"source-{attempt}")
    require(isinstance(raw, dict) and raw.get("status") == "frozen_candidate_source", f"source-{attempt}: status differs")
    require(raw.get("parent") == PARENT, f"source-{attempt}: parent differs")
    require(raw.get("overlay_paths") == list(OVERLAYS), f"source-{attempt}: overlay set differs")
    baseline_raw = read_json(BASELINE_SOURCE_MANIFEST, "0910 source manifest")
    require(baseline_raw.get("parent") == BASELINE_PARENT, "0910 source manifest parent differs")
    baseline_files = {row["path"]: row for row in baseline_raw.get("files", [])}
    files = raw.get("files")
    require(isinstance(files, list), f"source-{attempt}: inventory missing")
    current = {row.get("path"): row for row in files if isinstance(row, dict)}
    require(len(current) == len(files), f"source-{attempt}: duplicate or malformed inventory path")
    expected_paths = set(baseline_files) | {".owned-0913"} | set(OVERLAYS)
    require(set(current) == expected_paths, f"source-{attempt}: inventory is not baseline plus declared overlays/marker")
    for name, row in baseline_files.items():
        require(name in current, f"source-{attempt}: baseline file missing: {name}")
        if name not in OVERLAYS:
            require((current[name].get("bytes"), current[name].get("sha256")) == (row.get("bytes"), row.get("sha256")), f"source-{attempt}: unrequested source changed: {name}")
    for name, row in current.items():
        integer(row.get("bytes"), f"source-{attempt}.{name}.bytes")
        valid_sha(row.get("sha256"), f"source-{attempt}.{name}.sha256")
    baseline_descriptor = descriptor(BASELINE_SOURCE_MANIFEST)
    _, checked_baseline_descriptor = resolve_descriptor(raw.get("baseline_source_manifest"), f"source-{attempt}.baseline_source_manifest")
    require(checked_baseline_descriptor == baseline_descriptor, f"source-{attempt}: baseline manifest binding differs")
    resolve_descriptor(raw.get("overlay_patch"), f"source-{attempt}.overlay_patch")
    overlays = raw.get("overlays")
    require(isinstance(overlays, list) and {row.get("path") for row in overlays if isinstance(row, dict)} == set(OVERLAYS), f"source-{attempt}: overlay trace differs")
    return {"descriptor": descriptor(path), "raw": raw, "files": len(files), "baseline_files": len(baseline_files)}


def check_start(folder: Path, receipt: Mapping[str, Any], label: str) -> None:
    start = read_json(folder / "start.json", f"{label}.start")
    require(isinstance(start, dict), f"{label}.start: malformed")
    for field in ("argv", "start_utc"):
        require(start.get(field) == receipt.get(field), f"{label}: start/{field} binding differs")


def check_build(packet: Path, plan: Mapping[str, Any], attempt: str, source: Mapping[str, Any]) -> dict[str, Any]:
    folder = packet / f"build-{attempt}"
    receipt = read_json(folder / "command.json", f"build-{attempt}.command")
    require(receipt.get("exit_code") == 0, f"build-{attempt}: candidate build failed")
    check_start(folder, receipt, f"build-{attempt}")
    build_started = parse_timestamp(receipt.get("start_utc"), f"build-{attempt}.start_utc")
    build_ended = parse_timestamp(receipt.get("end_utc"), f"build-{attempt}.end_utc")
    require(build_ended >= build_started, f"build-{attempt}: ends before it starts")
    argv = receipt.get("argv")
    require(isinstance(argv, list), f"build-{attempt}: argv missing")
    required = ["cargo", "build", "--locked", "--offline", "--release", "--example", "pptx_owned_benchmark", "-j", "2"]
    require(all(item in argv for item in required), f"build-{attempt}: release command flags differ")
    require(receipt.get("source") == source["descriptor"], f"build-{attempt}: source binding differs")
    require(receipt.get("helper") == descriptor(packet / "run.py"), f"build-{attempt}: helper binding differs")
    require(receipt.get("plan") == descriptor(packet / "capture-plan.json"), f"build-{attempt}: plan binding differs")
    require(isinstance(receipt.get("toolchain"), dict), f"build-{attempt}: toolchain receipt missing")
    for name in ("cargo", "rustc"):
        row = receipt["toolchain"].get(name)
        require(isinstance(row, dict) and isinstance(row.get("version"), str) and row["version"], f"build-{attempt}: {name} identity missing")
        require(isinstance(row.get("launcher"), str) and row["launcher"], f"build-{attempt}: {name} launcher missing")
    baseline_build_path = packet.parent / "change-0910" / "build-v1" / "command.json"
    if baseline_build_path.is_file():
        baseline_build = read_json(baseline_build_path, "0910 build command")
        require(receipt["toolchain"] == baseline_build.get("toolchain"), f"build-{attempt}: toolchain differs from retained 0910 build")
        require(receipt.get("removed_environment_keys") == baseline_build.get("removed_environment_keys"), f"build-{attempt}: removed environment differs")
        current_env = receipt.get("environment_overrides")
        baseline_env = baseline_build.get("environment_overrides")
        require(isinstance(current_env, dict) and isinstance(baseline_env, dict), f"build-{attempt}: environment override receipt malformed")
        for key in ("CARGO_BUILD_JOBS", "CARGO_INCREMENTAL", "CARGO_PROFILE_RELEASE_LTO", "CARGO_PROFILE_RELEASE_CODEGEN_UNITS", "CARGO_PROFILE_RELEASE_DEBUG", "LC_ALL", "LANG", "TZ"):
            require(current_env.get(key) == baseline_env.get(key), f"build-{attempt}: build flag environment differs for {key}")
    require(receipt.get("logs") and len(receipt["logs"]) == 2, f"build-{attempt}: stdout/stderr receipts missing")
    for log in receipt["logs"]:
        resolve_descriptor(log, f"build-{attempt}.log")
    candidate = descriptor_shape(receipt.get("binary"), f"build-{attempt}.binary")
    retained = descriptor_shape(receipt.get("retained_binary"), f"build-{attempt}.retained_binary")
    _, retained_actual = resolve_descriptor(receipt["retained_binary"], f"build-{attempt}.retained_binary")
    require(candidate["bytes"] == retained["bytes"] and candidate["sha256"] == retained["sha256"], f"build-{attempt}: candidate retention identity differs")
    baseline = baseline_binary_descriptor()
    require(receipt.get("baseline_binary") == baseline, f"build-{attempt}: baseline binary binding differs")
    optional_descriptor(receipt["binary"], f"build-{attempt}.binary")
    handoff = read_json(folder / "binary.json", f"build-{attempt}.retention")
    require(handoff == {"original": receipt["binary"], "retained": receipt["retained_binary"]}, f"build-{attempt}: retained binary handoff differs")
    return {"receipt": receipt, "candidate": retained_actual, "baseline": baseline, "folder": folder}


def check_command(
    packet: Path,
    plan: Mapping[str, Any],
    leg: Mapping[str, Any],
    source: Mapping[str, Any],
    build: Mapping[str, Any],
) -> dict[str, Any]:
    label = str(leg["label"])
    folder = packet / "capture" / label
    receipt = read_json(folder / "command.json", f"{label}.command")
    require(receipt.get("exit_code") == 0, f"{label}: child failed; failed captures are rejected")
    check_start(folder, receipt, label)
    require(receipt.get("leg") == leg, f"{label}: leg differs from frozen plan")
    require(receipt.get("variant") == leg["variant"], f"{label}: variant differs")
    require(receipt.get("source") == source["descriptor"], f"{label}: source binding differs")
    require(receipt.get("helper") == descriptor(packet / "capture.py"), f"{label}: capture helper binding differs")
    require(receipt.get("build_helper") == descriptor(packet / "run.py"), f"{label}: build helper binding differs")
    require(receipt.get("host_environment") == descriptor(packet / "environment-after-recovery.json"), f"{label}: capture host environment binding differs")
    require(receipt.get("process_tools") == descriptor(packet / "process-tools.json"), f"{label}: process tool metadata binding differs")
    require(receipt.get("environment_is_complete") is True, f"{label}: capture environment is not complete")
    require(receipt.get("environment_overrides") == {"PATH": "/usr/bin:/bin", "LC_ALL": "C", "LANG": "C", "TZ": "UTC"}, f"{label}: capture environment differs")
    require(receipt.get("plan") == descriptor(packet / "capture-plan.json"), f"{label}: plan binding differs")
    baseline = build["baseline"]
    candidate = build["candidate"]
    expected_binary = baseline if leg["variant"] == "baseline" else candidate
    require(receipt.get("binary") == expected_binary, f"{label}: executable binding differs")
    require(receipt.get("baseline_binary") == baseline, f"{label}: baseline binding differs")
    require(receipt.get("candidate_binary") == candidate, f"{label}: candidate binding differs")
    argv = receipt.get("argv")
    require(isinstance(argv, list) and all(isinstance(item, str) for item in argv), f"{label}: argv malformed")
    require(argv[:2] == ["/usr/bin/time", "-f"] and "/usr/bin/time" in argv, f"{label}: /usr/bin/time missing")
    require("taskset" in argv and "-c" in argv and str(CPU) in argv, f"{label}: CPU pinning missing")
    artifact_folder = folder / "artifacts"
    require(len(argv) >= 5 and argv[-4:] == [leg["route"], str(WARMUP), str(SAMPLES), str(artifact_folder)], f"{label}: benchmark schedule differs")
    require(receipt.get("logs") and len(receipt["logs"]) == 2, f"{label}: stdout/stderr receipts missing")
    for log in receipt["logs"]:
        resolve_descriptor(log, f"{label}.log")
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, list), f"{label}: artifact receipts missing")
    by_name: dict[str, tuple[Path, dict[str, Any]]] = {}
    for raw in artifacts:
        path, checked = resolve_descriptor(raw, f"{label}.artifact")
        require(path.parent == artifact_folder.resolve(), f"{label}: artifact escaped output directory")
        require(path.name not in by_name, f"{label}: duplicate artifact {path.name}")
        by_name[path.name] = (path, checked)
    for name in ("source.pptx", "output.pptx", "report.json"):
        require(name in by_name, f"{label}: missing {name}")
    process_path, process_checked = resolve_descriptor(receipt.get("process_metrics"), f"{label}.process_metrics")
    require(process_path == (folder / "process.txt").resolve(), f"{label}: process metrics path differs")
    rss = parse_process_metrics(process_path, f"{label}.process_metrics")
    report_path, report_artifact = by_name["report.json"]
    report = read_json(report_path, f"{label}.report")
    check_report_scope(report, label, str(leg["route"]))
    expected_corpus(report.get("corpus"), f"{label}.corpus")
    source_artifact = by_name["source.pptx"][1]
    output_artifact = by_name["output.pptx"][1]
    require(source_artifact["bytes"] == SOURCE_BYTES and source_artifact["sha256"] == SOURCE_SHA256, f"{label}: source archive identity differs")
    require(output_artifact["bytes"] == OUTPUT_BYTES and output_artifact["sha256"] == OUTPUT_SHA256, f"{label}: output archive identity differs")
    require(report.get("source") == {"file": "source.pptx", "bytes": SOURCE_BYTES, "sha256": SOURCE_SHA256}, f"{label}: report source binding differs")
    require(report.get("output") == {"file": "output.pptx", "bytes": OUTPUT_BYTES, "sha256": OUTPUT_SHA256}, f"{label}: report output binding differs")
    sink_capacity = integer(report.get("sink_capacity_bytes"), f"{label}.sink_capacity_bytes", positive=True)
    sink_limit = integer(report.get("sink_max_write_bytes"), f"{label}.sink_max_write_bytes", positive=True)
    require(sink_capacity == 2 * SOURCE_BYTES + SINK_MAX_WRITE and sink_limit == SINK_MAX_WRITE, f"{label}: sink budget differs")
    samples = report.get("samples")
    require(isinstance(samples, list) and len(samples) == SAMPLES, f"{label}: sample vector cardinality differs")
    checked_samples = [check_sample(item, f"{label}.samples[{index}]") for index, item in enumerate(samples)]
    require([item["index"] for item in checked_samples] == list(range(SAMPLES)), f"{label}: sample indices are not ordered")
    expected_calls = (OUTPUT_BYTES + SINK_MAX_WRITE - 1) // SINK_MAX_WRITE
    for index, sample in enumerate(samples):
        calls = integer(sample.get("sink_write_calls"), f"{label}.samples[{index}].sink_write_calls", positive=True)
        largest = integer(sample.get("sink_largest_write"), f"{label}.samples[{index}].sink_largest_write", positive=True)
        require(calls == expected_calls and largest == min(OUTPUT_BYTES, SINK_MAX_WRITE), f"{label}: sink write guard differs")
        checked_samples[index]["sink_write_calls"] = calls
        checked_samples[index]["sink_largest_write"] = largest
    started = parse_timestamp(receipt.get("start_utc"), f"{label}.start_utc")
    ended = parse_timestamp(receipt.get("end_utc"), f"{label}.end_utc")
    require(ended >= started, f"{label}: process ended before start")
    return {
        "label": label,
        "block": leg["block"],
        "pair": leg["pair"],
        "route": leg["route"],
        "variant": leg["variant"],
        "position": leg["position"],
        "folder": relative(folder),
        "command": descriptor(folder / "command.json"),
        "binary": expected_binary,
        "source": source_artifact,
        "output": output_artifact,
        "report": descriptor(report_path),
        "process_metrics": process_checked,
        "rss": rss,
        "samples": checked_samples,
        "corpus": report["corpus"],
        "started": started,
        "ended": ended,
    }


def variant_summary(legs: Sequence[Mapping[str, Any]], route: str, variant: str) -> dict[str, Any]:
    selected = [leg for leg in legs if leg["route"] == route and leg["variant"] == variant]
    require(len(selected) == PAIRS_PER_ROUTE, f"{route}/{variant}: expected {PAIRS_PER_ROUTE} processes")
    out: dict[str, Any] = {"legs": [leg["label"] for leg in selected], "samples": PAIRS_PER_ROUTE * SAMPLES}
    for field in ALL_TIMINGS:
        out[field] = summary([sample[field] for leg in selected for sample in leg["samples"]], f"{route}.{variant}.{field}")
    out["max_rss_kib"] = summary([int(leg["rss"]["max_rss_kib"]) for leg in selected], f"{route}.{variant}.max_rss_kib")
    return out


def paired_route(legs: Sequence[Mapping[str, Any]], route: str) -> dict[str, Any]:
    selected = [leg for leg in legs if leg["route"] == route]
    by_pair: dict[int, dict[str, Mapping[str, Any]]] = {index: {} for index in range(PAIRS_PER_ROUTE)}
    for leg in selected:
        by_pair[int(leg["pair"])][str(leg["variant"])] = leg
    rows: list[dict[str, Any]] = []
    ratios: dict[str, list[float]] = {field: [] for field in ALL_TIMINGS}
    for pair in range(PAIRS_PER_ROUTE):
        row_pair = by_pair[pair]
        require(set(row_pair) == set(VARIANTS), f"{route}/pair-{pair}: missing variant")
        baseline = row_pair["baseline"]
        candidate = row_pair["candidate"]
        row: dict[str, Any] = {
            "pair": pair,
            "order": "AB" if baseline["position"] == 0 else "BA",
            "baseline": baseline["label"],
            "candidate": candidate["label"],
        }
        for field in ALL_TIMINGS:
            baseline_median = median([sample[field] for sample in baseline["samples"]], f"{route}/pair-{pair}.baseline.{field}")
            candidate_median = median([sample[field] for sample in candidate["samples"]], f"{route}/pair-{pair}.candidate.{field}")
            value = ratio(baseline_median, candidate_median, f"{route}/pair-{pair}.{field}")
            row[field] = {"baseline_median": baseline_median, "candidate_median": candidate_median, "baseline_over_candidate": value}
            ratios[field].append(value)
        row["baseline_max_rss_kib"] = baseline["rss"]["max_rss_kib"]
        row["candidate_max_rss_kib"] = candidate["rss"]["max_rss_kib"]
        rows.append(row)
    return {
        "route": route,
        "pairs": rows,
        "median_ratio_bootstrap": {field: bootstrap_median(values, f"{route}.{field}") for field, values in ratios.items()},
    }


def drift_analysis(legs: Sequence[Mapping[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for route in ROUTES:
        result[route] = {}
        for variant in VARIANTS:
            result[route][variant] = {}
            for order in ("AB", "BA"):
                selected = [leg for leg in legs if leg["route"] == route and leg["variant"] == variant and ((leg["pair"] % 2 == 0) == (order == "AB"))]
                require(len(selected) == PAIRS_PER_ROUTE // 2, f"{route}/{variant}/{order}: expected six processes")
                result[route][variant][order] = {field: summary([sample[field] for leg in selected for sample in leg["samples"]], f"drift.{route}.{variant}.{order}.{field}") for field in ALL_TIMINGS}
            result[route][variant]["BA_over_AB_p50"] = {field: ratio(result[route][variant]["BA"][field]["p50"], result[route][variant]["AB"][field]["p50"], f"drift.{route}.{variant}.{field}") for field in ALL_TIMINGS}
    return result


def regression_flags(routes: Mapping[str, Any], paired: Mapping[str, Any]) -> dict[str, Any]:
    flags: list[dict[str, Any]] = []
    for route in ROUTES:
        baseline = routes[route]["baseline"]
        candidate = routes[route]["candidate"]
        for field in ALL_TIMINGS + ("max_rss_kib",):
            baseline_p50 = baseline[field]["p50"]
            candidate_p50 = candidate[field]["p50"]
            if candidate_p50 > baseline_p50 * 1.05:
                flags.append({"route": route, "metric": field, "kind": "candidate_p50_regression_over_5_percent", "baseline_p50": baseline_p50, "candidate_p50": candidate_p50, "candidate_over_baseline": candidate_p50 / baseline_p50})
            baseline_p95 = baseline[field]["p95"]
            candidate_p95 = candidate[field]["p95"]
            if candidate_p95 > baseline_p95 * 1.05:
                flags.append({"route": route, "metric": field, "kind": "candidate_p95_regression_over_5_percent", "baseline_p95": baseline_p95, "candidate_p95": candidate_p95, "candidate_over_baseline": candidate_p95 / baseline_p95})
            baseline_p99 = baseline[field]["p99"]
            candidate_p99 = candidate[field]["p99"]
            if candidate_p99 > baseline_p99 * 1.05:
                flags.append({"route": route, "metric": field, "kind": "candidate_p99_regression_over_5_percent", "baseline_p99": baseline_p99, "candidate_p99": candidate_p99, "candidate_over_baseline": candidate_p99 / baseline_p99})
    owned_total = paired["owned"]["median_ratio_bootstrap"]["total_ns"]
    return {
        "threshold": 1.05,
        "flags": flags,
        "owned_total_bootstrap_lower_bound": owned_total["lower"],
        "owned_total_bootstrap_upper_bound": owned_total["upper"],
        "owned_total_ci_lower_above_one": owned_total["lower"] > 1.0,
        "owned_total_median": owned_total["statistic"],
        "owned_total_median_practical_target_met": owned_total["statistic"] >= 1.05,
        "owned_total_gate_pass": owned_total["statistic"] >= 1.05 and owned_total["lower"] > 1.0,
        "interpretation": "The owned route meets the descriptive gate only when its baseline-over-candidate total median is at least 1.05 and the paired bootstrap lower bound is above 1.0; route-specific >5% p50/p95/p99 or RSS flags still require review. These are controls within each route, not a cross-route claim.",
    }


def analyze(packet: Path) -> dict[str, Any]:
    require(packet.resolve() == HERE.resolve(), "analysis packet must be the change-0913 result directory")
    plan, plan_legs, attempt = validate_plan(packet)
    source = source_manifest_checks(packet, attempt)
    build = check_build(packet, plan, attempt, source)
    parsed = [check_command(packet, plan, leg, source, build) for leg in plan_legs]
    require(len(parsed) == LEGS, f"capture: expected {LEGS} child processes")
    require({leg["label"] for leg in parsed} == {leg["label"] for leg in plan_legs}, "capture: labels differ from plan")
    require(len({(leg["source"]["bytes"], leg["source"]["sha256"]) for leg in parsed}) == 1, "capture: source artifact differs across children")
    require(len({(leg["output"]["bytes"], leg["output"]["sha256"]) for leg in parsed}) == 1, "capture: output artifact differs across children")
    require(len({canonical(leg["corpus"]) for leg in parsed}) == 1, "capture: corpus differs across children")
    require(parsed[0]["source"]["bytes"] == SOURCE_BYTES and parsed[0]["source"]["sha256"] == SOURCE_SHA256, "capture: source guard differs")
    require(parsed[0]["output"]["bytes"] == OUTPUT_BYTES and parsed[0]["output"]["sha256"] == OUTPUT_SHA256, "capture: output guard differs")
    previous_end: datetime | None = None
    for leg in parsed:
        if previous_end is not None:
            require(leg["started"] >= previous_end, f"{leg['label']}: process overlaps previous child")
        previous_end = leg["ended"]
    routes = {route: {variant: variant_summary(parsed, route, variant) for variant in VARIANTS} for route in ROUTES}
    paired = {route: paired_route(parsed, route) for route in ROUTES}
    drift = drift_analysis(parsed)
    flags = regression_flags(routes, paired)
    return {
        "schema": SCHEMA,
        "status": "qualified_candidate_control",
        "production_change": True,
        "performance_claim": False,
        "scope": {
            "experiment": "retained 0910 baseline executable versus candidate executable built from the frozen source plus owned direct-read overlays",
            "warmup_iterations_per_process": WARMUP,
            "measured_samples_per_process": SAMPLES,
            "paired_processes_per_route": PAIRS_PER_ROUTE,
            "child_processes": LEGS,
            "routes": list(ROUTES),
            "variants": list(VARIANTS),
            "cpu": CPU,
            "timed_phases": list(ALL_TIMINGS),
            "ratio_direction": "baseline over candidate, within each route",
            "patch_replay_in_primary_timer": False,
        },
        "custody": {
            "capture_plan": descriptor(packet / "capture-plan.json"),
            "source_manifest": source["descriptor"],
            "baseline_binary": build["baseline"],
            "candidate_binary": build["candidate"],
            "build": descriptor(packet / f"build-{attempt}/command.json"),
            "binary_handoff": descriptor(packet / f"build-{attempt}/binary.json"),
            "plan": {"status": plan["status"], "parent": plan["parent"], "pairs_per_route": PAIRS_PER_ROUTE, "warmup": WARMUP, "samples": SAMPLES, "cpu": CPU},
            "leg_count": len(parsed),
            "serial_process_windows": True,
        },
        "corpus": {"manifest": parsed[0]["corpus"], "source": parsed[0]["source"], "same_across_legs": True},
        "output": {"artifact": parsed[0]["output"], "same_across_legs": True, "per_leg": {leg["label"]: leg["output"] for leg in parsed}},
        "routes": routes,
        "paired_by_route": paired,
        "drift_by_route_variant_order": drift,
        "regression_flags": flags,
        "legs": [{key: value for key, value in leg.items() if key not in ("started", "ended", "corpus")} for leg in parsed],
        "checks": {
            "plan_frozen_before_capture": True,
            "twelve_ab_ba_pairs_per_route": True,
            "four_variants_and_48_children": True,
            "twenty_samples_per_child": True,
            "baseline_and_candidate_binary_custody": True,
            "toolchain_and_command_logs_bound": True,
            "source_output_and_sink_guards_byte_equal": True,
            "corpus_manifest_equal_and_expected": True,
            "reported_sample_vectors_recomputed": True,
            "phase_sum_not_above_total": True,
            "serial_nonoverlap": True,
            "route_specific_bootstrap_seed_913_10000_draws": True,
            "owned_total_ci_and_practical_target_reported": True,
            "route_specific_five_percent_regression_flags_reported": True,
            "rss_is_whole_process": True,
        },
        "limitations": [
            "All quantiles, ratios, bootstrap intervals, RSS values, and regression flags are descriptive evidence for this generated fixture and shared host.",
            "The candidate is evaluated against a retained executable built from the same frozen baseline source plus the declared overlays; the two routes remain separate controls.",
            "No materialized-versus-owned cross-route performance claim, allocation claim, instruction claim, cold-cache claim, range-source claim, concurrency claim, producer claim, or physical-package generalization is made.",
            "Maximum RSS is a whole-process maximum including generation, warmups, guards, and measured samples; it is not operation peak live memory.",
        ],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True, help="new analysis JSON path")
    args = parser.parse_args()
    output = args.output.resolve()
    try:
        require(not output.exists() and not output.is_symlink(), f"refusing to overwrite analysis output: {output}")
        result = analyze(HERE)
        output.parent.mkdir(parents=True, exist_ok=True)
        with output.open("x", encoding="utf-8") as stream:
            json.dump(result, stream, indent=2, sort_keys=True)
            stream.write("\n")
    except (EvidenceError, OSError) as error:
        print(f"analyze.py: error: {error}", file=sys.stderr)
        return 2
    print("0913 matched PPTX binary-variant analysis passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
