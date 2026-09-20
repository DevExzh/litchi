#!/usr/bin/env python3
"""Audit and compare the frozen XLS worksheet-replay checkpoint captures.

The analyzer treats the raw probe stdout as evidence.  It verifies source,
probe, fixture, binary and build-manifest bindings before calculating any
statistics, compares every semantic result across A/A and A/B/B/A legs, and
then applies the packet's timing and allocation gates.  It writes
``analysis.json`` and ``analysis.md`` beside the captures and exits nonzero
for any binding, parity or hard-gate failure.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import math
import platform
import statistics
import subprocess
import sys
from pathlib import Path
from typing import Any


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
PLAN_PATH = PACKET / "plan.json"
CAPTURES = PACKET / "captures"


class AnalysisError(Exception):
    """A recoverable evidence failure recorded in the final report."""


def read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: Path, value: Any) -> None:
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def rel(path: Path) -> str:
    return path.relative_to(ROOT).as_posix()


def record(failures: list[str], message: str) -> None:
    failures.append(message)


def source_files(plan: dict[str, Any]) -> list[Path]:
    files: list[Path] = []
    for root in plan["source_roots"]:
        files.extend(path for path in (ROOT / root).rglob("*") if path.is_file())
    return sorted(files)


def current_source_hashes(plan: dict[str, Any]) -> dict[str, str]:
    return {rel(path): sha256_file(path) for path in source_files(plan)}


def git_source_hashes(plan: dict[str, Any], revision: str) -> dict[str, str | None]:
    result: dict[str, str | None] = {}
    for path in source_files(plan):
        name = rel(path)
        exists = subprocess.run(
            ["git", "cat-file", "-e", f"{revision}:{name}"],
            cwd=ROOT,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        if exists.returncode:
            result[name] = None
        else:
            result[name] = hashlib.sha256(
                subprocess.check_output(["git", "show", f"{revision}:{name}"], cwd=ROOT)
            ).hexdigest()
    return result


def current_probe_hashes(plan: dict[str, Any]) -> dict[str, str]:
    result: dict[str, str] = {}
    for root in plan["probe_roots"]:
        for path in sorted((ROOT / root).rglob("*")):
            if path.is_file():
                result[rel(path)] = sha256_file(path)
    return result


def current_tool_hashes() -> dict[str, str]:
    return {
        name: sha256_file(PACKET / name)
        for name in ("pilot.py", "analyze.py", "plan.json")
    }


def current_corpus_hashes(plan: dict[str, Any]) -> dict[str, dict[str, int | str]]:
    cases = list(plan["cases"]) + list(plan["budget_fence"]["cases"])
    result: dict[str, dict[str, int | str]] = {}
    for case in cases:
        path = ROOT / case["path"]
        if case["path"] not in result:
            result[case["path"]] = {
                "bytes": path.stat().st_size,
                "sha256": sha256_file(path),
            }
    return result


def audit_freeze(
    plan: dict[str, Any], failures: list[str], offline: bool
) -> dict[str, Any] | None:
    path = CAPTURES / "freeze.json"
    if not path.is_file():
        record(failures, f"freeze: missing {path}")
        return None
    try:
        frozen = read_json(path)
    except (OSError, json.JSONDecodeError) as error:
        record(failures, f"freeze: invalid {path}: {error}")
        return None
    if frozen.get("status") != "frozen":
        record(failures, "freeze: status is not frozen")
    if frozen.get("plan_sha256") != sha256_file(PLAN_PATH):
        record(failures, "freeze: plan hash differs")
    case_manifest = ROOT / plan["case_manifest"]
    if frozen.get("case_manifest_sha256") != sha256_file(case_manifest):
        record(failures, "freeze: case manifest hash differs")
    current_tools = current_tool_hashes()
    if frozen.get("tools_sha256_start") != current_tools or frozen.get("tools_sha256_end") != current_tools:
        record(failures, "freeze: tooling hash differs")
    current = current_source_hashes(plan)
    if frozen.get("source_sha256_start") != current or frozen.get("source_sha256_end") != current:
        record(failures, "freeze: source hash differs")
    current_probes = current_probe_hashes(plan)
    if frozen.get("probe_sha256") != current_probes or frozen.get("probe_sha256_end") != current_probes:
        record(failures, "freeze: probe hash differs")
    current_corpus = current_corpus_hashes(plan)
    if frozen.get("corpus") != current_corpus or frozen.get("corpus_end") != current_corpus:
        record(failures, "freeze: fixture hash differs")
    case_manifest_sha = sha256_file(case_manifest)
    if frozen.get("case_manifest_sha256_end") != case_manifest_sha:
        record(failures, "freeze: case manifest end hash differs")
    records = frozen.get("binaries", {})
    witnesses = cleanup_witnesses()
    for phase in ("baseline", "candidate"):
        for kind in plan["binaries"]:
            binary = binary_path(plan, phase, kind)
            frozen_binary = records.get(phase, {}).get(kind, {})
            binary_identity(
                binary,
                frozen_binary.get("sha256"),
                frozen_binary.get("bytes"),
                offline,
                witnesses,
                failures,
                f"freeze: {phase} {kind}",
            )
    if frozen.get("binaries_end") != records:
        record(failures, "freeze: binary end binding differs")
    builds = frozen.get("build_manifests", {})
    for phase in ("baseline", "candidate"):
        build = builds.get(phase)
        if not build or not Path(build.get("path", "")).is_file():
            record(failures, f"freeze: missing {phase} build manifest")
    if frozen.get("build_manifests_end") != builds:
        record(failures, "freeze: build manifest end binding differs")
    return frozen


def binary_path(plan: dict[str, Any], phase: str, kind: str) -> Path:
    return Path(plan["binary_root"]) / phase / plan["binaries"][kind]


def cleanup_witnesses() -> list[dict[str, Any]]:
    witnesses: list[dict[str, Any]] = []
    for filename in ("cleanup.json", "cleanup-witness.json", "cleanup-receipt.json"):
        path = PACKET / filename
        if not path.is_file() or path.is_symlink():
            continue
        try:
            value = read_json(path)
        except (OSError, json.JSONDecodeError):
            continue

        def walk(item: Any) -> None:
            if isinstance(item, dict):
                raw_path = item.get("path")
                digest = item.get("sha256", item.get("binary_sha256"))
                size = item.get("bytes", item.get("binary_bytes"))
                if isinstance(raw_path, str) and isinstance(digest, str):
                    witnesses.append({"path": raw_path, "sha256": digest, "bytes": size})
                for child in item.values():
                    walk(child)
            elif isinstance(item, list):
                for child in item:
                    walk(child)

        walk(value)
    return witnesses


def resolve_witness_path(raw_path: str) -> Path:
    path = Path(raw_path)
    return (ROOT / path).resolve() if not path.is_absolute() else path.resolve()


def binary_identity(
    path: Path,
    expected_sha: str | None,
    expected_bytes: int | None,
    offline: bool,
    witnesses: list[dict[str, Any]],
    failures: list[str],
    label: str,
) -> str | None:
    if path.is_file() and not path.is_symlink():
        actual_sha = sha256_file(path)
        actual_bytes = path.stat().st_size
        if expected_sha is not None and actual_sha != expected_sha:
            record(failures, f"{label}: live binary hash differs")
            return None
        if expected_bytes is not None and actual_bytes != expected_bytes:
            record(failures, f"{label}: live binary size differs")
            return None
        return "live-binary"
    if not offline:
        record(failures, f"{label}: binary is absent; rerun with --offline and an exact cleanup witness")
        return None
    for witness in witnesses:
        if (
            resolve_witness_path(str(witness["path"])) == path.resolve()
            and witness.get("sha256") == expected_sha
            and witness.get("bytes") == expected_bytes
        ):
            return "exact-cleanup-witness"
    record(failures, f"{label}: absent binary has no exact cleanup witness")
    return None


def compare_map(
    failures: list[str], label: str, expected: Any, actual: Any
) -> bool:
    if expected != actual:
        record(failures, f"{label}: expected binding differs")
        return False
    return True


def audit_manifest(
    plan: dict[str, Any],
    task: str,
    required_binary_kinds: list[str],
    failures: list[str],
    offline: bool,
) -> dict[str, Any] | None:
    path = CAPTURES / task / "manifest.json"
    if not path.is_file():
        record(failures, f"{task}: missing manifest {path}")
        return None
    try:
        manifest = read_json(path)
    except (OSError, json.JSONDecodeError) as error:
        record(failures, f"{task}: invalid manifest {path}: {error}")
        return None
    if manifest.get("task") != task:
        record(failures, f"{task}: manifest task field is {manifest.get('task')!r}")
    if manifest.get("status") != "complete":
        record(failures, f"{task}: manifest status is {manifest.get('status')!r}")
    if manifest.get("plan_sha256") != sha256_file(PLAN_PATH):
        record(failures, f"{task}: plan hash does not match capture manifest")
    freeze_path = CAPTURES / "freeze.json"
    if not freeze_path.is_file():
        record(failures, f"{task}: missing freeze before capture")
    else:
        if manifest.get("freeze_sha256") != sha256_file(freeze_path):
            record(failures, f"{task}: freeze hash does not match capture")
        frozen = read_json(freeze_path)
        if manifest.get("tools_sha256_start") != frozen.get("tools_sha256_start"):
            record(failures, f"{task}: tooling start hash differs from freeze")
        if manifest.get("tools_sha256_end") != frozen.get("tools_sha256_end"):
            record(failures, f"{task}: tooling end hash differs from freeze")
        if manifest.get("source_sha256_start") != frozen.get("source_sha256_start"):
            record(failures, f"{task}: source start hash differs from freeze")
        if manifest.get("source_sha256_end") != frozen.get("source_sha256_end"):
            record(failures, f"{task}: source end hash differs from freeze")
        if manifest.get("probe_sha256") != frozen.get("probe_sha256"):
            record(failures, f"{task}: probe start hash differs from freeze")
        if manifest.get("probe_sha256_end") != frozen.get("probe_sha256_end"):
            record(failures, f"{task}: probe end hash differs from freeze")
        if manifest.get("corpus") != frozen.get("corpus"):
            record(failures, f"{task}: fixture start hash differs from freeze")
        if manifest.get("corpus_end") != frozen.get("corpus_end"):
            record(failures, f"{task}: fixture end hash differs from freeze")
        if manifest.get("case_manifest_sha256") != frozen.get("case_manifest_sha256"):
            record(failures, f"{task}: case manifest start hash differs from freeze")
        if manifest.get("case_manifest_sha256_end") != frozen.get("case_manifest_sha256_end"):
            record(failures, f"{task}: case manifest end hash differs from freeze")
        if manifest.get("binaries") != frozen.get("binaries"):
            record(failures, f"{task}: binary start binding differs from freeze")
        if manifest.get("binaries_end") != frozen.get("binaries_end"):
            record(failures, f"{task}: binary end binding differs from freeze")
        if manifest.get("build_manifests") != frozen.get("build_manifests"):
            record(failures, f"{task}: build manifest start binding differs from freeze")
        if manifest.get("build_manifests_end") != frozen.get("build_manifests_end"):
            record(failures, f"{task}: build manifest end binding differs from freeze")
        if manifest.get("tools_sha256_end") != current_tool_hashes():
            record(failures, f"{task}: tooling end hash differs")
    if manifest.get("tools_sha256_start") != current_tool_hashes():
        record(failures, f"{task}: tooling start hash differs from current files")
    case_manifest_path = ROOT / plan["case_manifest"]
    case_manifest_sha = sha256_file(case_manifest_path)
    if manifest.get("case_manifest_sha256") != case_manifest_sha:
        record(failures, f"{task}: case manifest hash differs")
    if manifest.get("case_manifest_sha256_end") != case_manifest_sha:
        record(failures, f"{task}: case manifest end hash differs")
    try:
        resolved = subprocess.check_output(
            ["git", "rev-parse", plan["baseline_revision"]], cwd=ROOT, text=True
        ).strip()
        if manifest.get("baseline_revision_resolved") != resolved:
            record(failures, f"{task}: baseline revision binding differs")
    except subprocess.CalledProcessError as error:
        record(failures, f"{task}: cannot resolve baseline revision: {error}")

    current = current_source_hashes(plan)
    baseline = git_source_hashes(plan, plan["baseline_revision"])
    recorded_baseline = manifest.get("baseline_source_sha256")
    if recorded_baseline != baseline:
        record(failures, f"{task}: archived baseline source map differs from git {plan['baseline_revision']}")

    start = manifest.get("source_sha256_start")
    end = manifest.get("source_sha256_end")
    if start != current:
        record(failures, f"{task}: source start hash map differs from frozen candidate sources")
    if end != current:
        record(failures, f"{task}: end source hash map differs from current candidate sources")
    current_probes = current_probe_hashes(plan)
    if manifest.get("probe_sha256") != current_probes:
        record(failures, f"{task}: immutable probe hash map differs")
    if manifest.get("probe_sha256_end") != current_probes:
        record(failures, f"{task}: immutable probe end hash map differs")
    current_corpus = current_corpus_hashes(plan)
    if manifest.get("corpus") != current_corpus:
        record(failures, f"{task}: fixture hash map differs")
    if manifest.get("corpus_end") != current_corpus:
        record(failures, f"{task}: fixture end hash map differs")

    binaries = manifest.get("binaries", {})
    if manifest.get("binaries_end") != binaries:
        record(failures, f"{task}: binary start/end bindings differ")
    witnesses = cleanup_witnesses()
    for phase in ("baseline", "candidate"):
        phase_records = binaries.get(phase, {})
        for kind in plan["binaries"]:
            path_record = phase_records.get(kind, {})
            path = binary_path(plan, phase, kind)
            exists = path.is_file()
            if path_record.get("available") != exists and exists:
                record(failures, f"{task}: {phase} {kind} availability binding differs")
            binary_identity(
                path,
                path_record.get("sha256"),
                path_record.get("bytes"),
                offline,
                witnesses,
                failures,
                f"{task}: {phase} {kind}",
            )
        for kind in required_binary_kinds:
            required = binary_path(plan, phase, kind)
            binary_record = binaries.get(phase, {}).get(kind, {})
            if binary_identity(
                required,
                binary_record.get("sha256"),
                binary_record.get("bytes"),
                offline,
                witnesses,
                failures,
                f"{task}: required {phase} {kind}",
            ) is None:
                pass

    for phase in ("baseline", "candidate"):
        build_record = manifest.get("build_manifests", {}).get(phase)
        if build_record is None:
            record(failures, f"{task}: missing {phase} builds.json")
            continue
        build_path = Path(build_record.get("path", ""))
        if not build_path.is_file():
            record(failures, f"{task}: missing {phase} builds.json {build_path}")
            continue
        if build_record.get("sha256") != sha256_file(build_path):
            record(failures, f"{task}: {phase} builds.json hash differs")
            continue
        try:
            build_rows = read_json(build_path)
        except (OSError, json.JSONDecodeError) as error:
            record(failures, f"{task}: invalid {build_path}: {error}")
            continue
        if build_rows != build_record.get("records"):
            record(failures, f"{task}: {phase} builds.json content changed after capture")
        if not isinstance(build_rows, list):
            record(failures, f"{task}: {phase} builds.json is not a list")
            continue
        expected_build_sources = (
            {
                name: digest
                for name, digest in git_source_hashes(plan, plan["baseline_revision"]).items()
                if name.endswith(".rs") and digest is not None
            }
            if phase == "baseline"
            else {
                name: digest
                for name, digest in current_source_hashes(plan).items()
                if name.endswith(".rs")
            }
        )
        for row in build_rows:
            if row.get("source_sha256") != expected_build_sources:
                record(failures, f"{task}: {phase} build source hash map differs")
            kind = row.get("kind", row.get("probe", row.get("binary")))
            name = row.get("binary") or row.get("name")
            matched_kind = next(
                (
                    candidate_kind
                    for candidate_kind, candidate_name in plan["binaries"].items()
                    if candidate_name == name or candidate_kind == kind
                ),
                None,
            )
            if matched_kind is None:
                continue
            binary = binary_path(plan, phase, matched_kind)
            expected_hash = row.get("binary_sha256")
            expected_bytes = row.get("binary_bytes")
            if expected_hash is None:
                expected_hash = binaries.get(phase, {}).get(matched_kind, {}).get("sha256")
            if expected_bytes is None:
                expected_bytes = binaries.get(phase, {}).get(matched_kind, {}).get("bytes")
            if binary_identity(
                binary,
                expected_hash,
                expected_bytes,
                offline,
                witnesses,
                failures,
                f"{task}: build {phase} {name}",
            ) is None:
                pass

    if manifest.get("build_manifests_end") != manifest.get("build_manifests"):
        record(failures, f"{task}: build manifest start/end bindings differ")

    commands = manifest.get("commands")
    if not isinstance(commands, list):
        record(failures, f"{task}: commands is not a list")
    else:
        for command in commands:
            if command.get("exit_code") != 0:
                record(failures, f"{task}: command recorded nonzero exit: {command.get('command')}")

    raw = manifest.get("raw_sha256", {})
    if not isinstance(raw, dict):
        record(failures, f"{task}: raw_sha256 is not a map")
    else:
        for name, expected in raw.items():
            output = ROOT / name
            if not output.is_file():
                record(failures, f"{task}: missing raw output {output}")
            elif sha256_file(output) != expected:
                record(failures, f"{task}: raw output hash differs {output}")
    return manifest


def percentile(values: list[float], fraction: float) -> float:
    ordered = sorted(values)
    return ordered[max(0, math.ceil(fraction * len(ordered)) - 1)]


def stats(values: list[float]) -> dict[str, float | int]:
    if not values:
        raise AnalysisError("cannot summarize an empty timing vector")
    return {
        "n": len(values),
        "p50": statistics.median(values),
        "mean": statistics.mean(values),
        "p95": percentile(values, 0.95),
        "p99": percentile(values, 0.99),
        "minimum": min(values),
        "maximum": max(values),
    }


def percent(candidate: float, baseline: float) -> float:
    if baseline == 0:
        return 0.0 if candidate == 0 else float("inf")
    return (candidate / baseline - 1.0) * 100.0


def timing_gate(
    metric: str, baseline: float, candidate: float, plan: dict[str, Any]
) -> dict[str, Any]:
    limit = float(plan["hard_gates"]["timing_percent"])
    absolute_limit = float(plan["hard_gates"]["timing_absolute_ns"])
    delta = candidate - baseline
    change = percent(candidate, baseline)
    percent_pass = change <= limit
    warm_query = metric in set(plan["hard_gates"]["absolute_exception_metrics"])
    absolute_exception = warm_query and delta <= absolute_limit
    return {
        "baseline": baseline,
        "candidate": candidate,
        "delta_ns": delta,
        "percent": change,
        "percent_pass": percent_pass,
        "absolute_warm_query_exception": absolute_exception,
        "pass": percent_pass or absolute_exception,
    }


def native_file(
    directory: Path,
    leg: str,
    case_name: str,
    mode: str,
    failures: list[str],
) -> dict[str, Any] | None:
    path = directory / f"{leg}-{case_name}-{mode}.json"
    if not path.is_file():
        record(failures, f"native: missing {path}")
        return None
    try:
        value = read_json(path)
        if not isinstance(value, dict):
            record(failures, f"native: report is not an object {path}")
            return None
        return value
    except (OSError, json.JSONDecodeError) as error:
        record(failures, f"native: invalid {path}: {error}")
        return None


def validate_native_report(
    value: dict[str, Any], case: dict[str, Any], mode: str, plan: dict[str, Any], label: str, failures: list[str]
) -> bool:
    ok = True
    native = plan["native"]
    if value.get("schema_version") != 1 or value.get("probe") != "change-0686-xls-index-budget-retry":
        record(failures, f"{label}: probe schema identity differs")
        ok = False
    if value.get("fresh_owner_per_sample") is not True:
        record(failures, f"{label}: fresh-owner sampling is not asserted")
        ok = False
    if value.get("mode") != mode:
        record(failures, f"{label}: mode mismatch")
        ok = False
    if value.get("input_sha256") != sha256_file(ROOT / case["path"]):
        record(failures, f"{label}: fixture SHA-256 mismatch")
        ok = False
    fields = {
        "worksheet": case["sheet"],
        "row": case["row"],
        "column": case["column"],
        "max_query_index_bytes": case["budget"],
        "queries": native["queries"],
        "warmups": native["warmups"],
        "samples": native["samples"],
    }
    for field, expected in fields.items():
        if value.get(field) != expected:
            record(failures, f"{label}: {field} mismatch")
            ok = False
    records = value.get("records")
    if not isinstance(records, list) or len(records) != native["samples"]:
        record(failures, f"{label}: sample count mismatch")
        return False
    if not all(isinstance(item, dict) for item in records):
        record(failures, f"{label}: sample record is not an object")
        ok = False
    elif [item.get("sample") for item in records] != list(
        range(native["warmups"], native["warmups"] + native["samples"])
    ):
        record(failures, f"{label}: sample ordinals mismatch")
        ok = False
    for record_value in records:
        if not isinstance(record_value, dict):
            continue
        open_record = record_value.get("open")
        if not isinstance(open_record, dict) or not isinstance(open_record.get("outcome"), dict):
            record(failures, f"{label}: owner open record is invalid")
            ok = False
        elif open_record["outcome"].get("status") != "ok":
            record(failures, f"{label}: owner open failed")
            ok = False
        if not isinstance(open_record, dict) or not isinstance(open_record.get("elapsed_ns"), int) or open_record["elapsed_ns"] < 0:
            record(failures, f"{label}: owner open elapsed time is invalid")
            ok = False
        queries = record_value.get("queries")
        if not isinstance(queries, list) or len(queries) != native["queries"]:
            record(failures, f"{label}: query count mismatch")
            ok = False
            continue
        if not all(isinstance(query, dict) for query in queries):
            record(failures, f"{label}: query record is not an object")
            ok = False
            continue
        if [query.get("ordinal") for query in queries] != list(range(native["queries"])):
            record(failures, f"{label}: query ordinals mismatch")
            ok = False
        for query in queries:
            if not isinstance(query.get("elapsed_ns"), int) or query["elapsed_ns"] < 0:
                record(failures, f"{label}: query elapsed time is invalid")
                ok = False
            if not isinstance(query.get("outcome"), dict):
                record(failures, f"{label}: query outcome is invalid")
                ok = False
        if not record_value.get("all_queries_agree"):
            record(failures, f"{label}: repeated outcomes disagree")
            ok = False
        if any(not query.get("agrees_with_first") for query in queries):
            record(failures, f"{label}: query agreement flag is false")
            ok = False
    return ok


def native_metrics(report: dict[str, Any]) -> dict[str, list[float]]:
    records = report["records"]
    return {
        "open": [float(item["open"]["elapsed_ns"]) for item in records],
        "q1": [float(item["queries"][0]["elapsed_ns"]) for item in records],
        "q2": [float(item["queries"][1]["elapsed_ns"]) for item in records],
        "q3": [float(item["queries"][2]["elapsed_ns"]) for item in records],
        "q8": [float(item["queries"][7]["elapsed_ns"]) for item in records],
        "q3-to-q8-mean": [
            statistics.mean(float(query["elapsed_ns"]) for query in item["queries"][2:])
            for item in records
        ],
        "open-plus-eight": [
            float(item["open"]["elapsed_ns"])
            + sum(float(query["elapsed_ns"]) for query in item["queries"])
            for item in records
        ],
    }


def report_outcomes(report: dict[str, Any]) -> list[Any]:
    return [
        query["outcome"]
        for item in report["records"]
        for query in item["queries"]
    ]


def compare_native(plan: dict[str, Any], failures: list[str]) -> list[dict[str, Any]]:
    directory = CAPTURES / "native"
    cases = {case["case"]: case for case in plan["cases"]}
    rows: list[dict[str, Any]] = []
    for case_name, case in cases.items():
        for mode in ("owned", "file"):
            reports: dict[str, dict[str, Any]] = {}
            reports_valid = True
            for leg in plan["native"]["legs"]:
                value = native_file(directory, leg, case_name, mode, failures)
                if value is not None:
                    reports[leg] = value
                    reports_valid &= validate_native_report(
                        value, case, mode, plan, f"native/{leg}/{case_name}/{mode}", failures
                    )
            if not reports_valid or set(reports) != set(plan["native"]["legs"]):
                continue
            outcomes = {leg: report_outcomes(value) for leg, value in reports.items()}
            outcome_equal = all(value == outcomes[plan["native"]["legs"][0]] for value in outcomes.values())
            if not outcome_equal:
                record(failures, f"native/{case_name}/{mode}: semantic outcome mismatch across legs")
            expected = case["expected_status"]
            actual_statuses = {
                item["outcome"]["status"]
                for item in reports["a1"]["records"]
                for item in item["queries"]
            }
            if actual_statuses != {expected}:
                record(
                    failures,
                    f"native/{case_name}/{mode}: expected status {expected}, got {sorted(actual_statuses)}",
                )
            vectors = {leg: native_metrics(value) for leg, value in reports.items()}
            timing: dict[str, dict[str, dict[str, float | int]]] = {}
            paired: dict[str, dict[str, dict[str, Any]]] = {}
            control: dict[str, dict[str, Any]] = {}
            for metric in plan["native"]["timing_metrics"]:
                timing[metric] = {leg: stats(values[metric]) for leg, values in vectors.items()}
                paired[metric] = {}
                for pair, candidate_leg, baseline_leg in (
                    ("b1_a1", "b1", "a1"),
                    ("b2_a2", "b2", "a2"),
                ):
                    base_stats = timing[metric][baseline_leg]
                    candidate_stats = timing[metric][candidate_leg]
                    paired[metric][pair] = {
                        statistic: timing_gate(
                            metric,
                            float(base_stats[statistic]),
                            float(candidate_stats[statistic]),
                            plan,
                        )
                        for statistic in ("p50", "mean")
                    }
                    paired[metric][pair]["pass"] = all(
                        paired[metric][pair][statistic]["pass"] for statistic in ("p50", "mean")
                    )
                aa = timing_gate(
                    metric,
                    float(timing[metric]["aa1"]["p50"]),
                    float(timing[metric]["aa2"]["p50"]),
                    plan,
                )
                aa_mean = timing_gate(
                    metric,
                    float(timing[metric]["aa1"]["mean"]),
                    float(timing[metric]["aa2"]["mean"]),
                    plan,
                )
                control[metric] = {"p50": aa, "mean": aa_mean, "pass": aa["pass"] and aa_mean["pass"]}
            timing_pass = all(
                pair["pass"] for metric in paired.values() for pair in metric.values()
            )
            rows.append(
                {
                    "case": case_name,
                    "mode": mode,
                    "outcome_equal": outcome_equal,
                    "expected_status": expected,
                    "timing": timing,
                    "paired": paired,
                    "aa_control": control,
                    "timing_gate_pass": timing_pass,
                    "outcome": outcomes["a1"][0] if outcomes["a1"] else None,
                }
            )
    return rows


def parse_repeat(path: Path, failures: list[str]) -> dict[str, int] | None:
    try:
        fields = dict(item.split("=", 1) for item in path.read_text().strip().split("\t"))
        if set(fields) != {"repeats", "found", "nanos"}:
            raise ValueError(f"unexpected fields: {sorted(fields)}")
        result = {key: int(fields[key]) for key in ("repeats", "found", "nanos")}
        if any(value < 0 for value in result.values()):
            raise ValueError("negative repeat counter")
        return result
    except (OSError, ValueError, KeyError) as error:
        record(failures, f"repeat: invalid {path}: {error}")
        return None


def compare_repeat(plan: dict[str, Any], failures: list[str]) -> list[dict[str, Any]]:
    directory = CAPTURES / "repeat"
    cases = {case["case"]: case for case in plan["cases"]}
    rows: list[dict[str, Any]] = []
    for case_name in plan["repeat_cases"]:
        case = cases[case_name]
        for mode in ("owned", "file"):
            data: dict[str, list[dict[str, int]]] = {}
            complete = True
            for leg in plan["native"]["legs"]:
                values: list[dict[str, int]] = []
                for sample in range(plan["repeat"]["samples_per_leg"]):
                    path = directory / f"{case_name}-{mode}-{leg}-{sample}.tsv"
                    if not path.is_file():
                        record(failures, f"repeat: missing {path}")
                        complete = False
                        continue
                    value = parse_repeat(path, failures)
                    if value is None:
                        complete = False
                        continue
                    if value["repeats"] != plan["repeat"]["repetitions"]:
                        record(failures, f"repeat/{path.name}: repetition count mismatch")
                    values.append(value)
                data[leg] = values
            if not complete or any(len(values) != plan["repeat"]["samples_per_leg"] for values in data.values()):
                continue
            expected_found = (
                plan["repeat"]["repetitions"] if case["expected_status"] == "value" else 0
            )
            for leg, values in data.items():
                if any(value["found"] != expected_found for value in values):
                    record(failures, f"repeat/{case_name}/{mode}/{leg}: found count mismatch")
            found_equal = len({value["found"] for values in data.values() for value in values}) == 1
            if not found_equal:
                record(failures, f"repeat/{case_name}/{mode}: found count differs across legs")
            timing = {
                leg: stats([value["nanos"] / plan["repeat"]["repetitions"] for value in values])
                for leg, values in data.items()
            }
            paired: dict[str, dict[str, Any]] = {}
            for pair, candidate_leg, baseline_leg in (
                ("b1_a1", "b1", "a1"),
                ("b2_a2", "b2", "a2"),
            ):
                paired[pair] = {
                    statistic: timing_gate(
                        "repeat-q8",
                        float(timing[baseline_leg][statistic]),
                        float(timing[candidate_leg][statistic]),
                        plan,
                    )
                    for statistic in ("p50", "mean")
                }
                paired[pair]["pass"] = all(
                    paired[pair][statistic]["pass"] for statistic in ("p50", "mean")
                )
            aa = {
                statistic: timing_gate(
                    "repeat-q8",
                    float(timing["aa1"][statistic]),
                    float(timing["aa2"][statistic]),
                    plan,
                )
                for statistic in ("p50", "mean")
            }
            rows.append(
                {
                    "case": case_name,
                    "mode": mode,
                    "expected_found": expected_found,
                    "found_equal": found_equal,
                    "timing": timing,
                    "paired": paired,
                    "aa_control": {**aa, "pass": all(value["pass"] for value in aa.values())},
                    "timing_gate_pass": all(value["pass"] for value in paired.values()),
                }
            )
    return rows


def benefit_gate(
    rows: list[dict[str, Any]], plan: dict[str, Any], failures: list[str], kind: str
) -> dict[str, Any]:
    criterion = plan["benefit_gates"][kind]
    matching = [
        row
        for row in rows
        if row.get("case") == criterion["case"] and row.get("mode") == criterion["mode"]
    ]
    if len(matching) != 1:
        record(failures, f"benefit/{kind}: expected one matching gate row")
        return {"pass": False, "reason": "missing target row"}
    row = matching[0]
    checks: dict[str, Any] = {}
    minimum = float(criterion["minimum_improvement_percent"])
    if kind == "native":
        source = row["paired"][criterion["metric"]]
        for pair in ("b1_a1", "b2_a2"):
            checks[pair] = {}
            for statistic in criterion["statistics"]:
                change = float(source[pair][statistic]["percent"])
                improvement = -change
                checks[pair][statistic] = {
                    "percent_change": change,
                    "improvement_percent": improvement,
                    "pass": improvement >= minimum,
                }
    else:
        source = row["paired"]
        for pair in ("b1_a1", "b2_a2"):
            checks[pair] = {}
            for statistic in criterion["statistics"]:
                change = float(source[pair][statistic]["percent"])
                improvement = -change
                checks[pair][statistic] = {
                    "percent_change": change,
                    "improvement_percent": improvement,
                    "pass": improvement >= minimum,
                }
    passed = all(value[statistic]["pass"] for value in checks.values() for statistic in value)
    if not passed:
        record(failures, f"benefit/{kind}: 54016-late owned warm q8 improvement is below {minimum:.1f}%")
    return {"criterion": criterion, "checks": checks, "pass": passed}


ALLOC_METRICS = {
    "allocation_calls",
    "allocated_bytes",
    "deallocation_calls",
    "deallocated_bytes",
    "peak_live_delta",
    "retained_live_delta",
}


def compare_allocator(plan: dict[str, Any], failures: list[str]) -> list[dict[str, Any]]:
    directory = CAPTURES / "allocator"
    rows: list[dict[str, Any]] = []
    strict = set(plan["allocator"]["strict_operations"])
    q2_allowance = plan["allocator"]["q2_allowance"]
    cases = {case["case"]: case for case in plan["cases"]}
    for case_name, case in cases.items():
        for mode in ("owned", "file"):
            for operation in plan["allocator"]["operations"]:
                phase_values: dict[str, list[dict[str, Any]]] = {}
                complete = True
                for phase in ("baseline", "candidate"):
                    values: list[dict[str, Any]] = []
                    for repeat in range(plan["allocator"]["repeats"]):
                        path = directory / f"{phase}-{case_name}-{mode}-{operation}-{repeat}.json"
                        if not path.is_file():
                            record(failures, f"allocator: missing {path}")
                            complete = False
                            continue
                        try:
                            value = read_json(path)
                            if not isinstance(value, dict):
                                raise ValueError("report is not an object")
                            if value.get("mode") != mode or value.get("operation") != operation:
                                raise ValueError("report mode or operation differs from its filename")
                            if not ALLOC_METRICS.issubset(value):
                                raise ValueError("allocation metric fields are incomplete")
                            if any(not isinstance(value[metric], int) for metric in ALLOC_METRICS):
                                raise ValueError("allocation metric is not an integer")
                            values.append(value)
                        except (OSError, ValueError, json.JSONDecodeError) as error:
                            record(failures, f"allocator: invalid {path}: {error}")
                            complete = False
                    phase_values[phase] = values
                if not complete or any(
                    len(values) != plan["allocator"]["repeats"] for values in phase_values.values()
                ):
                    continue
                baseline = phase_values["baseline"]
                candidate = phase_values["candidate"]
                if any(value != baseline[0] for value in baseline[1:]):
                    record(failures, f"allocator/{case_name}/{mode}/{operation}: baseline repeats differ")
                if any(value != candidate[0] for value in candidate[1:]):
                    record(failures, f"allocator/{case_name}/{mode}/{operation}: candidate repeats differ")
                base = baseline[0]
                after = candidate[0]
                nonmetrics_base = {key: value for key, value in base.items() if key not in ALLOC_METRICS}
                nonmetrics_after = {key: value for key, value in after.items() if key not in ALLOC_METRICS}
                metadata_equal = nonmetrics_base == nonmetrics_after
                if not metadata_equal:
                    record(
                        failures,
                        f"allocator/{case_name}/{mode}/{operation}: outcome metadata differs",
                    )
                allowance = (
                    {metric: 0 for metric in ALLOC_METRICS}
                    if operation in strict
                    or case["budget"] == 0
                    or case["expected_status"] == "error"
                    else q2_allowance
                )
                deltas = {metric: after.get(metric, 0) - base.get(metric, 0) for metric in ALLOC_METRICS}
                metric_pass = all(
                    delta == 0
                    if operation in strict
                    or case["budget"] == 0
                    or case["expected_status"] == "error"
                    else delta <= int(allowance.get(metric, 0))
                    for metric, delta in deltas.items()
                )
                if not metric_pass:
                    record(
                        failures,
                        f"allocator/{case_name}/{mode}/{operation}: delta exceeds allowance {deltas}",
                    )
                rows.append(
                    {
                        "case": case_name,
                        "mode": mode,
                        "operation": operation,
                        "baseline": base,
                        "candidate": after,
                        "deltas": deltas,
                        "allowance": allowance,
                        "metadata_equal": metadata_equal,
                        "strict": operation in strict
                        or case["budget"] == 0
                        or case["expected_status"] == "error",
                        "gate_pass": metadata_equal and metric_pass,
                    }
                )
    return rows


def compare_budget_fence(plan: dict[str, Any], failures: list[str]) -> list[dict[str, Any]]:
    directory = CAPTURES / "budget-fence"
    rows: list[dict[str, Any]] = []
    for case in plan["budget_fence"]["cases"]:
        for budget in case["budgets"]:
            reports: dict[str, dict[str, Any]] = {}
            for phase in ("baseline", "candidate"):
                path = directory / f"{phase}-{case['case']}-{budget}.json"
                if not path.is_file():
                    record(failures, f"budget-fence: missing {path}")
                    continue
                try:
                    reports[phase] = read_json(path)
                except (OSError, json.JSONDecodeError) as error:
                    record(failures, f"budget-fence: invalid {path}: {error}")
            if set(reports) != {"baseline", "candidate"}:
                continue
            baseline = reports["baseline"]
            candidate = reports["candidate"]
            baseline_valid = validate_budget_report(
                baseline,
                case,
                budget,
                plan,
                plan["budget_fence"]["queries"],
                f"budget-fence/baseline/{case['case']}/{budget}",
                failures,
            )
            candidate_valid = validate_budget_report(
                candidate,
                case,
                budget,
                plan,
                plan["budget_fence"]["queries"],
                f"budget-fence/candidate/{case['case']}/{budget}",
                failures,
            )
            if not baseline_valid or not candidate_valid:
                continue
            base_queries = baseline.get("queries", [])
            candidate_queries = candidate.get("queries", [])
            outcomes_equal = (
                baseline.get("open_error") == candidate.get("open_error")
                and
                len(base_queries) == len(candidate_queries)
                and all(
                    base_queries[index].get("outcome") == candidate_queries[index].get("outcome")
                    for index in range(len(base_queries))
                )
                and baseline.get("all_queries_agree")
                and candidate.get("all_queries_agree")
            )
            if not outcomes_equal:
                record(failures, f"budget-fence/{case['case']}/{budget}: semantic outcome mismatch")
            base_metrics = [query.get("metrics", {}) for query in base_queries]
            candidate_metrics = [query.get("metrics", {}) for query in candidate_queries]
            metric_deltas: list[dict[str, int]] = []
            for before, after in zip(base_metrics, candidate_metrics):
                metric_deltas.append(
                    {
                        key: int(after.get(key, 0)) - int(before.get(key, 0))
                        for key in sorted(set(before) | set(after))
                    }
                )
            rows.append(
                {
                    "case": case["case"],
                    "budget": budget,
                    "outcomes_equal": outcomes_equal,
                    "baseline_outcomes": [query.get("outcome") for query in base_queries],
                    "candidate_outcomes": [query.get("outcome") for query in candidate_queries],
                    "baseline_metrics": base_metrics,
                    "candidate_metrics": candidate_metrics,
                    "candidate_minus_baseline_metrics": metric_deltas,
                    "route_changed": any(any(value != 0 for value in delta.values()) for delta in metric_deltas),
                }
            )
    return rows


def validate_budget_report(
    value: dict[str, Any],
    case: dict[str, Any],
    budget: int,
    plan: dict[str, Any],
    queries_expected: int,
    label: str,
    failures: list[str],
) -> bool:
    ok = True
    if not isinstance(value, dict):
        record(failures, f"{label}: report is not an object")
        return False
    if value.get("input_sha256") != sha256_file(ROOT / case["path"]):
        record(failures, f"{label}: fixture SHA-256 mismatch")
        ok = False
    for field, expected in (
        ("worksheet", case["sheet"]),
        ("row", case["row"]),
        ("column", case["column"]),
        ("max_query_index_bytes", budget),
        ("queries_requested", queries_expected),
    ):
        if value.get(field) != expected:
            record(failures, f"{label}: {field} mismatch")
            ok = False
    queries = value.get("queries")
    if not isinstance(queries, list) or len(queries) != queries_expected:
        record(failures, f"{label}: counted query vector length mismatch")
        return False
    if not all(isinstance(query, dict) for query in queries):
        record(failures, f"{label}: counted query record is not an object")
        return False
    if [query.get("ordinal") for query in queries] != list(range(queries_expected)):
        record(failures, f"{label}: counted query ordinals differ")
        ok = False
    if not value.get("all_queries_agree"):
        record(failures, f"{label}: counted query outcomes disagree")
        ok = False
    expected_status = case["expected_status"]
    statuses = {query.get("outcome", {}).get("status") for query in queries}
    if statuses != {expected_status}:
        record(failures, f"{label}: expected status {expected_status}, got {sorted(statuses)}")
        ok = False
    required_metrics = {"read_calls", "read_bytes", "version_calls", "len_calls"}
    open_metrics = value.get("open_metrics", {})
    if set(open_metrics) != required_metrics:
        record(failures, f"{label}: open metric fields differ")
        ok = False
    for query in queries:
        if set(query.get("metrics", {})) != required_metrics:
            record(failures, f"{label}: query metric fields differ")
            ok = False
    return ok


def compare_budget_primary(plan: dict[str, Any], failures: list[str]) -> list[dict[str, Any]]:
    directory = CAPTURES / "budget-primary"
    cases = {case["case"]: case for case in plan["cases"]}
    rows: list[dict[str, Any]] = []
    for case_name, case in cases.items():
        reports: dict[str, dict[str, Any]] = {}
        reports_valid = True
        for phase in ("baseline", "candidate"):
            path = directory / f"{phase}-{case_name}.json"
            if not path.is_file():
                record(failures, f"budget-primary: missing {path}")
                continue
            try:
                reports[phase] = read_json(path)
            except (OSError, json.JSONDecodeError) as error:
                record(failures, f"budget-primary: invalid {path}: {error}")
                continue
            reports_valid &= validate_budget_report(
                reports[phase],
                case,
                case["budget"],
                plan,
                plan["budget_primary"]["queries"],
                f"budget-primary/{phase}/{case_name}",
                failures,
            )
        if not reports_valid or set(reports) != {"baseline", "candidate"}:
            continue
        baseline = reports["baseline"]
        candidate = reports["candidate"]
        before_queries = baseline.get("queries", [])
        after_queries = candidate.get("queries", [])
        outcomes_equal = (
            baseline.get("open_error") == candidate.get("open_error")
            and len(before_queries) == len(after_queries)
            and all(
                before.get("outcome") == after.get("outcome")
                for before, after in zip(before_queries, after_queries)
            )
        )
        if not outcomes_equal:
            record(failures, f"budget-primary/{case_name}: semantic outcome mismatch")
        metric_deltas = []
        for before, after in zip(before_queries, after_queries):
            before_metrics = before.get("metrics", {})
            after_metrics = after.get("metrics", {})
            metric_deltas.append(
                {
                    key: int(after_metrics.get(key, 0)) - int(before_metrics.get(key, 0))
                    for key in sorted(set(before_metrics) | set(after_metrics))
                }
            )
        rows.append(
            {
                "case": case_name,
                "budget": case["budget"],
                "outcomes_equal": outcomes_equal,
                "baseline_open_metrics": baseline.get("open_metrics"),
                "candidate_open_metrics": candidate.get("open_metrics"),
                "baseline_query_metrics": [query.get("metrics") for query in before_queries],
                "candidate_query_metrics": [query.get("metrics") for query in after_queries],
                "candidate_minus_baseline_query_metrics": metric_deltas,
            }
        )
    return rows


def build_summary(manifests: dict[str, dict[str, Any] | None]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    seen: set[tuple[str, str]] = set()
    for manifest in manifests.values():
        if not manifest:
            continue
        for phase, value in manifest.get("build_manifests", {}).items():
            if not value:
                continue
            rows = value.get("records", [])
            if not isinstance(rows, list):
                continue
            for row in rows:
                name = row.get("binary") or row.get("name") or row.get("kind")
                key = (phase, str(name))
                if key in seen:
                    continue
                seen.add(key)
                result.append(
                    {
                        "phase": phase,
                        "binary": name,
                        "seconds": row.get("seconds"),
                        "source_sha256": row.get("source_sha256"),
                        "binary_sha256": row.get("binary_sha256"),
                    }
                )
    return result


def markdown_report(report: dict[str, Any]) -> str:
    gates = report["hard_gates"]
    lines = [
        "# Change 0723 worksheet replay checkpoint pilot",
        "",
        f"Status: **{'PASS' if report['passed'] else 'FAIL'}**",
        "",
        "The packet binds the baseline revision, current candidate source, immutable probes,",
        "raw fixture hashes, frozen binaries and per-phase build manifests before comparing",
        "fresh-owner query timings. A/A runs precede A/B/B/A when the staged commands are used.",
        "",
        "## Capture counts",
        "",
        f"- Native groups: {len(report['native'])}; expected 24.",
        f"- Repeated-query groups: {len(report['repeat'])}; expected 16, with nine 50,000-query processes per leg.",
        f"- Allocator groups: {len(report['allocator'])}; expected 96 per phase, with three repeats.",
        f"- Primary counted-I/O groups: {len(report['budget_primary'])}; expected 12.",
        f"- Budget-fence observations: {len(report['budget_fence'])}; logical charge fences include 224 and 264 bytes.",
        "",
        "## Hard gates",
        "",
        f"- Native semantic and p50/mean timing gates: {'PASS' if gates['native'] else 'FAIL'}.",
        f"- Repeated-query semantic and p50/mean timing gates: {'PASS' if gates['repeat'] else 'FAIL'}.",
        f"- Native 54016-late owned q8 benefit gate: {'PASS' if gates['native_benefit'] else 'FAIL'}.",
        f"- Repeated 54016-late owned benefit gate: {'PASS' if gates['repeat_benefit'] else 'FAIL'}.",
        f"- Allocator outcome and fixed-field bounds: {'PASS' if gates['allocator'] else 'FAIL'}.",
        f"- Primary counted-I/O semantic parity: {'PASS' if gates['budget_primary'] else 'FAIL'}.",
        f"- Budget-fence semantic parity: {'PASS' if gates['budget_fence'] else 'FAIL'}.",
        f"- Binding and command audit: {'PASS' if gates['bindings'] else 'FAIL'}.",
        "",
        "## Selected timing rows",
        "",
        "| Case | Mode | Metric | B1/A1 p50 | B2/A2 p50 | B1/A1 mean | B2/A2 mean |",
        "|---|---|---|---:|---:|---:|---:|",
    ]
    for row in report["native"]:
        if row.get("case") is None:
            continue
        for metric in ("q8", "open-plus-eight"):
            paired = row["paired"][metric]
            lines.append(
                f"| {row['case']} | {row['mode']} | {metric} | "
                f"{paired['b1_a1']['p50']['percent']:+.2f}% | "
                f"{paired['b2_a2']['p50']['percent']:+.2f}% | "
                f"{paired['b1_a1']['mean']['percent']:+.2f}% | "
                f"{paired['b2_a2']['mean']['percent']:+.2f}% |"
            )
    lines.extend(
        [
            "",
            "## Budget fences",
            "",
            "| Case | Budget | Semantic parity | Route metrics changed |",
            "|---|---:|---|---|",
        ]
    )
    for row in report["budget_fence"]:
        lines.append(
            f"| {row['case']} | {row['budget']} | "
            f"{'PASS' if row['outcomes_equal'] else 'FAIL'} | "
            f"{'yes' if row['route_changed'] else 'no'} |"
        )
    lines.extend(["", "## Failures", ""])
    if report["failures"]:
        lines.extend(f"- {failure}" for failure in report["failures"])
    else:
        lines.append("None.")
    return "\n".join(lines) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--allow-missing-build-manifests",
        action="store_true",
        help="diagnostic mode for local dry runs; the normal evidence gate requires both builds.json files",
    )
    parser.add_argument(
        "--offline",
        action="store_true",
        help="accept absent binaries only when cleanup.json carries an exact path, SHA-256 and byte-count witness",
    )
    args = parser.parse_args()
    plan = read_json(PLAN_PATH)
    failures: list[str] = []
    frozen = audit_freeze(plan, failures, args.offline)
    manifests: dict[str, dict[str, Any] | None] = {}
    required = {
        "native": ["native"],
        "repeat": ["repeat"],
        "allocator": ["allocator"],
        "budget-primary": ["budget"],
        "budget-fence": ["budget"],
    }
    for task, kinds in required.items():
        manifests[task] = audit_manifest(plan, task, kinds, failures, args.offline)
    if args.allow_missing_build_manifests:
        failures = [failure for failure in failures if "builds.json" not in failure]

    native_rows = compare_native(plan, failures)
    repeat_rows = compare_repeat(plan, failures)
    allocator_rows = compare_allocator(plan, failures)
    budget_primary_rows = compare_budget_primary(plan, failures)
    budget_rows = compare_budget_fence(plan, failures)
    native_benefit = benefit_gate(native_rows, plan, failures, "native")
    repeat_benefit = benefit_gate(repeat_rows, plan, failures, "repeat")

    native_pass = bool(native_rows) and all(
        row["outcome_equal"] and row["timing_gate_pass"] for row in native_rows
    ) and len(native_rows) == plan["native"]["groups"]
    repeat_pass = bool(repeat_rows) and all(
        row["found_equal"] and row["timing_gate_pass"] for row in repeat_rows
    ) and len(repeat_rows) == plan["repeat"]["groups"]
    allocator_pass = bool(allocator_rows) and all(row["gate_pass"] for row in allocator_rows)
    allocator_pass &= len(allocator_rows) == (
        len(plan["cases"])
        * 2
        * len(plan["allocator"]["operations"])
    )
    budget_primary_pass = bool(budget_primary_rows) and all(
        row["outcomes_equal"] for row in budget_primary_rows
    )
    budget_primary_pass &= len(budget_primary_rows) == len(plan["cases"])
    budget_pass = bool(budget_rows) and all(row["outcomes_equal"] for row in budget_rows)
    expected_budget_rows = sum(len(case["budgets"]) for case in plan["budget_fence"]["cases"])
    budget_pass &= len(budget_rows) == expected_budget_rows
    binding_failures = [
        failure
        for failure in failures
        if any(
            token in failure
            for token in (
                "manifest",
                "binding",
                "hash",
                "source",
                "probe",
                "fixture",
                "binary",
                "builds.json",
                "raw output",
                "command recorded",
            )
        )
    ]
    report = {
        "schema_version": 1,
        "packet": "change-0723",
        "generated_utc": dt.datetime.now(dt.timezone.utc).isoformat(),
        "repository_head": subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=ROOT, text=True
        ).strip(),
        "baseline_revision": plan["baseline_revision"],
        "manifests": manifests,
        "builds": build_summary(manifests),
        "native": native_rows,
        "repeat": repeat_rows,
        "allocator": allocator_rows,
        "budget_primary": budget_primary_rows,
        "budget_fence": budget_rows,
        "benefit": {"native": native_benefit, "repeat": repeat_benefit},
        "failures": failures,
        "hard_gates": {
            "bindings": not binding_failures,
            "native": native_pass,
            "repeat": repeat_pass,
            "native_benefit": native_benefit["pass"],
            "repeat_benefit": repeat_benefit["pass"],
            "allocator": allocator_pass,
            "budget_primary": budget_primary_pass,
            "budget_fence": budget_pass,
        },
        "passed": not failures
        and native_pass
        and repeat_pass
        and native_benefit["pass"]
        and repeat_benefit["pass"]
        and allocator_pass
        and budget_primary_pass
        and budget_pass,
    }
    write_json(PACKET / "analysis.json", report)
    (PACKET / "analysis.md").write_text(markdown_report(report), encoding="utf-8")
    print(
        "PASS" if report["passed"] else "FAIL",
        f"native={len(native_rows)} repeat={len(repeat_rows)} allocator={len(allocator_rows)} budget-primary={len(budget_primary_rows)} budget-fence={len(budget_rows)}",
    )
    if failures:
        for failure in failures:
            print(f"FAILURE: {failure}", file=sys.stderr)
    return 0 if report["passed"] else 1


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, subprocess.CalledProcessError, AnalysisError, json.JSONDecodeError) as error:
        print(f"analyze.py: {error}", file=sys.stderr)
        raise SystemExit(1)
