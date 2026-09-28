#!/usr/bin/env python3
"""Offline replay for the 0801 no-replay native and Callgrind lanes.

The root capture driver owns all children.  This module is intentionally a
reader: it validates the frozen inputs, receipt order, artifact identities and
probe reports, then computes deterministic summaries from retained artifacts.
It never builds, runs a probe, invokes a profiler, or uses a timing command.

Native samples are process-level elapsed values.  Their p50 is the frozen
nearest-rank statistic and the six block pairs are summarized with the frozen
median bootstrap.  Callgrind values are guest instruction/branch counters;
they are retained as mechanism diagnostics and are never presented as native
latency or allocator-call measurements.
"""

from __future__ import annotations

import argparse
import collections
import hashlib
import json
import math
import random
import re
import statistics
import sys
from pathlib import Path
from typing import Any, Iterable, Mapping, NoReturn, Sequence

import custody as c
from callgrind_parser import (
    DEFAULT_EVENTS,
    CallgrindError,
    aggregate_edges,
    function_view,
    incoming_edges,
    parse_callgrind,
    serializable_child,
    serializable_edge,
    vector_add,
    vector_equal,
    vector_mapping,
    vector_sum,
    vector_zero,
)


HERE = Path(__file__).resolve().parent
SCHEMA = "litchi.performance.0801.analysis.v1"
PLAN_SCHEMA = "litchi.performance.0801.v1"
EVENTS = tuple(DEFAULT_EVENTS)
LEGS = ("before", "after")
MODES = ("construct", "consume")
OWNER_PATTERN = "attribute_boundary_probe::{leg}_{mode}"
BOOTSTRAP_SEED = 801080
BOOTSTRAP_RESAMPLES = 10_000
BOOTSTRAP_LOW_RANK = 250
BOOTSTRAP_HIGH_RANK = 9_749
DIAGNOSTIC_RATIO = 1.05
DIAGNOSTIC_CI_LOW = 1.0
BASELINE_ITERATOR_SIZE = 120
CLONE_ADVANCES = (0, 1, 2, 3, 4, 5, 32, 33)

HEX_RE = re.compile(r"^[0-9a-f]{64}$")
REVISION_RE = re.compile(r"^[0-9a-f]{40}$")
SAFE_CASE_RE = re.compile(r"^[A-Za-z0-9_.-]+$")

EXPECTED_CASE_LABELS = (
    "distinct-0", "distinct-1", "distinct-2", "distinct-3", "distinct-4",
    "distinct-5", "distinct-8", "distinct-9", "distinct-16", "distinct-17",
    "distinct-32", "distinct-33", "distinct-64",
    "duplicate-valid-after-1", "duplicate-valid-after-2",
    "duplicate-valid-after-4", "duplicate-valid-after-5",
    "duplicate-valid-after-32", "duplicate-valid-after-33",
    "duplicate-long-quoted-after-1", "duplicate-long-quoted-after-2",
    "duplicate-long-quoted-after-33", "duplicate-long-unterminated-after-1",
    "duplicate-long-unterminated-after-2", "duplicate-long-unterminated-after-33",
    "duplicate-unquoted-after-1", "duplicate-unquoted-after-33",
    "syntax-flag-after-0", "syntax-flag-after-2", "syntax-flag-after-4",
    "syntax-flag-after-33", "syntax-unique-tail-after-0",
    "syntax-unique-tail-after-2", "syntax-unique-tail-after-4",
    "syntax-unique-tail-after-33", "syntax-equals-value-after-0",
    "syntax-equals-value-after-2", "syntax-equals-value-after-4",
    "syntax-equals-value-after-33",
)


class EvidenceError(ValueError):
    """A retained artifact is missing, stale, malformed, or contradictory."""


def fail(message: str) -> NoReturn:
    raise EvidenceError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read_json(path: Path, label: str) -> Any:
    require(path.is_file() and not path.is_symlink(), f"{label}: missing JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"{label}: invalid JSON: {error}")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1 << 20), b""):
                digest.update(chunk)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def relative(path: Path) -> str:
    try:
        return str(path.resolve().relative_to(HERE.resolve()))
    except ValueError:
        return str(path)


def is_sha(value: Any) -> bool:
    return isinstance(value, str) and HEX_RE.fullmatch(value) is not None


def finite(value: Any, label: str) -> float:
    require(isinstance(value, (int, float)) and not isinstance(value, bool),
            f"{label}: expected a number")
    result = float(value)
    require(math.isfinite(result), f"{label}: number is not finite")
    return result


def positive_int(value: Any, label: str) -> int:
    require(isinstance(value, int) and not isinstance(value, bool) and value > 0,
            f"{label}: expected a positive integer")
    return value


def nonnegative_int(value: Any, label: str) -> int:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label}: expected a non-negative integer")
    return value


def packet_path(raw: Any, label: str) -> Path:
    """Resolve an artifact path and keep packet artifacts packet-local."""

    require(isinstance(raw, str) and raw, f"{label}: path is missing")
    marker = f"/{HERE.name}/"
    if marker in raw:
        path = HERE / raw.split(marker, 1)[1]
    else:
        candidate = Path(raw)
        path = candidate if candidate.is_absolute() else HERE / candidate
    path = path.resolve()
    try:
        path.relative_to(HERE.resolve())
    except ValueError as error:
        raise EvidenceError(f"{label}: path escapes packet: {raw}") from error
    return path


def artifact(value: Any, label: str, *, allow_missing: bool = False) -> tuple[Path, dict[str, Any]]:
    require(isinstance(value, dict), f"{label}: artifact descriptor is missing")
    path = packet_path(value.get("path"), label)
    size = value.get("bytes", value.get("size"))
    nonnegative_int(size, f"{label}.bytes")
    digest = value.get("sha256", value.get("digest"))
    require(is_sha(digest), f"{label}.sha256: invalid digest")
    if not path.is_file() or path.is_symlink():
        require(allow_missing, f"{label}: artifact is missing: {path}")
        return path, {"path": relative(path), "bytes": size, "sha256": digest}
    require(path.stat().st_size == size, f"{label}: byte count changed")
    require(sha256(path) == digest, f"{label}: SHA-256 changed")
    return path, {"path": relative(path), "bytes": size, "sha256": digest}


def external_path(raw: Any, label: str) -> Path:
    require(isinstance(raw, str) and raw, f"{label}: path is missing")
    path = Path(raw)
    require(path.is_absolute(), f"{label}: path is not absolute")
    return path.resolve()


def external_artifact(value: Any, label: str, *, allow_missing: bool = False) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}: binary descriptor is missing")
    path = external_path(value.get("path"), label)
    size = value.get("bytes", value.get("size"))
    positive_int(size, f"{label}.bytes")
    digest = value.get("sha256", value.get("digest"))
    require(is_sha(digest), f"{label}.sha256: invalid digest")
    if path.is_file() and not path.is_symlink():
        require(path.stat().st_size == size, f"{label}: binary byte count changed")
        require(sha256(path) == digest, f"{label}: binary SHA-256 changed")
    else:
        require(allow_missing, f"{label}: binary is missing: {path}")
        cleanup_path = HERE / "cleanup.json"
        cleanup = read_json(cleanup_path, "cleanup witness")
        require(isinstance(cleanup, dict)
                and set(cleanup) == {"removed_binaries", "removed_failed_binaries", "removed_target_bytes",
                                     "target", "target_removed"},
                "cleanup witness schema changed")
        require(cleanup.get("target") == str(c.TARGET)
                and cleanup.get("target_removed") is True
                and not c.TARGET.exists(),
                "cleanup witness does not prove the owned target was removed")
        nonnegative_int(cleanup.get("removed_target_bytes"),
                        "cleanup removed_target_bytes")
        expected = {"path": str(path), "bytes": size, "sha256": digest}
        require(cleanup.get("removed_binaries") == [expected],
                f"{label}: cleanup binary identity differs")
        failed = read_json(HERE / "build-failed-0/relocation.json", "failed build relocation")
        require(cleanup.get("removed_failed_binaries") == [failed["binary"]],
                "cleanup failed-build binary identity differs")
    return {"path": str(path), "bytes": size, "sha256": digest}


def artifact_identity(path: Path) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing artifact: {path}")
    return {"path": relative(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def descriptor_for_existing(path: Path, label: str) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"{label}: missing: {path}")
    return artifact_identity(path)


def plan_and_cases() -> tuple[dict[str, Any], Path, list[Any], Path, list[str]]:
    plan_path = HERE / "plan.json"
    plan = read_json(plan_path, "plan")
    require(isinstance(plan, dict) and plan.get("schema") == PLAN_SCHEMA,
            "plan schema changed")
    require(plan.get("cpu") == 12, "analysis CPU changed")
    require(plan.get("legs") == list(LEGS), "leg set changed")
    require(plan.get("modes") == list(MODES), "mode set changed")
    require(plan.get("case_count") == 39, "case count changed")
    require(plan.get("distinct_counts") == [0, 1, 2, 3, 4, 5, 8, 9, 16, 17, 32, 33, 64],
            "distinct case schedule changed")
    require(plan.get("duplicate_counts") == [1, 2, 4, 5, 32, 33],
            "duplicate case schedule changed")
    require(plan.get("long_duplicate_counts") == [1, 2, 33]
            and plan.get("long_value_bytes") == 4096,
            "long duplicate case schedule changed")
    require(plan.get("syntax_prefix_counts") == [0, 2, 4, 33]
            and plan.get("syntax_tails") == ["flag", "tail=x", "=\"1\""],
            "syntax case schedule changed")

    native = plan.get("native")
    require(isinstance(native, dict), "native plan is missing")
    require(native.get("blocks") == 6 and native.get("samples") == 30
            and native.get("warmup") == 3 and native.get("iterations") == 4096,
            "native plan changed")
    require(native.get("orders") == [
        ["before", "after"], ["after", "before"], ["before", "after"],
        ["after", "before"], ["after", "before"], ["before", "after"],
    ], "native order changed")

    profiles = plan.get("profiles")
    require(isinstance(profiles, dict), "profile plan is missing")
    require(profiles.get("events") == list(EVENTS), "Callgrind event set changed")
    require(profiles.get("branch_sim") is True and profiles.get("repeats") == 2
            and profiles.get("samples") == 1 and profiles.get("warmup") == 0
            and profiles.get("iterations") == 1 and profiles.get("positive_dumps") == 1
            and profiles.get("termination_empty") is True,
            "Callgrind profile policy changed")
    require(profiles.get("orders") == [["before", "after"], ["after", "before"]],
            "profile order changed")
    require(profiles.get("owner_pattern") == OWNER_PATTERN,
            "profile owner pattern changed")

    analysis = plan.get("analysis")
    require(isinstance(analysis, dict)
            and analysis.get("bootstrap_seed") == BOOTSTRAP_SEED
            and analysis.get("resamples") == BOOTSTRAP_RESAMPLES
            and analysis.get("process_p50") == "nearest rank ceil(n/2)-1"
            and analysis.get("statistic") == "median of six paired process-p50 ratios"
            and analysis.get("zero_based_endpoints") == [BOOTSTRAP_LOW_RANK, BOOTSTRAP_HIGH_RANK],
            "analysis policy changed")
    diagnostic = analysis.get("diagnostic_regression")
    require(isinstance(diagnostic, dict)
            and diagnostic.get("ratio_above") == DIAGNOSTIC_RATIO
            and diagnostic.get("ci_low_above") == DIAGNOSTIC_CI_LOW,
            "diagnostic regression policy changed")

    cases_path = HERE / "cases.json"
    cases_value = read_json(cases_path, "cases")
    if isinstance(cases_value, dict):
        cases = cases_value.get("cases")
        require(isinstance(cases, list), "cases.json object has no cases array")
    else:
        cases = cases_value
    require(isinstance(cases, list) and len(cases) == len(EXPECTED_CASE_LABELS),
            "cases.json cardinality changed")

    labels: list[str] = []
    for index, case in enumerate(cases):
        label = case_label(case, f"cases[{index}]")
        require(SAFE_CASE_RE.fullmatch(label) is not None,
                f"cases[{index}]: unsafe case label {label!r}")
        require(label not in labels, f"cases[{index}]: duplicate case label {label!r}")
        labels.append(label)
    require(tuple(labels) == EXPECTED_CASE_LABELS,
            "cases.json order or case schedule changed")
    return plan, plan_path, cases, cases_path, labels


def case_label(case: Any, label: str) -> str:
    if isinstance(case, str):
        require(case, f"{label}: empty case label")
        return case
    require(isinstance(case, dict), f"{label}: case is not a string or object")
    for key in ("case", "name", "id", "label", "key"):
        value = case.get(key)
        if isinstance(value, str) and value:
            return value
    fail(f"{label}: object has no string case label")


def build_info() -> tuple[dict[str, Any], Path, dict[str, Any]]:
    path = HERE / "build" / "build.json"
    build = read_json(path, "build")
    require(isinstance(build, dict), "build.json is malformed")
    binary = external_artifact(build.get("binary"), "build binary", allow_missing=True)
    source = build.get("source")
    if isinstance(source, dict):
        # A build source descriptor may be retained as a packet artifact or a
        # manifest in the target.  Bind it when present, without assuming its
        # representation is identical to source.json.
        if isinstance(source.get("path"), str) and source.get("path"):
            raw = source["path"]
            if "/change-0801/" in raw or not Path(raw).is_absolute():
                _, checked = artifact(source, "build source", allow_missing=True)
                build_source = checked
            else:
                source_path = Path(raw)
                if source_path.is_file():
                    build_source = external_file_descriptor(source, "build source")
                else:
                    build_source = dict(source)
        else:
            build_source = dict(source)
    else:
        build_source = None
    return {
        "path": relative(path),
        "sha256": sha256(path),
        "binary": binary,
        "source": build_source,
    }, path, binary


def external_file_descriptor(value: Mapping[str, Any], label: str) -> dict[str, Any]:
    path = external_path(value.get("path"), label)
    size = value.get("bytes", value.get("size"))
    nonnegative_int(size, f"{label}.bytes")
    digest = value.get("sha256", value.get("digest"))
    require(is_sha(digest), f"{label}.sha256: invalid digest")
    require(path.is_file() and not path.is_symlink(), f"{label}: missing: {path}")
    require(path.stat().st_size == size and sha256(path) == digest,
            f"{label}: identity changed")
    return {"path": str(path), "bytes": size, "sha256": digest}


def frozen_source() -> dict[str, Any]:
    path = HERE / "source.json"
    value = read_json(path, "source")
    require(isinstance(value, dict), "source.json is malformed")
    revision = value.get("revision")
    require(isinstance(revision, str) and REVISION_RE.fullmatch(revision),
            "source revision is malformed")
    files = value.get("files")
    require(isinstance(files, dict) and files, "source file census is missing")
    for name, digest in files.items():
        require(isinstance(name, str) and name and is_sha(digest),
                "source file census contains an invalid digest")
    return {
        "path": "source.json",
        "sha256": sha256(path),
        "bytes": path.stat().st_size,
        "revision": revision,
        "files": dict(files),
    }


def expected_stem(block: int, case: str, mode: str, leg: str) -> str:
    return f"{block}-{case}-{mode}-{leg}"


def expected_native_jobs(plan: Mapping[str, Any], cases: Mapping[str, Any]) -> list[dict[str, Any]]:
    jobs: list[dict[str, Any]] = []
    for block, order in enumerate(plan["native"]["orders"]):
        for case in cases:
            for mode in MODES:
                for leg in order:
                    jobs.append({
                        "block": block, "case": case, "mode": mode, "leg": leg,
                        "samples": 30, "warmup": 3, "iterations": 4096,
                        "stem": expected_stem(block, case, mode, leg),
                    })
    return jobs


def expected_profile_jobs(plan: Mapping[str, Any], cases: Mapping[str, Any]) -> list[dict[str, Any]]:
    jobs: list[dict[str, Any]] = []
    for repeat, order in enumerate(plan["profiles"]["orders"]):
        for case in cases:
            for mode in MODES:
                for leg in order:
                    jobs.append({
                        "repeat": repeat, "case": case, "mode": mode, "leg": leg,
                        "samples": 1, "warmup": 0, "iterations": 1,
                        "stem": expected_stem(repeat, case, mode, leg),
                    })
    return jobs


def descriptor_from_row(value: Any, label: str) -> tuple[Path, dict[str, Any]]:
    return artifact(value, label)


def compare_binary(value: Any, expected: Mapping[str, Any], label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}: binary descriptor missing")
    digest = value.get("sha256")
    size = value.get("bytes", value.get("size"))
    require(digest == expected["sha256"] and size == expected["bytes"],
            f"{label}: binary identity differs from build")
    return external_artifact(value, label, allow_missing=True)


def check_receipt_time(row: Mapping[str, Any], previous_end: float, label: str) -> float:
    started = finite(row.get("started"), f"{label}.started")
    ended = finite(row.get("ended"), f"{label}.ended")
    require(started <= ended and previous_end <= started,
            f"{label}: receipt timing order changed")
    return ended


def command_value(command: Any, label: str) -> list[str]:
    require(isinstance(command, list) and all(isinstance(item, str) for item in command),
            f"{label}: command is not a string list")
    return list(command)


def option_value(command: Sequence[str], option: str, label: str) -> str | None:
    prefix = option + "="
    matches = [item[len(prefix):] for item in command if item.startswith(prefix)]
    if matches:
        require(len(matches) == 1, f"{label}: duplicate {option}")
        return matches[0]
    if option in command:
        index = command.index(option)
        require(index + 1 < len(command), f"{label}: missing value for {option}")
        return command[index + 1]
    return None


def command_has_value(command: Sequence[str], option: str, expected: str, label: str) -> None:
    value = option_value(command, option, label)
    require(value == expected, f"{label}: {option} changed ({value!r} != {expected!r})")


def validate_probe_command(command: Any, job: Mapping[str, Any], binary: Mapping[str, Any],
                           report_path: Path, label: str, *, profile: bool) -> list[str]:
    value = command_value(command, label)
    require("taskset" in value and "12" in value,
            f"{label}: command is not pinned to CPU 12")
    require(str(binary["path"]) in value, f"{label}: build binary is absent from command")
    command_has_value(value, "--leg", str(job["leg"]), label)
    command_has_value(value, "--case", str(job["case"]), label)
    command_has_value(value, "--mode", str(job["mode"]), label)
    command_has_value(value, "--samples", str(job["samples"]), label)
    command_has_value(value, "--warmup", str(job["warmup"]), label)
    command_has_value(value, "--iterations", str(job["iterations"]), label)
    output = option_value(value, "--output", label)
    require(output is not None and packet_path(output, f"{label} output") == report_path,
            f"{label}: output path changed")
    if profile:
        owner = OWNER_PATTERN.format(leg=job["leg"], mode=job["mode"])
        require("--tool=callgrind" in value, f"{label}: Callgrind tool is missing")
        require("--branch-sim=yes" in value, f"{label}: branch simulation is missing")
        require("--collect-atstart=no" in value, f"{label}: collection-at-start changed")
        for option in ("--toggle-collect", "--zero-before", "--dump-after"):
            command_has_value(value, option, owner, label)
        output_option = option_value(value, "--callgrind-out-file", label)
        require(output_option is not None, f"{label}: Callgrind output is missing")
        require(packet_path(output_option, f"{label} Callgrind output")
                == HERE / "profiles" / f"{job['stem']}.callgrind",
                f"{label}: Callgrind output path changed")
    return value


def report_sample_elapsed(sample: Any, label: str) -> int:
    if isinstance(sample, dict):
        value = sample.get("elapsed_ns")
        if value is None and isinstance(sample.get("metrics"), dict):
            value = sample["metrics"].get("elapsed_ns")
    else:
        value = sample
    return positive_int(value, label)


def validate_iterator_sizes(value: Any, label: str) -> dict[str, int]:
    """Validate measured helper sizes without baking in the candidate layout."""

    require(isinstance(value, dict), f"{label}: iterator size evidence is missing")
    baseline = positive_int(value.get("baseline_checked_attributes"),
                           f"{label}.baseline_checked_attributes")
    candidate = positive_int(value.get("candidate_checked_attributes"),
                             f"{label}.candidate_checked_attributes")
    require(baseline == BASELINE_ITERATOR_SIZE,
            f"{label}: baseline iterator size changed ({baseline} != "
            f"{BASELINE_ITERATOR_SIZE})")
    return {
        "baseline_checked_attributes": baseline,
        "candidate_checked_attributes": candidate,
    }


def validate_probe_report(report: Any, job: Mapping[str, Any], case: Mapping[str, Any],
                          label: str, *, expected_samples: int) -> dict[str, Any]:
    require(isinstance(report, dict), f"{label}: report is not an object")
    require(report.get("schema") == "litchi.attribute-boundary-probe.v1",
            f"{label}: probe schema changed")
    require(report.get("tool") == "attribute-boundary-probe-0801",
            f"{label}: probe tool changed")
    for key, expected in (("case", job["case"]), ("mode", job["mode"]),
                          ("leg", job["leg"]), ("warmup", job["warmup"]),
                          ("samples_requested", expected_samples),
                          ("iterations", job["iterations"])):
        require(report.get(key) == expected, f"{label}: report {key} changed")
    require(report.get("category") == case.get("category"),
            f"{label}: report category changed")
    require(report.get("attribute_count") == case.get("attribute_count"),
            f"{label}: report attribute count changed")
    require(report.get("source") == case.get("source"),
            f"{label}: report source identity changed")
    oracle = report.get("semantic_oracle")
    require(isinstance(oracle, dict) and oracle.get("all_checks_passed") is True
            and oracle.get("baseline_matches_quick_xml") is True
            and oracle.get("candidate_matches_quick_xml") is True,
            f"{label}: semantic oracle failed")
    require(oracle.get("baseline") == case.get("expected_baseline"),
            f"{label}: baseline trace differs from frozen case oracle")
    clone_checks = oracle.get("clone_checks")
    require(isinstance(clone_checks, list) and len(clone_checks) == 8,
            f"{label}: clone oracle cardinality changed")
    require([item.get("advance") for item in clone_checks] == list(CLONE_ADVANCES),
            f"{label}: clone advance schedule changed")
    require(all(isinstance(item, dict) and item.get("all_checks_passed") is not False
                and item.get("baseline_matches_quick_xml") is True
                and item.get("candidate_matches_quick_xml") is True
                and item.get("terminal_behavior_matches") is True
                for item in clone_checks),
            f"{label}: clone oracle failed")
    expected_result = report.get("expected_result")
    require(isinstance(expected_result, dict), f"{label}: expected result is missing")
    samples = report.get("samples")
    require(isinstance(samples, list) and len(samples) == expected_samples,
            f"{label}: native sample count changed")
    elapsed: list[int] = []
    for index, sample in enumerate(samples):
        require(isinstance(sample, dict), f"{label}: sample {index} is not an object")
        require(sample.get("index") == index, f"{label}: sample index changed")
        elapsed_value = report_sample_elapsed(sample, f"{label} sample {index}.elapsed_ns")
        elapsed.append(elapsed_value)
        for key in ("checksum", "accepted", "error_marker"):
            require(sample.get(key) == expected_result.get(key),
                    f"{label}: sample {index} {key} differs from expected result")
    sizes = validate_iterator_sizes(report.get("iterator_sizes"), label)
    return {"elapsed": elapsed, "count": len(elapsed), "iterator_sizes": sizes}


def iterator_size_summary(entries: Sequence[Mapping[str, Any]]) -> dict[str, Any]:
    """Require stable measured layouts for every leg and retain observations."""

    require(entries, "iterator size evidence has no reports")
    by_leg: dict[str, dict[tuple[int, int], int]] = {
        leg: {} for leg in LEGS
    }
    for entry in entries:
        identity = entry["identity"]
        leg = identity["leg"]
        require(leg in by_leg, f"iterator size report has unknown leg: {leg}")
        sizes = entry.get("iterator_sizes")
        checked = validate_iterator_sizes(sizes, f"{leg} iterator size")
        pair = (checked["baseline_checked_attributes"],
                checked["candidate_checked_attributes"])
        by_leg[leg][pair] = by_leg[leg].get(pair, 0) + 1

    result: dict[str, Any] = {
        "baseline_expected": BASELINE_ITERATOR_SIZE,
        "candidate_size_is_measured": True,
        "candidate_size_predeclared": False,
        "by_leg": {},
    }
    for leg in LEGS:
        pairs = by_leg[leg]
        require(len(pairs) == 1,
                f"{leg}: iterator sizes are not stable across reports")
        (baseline, candidate), report_count = next(iter(pairs.items()))
        result["by_leg"][leg] = {
            "baseline_checked_attributes": baseline,
            "candidate_checked_attributes": candidate,
            "reports": report_count,
        }
    all_pairs = {
        (leg, row["baseline_checked_attributes"], row["candidate_checked_attributes"])
        for leg, row in result["by_leg"].items()
    }
    require(len(all_pairs) == 2, "iterator size evidence is missing a leg")
    return result


def load_native(plan: Mapping[str, Any], cases: Mapping[str, Any], binary: Mapping[str, Any],
                source: Mapping[str, Any]) -> list[dict[str, Any]]:
    directory = HERE / "native"
    receipts_path = directory / "receipts.json"
    rows = read_json(receipts_path, "native receipts")
    jobs = expected_native_jobs(plan, cases)
    require(isinstance(rows, list) and len(rows) == len(jobs),
            f"native receipt count changed ({len(rows) if isinstance(rows, list) else 'invalid'})")
    entries: list[dict[str, Any]] = []
    seen: set[tuple[int, str, str, str]] = set()
    previous_end = float("-inf")
    for index, (row, job) in enumerate(zip(rows, jobs)):
        label = f"native/{index}/{job['stem']}"
        require(isinstance(row, dict), f"{label}: receipt is not an object")
        for key in ("block", "case", "mode", "leg"):
            require(row.get(key) == job[key], f"{label}: {key} changed")
        identity = (job["block"], job["case"], job["mode"], job["leg"])
        require(identity not in seen, f"{label}: duplicate receipt identity")
        seen.add(identity)
        require(row.get("exit_code") == 0, f"{label}: child failed")
        previous_end = check_receipt_time(row, previous_end, label)
        compare_binary(row.get("binary"), binary, f"{label} binary")
        report_path, report_descriptor = descriptor_from_row(row.get("report"), f"{label} report")
        log_path, log_descriptor = descriptor_from_row(row.get("log"), f"{label} log")
        require(report_path.name == f"{job['stem']}.json", f"{label}: report stem changed")
        require(log_path.name == f"{job['stem']}.log", f"{label}: log stem changed")
        command = validate_probe_command(row.get("command"), job, binary, report_path, label,
                                         profile=False)
        report = read_json(report_path, f"{label} report")
        outcome = validate_probe_report(report, job, cases[job["case"]], label,
                                        expected_samples=job["samples"])
        if isinstance(report.get("source"), dict) and is_sha(report["source"].get("sha256")):
            # Keep the probe's own source marker as evidence; source.json is a
            # manifest and need not have the same digest representation.
            source_marker = report["source"]["sha256"]
        else:
            source_marker = report.get("source_sha256") if is_sha(report.get("source_sha256")) else None
        entries.append({
            "identity": dict(job),
            "command": command,
            "binary": dict(binary),
            "started": row["started"],
            "ended": row["ended"],
            "report": report_descriptor,
            "log": log_descriptor,
            "report_sha256": sha256(report_path),
            "source_sha256": source_marker,
            "report_value": report,
            **outcome,
        })
    require(len(seen) == len(jobs), "native receipt identities are incomplete")
    # If reports carry a source marker, all children must agree.  This catches
    # accidental mixing of a stale baseline and candidate binary.
    markers = {entry["source_sha256"] for entry in entries if entry["source_sha256"] is not None}
    require(len(markers) <= 1, "native report source markers disagree")
    del source
    return entries


def nearest_rank(values: Sequence[int | float], percentile: float) -> int | float:
    require(values, "nearest rank received no values")
    ordered = sorted(values)
    index = max(1, math.ceil(len(ordered) * percentile / 100.0)) - 1
    return ordered[index]


def value_stats(values: Sequence[int]) -> dict[str, Any]:
    require(values, "empty native sample vector")
    return {
        "count": len(values),
        "values": list(values),
        "min": min(values),
        "p50": nearest_rank(values, 50),
        "mean": statistics.mean(values),
        "p95": nearest_rank(values, 95),
        "p99": nearest_rank(values, 99),
        "max": max(values),
    }


def spread_percent(values: Sequence[int | float]) -> float:
    require(values, "empty spread vector")
    low, high = min(values), max(values)
    if low == high:
        return 0.0
    if low == 0:
        return float("inf")
    return (high - low) * 100.0 / abs(float(low))


def bootstrap_median(values: Sequence[float]) -> dict[str, Any]:
    require(values, "cannot bootstrap an empty ratio vector")
    rng = random.Random(BOOTSTRAP_SEED)
    estimates = [
        statistics.median(values[rng.randrange(len(values))] for _ in values)
        for _ in range(BOOTSTRAP_RESAMPLES)
    ]
    estimates.sort()
    return {
        "seed": BOOTSTRAP_SEED,
        "resamples": BOOTSTRAP_RESAMPLES,
        "statistic": "median",
        "confidence": 0.95,
        "low_rank": BOOTSTRAP_LOW_RANK,
        "high_rank": BOOTSTRAP_HIGH_RANK,
        "ci_low": estimates[BOOTSTRAP_LOW_RANK],
        "ci_high": estimates[BOOTSTRAP_HIGH_RANK],
    }


def ratio_row(before: float, after: float) -> dict[str, Any]:
    if before == 0:
        equal = after == 0
        return {
            "before": before, "after": after,
            "ratio": 1.0 if equal else None,
            "change_percent": 0.0 if equal else None,
            "relative_change_defined": False,
            "zero_baseline_equal": equal,
            "zero_to_nonzero": not equal,
            "over_5_percent": not equal,
        }
    ratio = after / before
    return {
        "before": before, "after": after, "ratio": ratio,
        "change_percent": (ratio - 1.0) * 100.0,
        "relative_change_defined": True,
        "zero_baseline_equal": False,
        "zero_to_nonzero": False,
        "over_5_percent": ratio > DIAGNOSTIC_RATIO,
    }


def native_summary(entries: Sequence[Mapping[str, Any]]) -> dict[str, Any]:
    groups: dict[str, Any] = {}
    grouped: dict[tuple[str, str, str], list[Mapping[str, Any]]] = collections.defaultdict(list)
    for entry in entries:
        identity = entry["identity"]
        grouped[(identity["case"], identity["mode"], identity["leg"])].append(entry)
    spread_flags: list[dict[str, Any]] = []
    for key in sorted(grouped):
        items = sorted(grouped[key], key=lambda item: item["identity"]["block"])
        values = [value_stats(item["elapsed"]) for item in items]
        process_p50 = [stats["p50"] for stats in values]
        spread = spread_percent(process_p50)
        if spread > 5.0:
            spread_flags.append({
                "case": key[0], "mode": key[1], "leg": key[2],
                "metric": "process_p50", "spread_percent": spread,
            })
        groups["/".join(key)] = {
            "case": key[0], "mode": key[1], "leg": key[2],
            "processes": [
                {
                    "block": item["identity"]["block"],
                    "stats": stat,
                    "report": item["report"]["path"],
                    "report_sha256": item["report_sha256"],
                }
                for item, stat in zip(items, values)
            ],
            "process_p50_distribution": {
                "count": len(process_p50),
                "values": process_p50,
                "min": min(process_p50),
                "max": max(process_p50),
                "median": statistics.median(process_p50),
                "spread_percent": spread,
                "flag_over_5_percent": spread > 5.0,
            },
        }

    paired: dict[str, Any] = {}
    regression_flags: list[dict[str, Any]] = []
    for case in sorted({key[0] for key in grouped}):
        for mode in MODES:
            pairs: list[dict[str, Any]] = []
            ratios: list[float] = []
            for block in range(6):
                before = next(item for item in entries
                              if item["identity"]["case"] == case
                              and item["identity"]["mode"] == mode
                              and item["identity"]["leg"] == "before"
                              and item["identity"]["block"] == block)
                after = next(item for item in entries
                             if item["identity"]["case"] == case
                             and item["identity"]["mode"] == mode
                             and item["identity"]["leg"] == "after"
                             and item["identity"]["block"] == block)
                before_p50 = nearest_rank(before["elapsed"], 50)
                after_p50 = nearest_rank(after["elapsed"], 50)
                row = ratio_row(float(before_p50), float(after_p50))
                row["block"] = block
                pairs.append(row)
                if row["ratio"] is not None:
                    ratios.append(float(row["ratio"]))
            ratio_median = statistics.median(ratios) if ratios else None
            bootstrap = (bootstrap_median(ratios) if ratios else {
                "seed": BOOTSTRAP_SEED, "resamples": BOOTSTRAP_RESAMPLES,
                "statistic": "median", "confidence": 0.95,
                "low_rank": BOOTSTRAP_LOW_RANK, "high_rank": BOOTSTRAP_HIGH_RANK,
                "ci_low": None, "ci_high": None,
            })
            regression = (ratio_median is not None and ratio_median > DIAGNOSTIC_RATIO
                          and bootstrap["ci_low"] is not None
                          and bootstrap["ci_low"] > DIAGNOSTIC_CI_LOW)
            row = {
                "case": case, "mode": mode, "blocks": len(pairs),
                "pairs": pairs, "ratios": ratios,
                "ratio_median": ratio_median,
                "change_percent_median": (None if ratio_median is None
                                           else (ratio_median - 1.0) * 100.0),
                "bootstrap": bootstrap,
                "diagnostic_regression": regression,
                "comparison": "after/before paired within each alternating block",
            }
            paired[f"{case}/{mode}"] = row
            if regression:
                regression_flags.append({
                    "case": case, "mode": mode,
                    "ratio_median": ratio_median,
                    "change_percent_median": row["change_percent_median"],
                    "bootstrap_ci_low": bootstrap["ci_low"],
                    "bootstrap_ci_high": bootstrap["ci_high"],
                })
    return {
        "groups": groups,
        "paired_by_case_mode": paired,
        "spread_flags_over_5_percent": spread_flags,
        "diagnostic_regression_flags": regression_flags,
        "diagnostic_regression_policy": {
            "ratio_above": DIAGNOSTIC_RATIO,
            "bootstrap_ci_low_above": DIAGNOSTIC_CI_LOW,
            "flags_are_diagnostic_only": True,
        },
        "adoption_claim": False,
        "public_workflow_speedup_claim": False,
    }


def profile_artifact_map(row: Mapping[str, Any], job: Mapping[str, Any], label: str) -> dict[str, Any]:
    artifacts = row.get("artifacts")
    require(isinstance(artifacts, dict), f"{label}: artifact map is missing")
    stem = job["stem"]
    expected = {
        f"{stem}.callgrind", f"{stem}.callgrind.1", f"{stem}.log", f"{stem}.json",
    }
    require(set(artifacts) == expected,
            f"{label}: artifact set differs from {sorted(expected)}")
    checked: dict[str, dict[str, Any]] = {}
    paths: dict[str, Path] = {}
    for name in sorted(expected):
        path, descriptor = artifact(artifacts[name], f"{label} {name}")
        require(path.name == name, f"{label}: artifact name changed for {name}")
        checked[name] = descriptor
        paths[name] = path
    return {"descriptors": checked, "paths": paths}


def owner_attribution(parsed: Mapping[str, Any], owner: str) -> dict[str, Any]:
    events = tuple(parsed["header"]["events"])
    functions = parsed["functions"]
    owner_ids = sorted(
        function_id for function_id, function in functions.items()
        if function.get("name") == owner
    )
    failures: list[str] = []
    selected: dict[str, Any] = {
        "name": owner,
        "ids": owner_ids,
        "owner_qualified": False,
        "qualification_failures": failures,
    }
    if len(owner_ids) != 1:
        failures.append(f"expected exactly one exact owner function, found {owner_ids}")
        return selected
    owner_id = owner_ids[0]
    incoming = incoming_edges(parsed, owner_id)
    positive = [edge for edge in incoming if edge["calls"] > 0 and any(edge["cost"])]
    if len(positive) != 1:
        failures.append(f"expected one positive owner incoming edge, found {len(positive)}")
    if positive and positive[0]["calls"] != 1:
        failures.append(f"positive owner incoming edge calls is {positive[0]['calls']}, expected 1")
    if positive and not vector_equal(positive[0]["cost"], parsed["header"]["summary"]):
        failures.append("positive owner incoming cost differs from Callgrind summary")
    view = function_view(parsed, owner_id)
    partition = vector_add(view["self"], view["direct_children_cost"])
    if positive and not vector_equal(partition, positive[0]["cost"]):
        failures.append("owner self plus direct-child inclusive cost does not partition owner edge")
    selected.update({
        "id": owner_id,
        "view": {
            "id": owner_id,
            "name": view["name"],
            "records": int(view["records"]),
            "self": vector_mapping(view["self"], events),
            "inclusive": vector_mapping(view["inclusive"], events),
            "calls_in": int(view["calls_in"]),
            "calls_out": int(view["calls_out"]),
            "direct_children": [serializable_child(child, events)
                                 for child in view["direct_children"]],
            "direct_children_cost": vector_mapping(view["direct_children_cost"], events),
        },
        "incoming": [serializable_edge(edge, events) for edge in incoming],
        "partition": vector_mapping(partition, events),
    })
    if not failures:
        selected["owner_qualified"] = True
    return selected


SELECTED_PATTERNS: dict[str, re.Pattern[str]] = {
    "checked_iterator": re.compile(r"(?:CheckedAttributes|checked_attributes)", re.IGNORECASE),
    "iter_state": re.compile(r"(?:IterState|iter_state)", re.IGNORECASE),
    "drop": re.compile(r"(?:drop_in_place|::drop(?:$|::)|Drop)", re.IGNORECASE),
    "allocator": re.compile(
        r"(?:__rust_(?:alloc|dealloc|realloc)|alloc::|(?:^|:)alloc(?:ate|ation)?|"
        r"(?:^|:)(?:malloc|calloc|realloc|free|dealloc)(?:$|@|:))",
        re.IGNORECASE,
    ),
    "lexical_matching": re.compile(
        r"(?:lexic|memcmp|memchr|slice::cmp|str::cmp|Name::cmp|::cmp(?:$|::)|Ord::cmp)",
        re.IGNORECASE,
    ),
}


def selected_self_census(parsed: Mapping[str, Any]) -> dict[str, Any]:
    events = tuple(parsed["header"]["events"])
    result: dict[str, Any] = {}
    functions = parsed["functions"]
    for category, pattern in SELECTED_PATTERNS.items():
        ids = [function_id for function_id, function in functions.items()
               if pattern.search(function.get("name", ""))]
        ids.sort(key=lambda function_id: (-functions[function_id]["self"][0],
                                          functions[function_id].get("name", ""), function_id))
        rows = []
        for function_id in ids:
            view = function_view(parsed, function_id)
            rows.append({
                "id": function_id,
                "name": view["name"],
                "records": int(view["records"]),
                "self": vector_mapping(view["self"], events),
                "calls_in": int(view["calls_in"]),
                "calls_out": int(view["calls_out"]),
            })
        self_total = vector_sum((functions[function_id]["self"] for function_id in ids), events)
        result[category] = {
            "pattern": pattern.pattern,
            "function_count": len(rows),
            "functions": rows,
            "self_total": vector_mapping(self_total, events),
        }
    return result


def conservation(parsed: Mapping[str, Any], label: str, *, termination: bool) -> dict[str, Any]:
    events = tuple(parsed["header"]["events"])
    summary = parsed["header"]["summary"]
    totals = parsed["header"]["totals"]
    self_total = parsed["self_total"]
    require(totals is not None, f"{label}: Callgrind totals line is missing")
    equal_summary = vector_equal(self_total, summary)
    equal_totals = vector_equal(self_total, totals)
    require(equal_summary and equal_totals,
            f"{label}: self totals do not match summary and totals for every event")
    if termination:
        require(summary == vector_zero(events) and self_total == vector_zero(events)
                and totals == vector_zero(events),
                f"{label}: termination counters are not all zero")
    return {
        "self_total": vector_mapping(self_total, events),
        "summary": vector_mapping(summary, events),
        "totals": vector_mapping(totals, events),
        "self_matches_summary": equal_summary,
        "self_matches_totals": equal_totals,
        "all_counters_conserved": True,
    }


def parse_profile_pair(paths: Mapping[str, Path], job: Mapping[str, Any], label: str) -> dict[str, Any]:
    positive_path = paths[f"{job['stem']}.callgrind.1"]
    terminal_path = paths[f"{job['stem']}.callgrind"]
    positive = parse_callgrind(positive_path, expected_events=EVENTS)
    terminal = parse_callgrind(terminal_path, expected_events=EVENTS, allow_empty=True)
    require(positive["header"]["part"] == 1, f"{label}: positive part changed")
    require(positive["header"]["trigger"] == "--dump-after="
            + OWNER_PATTERN.format(leg=job["leg"], mode=job["mode"]),
            f"{label}: positive dump trigger changed")
    require(any(positive["header"]["summary"]), f"{label}: positive summary is zero")
    require(terminal["header"]["part"] == 2, f"{label}: termination part changed")
    require(terminal["header"]["trigger"] == "Program termination",
            f"{label}: termination trigger changed")
    require(terminal_path.with_name(f"{job['stem']}.callgrind.2").exists() is False,
            f"{label}: unexpected additional Callgrind dump")
    positive_conservation = conservation(positive, f"{label} positive", termination=False)
    terminal_conservation = conservation(terminal, f"{label} termination", termination=True)
    owner = OWNER_PATTERN.format(leg=job["leg"], mode=job["mode"])
    owner_info = owner_attribution(positive, owner)
    return {
        "positive": {
            "file": relative(positive_path),
            "bytes": positive_path.stat().st_size,
            "sha256": sha256(positive_path),
            "summary": vector_mapping(positive["header"]["summary"], EVENTS),
            "totals": vector_mapping(positive["header"]["totals"], EVENTS),
            "conservation": positive_conservation,
        },
        "termination": {
            "file": relative(terminal_path),
            "bytes": terminal_path.stat().st_size,
            "sha256": sha256(terminal_path),
            "summary": vector_mapping(terminal["header"]["summary"], EVENTS),
            "totals": vector_mapping(terminal["header"]["totals"], EVENTS),
            "conservation": terminal_conservation,
        },
        "owner": owner_info,
        "selected_self_function_census": selected_self_census(positive),
        "parser": positive["statistics"],
    }


def load_profiles(plan: Mapping[str, Any], cases: Mapping[str, Any], binary: Mapping[str, Any]) -> list[dict[str, Any]]:
    directory = HERE / "profiles"
    rows = read_json(directory / "receipts.json", "profile receipts")
    jobs = expected_profile_jobs(plan, cases)
    require(isinstance(rows, list) and len(rows) == len(jobs),
            f"profile receipt count changed ({len(rows) if isinstance(rows, list) else 'invalid'})")
    entries: list[dict[str, Any]] = []
    seen: set[tuple[int, str, str, str]] = set()
    previous_end = float("-inf")
    for index, (row, job) in enumerate(zip(rows, jobs)):
        label = f"profiles/{index}/{job['stem']}"
        require(isinstance(row, dict), f"{label}: receipt is not an object")
        for key in ("block", "case", "mode", "leg"):
            if key == "block":
                # The profile lane calls this coordinate repeat; accept both
                # spellings in the receipt while retaining repeat in output.
                actual = row.get("repeat", row.get("block"))
                require(actual == job["repeat"], f"{label}: repeat changed")
            else:
                require(row.get(key) == job[key], f"{label}: {key} changed")
        identity = (job["repeat"], job["case"], job["mode"], job["leg"])
        require(identity not in seen, f"{label}: duplicate receipt identity")
        seen.add(identity)
        require(row.get("exit_code") == 0, f"{label}: child failed")
        previous_end = check_receipt_time(row, previous_end, label)
        compare_binary(row.get("binary"), binary, f"{label} binary")
        checked = profile_artifact_map(row, job, label)
        report_path = checked["paths"][f"{job['stem']}.json"]
        command = validate_probe_command(row.get("command"), job, binary, report_path, label,
                                         profile=True)
        report = read_json(report_path, f"{label} report")
        require(isinstance(report, dict), f"{label}: profile report is not an object")
        outcome = validate_probe_report(report, job, cases[job["case"]], label,
                                        expected_samples=1)
        entries.append({
            "identity": dict(job),
            "command": command,
            "binary": dict(binary),
            "started": row["started"],
            "ended": row["ended"],
            "report": checked["descriptors"][f"{job['stem']}.json"],
            "artifacts": checked["descriptors"],
            "report_value": report,
            "iterator_sizes": outcome["iterator_sizes"],
            "profile": parse_profile_pair(checked["paths"], job, label),
        })
    require(len(seen) == len(jobs), "profile receipt identities are incomplete")
    return entries


def profile_counter_summary(entries: Sequence[Mapping[str, Any]]) -> dict[str, Any]:
    grouped: dict[tuple[str, str, str], list[Mapping[str, Any]]] = collections.defaultdict(list)
    for entry in entries:
        identity = entry["identity"]
        grouped[(identity["case"], identity["mode"], identity["leg"])].append(entry)
    by_case_mode_leg: dict[str, Any] = {}
    for key in sorted(grouped):
        items = sorted(grouped[key], key=lambda item: item["identity"]["repeat"])
        values = [item["profile"]["positive"]["summary"] for item in items]
        aggregate = {
            event: {
                "values": [int(value[event]) for value in values],
                "sum": sum(int(value[event]) for value in values),
                "mean": statistics.mean(int(value[event]) for value in values),
            }
            for event in EVENTS
        }
        by_case_mode_leg["/".join(key)] = {
            "case": key[0], "mode": key[1], "leg": key[2],
            "repeats": [item["identity"]["repeat"] for item in items],
            "counters": aggregate,
        }

    paired: dict[str, Any] = {}
    for case in sorted({key[0] for key in grouped}):
        for mode in MODES:
            repeat_rows: list[dict[str, Any]] = []
            for repeat in range(2):
                before = next(item for item in entries
                              if item["identity"]["case"] == case
                              and item["identity"]["mode"] == mode
                              and item["identity"]["leg"] == "before"
                              and item["identity"]["repeat"] == repeat)
                after = next(item for item in entries
                             if item["identity"]["case"] == case
                             and item["identity"]["mode"] == mode
                             and item["identity"]["leg"] == "after"
                             and item["identity"]["repeat"] == repeat)
                before_c = before["profile"]["positive"]["summary"]
                after_c = after["profile"]["positive"]["summary"]
                deltas = {event: int(after_c[event]) - int(before_c[event]) for event in EVENTS}
                counter_ratios: dict[str, Any] = {}
                for event in EVENTS:
                    row = ratio_row(float(before_c[event]), float(after_c[event]))
                    counter_ratios[event] = row
                repeat_rows.append({
                    "repeat": repeat,
                    "before": before_c,
                    "after": after_c,
                    "delta_after_minus_before": deltas,
                    "ratios": counter_ratios,
                })
            event_summary: dict[str, Any] = {}
            for event in EVENTS:
                ratios = [row["ratios"][event]["ratio"] for row in repeat_rows
                          if row["ratios"][event]["ratio"] is not None]
                deltas = [row["delta_after_minus_before"][event] for row in repeat_rows]
                event_summary[event] = {
                    "delta_values": deltas,
                    "delta_median": statistics.median(deltas),
                    "ratio_values": ratios,
                    "ratio_median": statistics.median(ratios) if ratios else None,
                }
            paired[f"{case}/{mode}"] = {
                "case": case, "mode": mode, "repeats": repeat_rows,
                "counter_changes": event_summary,
                "counters_are_guest_diagnostics": True,
            }
    return {
        "by_case_mode_leg": by_case_mode_leg,
        "paired_by_case_mode": paired,
        "events": list(EVENTS),
        "profiles": len(entries),
        "counter_latency_claim": False,
    }


def selected_profile_census(entries: Sequence[Mapping[str, Any]]) -> dict[str, Any]:
    grouped: dict[tuple[str, str, str], list[Mapping[str, Any]]] = collections.defaultdict(list)
    for entry in entries:
        identity = entry["identity"]
        grouped[(identity["case"], identity["mode"], identity["leg"])].append(entry)
    result: dict[str, Any] = {}
    for key in sorted(grouped):
        categories: dict[str, Any] = {}
        for category in SELECTED_PATTERNS:
            rows: dict[str, dict[str, Any]] = {}
            for entry in grouped[key]:
                detail = entry["profile"]["selected_self_function_census"][category]
                for function in detail["functions"]:
                    name = function["name"]
                    target = rows.setdefault(name, {
                        "name": name,
                        "repeats": 0,
                        "self": {event: 0 for event in EVENTS},
                    })
                    target["repeats"] += 1
                    for event in EVENTS:
                        target["self"][event] += int(function["self"][event])
            categories[category] = {
                "pattern": SELECTED_PATTERNS[category].pattern,
                "functions": sorted(rows.values(), key=lambda item: (-item["self"]["Ir"], item["name"])),
                "profiles": len(grouped[key]),
                "self_total": {
                    event: sum(item["self"][event] for item in rows.values()) for event in EVENTS
                },
            }
        result["/".join(key)] = {
            "case": key[0], "mode": key[1], "leg": key[2], "categories": categories,
            "lexical_census_is_name_based": True,
            "allocator_call_claim": False,
        }
    return result


def profile_summary(entries: Sequence[Mapping[str, Any]]) -> dict[str, Any]:
    failures = []
    for entry in entries:
        owner = entry["profile"]["owner"]
        if not owner["owner_qualified"]:
            identity = entry["identity"]
            failures.append({
                "repeat": identity["repeat"], "case": identity["case"],
                "mode": identity["mode"], "leg": identity["leg"],
                "failures": list(owner["qualification_failures"]),
            })
    return {
        "counts": {
            "profiles": len(entries),
            "positive_dumps": len(entries),
            "termination_dumps": len(entries),
            "owner_qualified_profiles": sum(
                1 for entry in entries if entry["profile"]["owner"]["owner_qualified"]
            ),
            "owner_qualification_failures": len(failures),
        },
        "owner_qualification_failures": failures,
        "counter_summary": profile_counter_summary(entries),
        "selected_self_function_census": selected_profile_census(entries),
        "events": list(EVENTS),
        "counters_are_guest_diagnostics": True,
        "nested_costs_overlap": True,
        "fractions_reported": False,
        "allocator_call_claim": False,
        "public_opened_presentation_expectation": False,
    }


def report_profile_entries(entries: Sequence[Mapping[str, Any]]) -> list[dict[str, Any]]:
    result = []
    for entry in entries:
        identity = entry["identity"]
        result.append({
            "repeat": identity["repeat"], "case": identity["case"],
            "mode": identity["mode"], "leg": identity["leg"],
            "report": entry["report"], "artifacts": entry["artifacts"],
            "iterator_sizes": entry["iterator_sizes"],
            "positive": entry["profile"]["positive"],
            "termination": entry["profile"]["termination"],
            "owner": entry["profile"]["owner"],
            "selected_self_function_census": entry["profile"]["selected_self_function_census"],
            "parser": entry["profile"]["parser"],
        })
    return result


def report_native_entries(entries: Sequence[Mapping[str, Any]]) -> list[dict[str, Any]]:
    return [
        {
            "block": item["identity"]["block"],
            "case": item["identity"]["case"],
            "mode": item["identity"]["mode"],
            "leg": item["identity"]["leg"],
            "report": item["report"],
            "report_sha256": item["report_sha256"],
            "iterator_sizes": item["iterator_sizes"],
            "elapsed": item["elapsed"],
        }
        for item in entries
    ]


def analyze() -> dict[str, Any]:
    plan, plan_path, cases, cases_path, labels = plan_and_cases()
    case_map = {case_label(case, f"cases[{index}]"): case
                for index, case in enumerate(cases)}
    source = frozen_source()
    build, build_path, binary = build_info()
    native_entries = load_native(plan, case_map, binary, source)
    profile_entries = load_profiles(plan, case_map, binary)
    require(len(native_entries) == 936, "native child cardinality changed")
    require(sum(len(item["elapsed"]) for item in native_entries) == 28_080,
            "native sample cardinality changed")
    require(len(profile_entries) == 312, "profile child cardinality changed")
    sizes = iterator_size_summary([*native_entries, *profile_entries])

    native = native_summary(native_entries)
    profiles = profile_summary(profile_entries)
    cases_identity = artifact_identity(cases_path)
    plan_identity = artifact_identity(plan_path)
    return {
        "schema": SCHEMA,
        "packet": "change-0801",
        "plan": {
            "path": relative(plan_path), "sha256": plan_identity["sha256"],
            "bytes": plan_identity["bytes"], "schema": plan["schema"],
            "cpu": plan["cpu"], "events": list(EVENTS),
        },
        "cases": {
            "path": relative(cases_path), "sha256": cases_identity["sha256"],
            "bytes": cases_identity["bytes"], "count": len(cases),
            "labels": list(labels),
        },
        "source": source,
        "build": build,
        "iterator_sizes": sizes,
        "counts": {
            "native_children": len(native_entries),
            "native_samples": sum(len(item["elapsed"]) for item in native_entries),
            "profile_children": len(profile_entries),
            "reports": len(native_entries) + len(profile_entries),
            "callgrind_positive_dumps": len(profile_entries),
            "callgrind_termination_dumps": len(profile_entries),
        },
        "bootstrap": {
            "seed": BOOTSTRAP_SEED, "resamples": BOOTSTRAP_RESAMPLES,
            "statistic": "median", "confidence": 0.95,
            "low_rank": BOOTSTRAP_LOW_RANK, "high_rank": BOOTSTRAP_HIGH_RANK,
        },
        "native": {
            "blocks": 6, "samples": 30, "warmup": 3, "iterations": 4096,
            "process_p50": "nearest rank ceil(n/2)-1",
            "analysis": native,
            "receipts": report_native_entries(native_entries),
        },
        "profiles": {
            "repeats": 2, "samples": 1, "warmup": 0, "iterations": 1,
            "events": list(EVENTS), "owner_pattern": OWNER_PATTERN,
            "analysis": profiles,
            "receipts": report_profile_entries(profile_entries),
        },
        "claims": [
            "Native elapsed values are paired process diagnostics for the frozen hot micro-input; no public workflow speedup is claimed.",
            "Process p50 uses the frozen nearest-rank rule and six block pairs use the frozen deterministic median bootstrap.",
            "Callgrind Ir, branch, and branch-misprediction values are guest-counter mechanism diagnostics, not native latency.",
            "All five Callgrind counters are conserved from parsed self rows to each positive summary and totals line; termination dumps are zero.",
            "Exact owner qualification failures are retained and never converted into success.",
            "Selected allocator symbols are a lexical function-name census, not allocator API call counts.",
            "Inclusive descendant values overlap; no phase fractions are reported.",
            "No production-adoption decision is made by this analyzer.",
        ],
        "verification": {
            "frozen_plan_checked": True,
            "frozen_cases_checked": True,
            "source_manifest_checked": True,
            "build_binary_identity_checked": True,
            "native_receipt_order_checked": True,
            "native_all_samples_checked": True,
            "profile_receipt_order_checked": True,
            "callgrind_self_summary_totals_checked": True,
            "callgrind_termination_zero_checked": True,
            "owner_edge_one_call_and_partition_checked": True,
            "owner_qualification_failures_retained": True,
            "public_opened_presentation_expectation": False,
            "allocator_api_call_claim": False,
        },
    }


def markdown(report: Mapping[str, Any]) -> str:
    native = report["native"]["analysis"]
    profiles = report["profiles"]["analysis"]
    lines = [
        "# 0801 attribute-boundary native and Callgrind analysis",
        "",
        "This document is an offline replay of 936 native children and 312 "
        "Callgrind children. Native process p50 values use the frozen nearest-rank "
        "rule; Callgrind counters are guest diagnostics and do not measure native latency.",
        "",
        "## Native paired diagnostics",
        "",
        "| Case | Mode | Median after/before p50 | Change | CI low | CI high | Diagnostic regression |",
        "| --- | --- | ---: | ---: | ---: | ---: | --- |",
    ]
    for key in sorted(native["paired_by_case_mode"]):
        row = native["paired_by_case_mode"][key]
        boot = row["bootstrap"]
        ratio = row["ratio_median"]
        change = row["change_percent_median"]
        lines.append(
            f"| `{row['case']}` | `{row['mode']}` | "
            f"{('n/a' if ratio is None else f'{ratio:.6f}')} | "
            f"{('n/a' if change is None else f'{change:.3f}%')} | "
            f"{('n/a' if boot['ci_low'] is None else f'{boot['ci_low']:.6f}')} | "
            f"{('n/a' if boot['ci_high'] is None else f'{boot['ci_high']:.6f}')} | "
            f"{'yes' if row['diagnostic_regression'] else 'no'} |"
        )
    lines += [
        "",
        f"Diagnostic flags: {len(native['diagnostic_regression_flags'])}. "
        "A flag means the median ratio is above 1.05 and the bootstrap lower "
        "endpoint is above 1.0; it is not an adoption decision.",
        "",
        f"Process p50 spread flags above 5%: {len(native['spread_flags_over_5_percent'])}.",
        "",
        "## Callgrind profile diagnostics",
        "",
        f"The parser checked {profiles['counts']['profiles']} positive dumps and "
        f"{profiles['counts']['termination_dumps']} empty termination dumps over "
        f"the events `{', '.join(EVENTS)}`.",
        "",
        f"Exact owner qualification succeeded for "
        f"{profiles['counts']['owner_qualified_profiles']} profiles and failed for "
        f"{profiles['counts']['owner_qualification_failures']}. Failures remain in "
        "the JSON evidence and are not treated as successful attribution.",
        "",
        "The selected self-function census covers checked-iterator, IterState, "
        "drop, allocator-name, and lexical-matching symbols. The allocator group "
        "is lexical name matching only; it does not count allocator API calls. "
        "Inclusive graph costs overlap, so no fractions are reported.",
        "",
        "The full receipt identities, raw Callgrind conservation records, owner "
        "qualification state, and selected symbol rows are retained in `analysis.json`.",
        "",
    ]
    return "\n".join(lines)


def write_or_check(path: Path, content: str, check: bool) -> None:
    if check:
        require(path.is_file() and not path.is_symlink(), f"missing expected output: {path}")
        require(path.read_text(encoding="utf-8") == content,
                f"replayed output differs: {relative(path)}")
    else:
        path.write_text(content, encoding="utf-8")


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    modes = parser.add_mutually_exclusive_group(required=True)
    modes.add_argument("--write", action="store_true")
    modes.add_argument("--check", action="store_true")
    args = parser.parse_args(argv)
    try:
        result = analyze()
        encoded = json.dumps(result, indent=2, sort_keys=True) + "\n"
        summary = markdown(result)
        write_or_check(HERE / "analysis.json", encoded, args.check)
        write_or_check(HERE / "summary.md", summary, args.check)
        print(json.dumps({
            "mode": "check" if args.check else "write",
            "native_children": result["counts"]["native_children"],
            "native_samples": result["counts"]["native_samples"],
            "profile_children": result["counts"]["profile_children"],
        }, sort_keys=True), flush=True)
        return 0
    except (EvidenceError, CallgrindError, OSError, KeyError, TypeError, ValueError) as error:
        print(f"analyze.py: error: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
