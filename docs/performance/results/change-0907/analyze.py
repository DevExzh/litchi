#!/usr/bin/env python3
"""Fail-closed offline analysis of the 0907 PPTX multi-slide capture.

The packet contains four raw ``litchi-perf-baseline`` reports named
``eager-a.json``, ``source-a.json``, ``source-b.json``, and ``eager-b.json``.
This reader checks the capture plan and its root receipts, validates the raw
report cardinality and semantic evidence, and emits descriptive per-leg and
pooled latency summaries.  It never starts Cargo, the workload, a profiler,
or a build command.

The historical identity comparison is deliberately informational.  Current
corpus and output hashes are read from the four fresh reports; no historical
hash is used as a current expected value.  The two eager/source pairs are
summarised descriptively and no confidence interval or speedup claim is
computed from them.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
import math
from pathlib import Path
import re
import statistics as _statistics
import sys
from typing import Any, Mapping, Sequence


HERE = Path(__file__).resolve().parent
SCHEMA = "litchi.performance.0907.pptx-multi-slide-analysis.v1"
LABELS = ("eager-a", "source-a", "source-b", "eager-b")
LEG_KIND = {
    "eager-a": "eager",
    "source-a": "source_backed",
    "source-b": "source_backed",
    "eager-b": "eager",
}
LEG_CASE = {
    "eager-a": "pptx_eager_multi_slide_batch_edit_save",
    "source-a": "pptx_source_backed_multi_slide_batch_edit_save",
    "source-b": "pptx_source_backed_multi_slide_batch_edit_save",
    "eager-b": "pptx_eager_multi_slide_batch_edit_save",
}
EXPECTED_LEGS = [[label, LEG_KIND[label]] for label in LABELS]
SAMPLES = 200
WARMUP = 30
TOTAL_SAMPLES = len(LABELS) * SAMPLES
SINK_CEILING = 64 * 1024
EXPECTED_MATERIALIZATIONS = {"eager": 229, "source_backed": 9}
HEX_RE = re.compile(r"^[0-9a-f]{64}$")

# These are only historical report locations.  Their values are loaded from
# the retained reports at analysis time and are never treated as current
# expected corpus or output identities.
HISTORICAL_REPORTS = (
    "docs/performance/results/abba-pptx-multi-slide-batch-before-a.json",
    "docs/performance/results/abba-pptx-multi-slide-batch-before-b.json",
    "docs/performance/results/abba-pptx-multi-slide-batch-after-a.json",
    "docs/performance/results/abba-pptx-multi-slide-batch-after-b.json",
)


class EvidenceError(ValueError):
    """A retained artifact is missing, malformed, or contradictory."""


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
            for block in iter(lambda: stream.read(1 << 20), b""):
                digest.update(block)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def relative(path: Path, packet: Path) -> str:
    try:
        return str(path.resolve().relative_to(packet.resolve()))
    except ValueError as error:
        raise EvidenceError(f"path escapes packet: {path}") from error


def canonical(value: Any) -> str:
    return json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=False)


def descriptor(path: Path, packet: Path) -> dict[str, Any]:
    return {
        "path": relative(path, packet),
        "bytes": path.stat().st_size,
        "sha256": sha256(path),
    }


def valid_sha(value: Any, label: str) -> str:
    require(isinstance(value, str) and HEX_RE.fullmatch(value) is not None,
            f"{label}: expected a lowercase SHA-256 digest")
    return value


def integer(value: Any, label: str, *, nonnegative: bool = True) -> int:
    require(type(value) is int, f"{label}: expected an integer")
    if nonnegative:
        require(value >= 0, f"{label}: expected a non-negative integer")
    else:
        require(value > 0, f"{label}: expected a positive integer")
    return value


def finite_number(value: Any, label: str) -> float:
    require(isinstance(value, (int, float)) and not isinstance(value, bool),
            f"{label}: expected a number")
    result = float(value)
    require(math.isfinite(result), f"{label}: expected a finite number")
    return result


def midpoint(left: int, right: int) -> int:
    """Match the Rust harness's overflow-safe integer midpoint."""

    return left // 2 + right // 2 + (left % 2 + right % 2) // 2


def nearest_rank(samples: Sequence[int], percentile: int) -> int:
    require(samples, "nearest-rank statistic received no samples")
    index = min((percentile * len(samples) + 99) // 100 - 1, len(samples) - 1)
    return samples[index]


def summary(values: Sequence[int], label: str) -> dict[str, Any]:
    require(values, f"{label}: no values")
    clean = [integer(value, f"{label}[{index}]") for index, value in enumerate(values)]
    ordered = sorted(clean)
    count = len(ordered)
    mean = _statistics.fmean(ordered)
    return {
        "samples": count,
        "min": ordered[0],
        "p50": midpoint(ordered[(count - 1) // 2], ordered[count // 2]),
        "p95": nearest_rank(ordered, 95),
        "p99": nearest_rank(ordered, 99),
        "max": ordered[-1],
        "mean": mean,
    }


def check_permutation(value: Any, label: str) -> list[int]:
    require(isinstance(value, list) and len(value) == SAMPLES,
            f"{label}: expected {SAMPLES} sample indices")
    result = [integer(item, f"{label}[{index}]") for index, item in enumerate(value)]
    require(sorted(result) == list(range(SAMPLES)), f"{label}: not a permutation of 0..199")
    return result


def check_reported_statistics(elapsed: Mapping[str, Any], values: Sequence[int], label: str) -> dict[str, Any]:
    require(elapsed.get("unit") == "ns", f"{label}: elapsed unit is not ns")
    require(isinstance(elapsed.get("samples"), list), f"{label}: elapsed samples are missing")
    require(len(values) == SAMPLES, f"{label}: expected {SAMPLES} elapsed samples")
    require(all(value > 0 for value in values), f"{label}: latency samples must be positive")
    require(list(values) == sorted(values), f"{label}: elapsed samples are not sorted")
    for key, expected in (
        ("min", values[0]),
        ("p50", midpoint(values[(SAMPLES - 1) // 2], values[SAMPLES // 2])),
        ("p95", nearest_rank(values, 95)),
        ("p99", nearest_rank(values, 99)),
        ("max", values[-1]),
    ):
        require(elapsed.get(key) == expected, f"{label}: reported {key} differs from raw samples")
    for key in ("mean", "standard_deviation"):
        finite_number(elapsed.get(key), f"{label}.{key}")
    interval = elapsed.get("confidence_interval_95")
    require(isinstance(interval, dict), f"{label}: confidence interval is missing")
    finite_number(interval.get("lower"), f"{label}.confidence_interval_95.lower")
    finite_number(interval.get("upper"), f"{label}.confidence_interval_95.upper")
    return {
        "samples": SAMPLES,
        "sample_order_is_permutation": True,
        "min": values[0],
        "p50": midpoint(values[(SAMPLES - 1) // 2], values[SAMPLES // 2]),
        "p95": nearest_rank(values, 95),
        "p99": nearest_rank(values, 99),
        "max": values[-1],
        "mean": _statistics.fmean(values),
    }


def check_vector(value: Any, label: str, *, positive: bool = False) -> list[int]:
    require(isinstance(value, list) and len(value) == SAMPLES,
            f"{label}: expected {SAMPLES} values")
    result = [integer(item, f"{label}[{index}]") for index, item in enumerate(value)]
    if positive:
        require(all(item > 0 for item in result), f"{label}: expected every value to be positive")
    return result


def check_sink(result: Mapping[str, Any], label: str) -> dict[str, Any]:
    sink = result.get("sink")
    require(isinstance(sink, dict), f"{label}: sink summary is missing")
    accepted = integer(sink.get("accepted_bytes"), f"{label}.sink.accepted_bytes")
    calls = integer(sink.get("write_calls"), f"{label}.sink.write_calls")
    largest = integer(sink.get("largest_write"), f"{label}.sink.largest_write")
    require(calls > 0, f"{label}: sink has no writes")
    require(largest <= SINK_CEILING, f"{label}: sink largest write exceeds 64 KiB")
    buckets = sink.get("write_size_buckets")
    over = None
    if buckets is not None:
        require(isinstance(buckets, dict), f"{label}.sink.write_size_buckets is malformed")
        over = integer(buckets.get("bytes_over_65536"),
                       f"{label}.sink.write_size_buckets.bytes_over_65536")
        require(over == 0, f"{label}: sink reports writes over 64 KiB")
    return {
        "accepted_bytes": accepted,
        "write_calls": calls,
        "largest_write": largest,
        "ceiling_bytes": SINK_CEILING,
        "within_ceiling": True,
        "bytes_over_65536": over,
    }


def check_source(result: Mapping[str, Any], label: str, kind: str) -> dict[str, Any]:
    source = result.get("source")
    require(isinstance(source, dict), f"{label}: source summary is missing")
    fields = (
        "read_calls",
        "read_bytes",
        "ordinary_payload_read_calls",
        "ordinary_payload_read_bytes",
        "max_in_flight_reads",
    )
    summaries: dict[str, Any] = {}
    for field in fields:
        values = check_vector(source.get(field), f"{label}.source.{field}")
        summaries[field] = summary(values, f"{label}.source.{field}")
    for field in ("read_calls", "read_bytes", "ordinary_payload_read_calls", "ordinary_payload_read_bytes"):
        # The current selector contract must demonstrate source ingress for
        # every measured sample.  max_in_flight_reads may legitimately be 0
        # for a selector that does not overlap reads, so it is not gated here.
        values = check_vector(source.get(field), f"{label}.source.{field}")
        require(all(value > 0 for value in values),
                f"{label}.source.{field}: no source read in at least one sample")

    materializations = check_vector(
        source.get("ordinary_payload_materializations"),
        f"{label}.source.ordinary_payload_materializations",
    )
    expected = EXPECTED_MATERIALIZATIONS[kind]
    require(all(value == expected for value in materializations),
            f"{label}: materialization count is not exactly {expected} per sample")
    summaries["ordinary_payload_materializations"] = {
        "expected_per_sample": expected,
        "all_expected": True,
        "unique": sorted(set(materializations)),
        "summary": summary(materializations, f"{label}.source.ordinary_payload_materializations"),
    }
    return summaries


def report_file(packet: Path, label: str) -> tuple[Path, dict[str, Any]]:
    path = packet / f"{label}.json"
    report = read_json(path, label)
    require(isinstance(report, dict), f"{label}: report is not an object")
    return path, report


def raw_report(packet: Path, label: str) -> dict[str, Any]:
    path, report = report_file(packet, label)
    require(report.get("schema_version") == 1, f"{label}: unsupported report schema")
    tool = report.get("tool")
    require(isinstance(tool, dict) and tool.get("name") == "litchi-perf-baseline",
            f"{label}: unexpected harness tool")
    configuration = report.get("configuration")
    require(isinstance(configuration, dict), f"{label}: configuration is missing")
    require(configuration.get("samples_per_case") == SAMPLES,
            f"{label}: samples_per_case is not {SAMPLES}")
    require(configuration.get("warmup_iterations_per_case") == WARMUP,
            f"{label}: warmup_iterations_per_case is not {WARMUP}")
    require(configuration.get("cases") == [LEG_CASE[label]],
            f"{label}: selected case differs")

    results = report.get("results")
    require(isinstance(results, list) and len(results) == 1,
            f"{label}: expected exactly one result")
    result = results[0]
    require(isinstance(result, dict) and result.get("case") == LEG_CASE[label],
            f"{label}: result case differs")
    corpus = result.get("corpus")
    require(isinstance(corpus, dict), f"{label}: corpus manifest is missing")
    require(corpus.get("generator") == "litchi-pptx-source-edit-media-v1",
            f"{label}: unexpected PPTX corpus generator")
    require(corpus.get("package_format") == "PPTX/OPC/ZIP",
            f"{label}: unexpected package format")
    require(corpus.get("entry_count") == 229, f"{label}: corpus entry count is not 229")
    require(corpus.get("archive_member_count") == 445,
            f"{label}: corpus archive member count is not 445")
    require(corpus.get("shape") == "media-rich", f"{label}: corpus shape is not media-rich")
    require(corpus.get("compression") == "deflate", f"{label}: corpus compression is not deflate")
    archive_sha = valid_sha(corpus.get("archive_sha256"), f"{label}.corpus.archive_sha256")
    output_sha = valid_sha(result.get("output_sha256"), f"{label}.output_sha256")

    elapsed = result.get("elapsed_ns")
    require(isinstance(elapsed, dict), f"{label}: elapsed_ns is missing")
    elapsed_values = check_vector(elapsed.get("samples"), f"{label}.elapsed_ns.samples", positive=True)
    sample_order = check_permutation(elapsed.get("sample_order"), f"{label}.elapsed_ns.sample_order")
    elapsed_summary = check_reported_statistics(elapsed, elapsed_values, f"{label}.elapsed_ns")
    sink_summary = check_sink(result, label)
    source_summary = check_source(result, label, LEG_KIND[label])

    binary_identity = report.get("binary_identity")
    require(isinstance(binary_identity, dict), f"{label}: binary_identity is missing")
    binary_sha = valid_sha(binary_identity.get("binary_sha256"), f"{label}.binary_identity.binary_sha256")
    binary_bytes = integer(binary_identity.get("binary_bytes"), f"{label}.binary_identity.binary_bytes")
    require(binary_bytes > 0, f"{label}.binary_identity.binary_bytes: expected non-empty binary")
    require(binary_identity.get("executable") is True,
            f"{label}.binary_identity.executable: expected true")

    return {
        "label": label,
        "kind": LEG_KIND[label],
        "case": LEG_CASE[label],
        "report": descriptor(path, packet),
        "tool": {key: tool[key] for key in ("name", "profile", "target_os", "target_arch") if key in tool},
        "configuration": {
            "samples_per_case": SAMPLES,
            "warmup_iterations_per_case": WARMUP,
            "cases": [LEG_CASE[label]],
        },
        "corpus": corpus,
        "corpus_archive_sha256": archive_sha,
        "output_sha256": output_sha,
        "binary_sha256_in_report": binary_sha,
        "binary_bytes_in_report": binary_bytes,
        "elapsed_ns": elapsed_summary,
        "sample_order": sample_order,
        "sink": sink_summary,
        "source": source_summary,
    }


def path_from_descriptor(packet: Path, raw: Any, label: str) -> Path:
    require(isinstance(raw, dict), f"{label}: descriptor is malformed")
    raw_path = raw.get("path")
    require(isinstance(raw_path, str) and raw_path, f"{label}: descriptor path is missing")
    path = Path(raw_path)
    if not path.is_absolute():
        path = packet / path
    return path.resolve()


def verify_file_descriptor(packet: Path, raw: Any, expected_name: str, label: str) -> dict[str, Any]:
    path = path_from_descriptor(packet, raw, label)
    require(path.name == expected_name, f"{label}: descriptor names {path.name}, expected {expected_name}")
    # Root receipts retain absolute paths, while replay and analysis output
    # must remain stable.  The file itself must still be packet-local.
    try:
        path.relative_to(packet.resolve())
    except ValueError as error:
        raise EvidenceError(f"{label}: descriptor escapes packet: {path}") from error
    require(path.is_file() and not path.is_symlink(), f"{label}: described file is missing")
    bytes_expected = integer(raw.get("bytes"), f"{label}.bytes")
    digest_expected = valid_sha(raw.get("sha256"), f"{label}.sha256")
    require(path.stat().st_size == bytes_expected, f"{label}: byte count changed")
    require(sha256(path) == digest_expected, f"{label}: file SHA-256 changed")
    return {
        "path": relative(path, packet),
        "bytes": bytes_expected,
        "sha256": digest_expected,
    }


def check_argv(argv: Any, label: str, binary_path: str, case_label: str) -> None:
    require(isinstance(argv, list) and all(isinstance(item, str) for item in argv),
            f"{label}: argv is malformed")
    require(len(argv) >= 4 and argv[:3] == ["taskset", "-c", "12"],
            f"{label}: argv does not pin CPU 12 with taskset")
    require(argv[3] == binary_path, f"{label}: argv binary path differs")

    def option_value(option: str) -> str:
        require(option in argv, f"{label}: argv omits {option}")
        index = argv.index(option)
        require(index + 1 < len(argv), f"{label}: argv option {option} has no value")
        return argv[index + 1]

    require(option_value("--workers") == "1", f"{label}: workers argv differs")
    require(option_value("--warmup") == str(WARMUP), f"{label}: warmup argv differs")
    require(option_value("--samples") == str(SAMPLES), f"{label}: samples argv differs")
    require(option_value("--case") == LEG_CASE[case_label], f"{label}: case argv differs")
    output = option_value("--json")
    require(Path(output).name == f"{case_label}.json", f"{label}: JSON output argv differs")


def check_capture_custody(packet: Path) -> dict[str, Any]:
    plan = read_json(packet / "capture-plan.json", "capture-plan")
    require(isinstance(plan, dict), "capture-plan: expected an object")
    require(plan.get("legs") == EXPECTED_LEGS, "capture-plan: leg order differs")
    require(plan.get("warmup") == WARMUP, "capture-plan: warmup differs")
    require(plan.get("samples_per_leg") == SAMPLES, "capture-plan: sample count differs")
    require(plan.get("cpu") == 12, "capture-plan: CPU differs")
    environment = plan.get("environment")
    require(isinstance(environment, dict), "capture-plan: environment is missing")

    binary = read_json(packet / "binary.json", "binary")
    require(isinstance(binary, dict), "binary: expected an object")
    binary_sha = valid_sha(binary.get("sha256"), "binary.sha256")
    binary_bytes = integer(binary.get("bytes"), "binary.bytes")
    require(binary_bytes > 0, "binary.bytes: expected a non-empty binary")
    if "executable" in binary:
        require(binary.get("executable") is True, "binary: executable flag is not true")
    require(plan.get("binary") == binary, "capture-plan: binary descriptor differs")

    legs: dict[str, Any] = {}
    previous_finish: _datetime.datetime | None = None
    for label in LABELS:
        start = read_json(packet / f"{label}-start.json", f"{label}-start")
        run = read_json(packet / f"{label}-run.json", f"{label}-run")
        require(isinstance(start, dict) and isinstance(run, dict), f"{label}: receipt is malformed")
        binary_path = binary.get("path")
        require(isinstance(binary_path, str) and binary_path,
                "binary.path: expected an executable path")
        check_argv(start.get("argv"), label, binary_path, label)
        check_argv(run.get("argv"), f"{label}-run", binary_path, label)
        require(start.get("argv") == run.get("argv"), f"{label}: start/run argv differs")
        require(start.get("binary") == binary, f"{label}-start: binary differs")
        require(start.get("environment") == environment, f"{label}-start: environment differs")
        require(run.get("exit_code") == 0, f"{label}-run: workload did not exit zero")
        require(run.get("binary") == binary, f"{label}-run: binary differs")
        require(run.get("environment") == environment, f"{label}-run: environment differs")
        require(isinstance(start.get("started_utc"), str), f"{label}-start: started_utc is missing")
        require(start.get("started_utc") == run.get("started_utc"),
                f"{label}: start/run started_utc differs")
        require(isinstance(start.get("cwd"), str) and start.get("cwd"),
                f"{label}-start: cwd is missing")
        require(start.get("cwd") == run.get("cwd"), f"{label}: start/run cwd differs")

        result_descriptor = verify_file_descriptor(
            packet, run.get("result"), f"{label}.json", f"{label}-run.result"
        )
        stdout_descriptor = verify_file_descriptor(
            packet, run.get("stdout"), f"{label}.stdout", f"{label}-run.stdout"
        )
        stderr_descriptor = verify_file_descriptor(
            packet, run.get("stderr"), f"{label}.stderr", f"{label}-run.stderr"
        )
        report_descriptor = descriptor(packet / f"{label}.json", packet)
        require(result_descriptor == report_descriptor, f"{label}: receipt result does not bind raw report")
        for key in ("started_utc", "finished_utc"):
            require(isinstance(run.get(key), str), f"{label}-run: {key} is missing")
        try:
            started = _datetime.datetime.fromisoformat(str(run["started_utc"]))
            finished = _datetime.datetime.fromisoformat(str(run["finished_utc"]))
        except ValueError as error:
            raise EvidenceError(f"{label}-run: invalid UTC timestamp") from error
        require(finished >= started, f"{label}-run: finish precedes start")
        if previous_finish is not None:
            require(started >= previous_finish, f"{label}: legs overlap or are out of order")
        previous_finish = finished
        legs[label] = {
            "start_receipt": relative(packet / f"{label}-start.json", packet),
            "run_receipt": relative(packet / f"{label}-run.json", packet),
            "result": result_descriptor,
            "stdout": stdout_descriptor,
            "stderr": stderr_descriptor,
        }

    complete = read_json(packet / "capture-complete.json", "capture-complete")
    require(isinstance(complete, dict), "capture-complete: expected an object")
    require(complete.get("status") == "pass", "capture-complete: status is not pass")
    require(complete.get("binary_unchanged") is True, "capture-complete: binary changed")
    require(complete.get("workload_children") == len(LABELS),
            "capture-complete: workload child count differs")
    require(complete.get("retained_samples") == TOTAL_SAMPLES,
            "capture-complete: retained sample count differs")
    return {
        "binary": {
            "bytes": binary_bytes,
            "sha256": binary_sha,
            "executable": binary.get("executable"),
            "same_across_receipts": True,
        },
        "plan": {
            "legs": EXPECTED_LEGS,
            "warmup": WARMUP,
            "samples_per_leg": SAMPLES,
            "cpu": plan["cpu"],
            "environment": environment,
        },
        "legs": legs,
        "complete": {
            "status": "pass",
            "binary_unchanged": True,
            "workload_children": len(LABELS),
            "retained_samples": TOTAL_SAMPLES,
        },
    }


def historical_identities(packet: Path) -> dict[str, Any]:
    root = HERE.parents[3]
    reports: list[dict[str, Any]] = []
    corpus_hashes: set[str] = set()
    output_hashes: set[str] = set()
    for relative_path in HISTORICAL_REPORTS:
        path = root / relative_path
        if not path.is_file() or path.is_symlink():
            continue
        try:
            report = read_json(path, relative_path)
            result = report["results"][0]
            corpus_sha = result["corpus"]["archive_sha256"]
            output_sha = result["output_sha256"]
            valid_sha(corpus_sha, f"{relative_path}.corpus.archive_sha256")
            valid_sha(output_sha, f"{relative_path}.output_sha256")
        except (KeyError, TypeError, IndexError, EvidenceError, json.JSONDecodeError) as error:
            fail(f"historical report {relative_path} is malformed: {error}")
        reports.append({
            "path": relative_path,
            "bytes": path.stat().st_size,
            "sha256": sha256(path),
            "corpus_archive_sha256": corpus_sha,
            "output_sha256": output_sha,
        })
        corpus_hashes.add(corpus_sha)
        output_hashes.add(output_sha)
    return {
        "available": bool(reports),
        "reports": reports,
        "corpus_archive_sha256": sorted(corpus_hashes),
        "output_sha256": sorted(output_hashes),
    }


def pooled_ratios(legs: Mapping[str, Mapping[str, Any]]) -> dict[str, Any]:
    eager = [value for label in ("eager-a", "eager-b") for value in _elapsed_values(legs[label])]
    source = [value for label in ("source-a", "source-b") for value in _elapsed_values(legs[label])]
    eager_summary = summary(eager, "pooled eager elapsed_ns")
    source_summary = summary(source, "pooled source-backed elapsed_ns")

    def ratio(numerator: Mapping[str, Any], denominator: Mapping[str, Any]) -> dict[str, float]:
        return {
            key: float(numerator[key]) / float(denominator[key])
            for key in ("p50", "p95", "p99")
        }

    pairs = {}
    for pair, eager_label, source_label in (
        ("a", "eager-a", "source-a"),
        ("b", "eager-b", "source-b"),
    ):
        eager_leg = legs[eager_label]["elapsed_ns"]
        source_leg = legs[source_label]["elapsed_ns"]
        pairs[pair] = {
            "source_backed_over_eager": ratio(source_leg, eager_leg),
            "eager_over_source_backed": ratio(eager_leg, source_leg),
        }
    return {
        "eager": eager_summary,
        "source_backed": source_summary,
        "ratios": {
            "source_backed_over_eager": ratio(source_summary, eager_summary),
            "eager_over_source_backed": ratio(eager_summary, source_summary),
        },
        "pair_ratios": pairs,
        "claim": "descriptive_current_same_binary_path_comparison_no_ci_or_speedup_claim",
    }


def _elapsed_values(leg: Mapping[str, Any]) -> list[int]:
    # The raw vector is intentionally retained only long enough to calculate
    # pooled distributions.  The per-leg result already exposes its checked
    # summary and sample count.
    return list(leg["_raw_elapsed"])


def analyze(packet: Path) -> dict[str, Any]:
    packet = packet.resolve()
    require(packet.is_dir() and not packet.is_symlink(), f"packet is not a regular directory: {packet}")
    custody = check_capture_custody(packet)
    parsed: dict[str, Any] = {}
    for label in LABELS:
        parsed[label] = raw_report(packet, label)

    corpus_values = [parsed[label]["corpus"] for label in LABELS]
    corpus_same = all(canonical(value) == canonical(corpus_values[0]) for value in corpus_values[1:])
    require(corpus_same, "corpus manifest differs between capture legs")
    output_hashes = {label: parsed[label]["output_sha256"] for label in LABELS}
    output_stable = {label: True for label in LABELS}
    binary_report_hashes = {
        label: parsed[label]["binary_sha256_in_report"]
        for label in LABELS
    }
    require(len(binary_report_hashes) == len(LABELS),
            "raw reports do not all contain binary_identity")
    require(len(set(binary_report_hashes.values())) == 1,
            "raw reports were produced by different binaries")
    require(next(iter(binary_report_hashes.values())) == custody["binary"]["sha256"],
            "raw report binary_identity differs from receipt binary")
    binary_report_bytes = {
        label: parsed[label]["binary_bytes_in_report"] for label in LABELS
    }
    require(len(set(binary_report_bytes.values())) == 1,
            "raw reports contain different binary byte counts")
    require(next(iter(binary_report_bytes.values())) == custody["binary"]["bytes"],
            "raw report binary byte count differs from receipt binary")

    path_pairs: dict[str, dict[str, Any]] = {}
    for kind, first, second in (
        ("eager", "eager-a", "eager-b"),
        ("source_backed", "source-a", "source-b"),
    ):
        output_equal = parsed[first]["output_sha256"] == parsed[second]["output_sha256"]
        sink_equal = parsed[first]["sink"]["accepted_bytes"] == parsed[second]["sink"]["accepted_bytes"]
        require(output_equal, f"{kind}: output hashes differ within the repeated path")
        require(sink_equal, f"{kind}: sink accepted bytes differ within the repeated path")
        path_pairs[kind] = {
            "labels": [first, second],
            "output_sha256_equal": output_equal,
            "sink_accepted_bytes_equal": sink_equal,
            "output_sha256": parsed[first]["output_sha256"],
            "sink_accepted_bytes": parsed[first]["sink"]["accepted_bytes"],
        }

    historical = historical_identities(packet)
    historical_corpus = set(historical["corpus_archive_sha256"])
    historical_output = set(historical["output_sha256"])

    pooled_input: dict[str, Any] = {}
    leg_output: dict[str, Any] = {}
    for label in LABELS:
        item = parsed[label]
        # Internal raw values are removed before the document is serialized.
        item["_raw_elapsed"] = check_vector(
            read_json(packet / f"{label}.json", label)["results"][0]["elapsed_ns"]["samples"],
            f"{label}.elapsed_ns.samples",
        )
        pooled_input[label] = item
        leg_output[label] = {key: value for key, value in item.items() if key != "_raw_elapsed"}

    pooled = pooled_ratios(pooled_input)
    corpus_identity = {
        "archive_sha256": parsed[LABELS[0]]["corpus_archive_sha256"],
        "archive_bytes": corpus_values[0].get("archive_bytes"),
        "archive_member_count": corpus_values[0].get("archive_member_count"),
        "entry_count": corpus_values[0].get("entry_count"),
        "shape": corpus_values[0].get("shape"),
        "generator": corpus_values[0].get("generator"),
    }
    historical_corpus_match = bool(historical_corpus) and corpus_identity["archive_sha256"] in historical_corpus
    historical_output_match = {
        label: bool(historical_output) and output_hashes[label] in historical_output
        for label in LABELS
    }
    historical_match = {
        "reference_available": historical["available"],
        "corpus_archive_sha256": historical_corpus_match,
        "output_sha256": historical_output_match,
        "all_current_identities_match": historical_corpus_match
        and all(historical_output_match.values()),
    }
    return {
        "schema": SCHEMA,
        "status": "pass",
        "scope": {
            "capture": "four fresh same-binary legs: eager-a, source-a, source-b, eager-b",
            "samples_per_leg": SAMPLES,
            "warmup_iterations_per_leg": WARMUP,
            "retained_samples": TOTAL_SAMPLES,
            "latency_statistics": "per-leg and pooled descriptive elapsed_ns summaries",
            "comparison_claim": "descriptive current same-binary path comparison",
            "confidence_intervals": False,
            "speedup_claim": False,
        },
        "custody": custody,
        "binary": {
            "sha256": custody["binary"]["sha256"],
            "bytes": custody["binary"]["bytes"],
            "same_across_receipts": True,
            "same_across_raw_reports": True,
        },
        "historical_match": historical_match,
        "corpus": {
            "same_across_legs": True,
            "identity": corpus_identity,
            "manifest": corpus_values[0],
            "historical_match": historical_corpus_match,
        },
        "output": {
            "per_leg_sha256": output_hashes,
            "per_leg_stable": output_stable,
            "within_path_pair": path_pairs,
            "cross_path_hashes_equal": len(set(output_hashes.values())) == 1,
            "historical_match": historical_output_match,
            "historical_reference": historical,
        },
        "legs": leg_output,
        "pooled": pooled,
        "checks": {
            "root_receipts": True,
            "same_binary_receipts": True,
            "sample_count_200_per_leg": True,
            "warmup_count_30_per_leg": True,
            "sample_order_permutation_per_leg": True,
            "reported_percentiles_recomputed": True,
            "corpus_manifest_equal": True,
            "output_hash_present_per_leg": True,
            "output_hash_and_sink_bytes_repeat_consistent": True,
            "materializations_229_eager_9_source": True,
            "sink_largest_write_at_most_65536": True,
            "source_reads_present_per_sample": True,
        },
        "limitations": [
            "The output hash and semantic correctness gates are those recorded by the harness; this reader does not reopen the PPTX.",
            "Source-read and materialization vectors are retained evidence for the completed samples; they are not physical-I/O or allocation measurements.",
            "The pooled ratios combine two legs per path and are descriptive; no confidence interval or speedup claim is made from two pairs.",
            "Historical identity flags are informational and do not qualify or reject the fresh corpus or output.",
        ],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--packet", type=Path, required=True,
                        help="packet directory containing the four raw reports and receipts")
    parser.add_argument("--output", type=Path, required=True,
                        help="new JSON analysis path; existing paths are never overwritten")
    args = parser.parse_args()
    try:
        output = args.output.resolve()
        require(not output.exists() and not output.is_symlink(),
                f"refusing to overwrite analysis output: {output}")
        result = analyze(args.packet)
        output.parent.mkdir(parents=True, exist_ok=True)
        with output.open("x", encoding="utf-8") as stream:
            json.dump(result, stream, indent=2, sort_keys=True)
            stream.write("\n")
    except (EvidenceError, OSError) as error:
        print(f"analyze.py: error: {error}", file=sys.stderr)
        return 2
    print(f"0907 PPTX analysis passed: {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
