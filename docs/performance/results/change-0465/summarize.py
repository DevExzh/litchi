#!/usr/bin/env python3
"""Recompute the descriptive 0465 default and ODP append summary.

The summary is deliberately a projection of retained JSON reports.  It does
not compare revisions, turn allocator observations into a latency claim, or
infer a speedup.  ``verify.py`` performs the fail-closed report and custody
checks before accepting this projection.
"""

from __future__ import annotations

import hashlib
import json
import math
import statistics
from pathlib import Path
from typing import Any, Iterable, Mapping


ROOT = Path(__file__).resolve().parent
CHANGE = 465
SCHEMA = "litchi-0465-summary-v1"
REPORT_SCHEMA_VERSION = 1
NORMAL_CASE_COUNT = 37
NORMAL_RESULT_COUNT = 201
ODP_CASE = "odp_existing_append_lifecycle"
ODP_SHAPES = ("tiny", "medium", "large")
INSTRUMENTATIONS = ("normal", "allocator")
REPEATS = ("R1", "R2")


class SummaryError(ValueError):
    """Raised when retained reports cannot produce the frozen summary."""


def fail(message: str) -> None:
    raise SummaryError(message)


def load(path: Path, label: str) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"{label}: invalid JSON ({error})")
    raise AssertionError("unreachable")


def obj(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label}: expected object")
    return value


def nonempty_text(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        fail(f"{label}: expected non-empty string")
    return value


def integer(value: Any, label: str, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        fail(f"{label}: expected integer >= {minimum}")
    return value


def sha256(value: Any, label: str) -> str:
    value = nonempty_text(value, label)
    if len(value) != 64 or any(character not in "0123456789abcdef" for character in value):
        fail(f"{label}: expected lowercase SHA-256")
    return value


def digest_file(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def _midpoint(left: int, right: int) -> int:
    # This is the same overflow-safe midpoint used by the Rust harness.
    return left // 2 + right // 2 + (left % 2 + right % 2) // 2


def _nearest_rank(values: list[int], percentile: int) -> int:
    index = min((percentile * len(values) + 99) // 100 - 1, len(values) - 1)
    return values[index]


def _student_t_critical_95(degrees: int) -> float:
    values = (
        12.706, 4.303, 3.182, 2.776, 2.571, 2.447, 2.365, 2.306, 2.262, 2.228,
        2.201, 2.179, 2.160, 2.145, 2.131, 2.120, 2.110, 2.101, 2.093, 2.086,
        2.080, 2.074, 2.069, 2.064, 2.060, 2.056, 2.052, 2.048, 2.045, 2.042,
    )
    if degrees == 0:
        return 0.0
    if degrees <= len(values):
        return values[degrees - 1]
    z = 1.959963984540054
    d = float(degrees)
    z2 = z * z
    z3 = z2 * z
    z5 = z3 * z2
    z7 = z5 * z2
    return z + (z3 + z) / (4.0 * d) + (5.0 * z5 + 16.0 * z3 + 3.0 * z) / (96.0 * d * d) + (3.0 * z7 + 19.0 * z5 + 17.0 * z3 - 15.0 * z) / (384.0 * d * d * d)


def stats(values: Iterable[Any], label: str) -> dict[str, Any]:
    raw = [integer(value, f"{label}[{index}]", 1) for index, value in enumerate(values)]
    if not raw:
        fail(f"{label}: no retained samples")
    ordered = sorted(raw)
    count = len(ordered)
    mean = statistics.fmean(ordered)
    deviation = statistics.stdev(ordered) if count > 1 else 0.0
    margin = _student_t_critical_95(count - 1) * deviation / math.sqrt(count) if count > 1 else 0.0
    return {
        "count": count,
        "min": ordered[0],
        "p50": _midpoint(ordered[(count - 1) // 2], ordered[count // 2]),
        "p95": _nearest_rank(ordered, 95),
        "p99": _nearest_rank(ordered, 99),
        "max": ordered[-1],
        "mean": mean,
        "standard_deviation": deviation,
        "confidence_interval_95": {
            "method": "two-sided Student's t interval for the mean",
            "lower": max(0.0, mean - margin),
            "upper": mean + margin,
        },
    }


def _report_samples(report: Mapping[str, Any], label: str) -> int:
    configuration = obj(report.get("configuration"), f"{label}.configuration")
    return integer(configuration.get("samples_per_case"), f"{label}.configuration.samples_per_case", 0)


def _rows(report: Mapping[str, Any], label: str) -> list[dict[str, Any]]:
    rows = report.get("results")
    if not isinstance(rows, list):
        fail(f"{label}.results: expected array")
    return [obj(value, f"{label}.results[{index}]") for index, value in enumerate(rows)]


def _stats_from_row(row: Mapping[str, Any], path: str) -> dict[str, Any]:
    elapsed = obj(row.get("elapsed_ns"), f"{path}.elapsed_ns")
    if elapsed.get("unit") != "ns":
        fail(f"{path}.elapsed_ns.unit: expected ns")
    values = elapsed.get("samples")
    if not isinstance(values, list):
        fail(f"{path}.elapsed_ns.samples: expected array")
    return stats(values, f"{path}.elapsed_ns.samples")


def _odp_summary(row: Mapping[str, Any], path: str, repeat: str, instrumentation: str) -> dict[str, Any]:
    corpus = obj(row.get("corpus"), f"{path}.corpus")
    source = obj(row.get("source"), f"{path}.source")
    append = obj(source.get("odp_append"), f"{path}.source.odp_append")
    shape = nonempty_text(append.get("shape"), f"{path}.source.odp_append.shape")
    if shape not in ODP_SHAPES:
        fail(f"{path}.source.odp_append.shape: unsupported {shape!r}")
    lifecycle = append.get("lifecycle_ns")
    if not isinstance(lifecycle, list):
        fail(f"{path}.source.odp_append.lifecycle_ns: expected array")
    output = append.get("output_sha256")
    if not isinstance(output, list):
        fail(f"{path}.source.odp_append.output_sha256: expected array")
    elapsed = obj(row.get("elapsed_ns"), f"{path}.elapsed_ns")
    samples = elapsed.get("samples")
    order = elapsed.get("sample_order")
    if not isinstance(samples, list) or not isinstance(order, list) or len(samples) != len(order):
        fail(f"{path}.elapsed_ns: samples/sample_order shape differs")
    if len(lifecycle) != len(samples) or len(output) != len(samples):
        fail(f"{path}.source.odp_append: vectors do not match elapsed sample count")
    if sorted(order) != list(range(len(order))):
        fail(f"{path}.elapsed_ns.sample_order: not a permutation")
    sorted_lifecycle = [lifecycle[index] for index in order]
    if sorted_lifecycle != samples:
        fail(f"{path}.source.odp_append.lifecycle_ns: does not align with elapsed_ns.samples")
    for index, digest in enumerate(output):
        sha256(digest, f"{path}.source.odp_append.output_sha256[{index}]")
    source_archive = sha256(append.get("source_archive_sha256"), f"{path}.source.odp_append.source_archive_sha256")
    output_archive = sha256(append.get("output_archive_sha256"), f"{path}.source.odp_append.output_archive_sha256")
    source_bytes = integer(append.get("source_archive_bytes"), f"{path}.source.odp_append.source_archive_bytes", 1)
    output_bytes = integer(append.get("output_archive_bytes"), f"{path}.source.odp_append.output_archive_bytes", 1)
    if corpus.get("archive_sha256") != source_archive or corpus.get("archive_bytes") != source_bytes:
        fail(f"{path}: ODP source archive differs from corpus identity")
    return {
        "repeat": repeat,
        "instrumentation": instrumentation,
        "shape": shape,
        "source_archive_sha256": source_archive,
        "source_archive_bytes": source_bytes,
        "output_archive_sha256": output_archive,
        "output_archive_bytes": output_bytes,
        "source_slide_count": append.get("source_slide_count"),
        "output_slide_count": append.get("output_slide_count"),
        "append_count": append.get("append_count"),
        "lifecycle_ns": stats(lifecycle, f"{path}.source.odp_append.lifecycle_ns"),
        "latency_ns": _stats_from_row(row, path),
        "output_sha256": output_archive,
        "output_digests": sorted(set(output)),
    }


def _case_rows(report: Mapping[str, Any], label: str) -> dict[str, dict[str, Any]]:
    result = {}
    for index, row in enumerate(_rows(report, label)):
        case = nonempty_text(row.get("case"), f"{label}.results[{index}].case")
        if case in result:
            fail(f"{label}: duplicate case {case}")
        result[case] = row
    return result


def _load_reports(root: Path) -> tuple[dict[tuple[str, str], dict[str, Any]], dict[str, Any]]:
    reports: dict[tuple[str, str], dict[str, Any]] = {}
    for instrumentation in INSTRUMENTATIONS:
        for repeat in REPEATS:
            path = root / "captures" / f"{repeat}-{instrumentation}" / "report.json"
            if not path.is_file():
                # Accept the older nested spelling when inspecting a copied
                # bundle, while capture.py's frozen spelling remains the
                # canonical R1-normal/R2-normal layout.
                path = root / "captures" / repeat / instrumentation / "report.json"
            if not path.is_file():
                fail(f"missing captures/{repeat}/{instrumentation}/report.json")
            reports[(instrumentation, repeat)] = obj(load(path, path.as_posix()), path.as_posix())
    preflight_path = root / "captures" / "preflight" / "report.json"
    preflight = obj(load(preflight_path, str(preflight_path)), str(preflight_path))
    return reports, preflight


def _report_projection(report: Mapping[str, Any], label: str, repeat: str, instrumentation: str) -> dict[str, Any]:
    rows = _rows(report, label)
    config = obj(report.get("configuration"), f"{label}.configuration")
    expected_count = 3 if instrumentation == "allocator" else NORMAL_RESULT_COUNT
    if len(rows) != expected_count:
        fail(f"{label}.results: expected {expected_count}, got {len(rows)}")
    odp_rows = [row for row in rows if row.get("case") == ODP_CASE]
    if len(odp_rows) != 3:
        fail(f"{label}: expected three ODP append rows")
    odp = {}
    for index, row in enumerate(sorted(odp_rows, key=lambda item: item["corpus"]["shape"])):
        entry = _odp_summary(row, f"{label}.results[{index}].odp", repeat, instrumentation)
        odp[entry["shape"]] = entry
    return {
        "instrumentation": instrumentation,
        "repeat": repeat,
        "report_count": 1,
        "result_count": len(rows),
        "case_count": len(config.get("cases", [])) if isinstance(config.get("cases"), list) else None,
        "samples_per_case": _report_samples(report, label),
        "warmup_iterations_per_case": integer(config.get("warmup_iterations_per_case"), f"{label}.configuration.warmup_iterations_per_case", 0),
        "odp": odp,
    }


def build_summary(root: Path = ROOT) -> dict[str, Any]:
    reports, preflight = _load_reports(root)
    normal_r1 = _report_projection(reports[("normal", "R1")], "captures/R1/normal", "R1", "normal")
    normal_r2 = _report_projection(reports[("normal", "R2")], "captures/R2/normal", "R2", "normal")
    allocator_r1 = _report_projection(reports[("allocator", "R1")], "captures/R1/allocator", "R1", "allocator")
    allocator_r2 = _report_projection(reports[("allocator", "R2")], "captures/R2/allocator", "R2", "allocator")
    preflight_config = obj(preflight.get("configuration"), "captures/preflight.configuration")
    preflight_rows = _rows(preflight, "captures/preflight")
    if len(preflight_rows) != NORMAL_RESULT_COUNT:
        fail("captures/preflight: expected 201 results")
    if integer(preflight_config.get("samples_per_case"), "captures/preflight.samples", 0) != 1 or integer(preflight_config.get("warmup_iterations_per_case"), "captures/preflight.warmup", 0) != 0:
        fail("captures/preflight: expected one sample and zero warmups")
    return {
        "schema": SCHEMA,
        "change": CHANGE,
        "scope": "descriptive default matrix plus ODP append lifecycle; allocator rows are separate evidence",
        "preflight": {
            "report_count": 1,
            "result_count": len(preflight_rows),
            "case_count": len(preflight_config.get("cases", [])),
            "samples_per_case": preflight_config.get("samples_per_case"),
            "warmup_iterations_per_case": preflight_config.get("warmup_iterations_per_case"),
        },
        "full_normal": {
            "report_count": 2,
            "result_count": NORMAL_RESULT_COUNT,
            "case_count": NORMAL_CASE_COUNT,
            "samples_per_case": 15,
            "warmup_iterations_per_case": 3,
            "repeats": {"R1": normal_r1["odp"], "R2": normal_r2["odp"]},
        },
        "odp_append": {
            "normal_latency": {"R1": normal_r1["odp"], "R2": normal_r2["odp"]},
            "allocator_observations": {"R1": allocator_r1["odp"], "R2": allocator_r2["odp"]},
            "allocator_is_separate_from_latency": True,
            "no_speedup_claim": True,
        },
        "allocator_reports": {
            "report_count": 2,
            "result_count": 3,
            "samples_per_case": 15,
            "warmup_iterations_per_case": 3,
        },
    }


def main() -> int:
    try:
        print(json.dumps(build_summary(), ensure_ascii=False, indent=2, sort_keys=True))
    except (SummaryError, OSError, KeyError, TypeError, ValueError) as error:
        print(json.dumps({"schema": SCHEMA, "change": CHANGE, "status": "failed", "error": str(error)}, sort_keys=True))
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
