#!/usr/bin/env python3
"""Validate and summarize the paired 0481 DOCX borrowed-name measurements.

The executable reports are produced by the unchanged 0479 harness.  This
module keeps that report boundary in :mod:`report_schema` and adds the paired
arm protocol used for this change: control A1, candidate B1, candidate B2,
and control A2.  It derives every statistic from the retained sample vectors;
producer summaries are never trusted.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
import math
import re
from pathlib import Path
from typing import Any, Mapping

import report_schema as _schema


ROOT = Path(__file__).resolve().parent
REPORT_SCHEMA = _schema.REPORT_SCHEMA
SUMMARY_SCHEMA = "docx-borrowed-names-comparison-summary-v1"
CORPUS_MANIFEST_SCHEMA = _schema.CORPUS_MANIFEST_SCHEMA
COUNTS = _schema.COUNTS
MODES = _schema.MODES
INSTRUMENTATIONS = _schema.INSTRUMENTATIONS
REPEATS = (1, 2)
ARMS = ("control", "candidate")
PHASES = _schema.PHASES
SAMPLES = _schema.SAMPLES
WARMUPS = _schema.WARMUPS
CPU = _schema.CPU
REGRESSION_REVIEW_PERCENT = _schema.REGRESSION_REVIEW_PERCENT
MAX_DURABLE_BYTES = _schema.MAX_DURABLE_BYTES
OPAQUE_PATH = _schema.OPAQUE_PATH
OPAQUE_BYTES = _schema.OPAQUE_BYTES
OPAQUE_SHA256 = _schema.OPAQUE_SHA256
OPAQUE_CRC32 = _schema.OPAQUE_CRC32
MAIN_PATH = _schema.MAIN_PATH
GENERATOR = _schema.GENERATOR
FORMAT = _schema.FORMAT
WORD_NAMESPACE = _schema.WORD_NAMESPACE
HASH_SINK_MAX_WRITE = _schema.HASH_SINK_MAX_WRITE
SINK_ID = _schema.SINK_ID
SOURCE_ID = _schema.SOURCE_ID
PROCESS_FIELDS = _schema.PROCESS_FIELDS
ALLOC_FIELDS = _schema.ALLOC_FIELDS
SINK_FIELDS = _schema.SINK_FIELDS
HIST_FIELDS = _schema.HIST_FIELDS
READ_FIELDS = _schema.READ_FIELDS
AnalysisError = _schema.AnalysisError

# Keep the report boundary independently named in this file so tests and the
# verifier can use it without relying on star-import behaviour for private
# helpers.  The implementation remains the frozen report_schema.py.
read_json = _schema.read_json
write_json = _schema.write_json
integer = _schema.integer
u64 = _schema.u64
positive = _schema.positive
digest = _schema.digest
fail = _schema.fail
_fragment = _schema._fragment
_source_xml = _schema._source_xml
_candidate_xml = _schema._candidate_xml
_semantic = _schema._semantic
_corpus = _schema._corpus
_total_sample = _schema._total_sample
_phase_sample = _schema._phase_sample
_stats = _schema._stats
_percent = _schema._percent
_read_equal = _schema._read_equal
_read_add = _schema._read_add


def _sync_root() -> None:
    """Make the frozen schema helper resolve artifacts under this root."""
    _schema.ROOT = ROOT


def expected_captures() -> list[dict[str, Any]]:
    """Return the exact A1/B1/B2/A2 capture order frozen by capture.py."""
    forward = [
        (instrumentation, count, mode)
        for instrumentation in INSTRUMENTATIONS
        for count in COUNTS
        for mode in MODES
    ]
    rows: list[dict[str, Any]] = []
    sequences = (
        ("control", 1, forward),
        ("candidate", 1, forward),
        ("candidate", 2, list(reversed(forward))),
        ("control", 2, list(reversed(forward))),
    )
    for arm, repeat, sequence in sequences:
        for instrumentation, count, mode in sequence:
            rows.append({
                "label": f"{arm}-r{repeat}-{instrumentation}-{count}-{mode}",
                "arm": arm,
                "instrumentation": instrumentation,
                "count": count,
                "mode": mode,
                "repeat": repeat,
            })
    return rows


def _timestamp(value: Any, label: str) -> _datetime.datetime:
    if not isinstance(value, str) or not value:
        fail(f"{label}: timestamp is missing")
    try:
        parsed = _datetime.datetime.fromisoformat(value)
    except ValueError as error:
        fail(f"{label}: invalid timestamp: {error}")
    if parsed.tzinfo is None:
        fail(f"{label}: timestamp must include a timezone")
    return parsed


def protocol_rows(protocol: Mapping[str, Any]) -> list[dict[str, Any]]:
    """Validate the paired protocol and return its capture specifications."""
    if protocol.get("schema") != "docx-borrowed-names-comparison-v1":
        fail("protocol schema differs")
    integer(protocol.get("samples"), "protocol.samples", SAMPLES)
    integer(protocol.get("warmups"), "protocol.warmups", WARMUPS)
    integer(protocol.get("cpu"), "protocol.cpu", CPU)
    integer(protocol.get("append_count"), "protocol.append_count", 1)
    for key in (
        "normal_and_allocator_timings_separate",
        "process_rss_includes_setup_oracles_and_teardown",
        "phase_peaks_are_not_total_peaks",
    ):
        if protocol.get(key) is not True:
            fail(f"protocol.{key} must be true")
    if protocol.get("regression_review_percent") != REGRESSION_REVIEW_PERCENT:
        fail("protocol regression review threshold differs")
    if protocol.get("performance_claim") != "none":
        fail("protocol performance claim differs")
    for key in ("comparison", "scope"):
        if not isinstance(protocol.get(key), str) or not protocol[key]:
            fail(f"protocol.{key} is missing")
    environment = protocol.get("environment")
    required_environment = (
        "RUSTUP_TOOLCHAIN", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL",
        "CARGO_PROFILE_RELEASE_DEBUG", "RUSTFLAGS", "DEBUGINFOD_URLS", "LC_ALL",
    )
    if not isinstance(environment, dict) or set(environment) != set(required_environment):
        fail("protocol.environment fields differ")
    if not all(isinstance(value, str) for value in environment.values()):
        fail("protocol.environment values must be strings")
    scripts = protocol.get("scripts")
    if not isinstance(scripts, dict) or set(scripts) != {"common.py", "capture.py"}:
        fail("protocol.scripts fields differ")
    for name, value in scripts.items():
        digest(value, f"protocol.scripts.{name}")
    captures = protocol.get("captures")
    expected = expected_captures()
    if not isinstance(captures, list) or len(captures) != len(expected):
        fail(f"protocol must contain exactly {len(expected)} captures")
    rows: list[dict[str, Any]] = []
    expected_fields = {"label", "arm", "instrumentation", "count", "mode", "repeat", "argv"}
    for index, (actual, wanted) in enumerate(zip(captures, expected)):
        if not isinstance(actual, dict) or set(actual) != expected_fields:
            fail(f"protocol.captures[{index}] fields differ")
        for key, value in wanted.items():
            if actual.get(key) != value:
                fail(f"protocol.captures[{index}].{key} differs")
        argv = actual.get("argv")
        if not isinstance(argv, list) or not argv or not all(isinstance(item, str) for item in argv):
            fail(f"protocol.captures[{index}].argv is malformed")
        rows.append(dict(actual))
    if len({row["label"] for row in rows}) != len(rows):
        fail("protocol capture labels are not unique")
    return rows


def validate_report(report: Mapping[str, Any], spec: Mapping[str, Any], path: str = "report") -> dict[str, Any]:
    """Validate one report; extra arm/provenance fields in ``spec`` are ignored."""
    _sync_root()
    return _schema.validate_report(report, spec, path)


def _resource_rss(spec: Mapping[str, Any]) -> int:
    _sync_root()
    return _schema._resource_rss(spec)


def _metric_rows(rows: list[dict[str, Any]], mode: str, instrumentation: str, corpus: Mapping[str, Any]) -> dict[str, Any]:
    result = _schema._metric_rows(rows, mode, instrumentation, corpus)
    if mode == "total" and instrumentation == "allocator":
        result["total_peak_live_bytes_scope"] = "absolute operation-region allocator live-byte high-water mark"
    return result


def _validate_pilot_report(value: Any, label: str, corpora: Mapping[str, Any]) -> None:
    """Validate a one-sample all-count pilot against the frozen corpus."""
    if not isinstance(value, dict) or set(value) != {"schema", "version", "binary", "config", "cases"}:
        fail(f"{label}: report fields differ")
    instrumentation = "allocator" if "allocator" in label else "normal"
    mode = "phases" if label.endswith("-phases") else "total"
    if value.get("schema") != REPORT_SCHEMA or value.get("version") != 1:
        fail(f"{label}: report schema differs")
    expected_binary = {
        "binary": "litchi-perf-baseline" if instrumentation == "normal" else "litchi-perf-baseline-alloc",
        "allocator": "Rust system allocator" if instrumentation == "normal" else "CountingSystemAllocator(std::alloc::System)",
        "instrumentation": "none" if instrumentation == "normal" else "system_allocator_operation_scoped",
        "counter_revision": None if instrumentation == "normal" else "serialized_region_peak_v3",
    }
    if value.get("binary") != expected_binary:
        fail(f"{label}: binary identity differs")
    config = value.get("config")
    required = {"counts", "samples", "warmups", "mode", "lifecycle_phases", "sink", "source"}
    if (
        not isinstance(config, dict)
        or set(config) != required
        or config.get("counts") != list(COUNTS)
        or config.get("samples") != 1
        or config.get("warmups") != 1
        or config.get("mode") != mode
        or config.get("lifecycle_phases") != list(PHASES)
        or config.get("sink") != SINK_ID
        or config.get("source") != SOURCE_ID
    ):
        fail(f"{label}: protocol differs")
    cases = value.get("cases")
    if not isinstance(cases, list) or [case.get("count") for case in cases if isinstance(case, dict)] != list(COUNTS):
        fail(f"{label}: cases differ")
    if len(cases) != len(COUNTS):
        fail(f"{label}: case count differs")
    for case in cases:
        count = case["count"]
        if set(case) != {"count", "corpus", "total_samples", "phase_samples"}:
            fail(f"{label}/{count}: case fields differ")
        corpus = case["corpus"]
        if corpus != corpora[str(count)]:
            fail(f"{label}/{count}: corpus differs from baseline")
        checked_corpus = _corpus(corpus, count, f"{label}/{count}.corpus")
        if mode == "total":
            require = case.get("phase_samples") is None and isinstance(case.get("total_samples"), list) and len(case["total_samples"]) == 1
            if not require:
                fail(f"{label}/{count}: total pilot sample shape differs")
            checked_sample = _total_sample(case["total_samples"][0], checked_corpus, instrumentation, f"{label}/{count}.total")
        else:
            require = case.get("total_samples") is None and isinstance(case.get("phase_samples"), list) and len(case["phase_samples"]) == 1
            if not require:
                fail(f"{label}/{count}: phase pilot sample shape differs")
            checked_sample = _phase_sample(case["phase_samples"][0], checked_corpus, instrumentation, f"{label}/{count}.phases")
        if checked_sample["sample"] != 0:
            fail(f"{label}/{count}: sample index differs")


def _baseline_corpus() -> tuple[dict[str, Any], str]:
    """Load the root-frozen corpus and independently re-derive its XML."""
    path = ROOT / "baseline-corpus.json"
    if not path.is_file():
        fail("baseline-corpus.json is missing")
    value = read_json(path)
    if not isinstance(value, dict) or set(value) != {"schema", "frozen_utc", "cases", "pilots"}:
        fail("baseline-corpus.json fields differ")
    if value["schema"] != CORPUS_MANIFEST_SCHEMA:
        fail("baseline-corpus.json schema differs")
    initial_path = ROOT / "initial-state.json"
    initial = read_json(initial_path)
    if not isinstance(initial, dict) or digest(initial.get("baseline_corpus_sha256"), "initial-state.baseline_corpus_sha256") != hashlib.sha256(path.read_bytes()).hexdigest():
        fail("baseline-corpus.json is not bound by initial-state.json")
    _timestamp(value["frozen_utc"], "baseline-corpus.json.frozen_utc")
    cases = value["cases"]
    if not isinstance(cases, dict) or set(cases) != {str(count) for count in COUNTS}:
        fail("baseline-corpus.json cases differ")
    corpora: dict[str, Any] = {}
    for count in COUNTS:
        corpora[str(count)] = _corpus(cases[str(count)], count, f"baseline-corpus.json.cases.{count}")
    pilots = value["pilots"]
    if not isinstance(pilots, dict) or set(pilots) != {f"pilot-{instrumentation}-{mode}" for instrumentation in INSTRUMENTATIONS for mode in MODES}:
        fail("baseline-corpus.json pilots differ")
    # These four references are retained historic 0479 provenance.  Their
    # files are intentionally absent from the 0481 bundle, so the baseline
    # boundary validates the reference shape without opening those paths.
    # Current control/candidate pilot reports are validated by verify.py and
    # are the reports that bind this comparison to the frozen corpus.
    for label, reference in pilots.items():
        if not isinstance(reference, dict) or set(reference) != {"path", "bytes", "sha256"}:
            fail(f"baseline-corpus.json.pilots.{label} reference differs")
        name = reference["path"]
        if not isinstance(name, str) or Path(name).name != name or Path(name).is_absolute():
            fail(f"baseline-corpus.json.pilots.{label} path differs")
        positive(reference["bytes"], f"baseline-corpus.json.pilots.{label}.bytes")
        digest(reference["sha256"], f"baseline-corpus.json.pilots.{label}.sha256")
    return {"schema": value["schema"], "frozen_utc": value["frozen_utc"], "cases": corpora, "pilots": value["pilots"]}, hashlib.sha256(path.read_bytes()).hexdigest()


def _corpus_manifest() -> tuple[dict[str, Any], str]:
    """Compatibility name used by the older evidence boundary."""
    return _baseline_corpus()


def _sample_identity_check(total: Mapping[str, Any], phase: Mapping[str, Any], label: str) -> None:
    _read_equal(total["source_reads"], phase["source_reads"], f"{label}: total/phase source reads differ")
    if total["sink"] != phase["sink"]:
        fail(f"{label}: total/phase sink differs")


def _pair_metric(before: Mapping[str, Any], after: Mapping[str, Any], metric: str) -> dict[str, Any]:
    left = before[metric]
    right = after[metric]
    return {
        "control": left,
        "candidate": right,
        "candidate_vs_control_percent": _percent(float(left["mean"]), float(right["mean"])),
    }


def _pair_metric_map(before: Mapping[str, Any], after: Mapping[str, Any], fields: tuple[str, ...]) -> dict[str, Any]:
    return {
        field: _pair_metric(before, after, field)
        for field in fields
        if field in before and field in after and isinstance(before[field], Mapping) and "mean" in before[field]
    }


def _review_flags(control: Mapping[str, Any], candidate: Mapping[str, Any], metric: str, scope: str) -> list[dict[str, Any]]:
    flags: list[dict[str, Any]] = []
    for statistic in ("mean", "p50", "p95", "p99"):
        change = _percent(float(control[statistic]), float(candidate[statistic]))
        if change is not None and abs(change) > REGRESSION_REVIEW_PERCENT:
            flags.append({
                "scope": scope, "metric": metric, "statistic": statistic,
                "candidate_vs_control_percent": change,
                "threshold_percent": REGRESSION_REVIEW_PERCENT,
            })
    return flags


def _pair_summary(
    control: Mapping[str, Any],
    candidate: Mapping[str, Any],
    mode: str,
    instrumentation: str,
    count: int,
    repeat: int,
) -> dict[str, Any]:
    result: dict[str, Any] = {
        "repeat": repeat,
        "count": count,
        "instrumentation": instrumentation,
        "mode": mode,
        "control_arm": "control",
        "candidate_arm": "candidate",
        "candidate_vs_control": {},
        "review_flags": [],
        "same_output_identity": control["corpus_identity"] == candidate["corpus_identity"],
    }
    if not result["same_output_identity"]:
        fail(f"pair {repeat}/{instrumentation}/{count}/{mode}: corpus identity differs")
    if mode == "total":
        for metric in ("elapsed_ns", "process_max_rss_kib"):
            result["candidate_vs_control"][metric] = _pair_metric(control, candidate, metric)
        result["review_flags"] += _review_flags(control["elapsed_ns"], candidate["elapsed_ns"], "latency", "total")
        rss_control = {key: value for key, value in control["process_max_rss_kib"].items()}
        rss_candidate = {key: value for key, value in candidate["process_max_rss_kib"].items()}
        result["review_flags"] += _review_flags(rss_control, rss_candidate, "process_max_rss_kib", "whole_process")
        for metric in ("allocation", "incremental_peak_live_bytes"):
            if control.get(metric) is not None and candidate.get(metric) is not None:
                if metric == "allocation":
                    result["candidate_vs_control"][metric] = _pair_metric_map(
                        control[metric], candidate[metric], tuple(ALLOC_FIELDS)
                    )
                else:
                    result["candidate_vs_control"][metric] = _pair_metric(control, candidate, metric)
        result["allocation_byte_scope"] = "allocated_bytes includes realloc new_size; no physical reallocated-byte counter is exposed"
    else:
        result["candidate_vs_control"]["phase_elapsed_ns"] = {}
        result["candidate_vs_control"]["phase_process"] = {}
        result["candidate_vs_control"]["phase_allocation"] = {}
        result["candidate_vs_control"]["process_max_rss_kib"] = _pair_metric(control, candidate, "process_max_rss_kib")
        result["review_flags"] += _review_flags(
            control["process_max_rss_kib"], candidate["process_max_rss_kib"],
            "process_max_rss_kib", "whole_process",
        )
        for phase in PHASES:
            result["candidate_vs_control"]["phase_elapsed_ns"][phase] = _pair_metric(
                control["phase_elapsed_ns"], candidate["phase_elapsed_ns"], phase
            )
            result["candidate_vs_control"]["phase_process"][phase] = {
                field: {
                    "control": control["process"][phase][field],
                    "candidate": candidate["process"][phase][field],
                    "candidate_vs_control_percent": _percent(
                        float(control["process"][phase][field]["mean"]),
                        float(candidate["process"][phase][field]["mean"]),
                    ),
                }
                for field in ("peak_rss_bytes", "rss_bytes")
            }
            result["review_flags"] += _review_flags(
                control["phase_elapsed_ns"][phase], candidate["phase_elapsed_ns"][phase],
                "latency", f"phase.{phase}",
            )
            result["review_flags"] += _review_flags(
                control["process"][phase]["peak_rss_bytes"], candidate["process"][phase]["peak_rss_bytes"],
                "peak_rss_bytes", f"phase.{phase}",
            )
        result["phase_peaks_are_not_total_peak"] = True
        if instrumentation == "allocator":
            for phase in PHASES:
                result["candidate_vs_control"]["phase_allocation"][phase] = {
                    field: {
                        "control": control["phase_allocation"][phase][field],
                        "candidate": candidate["phase_allocation"][phase][field],
                        "candidate_vs_control_percent": _percent(
                            float(control["phase_allocation"][phase][field]["mean"]),
                            float(candidate["phase_allocation"][phase][field]["mean"]),
                        ),
                    }
                    for field in ("allocation_calls", "deallocation_calls", "reallocation_calls", "allocated_bytes", "deallocated_bytes", "retained_delta", "region_peak_live_bytes")
                }
    return result


def _repeat_drifts(rows: Mapping[str, Mapping[str, Any]]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for arm in ARMS:
        for instrumentation in INSTRUMENTATIONS:
            for count in COUNTS:
                for mode in MODES:
                    first = rows[f"{arm}-r1-{instrumentation}-{count}-{mode}"]
                    second = rows[f"{arm}-r2-{instrumentation}-{count}-{mode}"]
                    changes: dict[str, float | None] = {
                        "process_max_rss_kib": _percent(first["process_max_rss_kib"]["mean"], second["process_max_rss_kib"]["mean"]),
                    }
                    if mode == "total":
                        for statistic in ("mean", "p50", "p95", "p99"):
                            changes[f"elapsed_{statistic}"] = _percent(first["elapsed_ns"][statistic], second["elapsed_ns"][statistic])
                    else:
                        for phase in PHASES:
                            for statistic in ("mean", "p50", "p95", "p99"):
                                changes[f"{phase}_elapsed_{statistic}"] = _percent(first["phase_elapsed_ns"][phase][statistic], second["phase_elapsed_ns"][phase][statistic])
                    result.append({
                        "arm": arm, "instrumentation": instrumentation, "count": count, "mode": mode,
                        "percent_changes": changes,
                        "absolute_review_flags": [key for key, value in changes.items() if value is not None and abs(value) > REGRESSION_REVIEW_PERCENT],
                    })
    return result


def derive() -> dict[str, Any]:
    """Validate all 48 formal reports and derive the paired summary in memory."""
    _sync_root()
    protocol = read_json(ROOT / "protocol.json")
    if not isinstance(protocol, dict):
        fail("protocol must be an object")
    specs = protocol_rows(protocol)
    corpus_manifest, corpus_manifest_sha256 = _corpus_manifest()
    reports: dict[str, dict[str, Any]] = {}
    rows: dict[str, dict[str, Any]] = {}
    identities: dict[int, dict[str, Any]] = {}
    for spec in specs:
        label = spec["label"]
        report_path = ROOT / "captures" / f"{label}.report.json"
        report = read_json(report_path)
        if not isinstance(report, dict):
            fail(f"{label}: report must be an object")
        checked = validate_report(report, spec, str(report_path))
        reports[label] = checked
        corpus = checked["corpus"]
        identity = json.loads(json.dumps(corpus, sort_keys=True))
        if spec["count"] in identities and identities[spec["count"]] != identity:
            fail(f"{label}: corpus identity changed across captures")
        identities[spec["count"]] = identity
        if corpus_manifest["cases"][str(spec["count"])] != identity:
            fail(f"{label}: corpus identity differs from baseline-corpus.json")
        metric = _metric_rows(checked["samples"], spec["mode"], spec["instrumentation"], corpus)
        metric["process_max_rss_kib"] = _stats([_resource_rss(spec)] * SAMPLES)
        rows[label] = {
            "capture": spec,
            "corpus_identity": {
                "count": spec["count"],
                "source_archive_bytes": corpus["source_archive_bytes"],
                "source_archive_sha256": corpus["source_archive_sha256"],
                "candidate_archive_bytes": corpus["candidate_archive_bytes"],
                "candidate_archive_sha256": corpus["candidate_archive_sha256"],
            },
            **metric,
        }

    pairs: list[dict[str, Any]] = []
    # The pair direction is always candidate versus control even for B2/A2,
    # whose execution order is reversed by the protocol.
    for repeat in REPEATS:
        for instrumentation in INSTRUMENTATIONS:
            for count in COUNTS:
                for mode in MODES:
                    control = rows[f"control-r{repeat}-{instrumentation}-{count}-{mode}"]
                    candidate = rows[f"candidate-r{repeat}-{instrumentation}-{count}-{mode}"]
                    if mode == "total":
                        phase_mode = "phases"
                    else:
                        phase_mode = "total"
                    pairs.append(_pair_summary(control, candidate, mode, instrumentation, count, repeat))
                    # The total and phase runs are distinct lifecycles, but
                    # their source and sink oracles must match sample-for-
                    # sample within each arm and repeat.
                    if mode == "total":
                        for arm in ARMS:
                            total = reports[f"{arm}-r{repeat}-{instrumentation}-{count}-total"]["samples"]
                            phases = reports[f"{arm}-r{repeat}-{instrumentation}-{count}-phases"]["samples"]
                            for index, (total_sample, phase_sample) in enumerate(zip(total, phases)):
                                _sample_identity_check(total_sample, phase_sample, f"{arm}/r{repeat}/{instrumentation}/{count}/sample{index}")
                    del phase_mode

    growth: list[dict[str, Any]] = []
    for arm in ARMS:
        for instrumentation in INSTRUMENTATIONS:
            for mode in MODES:
                if mode == "total":
                    growth.append({
                        "arm": arm, "instrumentation": instrumentation, "mode": mode,
                        "means": {str(count): rows[f"{arm}-r1-{instrumentation}-{count}-{mode}"]["elapsed_ns"]["mean"] for count in COUNTS},
                        "claim": "scaling_cost_observation",
                    })
                else:
                    growth.append({
                        "arm": arm, "instrumentation": instrumentation, "mode": mode,
                        "means": {
                            str(count): {
                                phase: rows[f"{arm}-r1-{instrumentation}-{count}-{mode}"]["phase_elapsed_ns"][phase]["mean"]
                                for phase in PHASES
                            } for count in COUNTS
                        },
                        "claim": "scaling_cost_observation",
                    })
    return {
        "schema": SUMMARY_SCHEMA,
        "protocol_schema": protocol["schema"],
        "captures": len(rows),
        "samples": len(rows) * SAMPLES,
        "arms": list(ARMS),
        "pairing": "A1=control-r1/B1=candidate-r1/B2=candidate-r2/A2=control-r2; candidate_vs_control direction",
        "rows": rows,
        "pairs": pairs,
        "repeat_drifts": _repeat_drifts(rows),
        "growth": growth,
        "corpus_manifest_sha256": corpus_manifest_sha256,
        "whole_process_rss_scope": "GNU /usr/bin/time -v Maximum resident set size (kbytes), including setup, corpus/oracles, warmups, measured lifecycles, report serialization, and teardown; not operation-local RSS",
        "operation_process_rss_scope": "report process.rss_bytes is a saturating RSS delta and process.peak_rss_bytes is the absolute VmHWM after the observer interval",
        "allocation_byte_counter_scope": "allocated_bytes includes the full new_size on realloc callbacks; no physical-copy or reallocated-byte counter is exposed",
        "uncertainty": "nearest-rank percentiles; mean interval uses t(29)=2.045; A1/B1/B2/A2 matched pairs and within-arm chronology retained; no optimization claim",
        "phase_peak_rule": "phase high-water marks are attribution evidence and are never summed or presented as a total-operation peak",
        "normal_allocation": "unavailable; normal binary does not install the counting allocator",
        "review_threshold": {"metric": "absolute latency/RSS percent change", "percent": REGRESSION_REVIEW_PERCENT},
    }


def main(argv: list[str] | None = None) -> int:
    global ROOT
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    ROOT = args.root.resolve()
    try:
        write_json(ROOT / "summary.json", derive())
    except (AnalysisError, OSError) as error:
        print(f"analyze.py: FAIL: {error}")
        return 1
    print("analyze.py: wrote summary.json")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
