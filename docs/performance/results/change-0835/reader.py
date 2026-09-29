#!/usr/bin/env python3
"""Independent admission reader for the 0835 filesystem baseline.

The native driver is deliberately kept outside this module.  This reader only
opens retained JSON, source snapshots, receipts, and report files; it never
starts a process, builds a binary, probes the filesystem, or interprets
latency as an optimization result.

The detailed evidence checks live in the committed 0834 reader.  Loading that
reader as a frozen dependency keeps the strict ZIP, cold-page-cache, PPTX
raw-read, and OPC output-proof rules byte-for-byte identical while this packet
owns its own origin, baseline freeze, build, receipt, qualification, and
formal-capture custody checks.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import math
import re
from pathlib import Path
from typing import Any


HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
BASE = "77adc4f1e24bdf76f5e34ed4dfc0113e25120c59"
STAGE = "baseline"
LEGACY_READER_SHA256 = "5570e355415bed275b6469af79dbc1217c70b6c00da5ce985d6e5917b404a16e"
QUALITY_GATES = ("fmt", "check", "test", "clippy", "doc", "boundaries")
CASES = (
    "opc_file_eager_open",
    "opc_file_source_open",
    "opc_file_eager_one_part_atomic_save",
    "opc_file_source_one_part_atomic_save",
    "pptx_file_eager_open_selected_slide_lifecycle",
    "pptx_file_source_open_selected_slide_lifecycle",
)
STATES = ("warm", "cold-verified")
REPORTS = 72
SAMPLES = 2160
CAPTURE_SAMPLES = 30
CAPTURE_WARMUP = 3
SHA256_RE = re.compile(r"[0-9a-f]{64}\Z")


class QualificationError(AssertionError):
    """An immutable packet or report admission invariant failed."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise QualificationError(message)


def read(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing regular JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise QualificationError(f"{path}: invalid JSON: {error}") from error


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1 << 20), b""):
                digest.update(block)
    except OSError as error:
        raise QualificationError(f"cannot hash {path}: {error}") from error
    return digest.hexdigest()


def _finite(value: Any, path: str = "json") -> None:
    if isinstance(value, float):
        require(math.isfinite(value), f"{path}: non-finite number")
    elif isinstance(value, dict):
        for key, child in value.items():
            require(isinstance(key, str), f"{path}: non-string key")
            _finite(child, f"{path}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            _finite(child, f"{path}[{index}]")


def _same(actual: Any, expected: Any, path: str) -> None:
    require(actual == expected, f"{path}: expected {expected!r}, got {actual!r}")


def _uint(value: Any, path: str, *, positive: bool = False) -> int:
    minimum = 1 if positive else 0
    require(type(value) is int and minimum <= value <= (1 << 64) - 1,
            f"{path}: expected {'positive' if positive else 'non-negative'} integer")
    return value


def _hash(value: Any, path: str) -> str:
    require(isinstance(value, str) and SHA256_RE.fullmatch(value) is not None,
            f"{path}: expected lowercase SHA-256")
    return value


def _descriptor(value: Any, base: Path, label: str, *, allow_missing: bool = False) -> Path:
    require(isinstance(value, dict), f"{label}: descriptor is missing")
    path_value = value.get("path")
    require(isinstance(path_value, str) and path_value, f"{label}.path: missing")
    _uint(value.get("bytes"), f"{label}.bytes")
    _hash(value.get("sha256"), f"{label}.sha256")
    candidate = Path(path_value)
    if not candidate.is_absolute():
        candidate = base / candidate
    require(not candidate.is_symlink(), f"{label}: descriptor is a symlink")
    candidate = candidate.resolve()
    if not candidate.exists():
        require(allow_missing, f"{label}: file is missing: {candidate}")
        return candidate
    require(candidate.is_file() and not candidate.is_symlink(),
            f"{label}: descriptor is not a regular file")
    _same(candidate.stat().st_size, value["bytes"], f"{label}.bytes")
    _same(sha(candidate), value["sha256"], f"{label}.sha256")
    return candidate


def _contains_descriptor(value: Any, expected: dict[str, Any]) -> bool:
    if isinstance(value, dict):
        if all(value.get(field) == expected.get(field)
               for field in ("path", "bytes", "sha256")):
            return True
        return any(_contains_descriptor(child, expected) for child in value.values())
    if isinstance(value, list):
        return any(_contains_descriptor(child, expected) for child in value)
    return False


def _load_legacy_reader() -> Any:
    """Load the retained 0834 strict report implementation without executing it."""

    legacy_path = ROOT / "docs/performance/results/change-0834/reader.py"
    require(legacy_path.is_file() and not legacy_path.is_symlink(),
            f"missing archived strict reader: {legacy_path}")
    _same(sha(legacy_path), LEGACY_READER_SHA256,
          f"{legacy_path}: archived strict reader hash differs from 0834 seal")
    spec = importlib.util.spec_from_file_location("litchi_reader_0834_frozen", legacy_path)
    require(spec is not None and spec.loader is not None,
            f"cannot load archived strict reader: {legacy_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    # The report implementation resolves source oracles through these module
    # globals.  Bind them to this packet before any validation occurs.
    module.HERE = HERE
    module.ROOT = ROOT
    module.BASE = BASE
    module.ACTIVE_STAGE = STAGE
    module.ACTIVE_SUFFIX = STAGE
    module._SOURCE_STAGE = STAGE
    module._SOURCE_ORACLE_CACHE = None
    return module


_LEGACY = _load_legacy_reader()


def _strict_stats(value: Any, path: str, expected_samples: int = 1) -> list[int]:
    """Validate Rust ``Statistics`` without losing acquisition order.

    ``samples`` is sorted by ``(elapsed_ns, original_sample_index)`` and
    ``sample_order`` maps each sorted position back to the acquisition index.
    The archived reader accidentally required that mapping to be the identity
    permutation, which only passes when the native run happened to be already
    sorted.  Formal capture must accept valid unsorted timings while retaining
    the tie-break rule and the full vector oracle.
    """

    require(isinstance(value, dict), f"{path}: elapsed statistics are missing")
    _same(value.get("unit"), "ns", f"{path}.unit")
    raw = value.get("samples")
    require(isinstance(raw, list) and len(raw) == expected_samples,
            f"{path}.samples: expected {expected_samples} values")
    samples = [_uint(item, f"{path}.samples[{index}]", positive=True)
               for index, item in enumerate(raw)]
    order = value.get("sample_order")
    require(isinstance(order, list) and len(order) == expected_samples,
            f"{path}.sample_order: missing or wrong length")
    sample_order = [_uint(item, f"{path}.sample_order[{index}]")
                    for index, item in enumerate(order)]
    _same(sorted(sample_order), list(range(expected_samples)),
          f"{path}.sample_order: not a permutation")
    for index, ((left_value, left_index), (right_value, right_index)) in enumerate(
        zip(zip(samples, sample_order), zip(samples[1:], sample_order[1:]))
    ):
        require((left_value, left_index) <= (right_value, right_index),
                f"{path}: samples/sample_order are not sorted at {index}")

    nearest = lambda rank: samples[min(expected_samples - 1,
                                       (rank * expected_samples + 99) // 100 - 1)]
    midpoint = lambda left, right: left // 2 + right // 2 + (left % 2 + right % 2) // 2
    _same(value.get("min"), samples[0], f"{path}.min")
    _same(value.get("p50"), midpoint(samples[(expected_samples - 1) // 2],
                                      samples[expected_samples // 2]), f"{path}.p50")
    _same(value.get("p95"), nearest(95), f"{path}.p95")
    _same(value.get("p99"), nearest(99), f"{path}.p99")
    _same(value.get("max"), samples[-1], f"{path}.max")
    mean_expected = sum(samples) / expected_samples
    mean = value.get("mean")
    require(type(mean) in (int, float) and math.isfinite(float(mean)) and
            math.isclose(float(mean), mean_expected, rel_tol=0.0, abs_tol=1e-6),
            f"{path}.mean: does not match sample vector")
    running_mean = 0.0
    squared = 0.0
    for index, item in enumerate(samples):
        count = index + 1
        delta = item - running_mean
        next_mean = running_mean + delta / count
        squared += delta * (item - next_mean)
        running_mean = next_mean
    deviation = math.sqrt(squared / (expected_samples - 1)) if expected_samples > 1 else 0.0
    _number_close(value.get("standard_deviation"), deviation,
                  f"{path}.standard_deviation")
    interval = value.get("confidence_interval_95")
    require(isinstance(interval, dict), f"{path}.confidence_interval_95: missing")
    _same(interval.get("method"), "two-sided Student's t interval for the mean",
          f"{path}.confidence_interval_95.method")
    critical = _LEGACY._student_t_critical_95(expected_samples - 1)
    margin = critical * deviation / math.sqrt(expected_samples) if expected_samples > 1 else 0.0
    _number_close(interval.get("lower"), max(0.0, mean_expected - margin),
                  f"{path}.confidence_interval_95.lower")
    _number_close(interval.get("upper"), mean_expected + margin,
                  f"{path}.confidence_interval_95.upper")
    return samples


def _number_close(actual: Any, expected: float, path: str) -> None:
    require(type(actual) in (int, float) and math.isfinite(float(actual)) and
            math.isclose(float(actual), expected, rel_tol=0.0, abs_tol=1e-6),
            f"{path}: differs from sample vector")


def reconstruct_acquisition_samples(statistics: dict[str, Any]) -> list[int]:
    """Return elapsed values in original acquisition order after strict checks."""

    values = _strict_stats(statistics, "statistics", len(statistics.get("samples", [])))
    order = statistics["sample_order"]
    result = [0] * len(values)
    for sorted_index, acquisition_index in enumerate(order):
        result[acquisition_index] = values[sorted_index]
    return result


_LEGACY._stats = _strict_stats
validate_report = _LEGACY.validate_report


def _origin(packet: Path) -> dict[str, Any]:
    origin_path = packet / "origin.json"
    origin = read(origin_path)
    _finite(origin, str(origin_path))
    require(isinstance(origin, dict), f"{origin_path}: origin is not an object")
    _same(origin.get("base"), BASE, f"{origin_path}.base")
    normative = origin.get("normative")
    unrelated = origin.get("unrelated")
    require(isinstance(normative, dict) and normative, f"{origin_path}.normative: missing")
    require(isinstance(unrelated, dict) and unrelated, f"{origin_path}.unrelated: missing")
    # Origin custody is useful only when its hashes are independently checked;
    # accepting a self-authored hash map would make the base receipt circular.
    for group_name, group in (("normative", normative), ("unrelated", unrelated)):
        for relative, digest in group.items():
            require(isinstance(relative, str) and relative and
                    not Path(relative).is_absolute() and ".." not in Path(relative).parts,
                    f"{origin_path}.{group_name}: invalid path {relative!r}")
            _hash(digest, f"{origin_path}.{group_name}.{relative}")
            candidate = ROOT / relative
            require(candidate.is_file() and not candidate.is_symlink(),
                    f"{origin_path}.{group_name}.{relative}: current file is missing")
            _same(sha(candidate), digest, f"{origin_path}.{group_name}.{relative}")
    return origin


def _source_freeze(packet: Path, origin: dict[str, Any]) -> dict[str, Any]:
    freeze_path = packet / "freeze-baseline.json"
    freeze = read(freeze_path)
    _finite(freeze, str(freeze_path))
    require(isinstance(freeze, dict), f"{freeze_path}: freeze is not an object")
    _same(freeze.get("stage"), STAGE, f"{freeze_path}.stage")
    source = freeze.get("source")
    require(isinstance(source, dict) and source, f"{freeze_path}.source: missing")
    driver_descriptor = freeze.get("driver")
    origin_descriptor = freeze.get("origin")
    driver_path = _descriptor(driver_descriptor, packet, f"{freeze_path}.driver")
    origin_path = _descriptor(origin_descriptor, packet, f"{freeze_path}.origin")
    _same(driver_path, packet / "driver.py", f"{freeze_path}.driver.path")
    _same(origin_path, packet / "origin.json", f"{freeze_path}.origin.path")
    # Every frozen inventory entry is checked.  The three owned source files
    # are retained in the packet; the rest are checked against the immutable
    # current checkout, matching the driver's custody model.
    for relative, digest in source.items():
        require(isinstance(relative, str) and relative and
                not Path(relative).is_absolute() and ".." not in Path(relative).parts,
                f"{freeze_path}.source: invalid path {relative!r}")
        _hash(digest, f"{freeze_path}.source.{relative}")
        frozen = packet / "sources" / STAGE / relative
        current = ROOT / relative
        candidate = frozen if frozen.exists() else current
        require(candidate.is_file() and not candidate.is_symlink(),
                f"{freeze_path}.source.{relative}: retained/current file is missing")
        _same(sha(candidate), digest, f"{freeze_path}.source.{relative}")
    return {"freeze": freeze, "freeze_path": freeze_path, "source": source,
            "driver": driver_path, "origin": origin}


def _receipt_file(receipt_path: Path, label: str, expected_exit: int) -> dict[str, Any]:
    require(receipt_path.is_file() and not receipt_path.is_symlink(),
            f"{label}: receipt is missing")
    receipt = read(receipt_path)
    require(isinstance(receipt, dict), f"{label}: receipt is not an object")
    _same(receipt.get("exit_code"), expected_exit, f"{label}.exit_code")
    _same(receipt.get("error"), None, f"{label}.error")
    _descriptor(receipt.get("log"), receipt_path.parent, f"{label}.log")
    return receipt


def _receipt_freeze(receipt: dict[str, Any], freeze_path: Path, label: str) -> None:
    freeze_descriptor = receipt.get("freeze")
    require(isinstance(freeze_descriptor, dict), f"{label}.freeze: missing")
    _same(freeze_descriptor.get("path"), str(freeze_path), f"{label}.freeze.path")
    _same(freeze_descriptor.get("bytes"), freeze_path.stat().st_size,
          f"{label}.freeze.bytes")
    _same(freeze_descriptor.get("sha256"), sha(freeze_path), f"{label}.freeze.sha256")


def _prepare_binding(packet: Path) -> dict[str, Any]:
    packet = packet.resolve()
    origin = _origin(packet)
    frozen = _source_freeze(packet, origin)
    build_path = packet / "build-baseline.json"
    build = read(build_path)
    require(isinstance(build, dict), f"{build_path}: build receipt is malformed")
    binary_descriptor = build.get("binary")
    binary_path = _descriptor(binary_descriptor, packet, f"{build_path}.binary",
                              allow_missing=True)
    _same(binary_path.name, "litchi-perf-baseline", f"{build_path}.binary.name")
    if not binary_path.exists():
        cleanup_path = packet / "cleanup.json"
        cleanup = read(cleanup_path)
        require(isinstance(cleanup, dict), f"{cleanup_path}: cleanup is malformed")
        _same(cleanup.get("status"), "pass", f"{cleanup_path}.status")
        require(_contains_descriptor(cleanup, binary_descriptor),
                f"{cleanup_path}: retained baseline binary descriptor is missing")
    build_receipt_descriptor = build.get("receipt")
    build_receipt_path = _descriptor(build_receipt_descriptor, packet,
                                     f"{build_path}.receipt")
    _same(build_receipt_path.name, "receipt.json", f"{build_path}.receipt.filename")
    _same(build_receipt_path.parent.name, "build-baseline",
          f"{build_path}.receipt.parent")
    build_receipt = _receipt_file(build_receipt_path, f"{build_path}.receipt", 0)
    _receipt_freeze(build_receipt, frozen["freeze_path"], f"{build_path}.receipt")

    quality_path = packet / "quality-reuse.json"
    quality = read(quality_path)
    require(isinstance(quality, dict), f"{quality_path}: quality receipt is malformed")
    _same(quality.get("status"), "pass", f"{quality_path}.status")
    _same(quality.get("reused"), True, f"{quality_path}.reused")
    _same(quality.get("source_matches_exactly"), True,
          f"{quality_path}.source_matches_exactly")
    _same(quality.get("source_count"), len(frozen["source"]),
          f"{quality_path}.source_count")
    _same(quality.get("gates"), list(QUALITY_GATES), f"{quality_path}.gates")
    prior = quality.get("prior_packet")
    _same(Path(prior).resolve(), ROOT / "docs/performance/results/change-0834",
          f"{quality_path}.prior_packet")
    scope = quality.get("scope")
    require(isinstance(scope, str) and "Byte-identical source" in scope and
            "No new full-suite run claimed" in scope,
            f"{quality_path}.scope: reuse scope is not disclosed")
    inputs = quality.get("inputs")
    require(isinstance(inputs, dict) and inputs, f"{quality_path}.inputs: missing")
    for key, descriptor in inputs.items():
        _descriptor(descriptor, packet, f"{quality_path}.inputs.{key}")
    for key in ("freeze-repaired-v3.json", "quality-v3.json", "audit.json", "seal.json"):
        require(key in inputs, f"{quality_path}.inputs: missing {key}")
    return {**frozen, "origin": origin, "build": build,
            "binary": binary_descriptor, "binary_path": binary_path,
            "build_receipt": build_receipt, "quality": quality}


def _receipt_contract(receipt: dict[str, Any], binding: dict[str, Any], label: str,
                      case_arg: str, samples: int = 1, warmup: int = 0,
                      cache_states: str = "warm,cold-verified") -> None:
    _receipt_freeze(receipt, binding["freeze_path"], label)
    argv = receipt.get("argv")
    require(isinstance(argv, list), f"{label}.argv: missing")
    binary = binding["binary"]
    require(isinstance(binary, dict) and isinstance(binary.get("path"), str),
            f"{label}: binary path is missing")
    _same(argv.count(binary["path"]), 1, f"{label}.argv: binary binding")
    require("--case" in argv, f"{label}.argv: missing --case")
    case_index = argv.index("--case")
    require(case_index + 1 < len(argv), f"{label}.argv: --case has no value")
    _same(argv[case_index + 1], case_arg, f"{label}.argv.--case")
    for option, expected in (("--samples", str(samples)), ("--warmup", str(warmup)),
                             ("--filesystem-cache", cache_states)):
        require(argv.count(option) == 1, f"{label}.argv: missing/duplicate {option}")
        index = argv.index(option)
        require(index + 1 < len(argv), f"{label}.argv: {option} has no value")
        _same(argv[index + 1], expected, f"{label}.argv.{option}")
    require("--filesystem-root" in argv and "--json" in argv,
            f"{label}.argv: selected root/report binding is missing")


def _report_descriptor(row: dict[str, Any], packet: Path, label: str,
                       expected_name: str) -> tuple[Path, str]:
    descriptor = row.get("report")
    require(isinstance(descriptor, dict), f"{label}.report: missing")
    path = _descriptor(descriptor, packet, f"{label}.report")
    _same(path.name, expected_name, f"{label}.report.filename")
    _same(sha(path), descriptor.get("sha256"), f"{label}.report.sha256")
    return path, descriptor["sha256"]


def _report_identity(checked: dict[str, Any], binding: dict[str, Any], label: str) -> None:
    identity = checked.get("binary")
    require(isinstance(identity, dict), f"{label}: checked binary identity missing")
    _same(identity.get("binary_sha256"), binding["binary"].get("sha256"),
          f"{label}.binary_identity.binary_sha256")
    _same(identity.get("binary_bytes"), binding["binary"].get("bytes"),
          f"{label}.binary_identity.binary_bytes")
    _same(identity.get("path"), binding["binary"].get("path"),
          f"{label}.binary_identity.path")


def _validate_single(report_path: Path, case: str, states: tuple[str, ...],
                     samples: int, warmup: int, binding: dict[str, Any],
                     expected_environment: dict[str, Any] | None,
                     seen_pids: set[int]) -> dict[str, Any]:
    checked = validate_report(report_path, case, states, samples, warmup,
                              binding["binary"], expected_environment=expected_environment,
                              seen_pids=seen_pids)
    _report_identity(checked, binding, str(report_path))
    _bind_sample_order(checked["report"], states, samples, str(report_path))
    return checked


def _validate_pair(report_path: Path, cases: tuple[str, str], binding: dict[str, Any],
                   expected_environment: dict[str, Any] | None,
                   seen_pids: set[int]) -> dict[str, Any]:
    # Keep bundle decomposition in one place so each case is checked by the
    # same strict report validator as a single-case qualification.
    value = read(report_path)
    _finite(value, str(report_path))
    require(isinstance(value, dict), f"{report_path}: pair report is not an object")
    configuration = value.get("configuration")
    require(isinstance(configuration, dict), f"{report_path}.configuration: missing")
    _same(configuration.get("cases"), list(cases), f"{report_path}.configuration.cases")
    evidence = value.get("filesystem_evidence")
    results = value.get("results")
    require(isinstance(evidence, list) and len(evidence) == 2,
            f"{report_path}: pair evidence count differs")
    require(isinstance(results, list) and len(results) == 4,
            f"{report_path}: pair result count differs")
    outputs: dict[str, dict[str, Any]] = {}
    environment = expected_environment
    for case in cases:
        single = dict(value)
        single_configuration = dict(configuration)
        single_configuration["cases"] = [case]
        single["configuration"] = single_configuration
        single["filesystem_evidence"] = [item for item in evidence
                                           if isinstance(item, dict) and item.get("case") == case]
        single["results"] = [item for item in results
                              if isinstance(item, dict) and item.get("case") == case]
        require(len(single["filesystem_evidence"]) == 1,
                f"{report_path}: missing/duplicate pair evidence for {case}")
        require(len(single["results"]) == 2,
                f"{report_path}: missing/duplicate pair results for {case}")
        checked = _validate_single_dict(single, case, binding, environment, seen_pids,
                                        str(report_path))
        environment = environment or checked["environment"]
        outputs[case] = checked["outputs"]
    return {"environment": environment, "outputs": outputs}


def _validate_single_dict(value: dict[str, Any], case: str, binding: dict[str, Any],
                          expected_environment: dict[str, Any] | None,
                          seen_pids: set[int], label: str) -> dict[str, Any]:
    checked = validate_report(value, case, STATES, 1, 0, binding["binary"],
                              expected_environment=expected_environment,
                              seen_pids=seen_pids)
    _report_identity(checked, binding, label)
    _bind_sample_order(checked["report"], STATES, 1, label)
    return checked


def _bind_sample_order(report: dict[str, Any], states: tuple[str, ...],
                       samples: int, label: str) -> None:
    """Bind sorted statistics back to the retained chronological evidence."""

    evidence = report.get("filesystem_evidence")
    results = report.get("results")
    require(isinstance(evidence, list) and isinstance(results, list),
            f"{label}: sample-order evidence is missing")
    for state in states:
        state_evidence = [item for item in evidence[0].get("samples", [])
                          if isinstance(item, dict) and item.get("cache_state") == state]
        # Bundle decomposition supplies one evidence record per case; a single
        # report has exactly one.  For this helper, the caller passes a report
        # already narrowed to one case.
        if not state_evidence:
            state_evidence = [item for record in evidence
                              if isinstance(record, dict)
                              for item in record.get("samples", [])
                              if isinstance(item, dict) and item.get("cache_state") == state]
        require(len(state_evidence) == samples,
                f"{label}: evidence count for {state} differs from {samples}")
        by_index: dict[int, dict[str, Any]] = {}
        for item in state_evidence:
            sample_index = _uint(item.get("sample_index"),
                                 f"{label}.{state}.sample_index")
            require(sample_index not in by_index,
                    f"{label}.{state}: duplicate evidence sample index {sample_index}")
            by_index[sample_index] = item
        _same(sorted(by_index), list(range(samples)),
              f"{label}.{state}: evidence sample-index permutation")
        result = next((item for item in results
                       if isinstance(item, dict) and item.get("cache_state") == state), None)
        require(isinstance(result, dict), f"{label}.{state}: timed result is missing")
        statistics = result.get("elapsed_ns")
        sorted_values = _strict_stats(statistics, f"{label}.{state}.elapsed_ns", samples)
        order = statistics["sample_order"]
        for sorted_index, acquisition_index in enumerate(order):
            sample = by_index[acquisition_index]
            _same(sample.get("elapsed_ns"), sorted_values[sorted_index],
                  f"{label}.{state}: sorted statistic/evidence mismatch at {sorted_index}")


def _validate_plan(packet: Path) -> list[dict[str, Any]]:
    path = packet / "measurement-plan.json"
    plan = read(path)
    _finite(plan, str(path))
    require(isinstance(plan, dict), f"{path}: plan is not an object")
    _same(plan.get("base"), BASE, f"{path}.base")
    _same(plan.get("purpose"), "Descriptive current-source route/cache baseline; no production optimization",
          f"{path}.purpose")
    _same(plan.get("expected_reports"), REPORTS, f"{path}.expected_reports")
    _same(plan.get("expected_samples"), SAMPLES, f"{path}.expected_samples")
    rows = plan.get("rows")
    require(isinstance(rows, list) and len(rows) == REPORTS, f"{path}.rows: expected {REPORTS}")
    require(sum(row.get("samples", 0) for row in rows if isinstance(row, dict)) == SAMPLES,
            f"{path}.rows: sample total differs")
    keys: set[tuple[int, str, str]] = set()
    for index, row in enumerate(rows):
        label = f"{path}.rows[{index}]"
        require(isinstance(row, dict), f"{label}: malformed")
        block = _uint(row.get("block"), f"{label}.block")
        require(block < 6, f"{label}.block: expected six counterbalanced blocks")
        case = row.get("case")
        state = row.get("cache_state")
        require(case in CASES and state in STATES, f"{label}: unknown case/state")
        key = (block, case, state)
        require(key not in keys, f"{label}: duplicate block/case/state")
        keys.add(key)
        _same(row.get("samples"), CAPTURE_SAMPLES, f"{label}.samples")
        _same(row.get("warmup"), CAPTURE_WARMUP, f"{label}.warmup")
    require(len(keys) == 72, f"{path}: block/case/state coverage differs")
    for block in range(6):
        require({(b, c, s) for b, c, s in keys if b == block} ==
                {(block, case, state) for case in CASES for state in STATES},
                f"{path}: block {block} does not cover every case/state")
    return rows


def validate_manifest(path: Path) -> dict[str, Any]:
    """Validate all seven fresh warm/cold qualification commands."""

    path = path.resolve()
    packet = path.parent
    value = read(path)
    _finite(value, str(path))
    require(isinstance(value, dict), f"{path}: qualification manifest is not an object")
    _same(value.get("status"), "commands_pass", f"{path}.status")
    rows = value.get("rows")
    require(isinstance(rows, list) and len(rows) == len(CASES) + 1,
            f"{path}.rows: expected six cases plus paired OPC save")
    binding = _prepare_binding(packet)
    expected_environment: dict[str, Any] | None = None
    seen_pids: set[int] = set()
    checked: list[dict[str, Any]] = []
    expected_rows = [(case, f"qualification-{index:02}", (case,))
                     for index, case in enumerate(CASES)]
    expected_rows.append((",".join(CASES[2:4]), "qualification-opc-pair", tuple(CASES[2:4])))
    for index, (row, (case_arg, command_label, row_cases)) in enumerate(zip(rows, expected_rows)):
        label = f"{path}.rows[{index}]"
        require(isinstance(row, dict), f"{label}: malformed")
        _same(row.get("case"), case_arg, f"{label}.case")
        _same(row.get("state"), "warm,cold-verified", f"{label}.state")
        _same(row.get("exit_code"), 0, f"{label}.exit_code")
        receipt_descriptor = row.get("receipt")
        receipt_path = _descriptor(receipt_descriptor, packet, f"{label}.receipt")
        _same(receipt_path.name, "receipt.json", f"{label}.receipt.filename")
        _same(receipt_path.parent.name, command_label, f"{label}.receipt.parent")
        receipt = _receipt_file(receipt_path, f"{label}.receipt", 0)
        _receipt_contract(receipt, binding, f"{label}.receipt", case_arg)
        report_path, report_digest = _report_descriptor(
            row, packet, label, f"{command_label}.json"
        )
        if len(row_cases) == 1:
            checked_report = _validate_single(report_path, row_cases[0], STATES, 1, 0,
                                               binding, expected_environment, seen_pids)
            expected_environment = expected_environment or checked_report["environment"]
            outputs = {row_cases[0]: checked_report["outputs"]}
        else:
            pair = _validate_pair(report_path, row_cases, binding, expected_environment, seen_pids)
            _LEGACY._pair_output_oracle(pair, row_cases, STATES, label)
            expected_environment = expected_environment or pair["environment"]
            outputs = pair["outputs"]
        checked.append({"index": index, "case": case_arg, "report": str(report_path),
                        "report_sha256": report_digest, "outputs": outputs})
    # Six single-case reports contain two cache-state children each.  The
    # paired OPC report contains two cases and two states, adding four more.
    require(len(seen_pids) == 16, f"{path}: expected sixteen fresh qualification children")
    return {"status": "pass", "qualification_valid": True, "reports": checked,
            "validated_sample_count": 16, "claim_authorized": False,
            "performance_claim": "none"}


def validate_capture(path: Path) -> dict[str, Any]:
    """Validate all 72 formal reports and their 2,160 retained samples."""

    path = path.resolve()
    packet = path.parent
    value = read(path)
    _finite(value, str(path))
    require(isinstance(value, dict), f"{path}: capture manifest is not an object")
    _same(value.get("status"), "commands_pass", f"{path}.status")
    rows = value.get("rows")
    require(isinstance(rows, list) and len(rows) == REPORTS,
            f"{path}.rows: expected {REPORTS} native rows")
    _same(value.get("report_count"), REPORTS, f"{path}.report_count")
    _same(value.get("sample_count"), SAMPLES, f"{path}.sample_count")
    plan_rows = _validate_plan(packet)
    binding = _prepare_binding(packet)
    expected_environment: dict[str, Any] | None = None
    seen_pids: set[int] = set()
    checked: list[dict[str, Any]] = []
    for index, (row, planned) in enumerate(zip(rows, plan_rows)):
        label = f"{path}.rows[{index}]"
        require(isinstance(row, dict), f"{label}: malformed")
        _same(row.get("plan"), planned, f"{label}.plan")
        result = row.get("result")
        require(isinstance(result, dict), f"{label}.result: missing")
        case = planned["case"]
        state = planned["cache_state"]
        _same(result.get("case"), case, f"{label}.result.case")
        _same(result.get("state"), state, f"{label}.result.state")
        _same(result.get("exit_code"), 0, f"{label}.result.exit_code")
        command_label = f"native-{index:03}"
        receipt_path = _descriptor(result.get("receipt"), packet, f"{label}.result.receipt")
        _same(receipt_path.name, "receipt.json", f"{label}.receipt.filename")
        _same(receipt_path.parent.name, command_label, f"{label}.receipt.parent")
        receipt = _receipt_file(receipt_path, f"{label}.receipt", 0)
        _receipt_contract(receipt, binding, f"{label}.receipt", case,
                          CAPTURE_SAMPLES, CAPTURE_WARMUP, state)
        report_path, report_digest = _report_descriptor(
            result, packet, label, f"{command_label}.json"
        )
        checked_report = _validate_single(report_path, case, (state,),
                                          CAPTURE_SAMPLES, CAPTURE_WARMUP, binding,
                                          expected_environment, seen_pids)
        expected_environment = expected_environment or checked_report["environment"]
        checked.append({"index": index, "block": planned["block"], "case": case,
                        "state": state, "report": str(report_path),
                        "report_sha256": report_digest,
                        "sample_count": CAPTURE_SAMPLES})
    require(len(seen_pids) == SAMPLES, f"{path}: fresh-child PID count differs from sample count")
    return {"status": "pass", "capture_valid": True, "reports": checked,
            "validated_sample_count": SAMPLES, "claim_authorized": False,
            "performance_claim": "none"}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("manifest", nargs="?", type=Path)
    parser.add_argument("--qualification", type=Path,
                        help="validate qualification.json")
    parser.add_argument("--capture", type=Path,
                        help="validate capture.json")
    parser.add_argument("--write", type=Path,
                        help="also write the validation object to this path")
    parser.add_argument("--check", action="store_true",
                        help="compare an existing validation object instead of writing it")
    args = parser.parse_args()
    require(not (args.qualification and args.capture),
            "choose exactly one of --qualification and --capture")
    path = args.qualification or args.capture or args.manifest or (HERE / "qualification.json")
    capture_mode = args.capture is not None or path.name == "capture.json"
    try:
        result = validate_capture(path.resolve()) if capture_mode else validate_manifest(path.resolve())
    except Exception as error:
        result = {"status": "invalid", "qualification_valid": False,
                  "capture_valid": False, "claim_authorized": False,
                  "performance_claim": "none", "error": str(error)}
        exit_code = 1
    else:
        exit_code = 0
    encoded = json.dumps(result, sort_keys=True, indent=2) + "\n"
    destination = args.write.resolve() if args.write else path.parent / (
        "capture-validation.json" if capture_mode else "qualification-validation.json"
    )
    if args.check:
        require(destination.is_file() and not destination.is_symlink(),
                f"validation output is missing: {destination}")
        _same(read(destination), result, f"{destination}: validation differs")
    elif args.write:
        require(destination.parent.is_dir(),
                f"validation output parent is missing: {destination.parent}")
        require(not destination.exists(), f"validation output already exists: {destination}")
        destination.write_text(encoded, encoding="utf-8")
    print(encoded, end="")
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
