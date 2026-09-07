#!/usr/bin/env python3
"""Validate one change-0458 ODP append attribution report.

The capture driver intentionally keeps this oracle independent from the Rust
runner.  It validates the report's frozen schema, corpus custody, correctness
gates, timing shape, and allocator evidence; it does not recompute timings or
interpret them as a performance result.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
from typing import Any, Mapping, Sequence


ROOT = Path(__file__).resolve().parent
SCHEMA = "litchi-odp-append-attribution-v1"
PHASE_NAMES = (
    "snapshot_open",
    "transaction",
    "add",
    "commit",
    "publication",
)
PHASE_SCOPES = (
    "Snapshot::from_bytes(input Vec)",
    "Snapshot::transaction()",
    "Transaction::add(title, body)",
    "Transaction::commit() including Patch construction and readback",
    "HashingDiscardSink::write_all(committed snapshot bytes)",
)
PRE_FLIGHT_GATES = (
    "source_manifest_bindings_verified",
    "output_manifest_bindings_verified",
    "untouched_members_verified",
    "opaque_member_compressed_identity_verified",
    "patch_replay_verified",
    "inverse_patch_verified",
    "stale_source_refusal_verified",
    "exact_noop_verified",
)
RUNTIME_GATES = (
    "source_bytes_identity_verified",
    "candidate_bytes_identity_verified",
    "commit_changed_verified",
    "patch_non_noop_verified",
    "sink_digest_verified",
    "sink_length_verified",
)
ALLOCATOR_FIELDS = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "failed_allocation_calls",
    "allocated_bytes",
    "deallocated_bytes",
    "live_bytes_before",
    "live_bytes_after",
    "peak_live_bytes_before",
    "peak_live_bytes_after",
    "region_peak_live_bytes",
)
ALLOCATOR_SAMPLE_KEYS = frozenset(("status", "scope", *ALLOCATOR_FIELDS))
PHASE_KEYS = frozenset(("name", "elapsed_ns", "allocation_metrics"))
# The Rust report omits unavailable normal-mode counters in both phase and
# lifecycle measurements; measured allocator samples remain mandatory.
PHASE_KEYS_NORMAL = frozenset(("name", "elapsed_ns"))
SINK_KEYS = frozenset(
    ("accepted_bytes", "write_calls", "largest_write", "write_size_buckets")
)
SINK_BUCKET_KEYS = frozenset(
    (
        "bytes_0",
        "bytes_1_to_512",
        "bytes_513_to_4096",
        "bytes_4097_to_16384",
        "bytes_16385_to_65536",
        "bytes_over_65536",
    )
)
MAX_U64 = (1 << 64) - 1

LIFECYCLE_SCOPE = (
    "input/title/body/sink construction outside; one Instant includes "
    "Snapshot::from_bytes, Snapshot::transaction, Transaction::add, "
    "Transaction::commit, and sink write_all; all hashes, gates, report "
    "assembly, and drops outside"
)
PHASES_SCOPE = (
    "input/title/body/sink construction outside; lifecycle_ns encloses five "
    "harness phase clocks and their boundary instrumentation; each phase "
    "clock includes only its named public call; hashes, gates, report "
    "assembly, and drops outside"
)


class OracleError(ValueError):
    """Raised when a report does not satisfy the frozen attribution contract."""


def _fail(path: str, message: str) -> None:
    raise OracleError(f"{path}: {message}")


def _mapping(value: Any, path: str) -> Mapping[str, Any]:
    if not isinstance(value, dict):
        _fail(path, "must be an object")
    return value


def _sequence(value: Any, path: str) -> Sequence[Any]:
    if not isinstance(value, list):
        _fail(path, "must be an array")
    return value


def _exact_keys(value: Mapping[str, Any], keys: Sequence[str] | frozenset[str], path: str) -> None:
    expected = set(keys)
    observed = set(value)
    missing = sorted(expected - observed)
    extra = sorted(observed - expected)
    if missing or extra:
        detail = []
        if missing:
            detail.append("missing " + ", ".join(missing))
        if extra:
            detail.append("unexpected " + ", ".join(extra))
        _fail(path, "; ".join(detail))


def _required(value: Mapping[str, Any], key: str, path: str) -> Any:
    if key not in value:
        _fail(path, f"missing {key}")
    return value[key]


def _string(value: Any, path: str, *, nonempty: bool = True) -> str:
    if not isinstance(value, str) or (nonempty and not value):
        _fail(path, "must be a non-empty string")
    return value


def _bool(value: Any, path: str) -> bool:
    if type(value) is not bool:
        _fail(path, "must be a boolean")
    return value


def _u64(value: Any, path: str, *, positive: bool = False) -> int:
    # bool is an int subclass, but is never a valid JSON integer here.
    if type(value) is not int or value < 0 or value > MAX_U64:
        _fail(path, "must be an unsigned 64-bit integer")
    if positive and value == 0:
        _fail(path, "must be greater than zero")
    return value


def _sha256(value: Any, path: str) -> str:
    value = _string(value, path)
    if len(value) != 64 or any(character not in "0123456789abcdef" for character in value):
        _fail(path, "must be a lowercase SHA-256 digest")
    return value


def _read_json(path: Path, label: str) -> Any:
    try:
        return json.loads(path.read_text())
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        _fail(label, f"cannot read JSON ({error})")


def _file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError as error:
        _fail(str(path), f"cannot read ({error})")
    return digest.hexdigest()


def _path_within_root(relative_path: str, path: str) -> Path:
    relative = Path(relative_path)
    if relative.is_absolute():
        _fail(path, "must be relative to the change directory")
    target = (ROOT / relative).resolve()
    try:
        target.relative_to(ROOT.resolve())
    except ValueError:
        _fail(path, "escapes the change directory")
    return target


def _prior_expectation(shape: str) -> tuple[Mapping[str, Any], Mapping[str, Any]]:
    """Return the authenticated prior corpus and its ODP identity summary."""
    bindings_path = ROOT / "prior-control-bindings.json"
    bindings = _mapping(_read_json(bindings_path, "prior-control-bindings.json"), "bindings")
    shapes = _mapping(_required(bindings, "shapes", "bindings"), "bindings.shapes")
    if set(shapes) != {"tiny", "medium", "large"}:
        _fail("bindings.shapes", "must contain exactly tiny, medium, and large")
    binding = _mapping(_required(shapes, shape, f"bindings.shapes[{shape!r}]"), f"bindings.shapes[{shape!r}]")
    _exact_keys(binding, ("original_path", "path", "sha256", "bytes"), f"bindings.shapes[{shape!r}]")
    relative = _string(binding["path"], f"bindings.shapes[{shape!r}].path")
    expected_relative = f"prior-control/{shape}.json"
    if relative != expected_relative:
        _fail(f"bindings.shapes[{shape!r}].path", f"must be {expected_relative!r}")
    prior_path = _path_within_root(relative, f"bindings.shapes[{shape!r}].path")
    expected_bytes = _u64(binding["bytes"], f"bindings.shapes[{shape!r}].bytes", positive=True)
    expected_sha = _sha256(binding["sha256"], f"bindings.shapes[{shape!r}].sha256")
    try:
        actual_bytes = prior_path.stat().st_size
    except OSError as error:
        _fail(relative, f"cannot stat authenticated prior report ({error})")
    if actual_bytes != expected_bytes:
        _fail(relative, f"byte count {actual_bytes} does not match binding {expected_bytes}")
    actual_sha = _file_sha256(prior_path)
    if actual_sha != expected_sha:
        _fail(relative, f"SHA-256 {actual_sha} does not match binding {expected_sha}")

    prior = _mapping(_read_json(prior_path, relative), relative)
    results = _sequence(_required(prior, "results", relative), f"{relative}.results")
    if len(results) != 1:
        _fail(f"{relative}.results", "must contain exactly one result")
    result = _mapping(results[0], f"{relative}.results[0]")
    corpus = _mapping(_required(result, "corpus", f"{relative}.results[0]"), f"{relative}.results[0].corpus")
    source = _mapping(_required(result, "source", f"{relative}.results[0]"), f"{relative}.results[0].source")
    odp = _mapping(
        _required(source, "odp_append", f"{relative}.results[0].source"),
        f"{relative}.results[0].source.odp_append",
    )
    for key in (
        "corpus_generator",
        "shape",
        "source_archive_sha256",
        "source_archive_bytes",
        "output_archive_sha256",
        "output_archive_bytes",
    ):
        if key not in odp:
            _fail(f"{relative}.results[0].source.odp_append", f"missing {key}")
    if odp["shape"] != shape:
        _fail(f"{relative}.results[0].source.odp_append.shape", "does not match lane shape")
    source_sha = _sha256(odp["source_archive_sha256"], f"{relative}.source.odp_append.source_archive_sha256")
    output_sha = _sha256(odp["output_archive_sha256"], f"{relative}.source.odp_append.output_archive_sha256")
    source_bytes = _u64(odp["source_archive_bytes"], f"{relative}.source.odp_append.source_archive_bytes", positive=True)
    output_bytes = _u64(odp["output_archive_bytes"], f"{relative}.source.odp_append.output_archive_bytes", positive=True)
    if corpus.get("shape") != shape:
        _fail(f"{relative}.results[0].corpus.shape", "does not match lane shape")
    if corpus.get("generator") != odp["corpus_generator"]:
        _fail(f"{relative}.results[0].corpus.generator", "does not match odp_append corpus_generator")
    if corpus.get("archive_sha256") != source_sha or corpus.get("archive_bytes") != source_bytes:
        _fail(f"{relative}.results[0].corpus", "source archive identity is inconsistent")
    return corpus, {
        "corpus_generator": odp["corpus_generator"],
        "source_sha256": source_sha,
        "source_bytes": source_bytes,
        "output_sha256": output_sha,
        "output_bytes": output_bytes,
    }


def _validate_lane_and_protocol(lane: Mapping[str, Any], protocol: Mapping[str, Any]) -> tuple[str, str, str, str, str, int, int]:
    for key in ("id", "repeat", "instrumentation", "shape", "scope"):
        if key not in lane:
            _fail("lane", f"missing {key}")
    lane_id = _string(lane["id"], "lane.id")
    repeat = _string(lane["repeat"], "lane.repeat")
    instrumentation = _string(lane["instrumentation"], "lane.instrumentation")
    shape = _string(lane["shape"], "lane.shape")
    scope = _string(lane["scope"], "lane.scope")
    if repeat not in {"R1", "R2", "diagnostic"}:
        _fail("lane.repeat", "must be R1, R2, or diagnostic")
    if instrumentation not in {"normal", "allocator"}:
        _fail("lane.instrumentation", "must be normal or allocator")
    if shape not in {"tiny", "medium", "large"}:
        _fail("lane.shape", "must be tiny, medium, or large")
    if scope not in {"lifecycle", "phases"}:
        _fail("lane.scope", "must be lifecycle or phases")
    warmup = _u64(_required(protocol, "warmup", "protocol"), "protocol.warmup")
    configured_samples = _u64(_required(protocol, "samples", "protocol"), "protocol.samples", positive=True)
    if warmup != 3:
        _fail("protocol.warmup", "must be exactly 3")
    if repeat == "diagnostic":
        if configured_samples not in {30, 100}:
            _fail("protocol.samples", "diagnostic validation accepts only 30 or 100")
        samples = 100
    else:
        if configured_samples != 30:
            _fail("protocol.samples", "normal capture validation requires exactly 30")
        samples = 30
    return lane_id, repeat, instrumentation, shape, scope, samples, warmup


def _validate_gates(value: Any, names: Sequence[str], path: str) -> None:
    gates = _mapping(value, path)
    _exact_keys(gates, names, path)
    for name in names:
        if not _bool(gates[name], f"{path}.{name}"):
            _fail(f"{path}.{name}", "must be true")


def _validate_allocator_identity(value: Any, instrumentation: str) -> None:
    allocator = _mapping(value, "report.allocator")
    if instrumentation == "normal":
        _exact_keys(allocator, ("allocator", "instrumentation"), "report.allocator")
        if allocator["allocator"] != "Rust system allocator":
            _fail("report.allocator.allocator", "normal lane must use the Rust system allocator")
        if allocator["instrumentation"] != "none":
            _fail("report.allocator.instrumentation", "normal lane must report none")
    else:
        _exact_keys(allocator, ("allocator", "instrumentation", "counter_revision"), "report.allocator")
        if allocator["allocator"] != "CountingSystemAllocator(std::alloc::System)":
            _fail("report.allocator.allocator", "allocator lane has the wrong allocator identity")
        if allocator["instrumentation"] != "system_allocator_operation_scoped":
            _fail("report.allocator.instrumentation", "allocator lane has the wrong instrumentation identity")
        if allocator["counter_revision"] != "serialized_region_peak_v3":
            _fail("report.allocator.counter_revision", "allocator lane has the wrong counter revision")


def _validate_hex_and_identity(
    row: Mapping[str, Any],
    path: str,
    expected: Mapping[str, Any],
) -> None:
    if _sha256(_required(row, "source_sha256", path), f"{path}.source_sha256") != expected["source_sha256"]:
        _fail(f"{path}.source_sha256", "does not match the authenticated source archive")
    if _u64(_required(row, "source_bytes", path), f"{path}.source_bytes", positive=True) != expected["source_bytes"]:
        _fail(f"{path}.source_bytes", "does not match the authenticated source archive")
    if _sha256(_required(row, "candidate_sha256", path), f"{path}.candidate_sha256") != expected["output_sha256"]:
        _fail(f"{path}.candidate_sha256", "does not match the authenticated expected output")
    if _u64(_required(row, "candidate_bytes", path), f"{path}.candidate_bytes", positive=True) != expected["output_bytes"]:
        _fail(f"{path}.candidate_bytes", "does not match the authenticated expected output")
    if _sha256(_required(row, "sink_sha256", path), f"{path}.sink_sha256") != expected["output_sha256"]:
        _fail(f"{path}.sink_sha256", "does not match the authenticated expected output")


def _bucket_for(length: int) -> str:
    if length == 0:
        return "bytes_0"
    if length <= 512:
        return "bytes_1_to_512"
    if length <= 4096:
        return "bytes_513_to_4096"
    if length <= 16384:
        return "bytes_4097_to_16384"
    if length <= 65536:
        return "bytes_16385_to_65536"
    return "bytes_over_65536"


def _validate_sink(value: Any, path: str, output_bytes: int) -> None:
    sink = _mapping(value, path)
    _exact_keys(sink, SINK_KEYS, path)
    if _u64(sink["accepted_bytes"], f"{path}.accepted_bytes", positive=True) != output_bytes:
        _fail(f"{path}.accepted_bytes", "must equal candidate_bytes")
    if _u64(sink["write_calls"], f"{path}.write_calls", positive=True) != 1:
        _fail(f"{path}.write_calls", "must be exactly one")
    if _u64(sink["largest_write"], f"{path}.largest_write", positive=True) != output_bytes:
        _fail(f"{path}.largest_write", "must equal candidate_bytes")
    buckets = _mapping(sink["write_size_buckets"], f"{path}.write_size_buckets")
    _exact_keys(buckets, SINK_BUCKET_KEYS, f"{path}.write_size_buckets")
    selected = _bucket_for(output_bytes)
    for name in SINK_BUCKET_KEYS:
        count = _u64(buckets[name], f"{path}.write_size_buckets.{name}")
        expected = 1 if name == selected else 0
        if count != expected:
            _fail(f"{path}.write_size_buckets.{name}", f"must be {expected}")


def _validate_metric(value: Any, path: str) -> dict[str, int]:
    metric = _mapping(value, path)
    _exact_keys(metric, ALLOCATOR_SAMPLE_KEYS, path)
    if metric["status"] != "measured":
        _fail(f"{path}.status", "allocator reports must have measured samples")
    if metric["scope"] != "operation_global_system_allocator":
        _fail(f"{path}.scope", "has the wrong allocator metric scope")
    numbers: dict[str, int] = {}
    for name in ALLOCATOR_FIELDS:
        numbers[name] = _u64(metric[name], f"{path}.{name}")

    # The counters include reallocations in both allocated/deallocated byte
    # totals, so this is the exact live-byte conservation equation.
    if numbers["live_bytes_before"] + numbers["allocated_bytes"] != numbers["live_bytes_after"] + numbers["deallocated_bytes"]:
        _fail(path, "live-byte arithmetic is inconsistent")
    if numbers["reallocation_calls"] > numbers["allocation_calls"]:
        _fail(path, "reallocation_calls exceeds allocation_calls")
    if numbers["peak_live_bytes_after"] < numbers["peak_live_bytes_before"]:
        _fail(path, "peak live bytes decrease")
    if numbers["peak_live_bytes_after"] < max(numbers["live_bytes_before"], numbers["live_bytes_after"]):
        _fail(path, "peak live bytes do not cover operation endpoints")
    region_peak = numbers["region_peak_live_bytes"]
    if region_peak < max(numbers["live_bytes_before"], numbers["live_bytes_after"]):
        _fail(path, "region peak does not cover operation endpoints")
    if region_peak > numbers["peak_live_bytes_after"]:
        _fail(path, "region peak exceeds process peak")
    return numbers


def _validate_phase_metrics(phases: Sequence[Any], path: str, instrumentation: str) -> None:
    previous: dict[str, int] | None = None
    for index, value in enumerate(phases):
        phase = _mapping(value, f"{path}[{index}]")
        expected_keys = PHASE_KEYS if instrumentation == "allocator" else PHASE_KEYS_NORMAL
        _exact_keys(phase, expected_keys, f"{path}[{index}]")
        if phase["name"] != PHASE_NAMES[index]:
            _fail(f"{path}[{index}].name", f"must be {PHASE_NAMES[index]!r}")
        _u64(phase["elapsed_ns"], f"{path}[{index}].elapsed_ns", positive=True)
        if instrumentation == "allocator":
            current = _validate_metric(phase["allocation_metrics"], f"{path}[{index}].allocation_metrics")
            if previous is not None:
                # The phase calls are sequential and no timed phase may begin
                # with a different allocator state.  Peak values are checked
                # for continuity, never added to the elapsed phase sum.
                if current["live_bytes_before"] != previous["live_bytes_after"]:
                    _fail(f"{path}[{index}]", "live-byte phase continuity is broken")
                if current["peak_live_bytes_before"] != previous["peak_live_bytes_after"]:
                    _fail(f"{path}[{index}]", "peak-byte phase continuity is broken")
            previous = current
        elif phase.get("allocation_metrics") is not None:
            _fail(f"{path}[{index}].allocation_metrics", "normal reports must carry null allocation metrics")


def _validate_row(
    value: Any,
    index: int,
    expected: Mapping[str, Any],
    scope: str,
    instrumentation: str,
) -> None:
    path = f"report.rows[{index}]"
    row = _mapping(value, path)
    if scope == "lifecycle":
        required = {
            "sample_index",
            "lifecycle_ns",
            "source_sha256",
            "source_bytes",
            "candidate_sha256",
            "candidate_bytes",
            "sink_sha256",
            "sink",
            "runtime_gates",
        }
        if instrumentation == "allocator":
            required.add("lifecycle_allocation_metrics")
    else:
        required = {
            "sample_index",
            "lifecycle_ns",
            "phase_sum_ns",
            "boundary_gap_ns",
            "phases",
            "source_sha256",
            "source_bytes",
            "candidate_sha256",
            "candidate_bytes",
            "sink_sha256",
            "sink",
            "runtime_gates",
        }
    _exact_keys(row, required, path)
    if _u64(row["sample_index"], f"{path}.sample_index") != index:
        _fail(f"{path}.sample_index", f"must be {index}")
    lifecycle_ns = _u64(row["lifecycle_ns"], f"{path}.lifecycle_ns", positive=True)
    _validate_hex_and_identity(row, path, expected)
    _validate_gates(row["runtime_gates"], RUNTIME_GATES, f"{path}.runtime_gates")
    _validate_sink(row["sink"], f"{path}.sink", expected["output_bytes"])

    if scope == "lifecycle":
        if instrumentation == "allocator":
            _validate_metric(row["lifecycle_allocation_metrics"], f"{path}.lifecycle_allocation_metrics")
        return

    phases = _sequence(row["phases"], f"{path}.phases")
    if len(phases) != len(PHASE_NAMES):
        _fail(f"{path}.phases", "must contain exactly five phase measurements")
    _validate_phase_metrics(phases, f"{path}.phases", instrumentation)
    phase_sum = sum(_u64(phase["elapsed_ns"], f"{path}.phases[{phase_index}].elapsed_ns", positive=True) for phase_index, phase in enumerate(phases))
    observed_sum = _u64(row["phase_sum_ns"], f"{path}.phase_sum_ns")
    boundary_gap = _u64(row["boundary_gap_ns"], f"{path}.boundary_gap_ns")
    if observed_sum != phase_sum:
        _fail(f"{path}.phase_sum_ns", "does not equal the five elapsed phase values")
    if observed_sum + boundary_gap != lifecycle_ns:
        _fail(path, "phase elapsed sum plus boundary gap does not equal lifecycle_ns")


def validate(report_path: str | Path, lane: Mapping[str, Any], protocol: Mapping[str, Any]) -> dict[str, Any]:
    """Validate ``report_path`` for the supplied capture lane.

    ``lane`` and ``protocol`` are the dictionaries used by ``capture.py``.
    A failed validation raises :class:`OracleError`; successful validation
    returns a deliberately small receipt suitable for embedding in a capture
    receipt.
    """
    if not isinstance(lane, dict):
        _fail("lane", "must be an object")
    if not isinstance(protocol, dict):
        _fail("protocol", "must be an object")
    _, repeat, instrumentation, shape, scope, samples, warmup = _validate_lane_and_protocol(lane, protocol)
    expected_corpus, expected_identity = _prior_expectation(shape)

    report_file = Path(report_path)
    report = _mapping(_read_json(report_file, str(report_file)), "report")
    top_keys = (
        "schema",
        "mode",
        "scope",
        "repeat",
        "shape",
        "warmup",
        "samples",
        "checked_iteration_count",
        "corpus_generator",
        "corpus",
        "identity",
        "allocator",
        "timing_scope",
        "phase_order",
        "phases",
        "preflight_gates",
        "rows",
    )
    _exact_keys(report, top_keys, "report")
    if report["schema"] != SCHEMA:
        _fail("report.schema", f"must be {SCHEMA!r}")
    if report["mode"] != scope:
        _fail("report.mode", "does not match lane scope")
    if report["scope"] != scope:
        _fail("report.scope", "does not match lane scope")
    if report["repeat"] != repeat:
        _fail("report.repeat", "does not match lane repeat")
    if report["shape"] != shape:
        _fail("report.shape", "does not match lane shape")
    if _u64(report["warmup"], "report.warmup") != warmup:
        _fail("report.warmup", "does not match protocol")
    if _u64(report["samples"], "report.samples", positive=True) != samples:
        _fail("report.samples", "does not match lane protocol")
    if _u64(report["checked_iteration_count"], "report.checked_iteration_count") != warmup + samples:
        _fail("report.checked_iteration_count", "must equal warmup + samples")
    if report["corpus_generator"] != expected_identity["corpus_generator"]:
        _fail("report.corpus_generator", "does not match authenticated corpus")
    if report["corpus"] != expected_corpus:
        _fail("report.corpus", "does not match the authenticated prior corpus manifest")

    identity = _mapping(report["identity"], "report.identity")
    _exact_keys(identity, ("source_archive_sha256", "source_archive_bytes", "expected_output_sha256", "expected_output_bytes"), "report.identity")
    if _sha256(identity["source_archive_sha256"], "report.identity.source_archive_sha256") != expected_identity["source_sha256"]:
        _fail("report.identity.source_archive_sha256", "does not match the authenticated source archive")
    if _u64(identity["source_archive_bytes"], "report.identity.source_archive_bytes", positive=True) != expected_identity["source_bytes"]:
        _fail("report.identity.source_archive_bytes", "does not match the authenticated source archive")
    if _sha256(identity["expected_output_sha256"], "report.identity.expected_output_sha256") != expected_identity["output_sha256"]:
        _fail("report.identity.expected_output_sha256", "does not match the authenticated expected output")
    if _u64(identity["expected_output_bytes"], "report.identity.expected_output_bytes", positive=True) != expected_identity["output_bytes"]:
        _fail("report.identity.expected_output_bytes", "does not match the authenticated expected output")

    _validate_allocator_identity(report["allocator"], instrumentation)
    if report["timing_scope"] != (LIFECYCLE_SCOPE if scope == "lifecycle" else PHASES_SCOPE):
        _fail("report.timing_scope", "does not match the selected timing mode")
    phase_order = _sequence(report["phase_order"], "report.phase_order")
    if phase_order != list(PHASE_NAMES):
        _fail("report.phase_order", "must list the five phases in execution order")
    descriptions = _sequence(report["phases"], "report.phases")
    if len(descriptions) != len(PHASE_NAMES):
        _fail("report.phases", "must contain exactly five descriptions")
    for index, description_value in enumerate(descriptions):
        description = _mapping(description_value, f"report.phases[{index}]")
        _exact_keys(description, ("name", "scope"), f"report.phases[{index}]")
        if description["name"] != PHASE_NAMES[index] or description["scope"] != PHASE_SCOPES[index]:
            _fail(f"report.phases[{index}]", "does not match the frozen phase description")
    _validate_gates(report["preflight_gates"], PRE_FLIGHT_GATES, "report.preflight_gates")

    rows = _sequence(report["rows"], "report.rows")
    if len(rows) != samples:
        _fail("report.rows", f"must contain exactly {samples} rows")
    for index, row in enumerate(rows):
        _validate_row(row, index, expected_identity, scope, instrumentation)
    return {
        "status": "pass",
        "rows": len(rows),
        "phase_count": len(PHASE_NAMES) if scope == "phases" else 0,
    }


def _cli_json(value: str, label: str) -> Any:
    try:
        candidate = Path(value)
        if candidate.is_file():
            return _read_json(candidate, label)
    except OSError:
        # Long inline JSON values are not valid filesystem paths; parse them
        # as JSON below rather than treating Path.is_file as validation.
        pass
    try:
        return json.loads(value)
    except (json.JSONDecodeError, TypeError) as error:
        _fail(label, f"must be a JSON file or JSON value ({error})")


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("paths", nargs="*", help="REPORT [LANE_JSON_OR_FILE PROTOCOL_JSON_OR_FILE]")
    parser.add_argument("--report", dest="report_option", help="report JSON path")
    parser.add_argument("--lane", dest="lane_option", help="lane JSON path or inline JSON")
    parser.add_argument("--protocol", dest="protocol_option", help="protocol JSON path or inline JSON")
    args = parser.parse_args(argv)
    values = [args.report_option, args.lane_option, args.protocol_option]
    if any(value is not None for value in values):
        if any(value is None for value in values):
            parser.error("--report, --lane, and --protocol must be supplied together")
        report_path = Path(args.report_option)
        lane_value = _cli_json(args.lane_option, "lane")
        protocol_value = _cli_json(args.protocol_option, "protocol")
    else:
        if len(args.paths) != 3:
            parser.error("provide REPORT, LANE_JSON_OR_FILE, and PROTOCOL_JSON_OR_FILE")
        report_path = Path(args.paths[0])
        lane_value = _cli_json(args.paths[1], "lane")
        protocol_value = _cli_json(args.paths[2], "protocol")
    try:
        result = validate(report_path, lane_value, protocol_value)
    except (OracleError, OSError, TypeError, ValueError) as error:
        parser.exit(1, f"oracle: {error}\n")
    print(json.dumps(result, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
