"""Independent readers for the 0840 fresh CFB emission experiment.

The probe owns generation and timing.  This module replays the retained JSON
and parses emitted CFB bytes without importing the probe or the production
writer.  In particular, hashes in a report are checked against independently
reconstructed input bytes and, when an artifact is retained, against an
independent OLE2 sector walk.
"""

from __future__ import annotations

import functools
import hashlib
import json
import math
import random
import statistics
import struct
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence


REPORT_SCHEMA = "litchi.execution-cfb-emission.v1"
ALLOCATION_SCHEMA = "litchi.execution-allocation-observer.v1"
GENERATOR = "tools/perf-baseline/src/lib.rs::write_fresh_doc+cfb-emission-v1"
BOOTSTRAP_SEED = 840084
BOOTSTRAP_RESAMPLES = 10_000
BOOTSTRAP_INDEXES = (249, 9749)
BOOTSTRAP_CONFIDENCE = 0.95
MAX_SAMPLES = 10_000
MAX_WARMUP = 10_000
CFB_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
DISECT = 0xFFFFFFFC
NOSTREAM = FREESECT
HEX = frozenset("0123456789abcdefABCDEF")


class ReaderError(ValueError):
    """Raised when retained evidence is missing or contradictory."""


def fail(message: str) -> None:
    raise ReaderError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def _json(path: Path) -> Any:
    path = Path(path)
    require(path.is_file() and not path.is_symlink(), f"missing report: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, ValueError) as error:
        fail(f"invalid JSON report {path}: {error}")


def _sha256(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def _sha256_file(path: Path) -> str:
    try:
        with Path(path).open("rb") as stream:
            digest = hashlib.sha256()
            for chunk in iter(lambda: stream.read(1 << 20), b""):
                digest.update(chunk)
            return digest.hexdigest()
    except OSError as error:
        fail(f"cannot hash {path}: {error}")


def _sha(value: Any, label: str) -> str:
    require(isinstance(value, str) and len(value) == 64 and all(c in HEX for c in value),
            f"{label} is not a SHA-256 hexadecimal string")
    return value


def _uint(value: Any, label: str, *, positive: bool = False) -> int:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a non-negative integer")
    if positive:
        require(value > 0, f"{label} is not positive")
    return value


def _finite(value: Any, label: str) -> float:
    require(isinstance(value, (int, float)) and not isinstance(value, bool),
            f"{label} is not numeric")
    result = float(value)
    require(math.isfinite(result), f"{label} is not finite")
    return result


def _sequence_sha256(items: Iterable[tuple[str, bytes]]) -> str:
    digest = hashlib.sha256()
    for name, value in items:
        digest.update(name.encode("utf-8"))
        digest.update(b"\0")
        digest.update(value)
        digest.update(b"\xff")
    return digest.hexdigest()


def _writer_text(kind: str, first: int, second: int, third: int) -> str:
    return (f"litchi-perf-baseline-{kind}-v1-{first:03}-{second:05}-{third:03} "
            "deterministic payload")


def _writer_payload_text(kind: str, first: int, second: int, third: int,
                         length: int) -> str:
    repeated = "litchi-perf-baseline-payload-heavy-v1 "
    text = _writer_text(kind, first, second, third)
    while len(text) < length:
        text += repeated
    return text[:length]


@functools.cache
def _payload_bytes(length: int, seed: int) -> bytes:
    # This is the probe's u64 xorshift generator, evaluated with explicit
    # masking so Python cannot accidentally widen the state.
    state = ((seed * 0x9E3779B97F4A7C15) + 0xD1B54A32D192ED03) & ((1 << 64) - 1)
    output = bytearray()
    for _ in range(length):
        state ^= (state << 13) & ((1 << 64) - 1)
        state &= (1 << 64) - 1
        state ^= state >> 7
        state &= (1 << 64) - 1
        state ^= (state << 17) & ((1 << 64) - 1)
        state &= (1 << 64) - 1
        output.append((state >> 24) & 0xFF)
    return bytes(output)


@functools.cache
def _small_payload(length: int, seed: int) -> bytes:
    block = b"litchi-perf-cfb-emission-mini-payload-v1\n"
    return bytes(block[(offset + seed) % len(block)] for offset in range(length))


@functools.cache
def _prepared_members(case: str) -> tuple[dict[str, Any], list[tuple[str, bytes]]]:
    """Reconstruct the probe input identity independently."""
    if case in {"doc-tiny", "doc-large", "doc-payload"}:
        count, payload_length = {
            "doc-tiny": (3, None),
            "doc-large": (512, None),
            "doc-payload": (128, 20_000),
        }[case]
        members: list[tuple[str, bytes]] = []
        for index in range(count):
            text = (_writer_text("doc", 0, index, 0) if payload_length is None
                    else _writer_payload_text("doc", 0, index, 0, payload_length))
            members.append((f"paragraph:{index:05}", text.encode("utf-8")))
        identity = {
            "generator": GENERATOR,
            "case": case,
            "format": "DOC/CFB",
            "sector_size": None,
            "paragraph_count": count,
            "stream_count": None,
            "logical_input_bytes": sum(len(value) for _, value in members),
            "input_sha256": _sequence_sha256(members),
            "members": [{"name": name, "bytes": len(value), "sha256": _sha256(value)}
                        for name, value in members],
            "preparation_boundary":
                "paragraph text generation and hashing occur before each measured operation",
        }
        return identity, members

    specs: dict[str, tuple[int, list[tuple[str, bytes]]]] = {
        "cfb-tiny": (512, [("MiniPayload", _small_payload(512, 0)),
                            ("LargePayload", _payload_bytes(8 * 1024, 1))]),
        "cfb-large": (512, [("MiniPayload", _small_payload(2 * 1024, 0)),
                             ("LargePayload", _payload_bytes(4 * 1024 * 1024, 1))]),
        "cfb-large-only": (512, [("LargePayload", _payload_bytes(4 * 1024 * 1024, 1))]),
        "cfb-mini-only": (512, [("MiniPayloadA", _small_payload(256, 0)),
                                 ("MiniPayloadB", _small_payload(1536, 1)),
                                 ("MiniPayloadC", _small_payload(3072, 2))]),
        "cfb-v4": (4096, [("MiniPayload", _small_payload(2 * 1024, 0)),
                           ("LargePayload", _payload_bytes(4 * 1024 * 1024, 1))]),
        "cfb-difat": (512, [("DifatPayload", _payload_bytes(8 * 1024 * 1024, 7))]),
    }
    require(case in specs, f"unknown probe case {case!r}")
    sector_size, members = specs[case]
    identity = {
        "generator": GENERATOR,
        "case": case,
        "format": "CFB/OLE2",
        "sector_size": sector_size,
        "paragraph_count": None,
        "stream_count": len(members),
        "logical_input_bytes": sum(len(value) for _, value in members),
        "input_sha256": _sequence_sha256(members),
        "members": [{"name": name, "bytes": len(value), "sha256": _sha256(value)}
                    for name, value in members],
        "preparation_boundary":
            "stream payload generation and hashing occur before each measured operation",
    }
    return identity, members


def _check_input(value: Any, case: str) -> dict[str, Any]:
    require(isinstance(value, dict), "report input is not an object")
    expected, members = _prepared_members(case)
    for key, expected_value in expected.items():
        require(value.get(key) == expected_value, f"input.{key} differs from deterministic corpus")
    actual_members = value.get("members")
    require(isinstance(actual_members, list) and actual_members == expected["members"],
            "input member manifest differs from deterministic corpus")
    identity = dict(expected)
    identity["member_hashes"] = [entry["sha256"] for entry in expected["members"]]
    identity["member_bytes"] = [entry["bytes"] for entry in expected["members"]]
    identity["member_names"] = [entry["name"] for entry in expected["members"]]
    identity["fingerprint"] = _sha256(json.dumps(identity, sort_keys=True,
                                                  separators=(",", ":")).encode())
    # Keep the independent bytes private to the parser while returning a
    # useful normalized identity to the custody/analysis callers.
    identity["_members"] = members
    return identity


def _check_verification(value: Any, identity: Mapping[str, Any], label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} verification is missing")
    require(value.get("valid_output") is True and value.get("exact_input_match") is True,
            f"{label} output verification failed")
    output_bytes = _uint(value.get("output_bytes"), f"{label}.output_bytes", positive=True)
    output_sha256 = _sha(value.get("output_sha256"), f"{label}.output_sha256")
    semantic_sha256 = _sha(value.get("semantic_sha256"), f"{label}.semantic_sha256")
    logical = _uint(value.get("logical_output_bytes"), f"{label}.logical_output_bytes")
    count = _uint(value.get("member_or_paragraph_count"),
                  f"{label}.member_or_paragraph_count")
    require(logical == identity["logical_input_bytes"], f"{label} logical output differs")
    require(count == len(identity["members"]), f"{label} member/paragraph count differs")
    if identity["format"] == "CFB/OLE2":
        require(semantic_sha256 == identity["input_sha256"],
                f"{label} CFB semantic hash differs from input streams")
    else:
        text = b"".join(value for _, value in identity["_members"])
        text = b"\r".join(value for _, value in identity["_members"]) + b"\r"
        require(semantic_sha256 == _sha256(text), f"{label} DOC semantic hash differs")
    return {
        "valid_output": True,
        "output_bytes": output_bytes,
        "output_sha256": output_sha256,
        "semantic_sha256": semantic_sha256,
        "logical_output_bytes": logical,
        "member_or_paragraph_count": count,
        "exact_input_match": True,
    }


def _config(value: Any, case: str, samples: int, warmup: int, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} config is missing")
    expected = {
        "case": case,
        "samples": samples,
        "warmup": warmup,
        "measured_operation":
            "fresh public writer construction, prepared-input registration, and write_to",
        "output_verification":
            "outside wall timer: reopen, semantic projection, hashes, and inventory",
    }
    for key, expected_value in expected.items():
        require(value.get(key) == expected_value, f"{label} config.{key} changed")
    return expected


def _timed_samples(value: Any, identity: Mapping[str, Any], samples: int,
                   label: str) -> tuple[list[int], list[dict[str, Any]]]:
    require(isinstance(value, list) and len(value) == samples,
            f"{label} timed sample count changed")
    walls: list[int] = []
    normalized: list[dict[str, Any]] = []
    for index, sample in enumerate(value):
        require(isinstance(sample, dict), f"{label} sample {index} is malformed")
        require(sample.get("sample") == index, f"{label} sample numbering changed")
        wall = _uint(sample.get("wall_ns"), f"{label} sample {index}.wall_ns", positive=True)
        output_bytes = _uint(sample.get("output_bytes"),
                             f"{label} sample {index}.output_bytes", positive=True)
        output_sha = _sha(sample.get("output_sha256"),
                          f"{label} sample {index}.output_sha256")
        verification = _check_verification(sample.get("verification"), identity,
                                            f"{label} sample {index}")
        require(output_bytes == verification["output_bytes"]
                and output_sha == verification["output_sha256"],
                f"{label} sample {index} output identity disagrees with verification")
        walls.append(wall)
        normalized.append({"sample": index, "wall_ns": wall,
                            "output_bytes": output_bytes, "output_sha256": output_sha,
                            "verification": verification})
    require(len({row["output_sha256"] for row in normalized}) == 1,
            f"{label} output bytes changed between timed samples")
    return walls, normalized


def _observer(value: Any, identity: Mapping[str, Any], label: str) -> tuple[dict[str, Any], dict[str, int]]:
    require(isinstance(value, dict), f"{label} observer summary is missing")
    names = ("write_calls", "write_bytes", "seek_calls", "backward_seek_calls",
             "backward_seek_bytes", "gap_events", "zero_filled_gap_bytes",
             "largest_gap_bytes", "output_bytes", "events_recorded")
    numbers = {name: _uint(value.get(name), f"{label}.{name}") for name in names}
    output_sha = _sha(value.get("output_sha256"), f"{label}.output_sha256")
    require(value.get("events_truncated") is False, f"{label} observer event list is truncated")
    events = value.get("events")
    require(isinstance(events, list) and len(events) == numbers["events_recorded"],
            f"{label} observer event count changed")
    write_events: list[dict[str, int]] = []
    seek_events: list[dict[str, int]] = []
    for index, event in enumerate(events):
        require(isinstance(event, dict), f"{label} event {index} is malformed")
        kind = event.get("kind")
        if kind == "write":
            row = {name: _uint(event.get(name), f"{label} event {index}.{name}")
                   for name in ("at", "bytes", "gap_before")}
            require(row["bytes"] > 0, f"{label} event {index} is an empty write")
            write_events.append(row)
        elif kind == "seek":
            row = {name: _uint(event.get(name), f"{label} event {index}.{name}")
                   for name in ("from", "to")}
            seek_events.append(row)
        else:
            fail(f"{label} event {index} has unknown kind {kind!r}")
    require(numbers["write_calls"] == len(write_events)
            and numbers["write_bytes"] == sum(row["bytes"] for row in write_events),
            f"{label} observer write counters disagree with event trace")
    gaps = [row["gap_before"] for row in write_events]
    nonzero_gaps = [gap for gap in gaps if gap]
    require(numbers["gap_events"] == len(nonzero_gaps)
            and numbers["zero_filled_gap_bytes"] == sum(nonzero_gaps)
            and numbers["largest_gap_bytes"] == (max(gaps) if gaps else 0),
            f"{label} observer gap counters disagree with event trace")
    backwards = [row["from"] - row["to"] for row in seek_events if row["to"] < row["from"]]
    require(numbers["seek_calls"] == len(seek_events)
            and numbers["backward_seek_calls"] == len(backwards)
            and numbers["backward_seek_bytes"] == sum(backwards),
            f"{label} observer seek counters disagree with event trace")
    require(numbers["output_bytes"] > 0 and output_sha and numbers["output_bytes"]
            == _uint(value.get("output_bytes"), f"{label}.output_bytes"),
            f"{label} observer output identity is invalid")
    # The summary is cross-checked against the independent verification by
    # validate_report; this function intentionally does not accept a report
    # hash as the parser's evidence.
    return ({**numbers, "output_sha256": output_sha, "events_recorded": len(events),
             "events_truncated": False, "events": events},
            {"write_calls": numbers["write_calls"], "write_bytes": numbers["write_bytes"],
             "seek_calls": numbers["seek_calls"],
             "backward_seek_calls": numbers["backward_seek_calls"],
             "backward_seek_bytes": numbers["backward_seek_bytes"],
             "gap_events": numbers["gap_events"],
             "zero_filled_gap_bytes": numbers["zero_filled_gap_bytes"],
             "largest_gap_bytes": numbers["largest_gap_bytes"]})


def _allocation_sample(value: Any, label: str) -> dict[str, int]:
    require(isinstance(value, dict), f"{label} allocation sample is missing")
    # The allocation binary uses the same verification shape, but reports no
    # timing fields.  Its counter names are deliberately explicit so a timing
    # report cannot be mistaken for allocator evidence.
    for forbidden in ("wall_ns", "cpu_ns"):
        require(forbidden not in value, f"{label} allocation sample carries {forbidden}")
    allocation = value.get("allocation")
    require(isinstance(allocation, dict), f"{label}.allocation is missing")
    require(allocation.get("status") == "measured"
            and allocation.get("scope") == "operation_global_system_allocator",
            f"{label} allocation observation is not measured")
    names = ("allocation_calls", "allocated_bytes", "live_bytes_before", "live_bytes_after",
             "peak_live_bytes_before", "peak_live_bytes_after", "region_peak_live_bytes",
             "deallocation_calls", "deallocated_bytes", "failed_allocation_calls")
    numbers = {name: _uint(allocation.get(name), f"{label}.allocation.{name}")
               for name in names}
    require(numbers["failed_allocation_calls"] == 0, f"{label} has failed allocations")
    require(numbers["live_bytes_after"] == numbers["live_bytes_before"]
            + numbers["allocated_bytes"] - numbers["deallocated_bytes"],
            f"{label} allocation live-byte conservation failed")
    require(numbers["peak_live_bytes_after"] >= numbers["peak_live_bytes_before"]
            and numbers["region_peak_live_bytes"] >= numbers["live_bytes_before"]
            and numbers["region_peak_live_bytes"] >= numbers["live_bytes_after"]
            and numbers["region_peak_live_bytes"] <= numbers["peak_live_bytes_after"],
            f"{label} allocation peak bounds failed")
    output_bytes = _uint(value.get("output_bytes"), f"{label}.output_bytes", positive=True)
    output_sha = _sha(value.get("output_sha256"), f"{label}.output_sha256")
    return {
        "allocation_calls": numbers["allocation_calls"],
        "allocated_bytes": numbers["allocated_bytes"],
        "region_peak_live_bytes_minus_entry":
            numbers["region_peak_live_bytes"] - numbers["live_bytes_before"],
        "retained_live_bytes_delta":
            numbers["live_bytes_after"] - numbers["live_bytes_before"],
        "_output_bytes": output_bytes,
        "_output_sha256": output_sha,
    }


def validate_report(path: Path, case: str, samples: int, warmup: int, *,
                    observer: bool = False, allocation: bool = False) -> dict[str, Any]:
    """Validate one strict 0840 report and return analysis-ready vectors."""
    report = _json(Path(path))
    require(isinstance(report, dict), f"{path} report is not an object")
    expected_schema = ALLOCATION_SCHEMA if allocation else REPORT_SCHEMA
    require(report.get("schema") == expected_schema, f"{path} schema changed")
    if allocation:
        require("mode" not in report, f"{path} allocation report carries a timing mode")
    else:
        expected_mode = "observe" if observer else "timed"
        require(report.get("mode") == expected_mode, f"{path} mode changed")
    label = str(path)
    if allocation:
        config = {
            "case": case,
            "samples": samples,
            "warmup": warmup,
            "measured_operation":
                "fresh public writer construction, prepared-input registration, and write_to",
            "verification":
                "outside allocation region: reopen, semantic projection, hashes, and inventory",
        }
        identity = _check_input(report.get("corpus"), case)
    else:
        config = _config(report.get("config"), case, samples, warmup, label)
        identity = _check_input(report.get("input"), case)
    observer_metrics: dict[str, int] | None = None
    allocation_metrics: dict[str, list[int]] | None = None
    if observer:
        require(samples == 1 and warmup == 0, f"{label} observer schedule changed")
        verification = _check_verification(report.get("verification"), identity,
                                            f"{label}.verification")
        summary, observer_metrics = _observer(report.get("observer"), identity, label)
        require(summary["output_bytes"] == verification["output_bytes"]
                and summary["output_sha256"] == verification["output_sha256"],
                f"{label} observer output differs from verification")
        walls: list[int] = []
        normalized_samples: list[dict[str, Any]] = []
        final = verification
    elif allocation:
        require(report.get("scope")
                == "global system allocator callbacks around the same fresh public writer operation; no wall or CPU timing",
                f"{label} allocation scope changed")
        require("input" not in report and isinstance(report.get("corpus"), dict),
                f"{label} allocation corpus envelope changed")
        identity = _check_input(report.get("corpus"), case)
        require(report.get("metrics") == {
            "allocator_identity": "CountingSystemAllocator(std::alloc::System)",
            "counter_revision": "serialized_region_peak_v3",
            "instrumentation_identity": "system_allocator_operation_scoped",
            "timing": "not measured by this binary",
        }, f"{label} allocation instrumentation identity changed")
        allocation_config = report.get("config")
        require(isinstance(allocation_config, dict), f"{label} allocation config is missing")
        require(allocation_config == {
            "case": case,
            "samples": samples,
            "warmup": warmup,
            "measured_operation":
                "fresh public writer construction, prepared-input registration, and write_to",
            "verification":
                "outside allocation region: reopen, semantic projection, hashes, and inventory",
        }, f"{label} allocation config changed")
        rows = report.get("samples")
        require(isinstance(rows, list) and len(rows) == samples,
                f"{label} allocation sample count changed")
        allocation_rows: list[dict[str, int]] = []
        normalized_samples = []
        for index, row in enumerate(rows):
            require(isinstance(row, dict) and row.get("sample") == index,
                    f"{label} allocation sample {index} is malformed")
            require(row.get("sample") == index, f"{label} allocation sample numbering changed")
            metrics = _allocation_sample(row, f"{label} sample {index}")
            allocation_rows.append(metrics)
            normalized_samples.append({"sample": index, "allocation": metrics})
        final_bytes = _uint(report.get("final_output_bytes"),
                            f"{label}.final_output_bytes", positive=True)
        final_sha = _sha(report.get("final_output_sha256"),
                         f"{label}.final_output_sha256")
        require(all(row["_output_sha256"] == final_sha and row["_output_bytes"] == final_bytes
                    for row in allocation_rows),
                f"{label} allocation output identity changed")
        allocation_metrics = {key: [row[key] for row in allocation_rows]
                              for key in ("allocation_calls", "allocated_bytes",
                                           "region_peak_live_bytes_minus_entry",
                                           "retained_live_bytes_delta")}
        walls = []
        verification = {
            "valid_output": True,
            "output_bytes": final_bytes,
            "output_sha256": final_sha,
            "semantic_sha256": None,
            "logical_output_bytes": None,
            "member_or_paragraph_count": None,
            "exact_input_match": True,
        }
    else:
        walls, normalized_samples = _timed_samples(report.get("samples"), identity,
                                                    samples, label)
        final = _check_verification(report.get("final_verification"), identity,
                                    f"{label}.final_verification")
        require(final["output_sha256"] == normalized_samples[-1]["output_sha256"]
                and final["output_bytes"] == normalized_samples[-1]["output_bytes"],
                f"{label} final output differs from final timed sample")
        verification = final
    artifact_identity = {
        "output_bytes": verification["output_bytes"],
        "output_sha256": verification["output_sha256"],
        "semantic_sha256": verification["semantic_sha256"],
        "logical_output_bytes": verification["logical_output_bytes"],
        "member_or_paragraph_count": verification["member_or_paragraph_count"],
    }
    result: dict[str, Any] = {
        "report": {"path": str(path), "bytes": Path(path).stat().st_size,
                   "sha256": _sha256_file(Path(path))},
        "case": case,
        "config": config,
        "corpus_identity": {key: value for key, value in identity.items() if key != "_members"},
        "corpus_fingerprint": identity["fingerprint"],
        "artifact_identity": artifact_identity,
        "verification": verification,
        "verification_ok": True,
        "wall_ns": None if observer or allocation else walls,
        "cpu_ns": None,
        "p50_ns": None if observer or allocation else nearest_rank(walls, 0.50),
        "p95_ns": None if observer or allocation else nearest_rank(walls, 0.95),
        "p99_ns": None if observer or allocation else nearest_rank(walls, 0.99),
        "mean_ns": None if observer or allocation else statistics.fmean(walls),
        "samples": normalized_samples,
        "observer_metrics": observer_metrics,
        "allocation_metrics": allocation_metrics,
    }
    return result


def parse_report(path: Path, case: str, samples: int, warmup: int, *,
                 observer: bool = False, allocation: bool = False) -> dict[str, Any]:
    """Compatibility wrapper for callers using the older reader spelling."""
    return validate_report(path, case, samples, warmup,
                           observer=observer, allocation=allocation)


def nearest_rank(values: Iterable[float], quantile: float) -> float:
    require(isinstance(quantile, (int, float)) and not isinstance(quantile, bool)
            and 0.0 <= float(quantile) <= 1.0, "quantile is outside [0, 1]")
    ordered = sorted(_finite(value, "quantile input") for value in values)
    require(ordered, "nearest-rank received no values")
    rank = max(1, math.ceil(float(quantile) * len(ordered)))
    return ordered[min(rank, len(ordered)) - 1]


def bootstrap(values: Sequence[float], *, seed: int = BOOTSTRAP_SEED,
              resamples: int = BOOTSTRAP_RESAMPLES,
              endpoint_indexes: tuple[int, int] = BOOTSTRAP_INDEXES) -> dict[str, Any]:
    require(values, "bootstrap received no values")
    require(resamples > max(endpoint_indexes) >= 0, "bootstrap endpoint indexes are invalid")
    numbers = [_finite(value, "bootstrap input") for value in values]
    rng = random.Random(seed)
    estimates = sorted(statistics.median(rng.choice(numbers) for _ in numbers)
                       for _ in range(resamples))
    return {
        "estimate": statistics.median(numbers),
        "lower": estimates[endpoint_indexes[0]],
        "upper": estimates[endpoint_indexes[1]],
        "seed": seed,
        "resamples": resamples,
        "confidence": BOOTSTRAP_CONFIDENCE,
        "statistic": "median",
        "endpoint_indexes": list(endpoint_indexes),
    }


def bootstrap_ratio(before: Sequence[float], after: Sequence[float],
                    name: str = "ratio") -> dict[str, Any]:
    require(len(before) == len(after) and before,
            f"{name} pair cardinality is empty or differs")
    old = [_finite(value, f"{name}.before") for value in before]
    new = [_finite(value, f"{name}.after") for value in after]
    ratios: list[float | None] = []
    for index, (left, right) in enumerate(zip(old, new)):
        require(left >= 0 and right >= 0, f"{name} pair {index} is negative")
        ratios.append(right / left if left > 0 else None)
    usable = [value for value in ratios if value is not None]
    interval = bootstrap(usable) if usable else None
    result: dict[str, Any] = {
        "name": name,
        "before": old,
        "after": new,
        "ratios": ratios,
        "block_ratios": ratios,
        "raw_distributions_retained": True,
    }
    if interval is None:
        result.update({"estimate": None, "ci95_low": None, "ci95_high": None,
                       "bootstrap": None})
    else:
        result.update({"estimate": interval["estimate"], "ci95_low": interval["lower"],
                       "ci95_high": interval["upper"], "bootstrap": interval})
    return result


def _u16(data: bytes, offset: int) -> int:
    require(offset >= 0 and offset + 2 <= len(data), "CFB u16 outside artifact")
    return struct.unpack_from("<H", data, offset)[0]


def _u32(data: bytes, offset: int) -> int:
    require(offset >= 0 and offset + 4 <= len(data), "CFB u32 outside artifact")
    return struct.unpack_from("<I", data, offset)[0]


def _sector(data: bytes, sector_size: int, sid: int) -> bytes:
    require(sid != FREESECT and sid != ENDOFCHAIN and sid != FATSECT and sid != DISECT,
            f"invalid CFB sector id {sid:#x}")
    start = (sid + 1) * sector_size
    end = start + sector_size
    require(0 <= sid and end <= len(data), f"CFB sector {sid} outside artifact")
    return data[start:end]


def _chain(fat: Sequence[int], start: int, limit: int, label: str) -> list[int]:
    if start == ENDOFCHAIN:
        return []
    require(start not in (FREESECT, FATSECT, DISECT), f"{label} starts with reserved sector")
    result: list[int] = []
    seen: set[int] = set()
    current = start
    while current != ENDOFCHAIN:
        require(current < len(fat), f"{label} points outside FAT")
        require(current not in seen, f"{label} contains a cycle")
        require(len(result) < limit, f"{label} exceeds bounded chain length")
        seen.add(current)
        result.append(current)
        current = fat[current]
    return result


def _directory_name(entry: bytes, index: int) -> str:
    length = _u16(entry, 0x40)
    require(length % 2 == 0 and 2 <= length <= 64, f"directory entry {index} name length invalid")
    name_bytes = entry[: length - 2]
    try:
        return name_bytes.decode("utf-16le")
    except UnicodeDecodeError as error:
        fail(f"directory entry {index} name is invalid UTF-16: {error}")


def parse_cfb_artifact(data: bytes | bytearray | memoryview) -> dict[str, Any]:
    """Parse a complete CFB artifact and return normalized physical metadata."""
    data = bytes(data)
    require(len(data) >= 512 and data[:8] == CFB_MAGIC, "CFB header signature is invalid")
    require(len(data) % 512 == 0, "CFB artifact is not sector-aligned")
    major = _u16(data, 0x1A)
    byte_order = _u16(data, 0x1C)
    sector_shift = _u16(data, 0x1E)
    mini_shift = _u16(data, 0x20)
    require(major in (3, 4) and byte_order == 0xFFFE,
            "CFB version or byte order is invalid")
    require(sector_shift in (9, 12) and mini_shift == 6,
            "CFB sector shifts are invalid")
    sector_size = 1 << sector_shift
    require(len(data) >= sector_size, "CFB artifact is shorter than its sector size")
    sector_count = len(data) // sector_size - 1
    num_directory = _u32(data, 0x28)
    num_fat = _u32(data, 0x2C)
    first_directory = _u32(data, 0x30)
    mini_cutoff = _u32(data, 0x38)
    first_minifat = _u32(data, 0x3C)
    num_minifat = _u32(data, 0x40)
    first_difat = _u32(data, 0x44)
    num_difat = _u32(data, 0x48)
    require(mini_cutoff == 4096, "CFB mini-stream cutoff changed")
    require(num_fat > 0 and num_fat <= sector_count, "CFB FAT count is invalid")
    require(num_difat <= sector_count, "CFB DIFAT count is invalid")
    difat: list[int] = []
    for index in range(109):
        sid = _u32(data, 0x4C + index * 4)
        if sid != FREESECT:
            difat.append(sid)
    difat_chain: list[int] = []
    current = first_difat
    for _ in range(num_difat):
        require(current not in (FREESECT, ENDOFCHAIN), "CFB DIFAT chain ended too early")
        require(current < sector_count, "CFB DIFAT sector is outside artifact")
        require(current not in difat_chain, "CFB DIFAT chain contains a cycle")
        difat_chain.append(current)
        block = _sector(data, sector_size, current)
        for index in range(sector_size // 4 - 1):
            sid = _u32(block, index * 4)
            if sid != FREESECT:
                difat.append(sid)
        current = _u32(block, sector_size - 4)
    require(current == ENDOFCHAIN if num_difat else current in (ENDOFCHAIN, FREESECT),
            "CFB DIFAT chain has an unexpected tail")
    require(len(difat) >= num_fat, "CFB DIFAT does not list every FAT sector")
    fat_sector_ids = difat[:num_fat]
    require(len(set(fat_sector_ids)) == len(fat_sector_ids), "CFB FAT sector list repeats")
    fat_values: list[int] = []
    for sid in fat_sector_ids:
        fat_values.extend(_u32(_sector(data, sector_size, sid), offset)
                          for offset in range(0, sector_size, 4))
    fat_values = fat_values[:sector_count]
    require(len(fat_values) == sector_count, "CFB FAT does not cover artifact sectors")
    for sid in fat_sector_ids:
        require(fat_values[sid] == FATSECT, f"CFB FAT sector {sid} lacks FATSECT marker")
    for sid in difat_chain:
        require(fat_values[sid] == DISECT, f"CFB DIFAT sector {sid} lacks DISECT marker")
    directory_chain = _chain(fat_values, first_directory, sector_count, "directory")
    require(num_directory == 0 or len(directory_chain) == num_directory,
            "CFB directory sector count differs from header")
    directory_bytes = b"".join(_sector(data, sector_size, sid) for sid in directory_chain)
    require(len(directory_bytes) % 128 == 0, "CFB directory is not entry-aligned")
    entries: list[dict[str, Any]] = []
    raw_entries: list[bytes] = []
    for index in range(len(directory_bytes) // 128):
        entry = directory_bytes[index * 128:(index + 1) * 128]
        object_type = entry[0x42]
        if object_type == 0:
            continue
        require(object_type in (1, 2, 5), f"directory entry {index} type is invalid")
        name = _directory_name(entry, index)
        start = _u32(entry, 0x74)
        low = _u32(entry, 0x78)
        high = _u32(entry, 0x7C) if major == 4 else 0
        size = low | (high << 32)
        entries.append({"index": index, "name": name, "object_type": object_type,
                        "left": _u32(entry, 0x44), "right": _u32(entry, 0x48),
                        "child": _u32(entry, 0x4C), "start_sector": start,
                        "stream_size": size})
        raw_entries.append(entry)
    roots = [entry for entry in entries if entry["object_type"] == 5]
    require(len(roots) == 1 and roots[0]["name"] == "Root Entry",
            "CFB root directory entry is missing or duplicated")
    root = roots[0]
    minifat_values: list[int] = []
    minifat_chain: list[int] = []
    if num_minifat:
        minifat_chain = _chain(fat_values, first_minifat, sector_count, "MiniFAT")
        require(len(minifat_chain) == num_minifat, "CFB MiniFAT sector count differs from header")
        minifat_bytes = b"".join(_sector(data, sector_size, sid) for sid in minifat_chain)
        minifat_values = [_u32(minifat_bytes, offset)
                          for offset in range(0, len(minifat_bytes), 4)]
    else:
        require(first_minifat in (FREESECT, ENDOFCHAIN), "CFB empty MiniFAT has a start sector")
    root_size = root["stream_size"]
    root_chain = _chain(fat_values, root["start_sector"], sector_count, "root mini stream")
    root_bytes = b"".join(_sector(data, sector_size, sid) for sid in root_chain)[:root_size]
    require(len(root_bytes) == root_size, "CFB root mini stream is shorter than declared")
    streams: list[dict[str, Any]] = []
    for entry in entries:
        if entry["object_type"] != 2:
            continue
        size = entry["stream_size"]
        if size == 0:
            require(entry["start_sector"] in (FREESECT, ENDOFCHAIN),
                    f"empty stream {entry['name']} has a sector")
            stream = b""
            storage = "empty"
            chain = []
        elif size < mini_cutoff:
            require(minifat_values, f"mini stream {entry['name']} lacks MiniFAT")
            needed = (size + (1 << mini_shift) - 1) // (1 << mini_shift)
            mini_chain = _chain(minifat_values, entry["start_sector"], len(minifat_values),
                                f"MiniFAT stream {entry['name']}")
            require(len(mini_chain) == needed, f"mini stream {entry['name']} chain length differs")
            parts = []
            for mini_sid in mini_chain:
                start = mini_sid * (1 << mini_shift)
                require(start + (1 << mini_shift) <= len(root_bytes),
                        f"mini stream {entry['name']} points outside root mini stream")
                parts.append(root_bytes[start:start + (1 << mini_shift)])
            stream = b"".join(parts)[:size]
            storage = "mini"
            chain = mini_chain
        else:
            needed = (size + sector_size - 1) // sector_size
            chain = _chain(fat_values, entry["start_sector"], sector_count,
                           f"regular stream {entry['name']}")
            require(len(chain) == needed, f"regular stream {entry['name']} chain length differs")
            stream = b"".join(_sector(data, sector_size, sid) for sid in chain)[:size]
            storage = "regular"
        require(len(stream) == size, f"stream {entry['name']} is shorter than declared")
        streams.append({"name": entry["name"], "bytes": size, "sha256": _sha256(stream),
                        "storage": storage, "start_sector": entry["start_sector"],
                        "chain_length": len(chain), "data": stream})
    # Retain physical directory order separately from the sorted inventory.
    # Input registration order is reconstructed explicitly by the case validator.
    sequence_streams = list(streams)
    streams.sort(key=lambda row: row["name"])
    # Apply the independent deterministic-input gate opportunistically for
    # the synthetic CFB cases.  Their stream names are disjoint from DOC's
    # public writer streams, so this does not impose a DOC layout assumption.
    synthetic_case: str | None = None
    actual_names = {row["name"] for row in streams}
    signatures = {
        "cfb-tiny": {"MiniPayload": 512, "LargePayload": 8 * 1024},
        "cfb-large": {"MiniPayload": 2 * 1024, "LargePayload": 4 * 1024 * 1024},
        "cfb-large-only": {"LargePayload": 4 * 1024 * 1024},
        "cfb-mini-only": {"MiniPayloadA": 256, "MiniPayloadB": 1536,
                          "MiniPayloadC": 3072},
        "cfb-v4": {"MiniPayload": 2 * 1024, "LargePayload": 4 * 1024 * 1024},
        "cfb-difat": {"DifatPayload": 8 * 1024 * 1024},
    }
    # CFB large and CFB v4 have the same stream names; sector size separates
    # them without materializing an input corpus for every candidate.
    actual_lengths = {row["name"]: row["bytes"] for row in streams}
    candidates = [candidate for candidate, signature in signatures.items()
                  if actual_lengths == signature and (
                      candidate not in ("cfb-large", "cfb-v4") or
                      sector_size == (4096 if candidate == "cfb-v4" else 512))]
    for candidate in candidates:
        _, expected_members = _prepared_members(candidate)
        if actual_names != {name for name, _ in expected_members}:
            continue
        expected_hashes = {name: _sha256(value) for name, value in expected_members}
        require(all(row["sha256"] == expected_hashes[row["name"]]
                    for row in streams),
                f"{candidate} CFB stream content differs from deterministic input")
        require(synthetic_case is None, "CFB artifact matches multiple synthetic corpora")
        synthetic_case = candidate
    public_streams = [{key: value for key, value in row.items() if key != "data"}
                      for row in streams]
    return {
        "format": "CFB/OLE2",
        "bytes": len(data),
        "sha256": _sha256(data),
        "header": {
            "major_version": major, "sector_size": sector_size,
            "mini_sector_size": 1 << mini_shift,
            "sector_count": sector_count, "directory_sector_count": num_directory,
            "fat_sector_count": num_fat, "mini_fat_sector_count": num_minifat,
            "difat_sector_count": num_difat, "first_directory_sector": first_directory,
            "first_mini_fat_sector": first_minifat, "first_difat_sector": first_difat,
            "mini_stream_cutoff": mini_cutoff,
        },
        "fat": {"sector_ids": fat_sector_ids, "entry_count": len(fat_values),
                "sha256": _sha256(b"".join(struct.pack("<I", value) for value in fat_values))},
        "directory": [{key: value for key, value in entry.items()} for entry in entries],
        "minifat": {"sector_ids": minifat_chain, "entry_count": len(minifat_values),
                     "sha256": _sha256(b"".join(struct.pack("<I", value)
                                                for value in minifat_values))},
        "root_mini_stream": {"bytes": root_size, "sha256": _sha256(root_bytes)},
        "streams": public_streams,
        "synthetic_case": synthetic_case,
        "stream_sequence_sha256": _sequence_sha256(
            (row["name"], row["data"]) for row in sequence_streams),
        "_stream_bytes": {row["name"]: row["data"] for row in streams},
    }


def validate_cfb_artifact(case: str, data: bytes | bytearray | memoryview) -> dict[str, Any]:
    """Parse an artifact and apply the case-specific deterministic stream gate.

    CFB synthetic cases use streams whose bytes are prepared independently by
    :func:`_prepared_members`, so their full stream names, lengths, hashes, and
    sequence digest can be checked here.  DOC output is a public writer
    projection rather than a byte-for-byte copy of paragraph input; for it the
    parser still requires a complete CFB inventory and returns the physical
    identity to the custody layer, which compares all DOC artifacts by bytes.
    """
    parsed = parse_cfb_artifact(data)
    expected, members = _prepared_members(case)
    if case.startswith("cfb-"):
        actual = {row["name"]: row for row in parsed["streams"]}
        require(set(actual) == {name for name, _ in members},
                f"{case} CFB stream inventory differs from prepared input")
        for name, payload in members:
            row = actual[name]
            require(row["bytes"] == len(payload) and row["sha256"] == _sha256(payload),
                    f"{case} CFB stream {name} differs from prepared input")
        require(_sequence_sha256((name, parsed["_stream_bytes"][name]) for name, _ in members) == expected["input_sha256"],
                f"{case} CFB stream sequence differs from prepared input")
    else:
        require(parsed["streams"], f"{case} DOC artifact has no streams")
    return parsed


def paired_metric(before: Sequence[float], after: Sequence[float],
                  name: str = "metric") -> dict[str, Any]:
    require(len(before) == len(after) and before,
            f"{name} pair cardinality is empty or differs")
    old = [_finite(value, f"{name}.before") for value in before]
    new = [_finite(value, f"{name}.after") for value in after]
    difference = bootstrap([right - left for left, right in zip(old, new)])
    result = bootstrap_ratio(old, new, name)
    result.update({"before_block_values": old, "after_block_values": new,
                   "block_deltas": [right - left for left, right in zip(old, new)],
                   "difference_bootstrap": difference,
                   "difference_estimate": difference["estimate"],
                   "difference_ci95_low": difference["lower"],
                   "difference_ci95_high": difference["upper"]})
    return result
