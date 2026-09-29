"""Independent readers for the 0839 cached-Part scheduling trial.

The capture driver owns process creation and the paired comparison.  This
module only replays retained JSON and raw ``/proc`` text.  It reconstructs the
payload oracle without importing the benchmark, so a report cannot make its
own verification fields authoritative.
"""

from __future__ import annotations

import hashlib
import json
import math
import random
import re
import statistics
import struct
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence


REPORT_SCHEMA = "litchi.execution-baseline.v1"
BOOTSTRAP_SEED = 839083
BOOTSTRAP_RESAMPLES = 10_000
BOOTSTRAP_CONFIDENCE = 0.95
BOOTSTRAP_INDEXES = (249, 9749)
MEMBER_COUNT = 32
CPU_TASK_LIMIT = 1_000_000
HISTOGRAM_BINS = 8
HEX = frozenset("0123456789abcdefABCDEF")
_ORACLE_CACHE: dict[str, dict[str, Any]] | None = None


class ReaderError(ValueError):
    """Raised when retained evidence is missing, stale, or contradictory."""


def fail(message: str) -> None:
    raise ReaderError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def _read_json(path: Path) -> Any:
    path = Path(path)
    require(path.is_file() and not path.is_symlink(), f"missing report: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, ValueError) as error:
        fail(f"invalid JSON report {path}: {error}")


def _sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with Path(path).open("rb") as stream:
            for chunk in iter(lambda: stream.read(1 << 20), b""):
                digest.update(chunk)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def _is_sha(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 64 and all(char in HEX for char in value)


def _nonnegative_int(value: Any, label: str) -> int:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a non-negative integer")
    return value


def _positive_int(value: Any, label: str) -> int:
    result = _nonnegative_int(value, label)
    require(result > 0, f"{label} is not positive")
    return result


def _finite(value: Any, label: str) -> float:
    require(isinstance(value, (int, float)) and not isinstance(value, bool),
            f"{label} is not numeric")
    result = float(value)
    require(math.isfinite(result), f"{label} is not finite")
    return result


def _shape_size(shape: str, index: int) -> int:
    require(shape in ("small", "large", "mixed"), f"unknown corpus shape: {shape!r}")
    if shape == "small" or (shape == "mixed" and index == MEMBER_COUNT - 1):
        return 4 * 1024
    return 256 * 1024


def _payload(shape: str, index: int) -> bytes:
    """Rebuild one member exactly as ``probe/main.rs`` does."""
    size = _shape_size(shape, index)
    label = f"litchi-0786-member-{index:02}-".encode("ascii")
    output = bytearray(size)
    raw_offset = index * 11
    offset = raw_offset if raw_offset <= 0xFF else 0
    for position in range(size):
        lane = position % 97
        label_byte = label[position % len(label)]
        output[position] = (label_byte + lane * 3 + offset) & 0xFF
    return bytes(output)


def payload_oracle() -> dict[str, dict[str, Any]]:
    """Return independent member hashes, sizes, logical bytes, and sequence hashes."""
    global _ORACLE_CACHE
    if _ORACLE_CACHE is not None:
        return _ORACLE_CACHE
    result: dict[str, dict[str, Any]] = {}
    for shape in ("small", "large", "mixed"):
        member_hashes: list[str] = []
        sizes: list[int] = []
        sequence = hashlib.sha256()
        logical_bytes = 0
        for index in range(MEMBER_COUNT):
            payload = _payload(shape, index)
            member_hashes.append(_sha256_bytes(payload))
            sizes.append(len(payload))
            sequence.update(struct.pack("<Q", index))
            sequence.update(payload)
            logical_bytes += len(payload)
        digest = sequence.hexdigest()
        result[shape] = {
            # These compact names retain compatibility with the 0788 reader.
            "members": member_hashes,
            "member_hashes": member_hashes,
            "sizes": sizes,
            "sequence": digest,
            "sequence_sha256": digest,
            "logical_bytes": logical_bytes,
        }
    _ORACLE_CACHE = result
    return result


def _case_value(case: Mapping[str, Any], name: str) -> Any:
    value = case.get(name)
    require(value is not None, f"case lacks {name}")
    return value


def _case_copy(case: Mapping[str, Any]) -> dict[str, Any]:
    return {
        name: _case_value(case, name)
        for name in ("route", "shape", "state", "task_floor", "workers")
    }


def _corpus_fingerprint(identity: Mapping[str, Any]) -> str:
    encoded = json.dumps(identity, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return _sha256_bytes(encoded)


def check_corpus(report: Mapping[str, Any], expected_shape: str,
                 oracle: Mapping[str, Mapping[str, Any]],
                 label: str = "report") -> dict[str, Any]:
    """Check the exact 32-member corpus manifest and return its identity."""
    corpus = report.get("corpus")
    require(isinstance(corpus, dict), f"{label} corpus is missing")
    require(corpus.get("shape") == expected_shape, f"{label} corpus shape changed")
    require(corpus.get("metadata_members") == ["[Content_Types].xml", "_rels/.rels"],
            f"{label} metadata member manifest changed")
    members = corpus.get("members")
    require(isinstance(members, list) and len(members) == MEMBER_COUNT,
            f"{label} member cardinality changed")
    expected = oracle[expected_shape]
    hashes = expected["member_hashes"]
    sizes = expected["sizes"]
    for index, member in enumerate(members):
        require(isinstance(member, dict), f"{label} member {index} is malformed")
        require(member.get("index") == index
                and member.get("bytes") == sizes[index]
                and member.get("sha256") == hashes[index],
                f"{label} member {index} payload/order hash changed")
        name = f"custom/member{index:02}.bin"
        require(member.get("opc_name") == name
                and member.get("opc_uri") == f"/{name}"
                and member.get("cfb_name") == f"Member{index:02}",
                f"{label} member {index} naming changed")
    require(corpus.get("selected_payload_member_count") == MEMBER_COUNT,
            f"{label} selected member count changed")
    require(_is_sha(corpus.get("opc_sha256")) and _is_sha(corpus.get("cfb_sha256")),
            f"{label} container identity is invalid")
    identity = {
        "shape": expected_shape,
        "metadata_members": list(corpus["metadata_members"]),
        "member_count": MEMBER_COUNT,
        "member_sizes": list(sizes),
        "member_hashes": list(hashes),
        "selected_payload_member_count": MEMBER_COUNT,
        "opc_sha256": corpus["opc_sha256"],
        "cfb_sha256": corpus["cfb_sha256"],
    }
    identity["fingerprint"] = _corpus_fingerprint(identity)
    return identity


def _resource_snapshot(value: Any, label: str) -> dict[str, int]:
    require(isinstance(value, dict), f"{label} is missing")
    result: dict[str, int] = {}
    for name in ("workers", "io_concurrency", "cpu_tasks"):
        result[name] = _nonnegative_int(value.get(name), f"{label}.{name}")
    return result


def _resource_contract(sample: Mapping[str, Any], case: Mapping[str, Any],
                       label: str) -> dict[str, Any]:
    resources = sample.get("resources", sample.get("resource_snapshots"))
    require(isinstance(resources, dict), f"{label} resources are missing")
    limits = resources.get("limits")
    require(isinstance(limits, dict), f"{label} resource limits are missing")
    limit_values = {
        name: _nonnegative_int(limits.get(name), f"{label}.limits.{name}")
        for name in ("workers", "io_concurrency", "cpu_tasks")
    }
    workers = _positive_int(_case_value(case, "workers"), "case.workers")
    require(limit_values["workers"] == workers
            and limit_values["io_concurrency"] == workers,
            f"{label} worker or I/O limit changed")
    require(limit_values["cpu_tasks"] == CPU_TASK_LIMIT,
            f"{label} CPU task limit changed")
    snapshots = {
        name: _resource_snapshot(resources.get(name), f"{label}.{name}")
        for name in ("before_operation", "after_operation", "after_drop")
    }
    for marker, snapshot in snapshots.items():
        for name, value in snapshot.items():
            require(value <= limit_values[name], f"{label}.{marker}.{name} exceeds limit")
    # The OPC-from-bytes control has no member-read scheduler and therefore a
    # zero increment; CFB and source-backed Parts charge one task per member.
    route = _case_value(case, "route")
    increment = 0 if route == "opc" else MEMBER_COUNT
    state = _case_value(case, "state")
    initial = 0 if state == "fresh" else MEMBER_COUNT
    require(snapshots["before_operation"]["cpu_tasks"] == initial
            and snapshots["after_operation"]["cpu_tasks"] == initial + increment
            and snapshots["after_drop"]["cpu_tasks"] == initial + increment,
            f"{label} CPU task lifecycle changed")
    require(snapshots["after_drop"]["workers"] == 0
            and snapshots["after_drop"]["io_concurrency"] == 0,
            f"{label} worker/I/O permits were not released")
    require(resources.get("worker_and_io_released") is True
            and resources.get("cpu_tasks_within_limit") is True,
            f"{label} resource witness failed")
    return {
        "limits": limit_values,
        "snapshots": snapshots,
        "worker_and_io_released": True,
        "cpu_tasks_within_limit": True,
    }


def _source_metrics(sample: Mapping[str, Any], case: Mapping[str, Any], feature: bool,
                    label: str) -> dict[str, Any]:
    value = sample.get("source_metrics", sample.get("source"))
    require(isinstance(value, dict), f"{label} source metrics are missing")
    route = _case_value(case, "route")
    available = feature and route in ("cfb", "parts")
    expected_availability = "source-metrics-feature" if available else (
        "not-applicable-opc-from-bytes" if route == "opc" else "unavailable-normal-build"
    )
    require(value.get("availability") == expected_availability,
            f"{label} source metric availability changed")
    fields = ("logical_calls", "requested_bytes", "returned_bytes", "short_reads",
              "active_reads_after_operation", "max_simultaneous_reads",
              "request_size_histogram")
    if not available:
        for name in fields:
            require(value.get(name) is None, f"{label} disabled source metric {name} is present")
        return {"availability": expected_availability, **{name: None for name in fields}}
    calls = _nonnegative_int(value.get("logical_calls"), f"{label}.logical_calls")
    requested = _nonnegative_int(value.get("requested_bytes"), f"{label}.requested_bytes")
    returned = _nonnegative_int(value.get("returned_bytes"), f"{label}.returned_bytes")
    short_reads = _nonnegative_int(value.get("short_reads"), f"{label}.short_reads")
    active = _nonnegative_int(value.get("active_reads_after_operation"),
                              f"{label}.active_reads_after_operation")
    maximum = _nonnegative_int(value.get("max_simultaneous_reads"),
                               f"{label}.max_simultaneous_reads")
    workers = _positive_int(_case_value(case, "workers"), "case.workers")
    require(maximum <= workers, f"{label} source concurrency exceeds workers")
    require(active == 0 and short_reads == 0, f"{label} source reads did not quiesce")
    histogram = value.get("request_size_histogram")
    require(isinstance(histogram, list) and len(histogram) == HISTOGRAM_BINS,
            f"{label} request histogram shape changed")
    histogram = [_nonnegative_int(item, f"{label}.request_size_histogram") for item in histogram]
    require(sum(histogram) == calls, f"{label} request histogram does not sum to calls")
    state = _case_value(case, "state")
    expected_calls = 64 if state == "fresh" else 0
    require(calls == expected_calls, f"{label} source call count changed")
    require(returned <= requested, f"{label} returned source bytes exceed requested bytes")
    if calls == 0:
        require(requested == 0 and returned == 0, f"{label} primed source bytes are nonzero")
        require(maximum == 0, f"{label} primed source concurrency is nonzero")
    else:
        require(requested > 0 and returned > 0, f"{label} fresh source bytes are empty")
    return {
        "availability": expected_availability,
        "logical_calls": calls,
        "requested_bytes": requested,
        "returned_bytes": returned,
        "short_reads": short_reads,
        "active_reads_after_operation": active,
        "max_simultaneous_reads": maximum,
        "request_size_histogram": histogram,
    }


def _verification(sample: Mapping[str, Any], expected: Mapping[str, Any],
                  oracle: Mapping[str, Mapping[str, Any]], label: str) -> dict[str, Any]:
    value = sample.get("verification")
    require(isinstance(value, dict), f"{label} verification is missing")
    expected_shape = _case_value(expected, "shape")
    expected_oracle = oracle[expected_shape]
    require(value.get("ordered") is True
            and value.get("all_member_sha256_match") is True
            and value.get("members") == MEMBER_COUNT
            and value.get("logical_bytes") == expected_oracle["logical_bytes"]
            and value.get("sequence_sha256") == expected_oracle["sequence_sha256"],
            f"{label} payload verification failed")
    return {
        "ordered": True,
        "all_member_sha256_match": True,
        "members": MEMBER_COUNT,
        "logical_bytes": expected_oracle["logical_bytes"],
        "sequence_sha256": expected_oracle["sequence_sha256"],
    }


def _allocation_contract(sample: Mapping[str, Any], label: str) -> tuple[dict[str, Any], dict[str, int]]:
    value = sample.get("allocation")
    require(isinstance(value, dict), f"{label} allocation metrics are missing")
    require(value.get("status") == "measured"
            and value.get("scope") == "operation_global_system_allocator",
            f"{label} allocation observation is not measured")
    names = ("allocation_calls", "deallocation_calls", "reallocation_calls",
             "failed_allocation_calls", "allocated_bytes", "deallocated_bytes",
             "live_bytes_before", "live_bytes_after", "peak_live_bytes_before",
             "peak_live_bytes_after", "region_peak_live_bytes")
    numbers = {name: _nonnegative_int(value.get(name), f"{label}.allocation.{name}")
               for name in names}
    require(numbers["failed_allocation_calls"] == 0,
            f"{label} observed failed allocation")
    require(numbers["live_bytes_after"]
            == numbers["live_bytes_before"] + numbers["allocated_bytes"]
            - numbers["deallocated_bytes"],
            f"{label} allocation live-byte conservation failed")
    require(numbers["peak_live_bytes_after"] >= numbers["peak_live_bytes_before"],
            f"{label} allocator high-water mark decreased")
    require(numbers["region_peak_live_bytes"] >= numbers["live_bytes_before"]
            and numbers["region_peak_live_bytes"] >= numbers["live_bytes_after"]
            and numbers["region_peak_live_bytes"] <= numbers["peak_live_bytes_after"],
            f"{label} operation peak bounds failed")
    derived = {
        "region_peak_live_bytes_minus_entry":
        numbers["region_peak_live_bytes"] - numbers["live_bytes_before"],
        "retained_live_bytes_delta":
        numbers["live_bytes_after"] - numbers["live_bytes_before"],
    }
    return dict(value), {**numbers, **derived}


def validate_report(path: Path, case: Mapping[str, Any], samples: int, warmup: int,
                    feature: bool = False, allocation: bool = False) -> dict[str, Any]:
    """Validate one retained probe report and return lossless timing vectors."""
    path = Path(path)
    report = _read_json(path)
    require(isinstance(report, dict), "report is not an object")
    label = str(path)
    expected_schema = ("litchi.execution-allocation-observer.v1"
                       if allocation else REPORT_SCHEMA)
    require(report.get("schema") == expected_schema, f"{label} report schema changed")
    expected = _case_copy(case)
    config = report.get("config")
    require(isinstance(config, dict), f"{label} config is missing")
    for name, expected_value in expected.items():
        require(config.get(name) == expected_value, f"{label} config.{name} differs from case")
    require(config.get("samples") == samples and config.get("warmup") == warmup
            and config.get("aggregate_parallel_bytes") == 64 * 1024
            and config.get("cpu_task_limit") == CPU_TASK_LIMIT,
            f"{label} report configuration changed")
    oracle = payload_oracle()
    corpus = check_corpus(report, expected["shape"], oracle, label)
    rows = report.get("samples")
    require(isinstance(rows, list) and len(rows) == samples,
            f"{label} sample count changed")
    metrics = report.get("metrics")
    require(isinstance(metrics, dict)
            and metrics.get("source_metrics_feature") is feature,
            f"{label} source metric feature flag changed")
    cpu_clock = metrics.get("cpu_clock")
    require(isinstance(cpu_clock, str), f"{label} CPU clock description is missing")
    if allocation:
        require(expected["route"] == "parts", f"{label} allocation route changed")
        require(metrics.get("instrumentation_identity")
                == "system_allocator_operation_scoped",
                f"{label} allocator instrumentation identity changed")
        require(metrics.get("counter_revision") == "serialized_region_peak_v3",
                f"{label} allocator counter revision changed")
        require(metrics.get("allocator_identity")
                == "CountingSystemAllocator(std::alloc::System)",
                f"{label} allocator identity changed")
        require(cpu_clock == "not-used-by-allocation-observer",
                f"{label} allocation report advertises CPU timing")
    walls: list[float] = []
    cpus: list[float] = []
    resources: list[dict[str, Any]] = []
    source_rows: list[dict[str, Any]] = []
    verifications: list[dict[str, Any]] = []
    allocation_rows: list[dict[str, Any]] = []
    allocation_derived: list[dict[str, int]] = []
    sequence_hashes: set[str] = set()
    for index, sample in enumerate(rows):
        require(isinstance(sample, dict), f"{label} sample {index} is malformed")
        if allocation:
            require("wall_ns" not in sample and "cpu_ns" not in sample,
                    f"{label} allocation sample carries timing fields")
            allocation_row, derived = _allocation_contract(
                sample, f"{label} sample {index}")
            allocation_rows.append(allocation_row)
            allocation_derived.append(derived)
        else:
            wall = _finite(sample.get("wall_ns"), f"{label} sample {index}.wall_ns")
            require(wall > 0, f"{label} sample {index}.wall_ns is not positive")
            walls.append(wall)
            cpu_value = sample.get("cpu_ns")
            if cpu_value is not None:
                cpu = _finite(cpu_value, f"{label} sample {index}.cpu_ns")
                require(cpu >= 0, f"{label} sample {index}.cpu_ns is negative")
                cpus.append(cpu)
        verifications.append(_verification(sample, expected, oracle, f"{label} sample {index}"))
        sequence_hashes.add(verifications[-1]["sequence_sha256"])
        resources.append(_resource_contract(sample, expected, f"{label} sample {index}"))
        source_rows.append(_source_metrics(sample, expected, feature,
                                           f"{label} sample {index}"))
    require(len(sequence_hashes) == 1, f"{label} output sequence changed across samples")
    require(all(not isinstance(row, dict) or row.get("status") == "measured"
                 for row in allocation_rows),
            f"{label} allocation status changed")
    require(not cpus or len(cpus) == len(rows), f"{label} CPU metric availability changed")
    if not allocation and "ProcessCPUTime" in cpu_clock:
        require(len(cpus) == len(rows), f"{label} ProcessCPUTime samples are incomplete")
    identity = dict(corpus)
    return {
        "report": {"path": str(path), "bytes": path.stat().st_size,
                    "sha256": _sha256_file(path)},
        "case": expected,
        "samples": len(rows),
        "warmup": warmup,
        "wall_ns": None if allocation else walls,
        "cpu_ns": None if allocation else (cpus if cpus else None),
        "p50_ns": None if allocation else nearest_rank(walls, 0.50),
        "p95_ns": None if allocation else nearest_rank(walls, 0.95),
        "p99_ns": None if allocation else nearest_rank(walls, 0.99),
        "mean_ns": None if allocation else statistics.fmean(walls),
        "corpus": identity,
        "corpus_identity": identity,
        "corpus_fingerprint": identity["fingerprint"],
        "resources": resources,
        "source_metrics": source_rows,
        "verification": verifications,
        "verification_ok": True,
        "allocations": allocation_rows if allocation else None,
        "allocation_metrics": ({
            "allocation_calls": [row["allocation_calls"] for row in allocation_derived],
            "allocated_bytes": [row["allocated_bytes"] for row in allocation_derived],
            "region_peak_live_bytes_minus_entry": [
                row["region_peak_live_bytes_minus_entry"] for row in allocation_derived
            ],
            "retained_live_bytes_delta": [
                row["retained_live_bytes_delta"] for row in allocation_derived
            ],
        } if allocation else None),
        "allocation_ok": allocation if allocation else None,
    }


def parse_report(path: Path, case: Mapping[str, Any], samples: int, warmup: int,
                 feature: bool = False, allocation: bool = False) -> dict[str, Any]:
    """Compatibility spelling for callers that used the 0788 reader."""
    return validate_report(path, case, samples, warmup, feature, allocation)


def nearest_rank(values: Iterable[float], quantile: float) -> float:
    """Return the nearest-rank quantile without interpolation."""
    require(isinstance(quantile, (int, float)) and not isinstance(quantile, bool)
            and 0.0 <= float(quantile) <= 1.0, "quantile is outside [0, 1]")
    ordered = sorted(_finite(value, "quantile input") for value in values)
    require(ordered, "nearest-rank received no values")
    rank = max(1, math.ceil(float(quantile) * len(ordered)))
    return ordered[min(rank, len(ordered)) - 1]


def bootstrap(values: Sequence[float], *, seed: int = BOOTSTRAP_SEED,
              resamples: int = BOOTSTRAP_RESAMPLES,
              endpoint_indexes: tuple[int, int] = BOOTSTRAP_INDEXES) -> dict[str, Any]:
    """Deterministic median bootstrap retained for paired summaries."""
    require(values, "bootstrap received no values")
    numbers = [_finite(value, "bootstrap input") for value in values]
    require(resamples > max(endpoint_indexes) >= 0, "bootstrap endpoint indexes are invalid")
    rng = random.Random(seed)
    estimates = [statistics.median(rng.choice(numbers) for _ in numbers)
                 for _ in range(resamples)]
    estimates.sort()
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
    """Summarize paired ``after / before`` ratios with a seeded bootstrap."""
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
        "bootstrap": interval,
    }
    if interval is None:
        result.update({"estimate": None, "ci95_low": None, "ci95_high": None})
    else:
        result.update({"estimate": interval["estimate"], "ci95_low": interval["lower"],
                       "ci95_high": interval["upper"]})
    return result


def paired_metric(before: Sequence[float], after: Sequence[float],
                  name: str = "metric") -> dict[str, Any]:
    """Retain paired differences and ratio bootstrap diagnostics."""
    require(len(before) == len(after) and before,
            f"{name} pair cardinality is empty or differs")
    old = [_finite(value, f"{name}.before") for value in before]
    new = [_finite(value, f"{name}.after") for value in after]
    deltas = [right - left for left, right in zip(old, new)]
    difference = bootstrap(deltas)
    result = bootstrap_ratio(old, new, name)
    result.update({"before_block_values": old, "after_block_values": new,
                   "block_deltas": deltas, "difference_bootstrap": difference,
                   "difference_estimate": difference["estimate"],
                   "difference_ci95_low": difference["lower"],
                   "difference_ci95_high": difference["upper"]})
    return result


def _proc_kib(text: str, field: str) -> list[int]:
    pattern = re.compile(rf"(?m)^\s*{re.escape(field)}:\s*(\d+)\s+kB(?:\s|$)")
    return [int(match.group(1)) for match in pattern.finditer(text)]


def _stat_identity(text: str) -> tuple[int, int, int]:
    close = text.rfind(")")
    require(close > 0, "raw /proc/stat is malformed")
    try:
        pid = int(text[:text.index(" ")])
        fields = text[close + 2:].split()
        # fields[0] is state; fields[1] is PPid and fields[19] is starttime
        return pid, int(fields[1]), int(fields[19])
    except (ValueError, IndexError) as error:
        fail(f"raw /proc/stat identity is malformed: {error}")


def parse_snapshot_rss(snapshot: Mapping[str, Any] | Mapping[str, str]) -> dict[str, Any]:
    """Parse one acknowledged raw ``/proc`` snapshot and compare RSS views.

    The result describes one point-in-time resident set observation.  It is
    deliberately not called a peak: peak residency needs a time-series or an
    exit-accounting source and is outside this parser's evidence.
    """
    raw_value: Any = snapshot.get("raw") if isinstance(snapshot, dict) else None
    raw = raw_value if isinstance(raw_value, dict) else snapshot
    require(isinstance(raw, dict), "process snapshot raw fields are missing")
    smaps = raw.get("smaps")
    rollup = raw.get("smaps_rollup")
    status = raw.get("status")
    stat = raw.get("stat")
    require(isinstance(smaps, str) and isinstance(rollup, str) and isinstance(status, str),
            "process snapshot lacks smaps, smaps_rollup, or status text")
    smaps_values = _proc_kib(smaps, "Rss")
    rollup_values = _proc_kib(rollup, "Rss")
    status_values = _proc_kib(status, "VmRSS")
    require(smaps_values, "smaps contains no Rss fields")
    require(len(rollup_values) == 1 and len(status_values) == 1,
            "smaps_rollup/status RSS fields are not unique")
    smaps_kib = sum(smaps_values)
    rollup_kib = rollup_values[0]
    status_kib = status_values[0]
    comparisons = {
        "smaps_vs_rollup_equal": smaps_kib == rollup_kib,
        "smaps_vs_status_equal": smaps_kib == status_kib,
        "rollup_vs_status_equal": rollup_kib == status_kib,
    }
    stat_pid: int | None = None
    stat_parent: int | None = None
    stat_start: int | None = None
    if stat is not None:
        require(isinstance(stat, str), "raw /proc/stat is not text")
        stat_pid, stat_parent, stat_start = _stat_identity(stat)
    top_pid = snapshot.get("pid") if isinstance(snapshot, dict) else None
    parent_pid = ((snapshot.get("ppid", snapshot.get("parent_pid")))
                  if isinstance(snapshot, dict) else None)
    require(top_pid is None or stat_pid is None or top_pid == stat_pid,
            "snapshot PID disagrees with raw /proc/stat")
    require(parent_pid is None or stat_parent is None or parent_pid == stat_parent,
            "snapshot parent PID disagrees with raw /proc/stat")
    result = {
        "pid": top_pid if top_pid is not None else stat_pid,
        "ppid": parent_pid if parent_pid is not None else stat_parent,
        "parent_pid": parent_pid if parent_pid is not None else stat_parent,
        "starttime": (snapshot.get("starttime") if isinstance(snapshot, dict)
                       and snapshot.get("starttime") is not None else stat_start),
        "phase": snapshot.get("phase") if isinstance(snapshot, dict) else None,
        "sample": snapshot.get("sample") if isinstance(snapshot, dict) else None,
        "smaps_sum_rss_kib": smaps_kib,
        "smaps_rollup_rss_kib": rollup_kib,
        "status_vmrss_kib": status_kib,
        "smaps_sum_rss_bytes": smaps_kib * 1024,
        "smaps_rollup_rss_bytes": rollup_kib * 1024,
        "status_vmrss_bytes": status_kib * 1024,
        "rss_bytes": smaps_kib * 1024,
        "comparisons": comparisons,
        "scope": "acknowledged single-process /proc snapshot; not a physical peak",
    }
    return result


def parse_acknowledged_rss(snapshot: Mapping[str, Any] | Mapping[str, str]) -> dict[str, Any]:
    """Alias emphasizing that the caller must have performed the handshake."""
    return parse_snapshot_rss(snapshot)


def parse_proc_snapshot(snapshot: Mapping[str, Any] | Mapping[str, str]) -> dict[str, Any]:
    """Compatibility spelling for callers that use the raw proc terminology."""
    return parse_snapshot_rss(snapshot)


def parse_snapshot_series(snapshots: Iterable[Mapping[str, Any]]) -> list[dict[str, Any]]:
    """Parse a retained series and enforce one benchmark PID identity."""
    rows = [parse_snapshot_rss(snapshot) for snapshot in snapshots]
    require(rows, "snapshot series is empty")
    identities = {(row["pid"], row["ppid"], row["starttime"]) for row in rows}
    require(len(identities) == 1, "snapshot series changes process identity")
    return rows
