#!/usr/bin/env python3
"""Run the 0358 XLS source-attribution ABBA evidence batch.

This driver deliberately does not build the profiler.  It accepts two
absolute paths to already-built ``xls_source_attribution`` binaries and runs
one fresh child for every retained sample.  The default batch is therefore
serial and bounded: 6 selectors * 4 ABBA legs * 500 children, with one
profiler sample and 20 profiler warmups in each child.

The profiler's implementation-local semantic oracle is frozen from the first
A1 sample for each selector cell.  Every subsequent observation, corpus
identity, binary identity, revision, and source-version result is checked
against that cell's oracle before the sample is appended to the streamed
JSONL artifact.  A malformed child or a failed gate rejects the whole batch;
there are no retries or trimmed samples.

Only the Python standard library is used.  A normal invocation is, for
example::

    python3 tools/xls_source_attribution_abba.py \
      --control-binary /path/to/control/xls_source_attribution \
      --candidate-binary /path/to/candidate/xls_source_attribution \
      --control-revision <control-revision> \
      --candidate-revision <candidate-revision> \
      --input test-data/ole/xls/ConditionalFormattingSamples.xls \
      --output-dir docs/performance/results/0358-xls-source-span-abba-20260901 \
      --tmpdir /home/zhuhe/CodeProjects/.cargo-targets/change-0358/tmp

The binary paths are resolved and hashed before any child is launched.  The
``--tmpdir`` path is used for the profiler's staged input; child stdout and
stderr are drained through bounded pipes, and known tmpfs locations are
rejected.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
import selectors as selector_io
import signal
import subprocess
import sys
import tempfile
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, BinaryIO, Iterable, Mapping, Sequence

try:  # pragma: no cover - resource is available on the supported Linux host
    import resource
except ImportError:  # pragma: no cover - keeps the module importable elsewhere
    resource = None  # type: ignore[assignment]


DRIVER_SCHEMA_VERSION = 1
PROFILER_SCHEMA_VERSION = 2
CHANGE_ID = "0358-xls-source-span-abba"
DRIVER_VERSION = "1"
DEFAULT_WARMUPS = 20
DEFAULT_SAMPLES = 500
DEFAULT_MEMORY_LIMIT_BYTES = 2 * 1024**3
HARD_MEMORY_LIMIT_BYTES = 2 * 1024**3
DEFAULT_TIMEOUT_SECONDS = 120.0
DEFAULT_CPU = 2
MAX_TEST_SAMPLES = 10
MAX_TEST_WARMUPS = 10
MAX_CHILD_OUTPUT_BYTES = 16 * 1024 * 1024
MAX_NORMALIZED_ROW_BYTES = 64 * 1024
MAX_SAMPLES_ARTIFACT_BYTES = 64 * 1024 * 1024
MAX_U64 = (1 << 64) - 1

LEGS = ("A1", "B1", "B2", "A2")
STAT_NAMES = ("p50", "mean", "p95", "p99")
COUNT_METRICS = frozenset(
    {
        "len_calls",
        "read_calls",
        "read_bytes",
        "range_union_bytes",
        "seek_calls",
        "version_calls",
    }
)
REQUIRED_METRICS = (
    "len_calls",
    "len_ns",
    "mutex_ns",
    "range_union_bytes",
    "range_union_ns",
    "read_bytes",
    "read_calls",
    "read_ns",
    "seek_calls",
    "version_calls",
    "version_ns",
)
METRIC_ALLOWLIST = frozenset(REQUIRED_METRICS)
OBSERVATION_KEYS = frozenset(
    {"kind", "worksheet_count", "worksheet_names", "cell"}
)


class DriverError(RuntimeError):
    """A fail-closed protocol, child, schema, identity, or gate error."""


@dataclass(frozen=True)
class Selector:
    key: str
    mode: str
    operation: str


SELECTORS = (
    Selector("file-source/open", "file-source", "open"),
    Selector("file-source/list", "file-source", "list"),
    Selector("file-source/one-cell", "file-source", "one-cell"),
    Selector("atomic-file/open", "atomic-file", "open"),
    Selector("atomic-file/list", "atomic-file", "list"),
    Selector("atomic-file/one-cell", "atomic-file", "one-cell"),
)
SELECTOR_BY_KEY = {selector.key: selector for selector in SELECTORS}


@dataclass
class GroupAccumulator:
    """All retained values for one selector/leg, bounded by 500 samples."""

    selector: Selector
    leg: str
    expected_samples: int
    elapsed_ns: list[int] = field(default_factory=list)
    metrics: dict[str, list[int]] = field(default_factory=dict)
    counters: dict[str, int] | None = None

    def add(self, elapsed_ns: int, metrics: Mapping[str, int]) -> None:
        if len(self.elapsed_ns) >= self.expected_samples:
            raise DriverError(
                f"{self.leg}/{self.selector.key} received more than "
                f"{self.expected_samples} samples"
            )
        if type(elapsed_ns) is not int or elapsed_ns < 0 or elapsed_ns > MAX_U64:
            raise DriverError("elapsed_ns must be a non-negative integer")
        for name, value in metrics.items():
            _require_nonnegative_int(value, f"metric {name}")
        names = set(metrics)
        if names != METRIC_ALLOWLIST:
            raise DriverError(
                f"{self.leg}/{self.selector.key} has an unexpected metric schema"
            )
        if self.metrics and names != set(self.metrics):
            raise DriverError(
                f"{self.leg}/{self.selector.key} changed its metric schema"
            )
        if not self.metrics:
            self.metrics = {name: [] for name in sorted(names)}
        if self.counters is None:
            self.counters = {
                name: value for name, value in metrics.items() if name in COUNT_METRICS
            }
        else:
            for name in COUNT_METRICS:
                if name in metrics and name in self.counters and metrics[name] != self.counters[name]:
                    raise DriverError(
                        f"{self.leg}/{self.selector.key} counter {name} "
                        "varied across retained samples"
                    )
        self.elapsed_ns.append(elapsed_ns)
        for name, value in metrics.items():
            self.metrics[name].append(value)

    @property
    def complete(self) -> bool:
        return len(self.elapsed_ns) == self.expected_samples

    def summary(self) -> dict[str, Any]:
        result: dict[str, Any] = {
            "leg": self.leg,
            "selector": self.selector.key,
            "mode": self.selector.mode,
            "operation": self.selector.operation,
            "samples": len(self.elapsed_ns),
            "expected_samples": self.expected_samples,
            "complete": self.complete,
            "elapsed_ns": percentile_stats(self.elapsed_ns) if self.elapsed_ns else None,
            "metrics": {
                name: percentile_stats(values) for name, values in sorted(self.metrics.items())
            },
            "counters": dict(sorted((self.counters or {}).items())),
        }
        return result


def _reject_duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise DriverError(f"JSON contains duplicate key {key!r}")
        result[key] = value
    return result


def _reject_nonfinite(value: str) -> None:
    raise DriverError(f"JSON contains non-finite value {value!r}")


def parse_json_bytes(raw: bytes, location: str) -> dict[str, Any]:
    """Parse one strict UTF-8 JSON object from a profiler child."""

    try:
        value = json.loads(
            raw.decode("utf-8"),
            object_pairs_hook=_reject_duplicate_keys,
            parse_constant=_reject_nonfinite,
        )
    except DriverError:
        raise
    except (
        UnicodeError,
        json.JSONDecodeError,
        RecursionError,
        ValueError,
        TypeError,
        OverflowError,
    ) as error:
        raise DriverError(f"{location} is not valid UTF-8 JSON: {error}") from error
    if not isinstance(value, dict):
        raise DriverError(f"{location} must contain a JSON object")
    return value


def canonical_json(value: Any) -> bytes:
    """Return deterministic JSON bytes for identity and artifact output."""

    try:
        return json.dumps(
            value,
            sort_keys=True,
            separators=(",", ":"),
            ensure_ascii=True,
            allow_nan=False,
        ).encode("utf-8")
    except (TypeError, ValueError, OverflowError, RecursionError) as error:
        raise DriverError(f"value is not canonical JSON: {error}") from error


def _require_mapping(value: Any, location: str) -> Mapping[str, Any]:
    if not isinstance(value, dict):
        raise DriverError(f"{location} must be an object")
    return value


def _require_string(value: Any, location: str) -> str:
    if not isinstance(value, str) or not value:
        raise DriverError(f"{location} must be a non-empty string")
    return value


def _require_nonnegative_int(
    value: Any, location: str, maximum: int = MAX_U64
) -> int:
    if type(value) is not int or value < 0 or value > maximum:
        raise DriverError(f"{location} must be a non-negative integer")
    return value


def _hash_file(path: Path, label: str) -> dict[str, Any]:
    try:
        resolved = path.expanduser().resolve(strict=True)
        stat = resolved.stat()
    except OSError as error:
        raise DriverError(f"cannot inspect {label} {path}: {error}") from error
    if not resolved.is_file():
        raise DriverError(f"{label} is not a regular file: {resolved}")
    digest = hashlib.sha256()
    total = 0
    try:
        with resolved.open("rb") as stream:
            while True:
                chunk = stream.read(1024 * 1024)
                if not chunk:
                    break
                digest.update(chunk)
                total += len(chunk)
    except OSError as error:
        raise DriverError(f"cannot hash {label} {resolved}: {error}") from error
    if total != stat.st_size:
        raise DriverError(f"{label} changed while being hashed: {resolved}")
    return {"path": str(resolved), "bytes": total, "sha256": digest.hexdigest()}


def _forbidden_tmp_path(path: Path) -> bool:
    forbidden = (Path("/tmp"), Path("/dev/shm"), Path("/run/shm"), Path("/var/tmp"))
    return any(path == item or item in path.parents for item in forbidden)


def _reject_tmpfs_filesystem(path: Path) -> None:
    """Reject Linux tmpfs/ramfs even when it is mounted outside /tmp."""

    if not sys.platform.startswith("linux") or not hasattr(os, "statfs"):
        return
    try:
        filesystem_type = os.statfs(path).f_type
    except OSError as error:
        raise DriverError(f"cannot inspect TMPDIR filesystem {path}: {error}") from error
    # Linux TMPFS_MAGIC and RAMFS_MAGIC.  Other filesystem types, including
    # ordinary disk-backed overlay filesystems, remain valid evidence storage.
    if filesystem_type in (0x01021994, 0x858458F6):
        raise DriverError(f"TMPDIR is on tmpfs/ramfs: {path}")


def ensure_tmpdir(path: Path) -> tuple[Path, bool]:
    """Create and validate an on-disk temporary directory.

    The boolean says whether the directory was created by this batch and can
    safely be removed afterward.
    """

    raw = path.expanduser()
    if not raw.is_absolute():
        raw = Path.cwd() / raw
    # Check the user-supplied spelling and every existing prefix before
    # resolution.  Checking only the resolved path would hide a symlink swap.
    prefix = raw
    while True:
        try:
            if prefix.is_symlink():
                raise DriverError(f"TMPDIR path contains a symlink: {prefix}")
        except OSError as error:
            raise DriverError(f"cannot inspect TMPDIR path {prefix}: {error}") from error
        if prefix.exists() or prefix == prefix.parent:
            break
        prefix = prefix.parent
    try:
        resolved = raw.resolve()
    except OSError as error:
        raise DriverError(f"cannot resolve TMPDIR {path}: {error}") from error
    if _forbidden_tmp_path(resolved):
        raise DriverError(f"TMPDIR must not be a known tmpfs path: {resolved}")
    created = False
    if resolved.exists():
        if not resolved.is_dir():
            raise DriverError(f"TMPDIR is not a non-symlink directory: {resolved}")
    else:
        try:
            resolved.mkdir(parents=True, exist_ok=False)
        except OSError as error:
            raise DriverError(f"cannot create TMPDIR {resolved}: {error}") from error
        created = True
    try:
        _reject_tmpfs_filesystem(resolved)
        with tempfile.NamedTemporaryFile(dir=resolved, prefix=".0358-probe-", delete=True):
            pass
    except OSError as error:
        raise DriverError(f"TMPDIR is not writable: {resolved}: {error}") from error
    except DriverError:
        if created:
            try:
                resolved.rmdir()
            except OSError:
                pass
        raise
    return resolved, created


def require_enforceable_limits(memory_limit_bytes: int) -> None:
    """Refuse evidence collection if child limits cannot be enforced."""

    if os.name != "posix" or resource is None:
        raise DriverError("evidence execution requires POSIX resource limits")
    if not hasattr(resource, "RLIMIT_AS"):
        raise DriverError("evidence execution requires resource.RLIMIT_AS")
    if type(memory_limit_bytes) is not int or memory_limit_bytes <= 0:
        raise DriverError("memory limit must be a positive integer")
    if memory_limit_bytes > HARD_MEMORY_LIMIT_BYTES:
        raise DriverError(f"memory limit cannot exceed {HARD_MEMORY_LIMIT_BYTES} bytes")
    try:
        address_hard = resource.getrlimit(resource.RLIMIT_AS)[1]
    except (OSError, ValueError) as error:
        raise DriverError(f"cannot inspect child resource limits: {error}") from error
    infinity = getattr(resource, "RLIM_INFINITY", -1)
    if address_hard != infinity and address_hard < memory_limit_bytes:
        raise DriverError("parent RLIMIT_AS hard limit is below requested child limit")


def percentile_stats(values: Sequence[int]) -> dict[str, int | float]:
    """Calculate the predeclared no-trimming statistics for all samples.

    p50 is the lower middle value for an even sample count.  p95 and p99 use
    nearest-rank selection (ceil(p*n)).  These definitions avoid interpolation
    and are stable for the required 500 retained samples.
    """

    if not values:
        raise DriverError("cannot summarize an empty sample set")
    if any(type(value) is not int or value < 0 or value > MAX_U64 for value in values):
        raise DriverError("statistics require non-negative integer samples")
    ordered = sorted(values)
    count = len(ordered)

    def nearest_rank(percent: float) -> int:
        index = max(1, math.ceil(percent * count)) - 1
        return ordered[index]

    return {
        "count": count,
        "min": ordered[0],
        "max": ordered[-1],
        "mean": sum(ordered) / count,
        "p50": ordered[(count - 1) // 2],
        "p95": nearest_rank(0.95),
        "p99": nearest_rank(0.99),
    }


def _value_comparison(left: int | float, right: int | float) -> dict[str, int | float | None]:
    delta = left - right
    if left == 0:
        improvement: float | None = 0.0 if right == 0 else None
    else:
        improvement = (delta * 100.0) / left
    return {"delta": delta, "improvement_percent": improvement, "left": left, "right": right}


def compare_stats(left: Mapping[str, Any], right: Mapping[str, Any]) -> dict[str, dict[str, int | float | None]]:
    result: dict[str, dict[str, int | float | None]] = {}
    for name in STAT_NAMES:
        left_value = left.get(name)
        right_value = right.get(name)
        if not isinstance(left_value, (int, float)) or isinstance(left_value, bool):
            raise DriverError(f"left statistic {name} is invalid")
        if not isinstance(right_value, (int, float)) or isinstance(right_value, bool):
            raise DriverError(f"right statistic {name} is invalid")
        result[name] = _value_comparison(left_value, right_value)
    return result


def expected_observation(semantic_oracle: Mapping[str, Any], selector: Selector) -> dict[str, Any]:
    """Project the frozen source oracle into the selected operation result."""

    source = _require_mapping(
        semantic_oracle.get("source_implementation_projection"),
        "semantic_oracle.source_implementation_projection",
    )
    source_count = _require_nonnegative_int(
        source.get("worksheet_count"), "source worksheet_count"
    )
    source_names = source.get("worksheet_names")
    if type(source_names) is not list or any(type(name) is not str for name in source_names):
        raise DriverError("source worksheet_names must be a string array")
    source_cell = source.get("selected_cell")
    if source_cell is not None and type(source_cell) is not str:
        raise DriverError("source selected_cell must be a string or null")
    if selector.operation == "open":
        return {
            "kind": "open",
            "worksheet_count": source_count,
            "worksheet_names": None,
            "cell": None,
        }
    if selector.operation == "list":
        return {
            "kind": "list",
            "worksheet_count": None,
            "worksheet_names": source_names,
            "cell": None,
        }
    if selector.operation == "one-cell":
        return {
            "kind": "one-cell",
            "worksheet_count": None,
            "worksheet_names": None,
            "cell": source_cell,
        }
    raise DriverError(f"unknown selector operation {selector.operation!r}")


def validate_observation(
    observation: Any, expected: Mapping[str, Any], location: str
) -> dict[str, Any]:
    """Validate observation types before comparing values.

    Python considers ``True == 1`` and ``1.0 == 1``.  A profiler result must
    not be able to exploit that equivalence when matching the frozen oracle.
    """

    value = _require_mapping(observation, location)
    if set(value) != OBSERVATION_KEYS:
        raise DriverError(f"{location} has an unexpected field set")
    if type(value["kind"]) is not str:
        raise DriverError(f"{location}.kind must be a string")
    if value["worksheet_count"] is not None:
        _require_nonnegative_int(value["worksheet_count"], f"{location}.worksheet_count")
    names = value["worksheet_names"]
    if names is not None:
        if type(names) is not list or any(type(name) is not str for name in names):
            raise DriverError(f"{location}.worksheet_names must be a string array or null")
    cell = value["cell"]
    if cell is not None and type(cell) is not str:
        raise DriverError(f"{location}.cell must be a string or null")
    expected_value = _require_mapping(expected, f"{location}.expected")
    if canonical_json(dict(value)) != canonical_json(dict(expected_value)):
        raise DriverError(
            f"{location} disagrees with the frozen source oracle: "
            f"expected {dict(expected_value)!r}, got {dict(value)!r}"
        )
    return dict(value)


def build_protocol(
    *,
    corpus: Mapping[str, Any],
    control_binary: Mapping[str, Any],
    candidate_binary: Mapping[str, Any],
    tmpdir: str,
    cpu: int | None,
    memory_limit_bytes: int,
    timeout_seconds: float,
    warmups: int = DEFAULT_WARMUPS,
    samples: int = DEFAULT_SAMPLES,
    test_mode: bool = False,
    cwd: str | None = None,
) -> dict[str, Any]:
    """Build the immutable protocol declaration before collection."""

    if type(warmups) is not int or type(samples) is not int:
        raise DriverError("warmups and samples must be integers")
    if warmups <= 0 or samples <= 0:
        raise DriverError("warmups and samples must be greater than zero")
    if type(memory_limit_bytes) is not int or memory_limit_bytes <= 0:
        raise DriverError("memory limit must be positive")
    if memory_limit_bytes > HARD_MEMORY_LIMIT_BYTES:
        raise DriverError(
            f"memory limit cannot exceed {HARD_MEMORY_LIMIT_BYTES} bytes"
        )
    if type(test_mode) is not bool:
        raise DriverError("test_mode must be a boolean")
    if not test_mode and (warmups != DEFAULT_WARMUPS or samples != DEFAULT_SAMPLES):
        raise DriverError(
            "evidence protocol requires exactly 20 warmups and 500 retained samples"
        )
    if test_mode and (warmups > MAX_TEST_WARMUPS or samples > MAX_TEST_SAMPLES):
        raise DriverError("test-mode protocol exceeds its explicit small-run limits")
    if type(timeout_seconds) not in (int, float) or isinstance(timeout_seconds, bool):
        raise DriverError("timeout must be a number")
    if timeout_seconds <= 0 or not math.isfinite(timeout_seconds):
        raise DriverError("timeout must be a finite positive number")
    return {
        "schema_version": DRIVER_SCHEMA_VERSION,
        "change": CHANGE_ID,
        "driver": {"name": "xls_source_attribution_abba", "version": DRIVER_VERSION},
        "corpus": dict(corpus),
        "binaries": {"control": dict(control_binary), "candidate": dict(candidate_binary)},
        "selector_order": [selector.key for selector in SELECTORS],
        "selectors": [
            {
                "id": selector.key,
                "mode": selector.mode,
                "operation": selector.operation,
                "worksheet_index": 1,
                "row": 1,
                "column": 0,
            }
            for selector in SELECTORS
        ],
        "collection": {
            "order": list(LEGS),
            "workers": 1,
            "serial": True,
            "fresh_child_per_sample": True,
            "warmups_per_child": warmups,
            "retained_samples_per_selector_per_leg": samples,
            "total_children": len(SELECTORS) * len(LEGS) * samples,
            "no_retries": True,
            "test_mode": test_mode,
            "evidence_eligible": not test_mode,
        },
        "cache_state": "read-only staged snapshot with warm page cache; cold and physical I/O are not measured",
        "runtime": {
            "cwd": cwd,
            "tmpdir": tmpdir,
            "cpu": cpu,
            "memory_limit_bytes": memory_limit_bytes,
            "timeout_seconds": timeout_seconds,
            "cargo_build_jobs": 1,
            "child_output_limit_bytes": MAX_CHILD_OUTPUT_BYTES,
            "child_capture": "nonblocking combined stdout/stderr pipes",
            "normalized_row_limit_bytes": MAX_NORMALIZED_ROW_BYTES,
            "samples_artifact_limit_bytes": MAX_SAMPLES_ARTIFACT_BYTES,
        },
        "statistics": {
            "retained_samples": "all samples; no trimming, retries, or winsorization",
            "p50": "lower middle order statistic for an even sample count",
            "p95": "nearest rank ceil(0.95*n)",
            "p99": "nearest rank ceil(0.99*n)",
            "mean": "arithmetic mean",
        },
        "gates": {
            "same_side_max_abs_percent": 5.0,
            "candidate_tail_max_regression_percent": 5.0,
            "claim_selector": "file-source/one-cell",
            "claim_minimum_p50_and_mean_improvement_percent": 1.0,
            "guards": [
                "file-source/open",
                "file-source/list",
                "atomic-file/open",
                "atomic-file/list",
                "atomic-file/one-cell",
            ],
            "claim_direction": "A1->B1 and A2->B2 must both pass",
            "same_side_pairs": ["A1->A2", "B1->B2"],
            "test_mode_can_accept": False,
        },
    }


def _validate_identity(
    observed: Any, expected: Mapping[str, Any], location: str
) -> str:
    value = _require_mapping(observed, location)
    actual_bytes = _require_nonnegative_int(value.get("bytes"), f"{location}.bytes")
    actual_hash = _require_string(value.get("sha256"), f"{location}.sha256")
    actual_path = _require_string(value.get("path"), f"{location}.path")
    expected_path = expected.get("path")
    if type(expected_path) is not str or not expected_path:
        raise DriverError(f"{location} expected identity has no canonical path")
    if actual_path != expected_path:
        raise DriverError(
            f"{location} path mismatch: expected canonical {expected_path!r}, "
            f"got {actual_path!r}"
        )
    if actual_bytes != expected["bytes"] or actual_hash != expected["sha256"]:
        raise DriverError(
            f"{location} identity mismatch: expected {expected['bytes']} bytes "
            f"{expected['sha256']}, got {actual_bytes} bytes {actual_hash}"
        )
    return actual_path


def validate_child_report(
    report: Mapping[str, Any],
    *,
    selector: Selector,
    leg: str,
    input_identity: Mapping[str, Any],
    binary_identity: Mapping[str, Any],
    warmups: int,
    semantic_oracles: dict[str, Mapping[str, Any]],
    seen_revisions: Mapping[str, str],
    seen_binary_revisions: Mapping[str, str] | None = None,
    expected_revision: str | None = None,
) -> tuple[dict[str, Any], Mapping[str, Any], str, str, str, str]:
    """Validate and normalize one profiler report.

    The return value is ``(record, semantic_oracles, revision, tool_revision,
    reported_binary_path, reported_input_path)``.  Each selector cell's oracle
    is established only by its first A1 sample; later calls must pass the same
    oracle for that cell.
    """

    if not isinstance(report, dict):
        raise DriverError("child report must be an object")
    schema_version = _require_nonnegative_int(report.get("schema_version"), "child schema_version")
    if schema_version != PROFILER_SCHEMA_VERSION:
        raise DriverError(
            f"{leg}/{selector.key} has unsupported profiler schema "
            f"{report.get('schema_version')!r}"
        )
    if report.get("mode") != selector.mode or report.get("operation") != selector.operation:
        raise DriverError(f"{leg}/{selector.key} child mode/operation mismatch")
    report_warmups = _require_nonnegative_int(report.get("warmups"), "child warmups")
    report_samples = _require_nonnegative_int(report.get("samples"), "child samples")
    if type(warmups) is not int or report_warmups != warmups or report_samples != 1:
        raise DriverError(f"{leg}/{selector.key} must report warmups={warmups}, samples=1")
    worksheet_index = _require_nonnegative_int(report.get("worksheet_index"), "child worksheet_index")
    row = _require_nonnegative_int(report.get("row"), "child row")
    column = _require_nonnegative_int(report.get("column"), "child column")
    if worksheet_index != 1 or row != 1 or column != 0:
        raise DriverError(f"{leg}/{selector.key} coordinate mismatch")

    reported_input_path = _validate_identity(report.get("input"), input_identity, "child input")
    reported_binary_path = _validate_identity(report.get("binary"), binary_identity, "child binary")

    revision = _require_string(report.get("revision"), "child revision")
    if expected_revision is not None and revision != expected_revision:
        raise DriverError(
            f"{leg}/{selector.key} revision mismatch: expected {expected_revision}, got {revision}"
        )
    previous_revision = seen_revisions.get(leg)
    if previous_revision is not None and previous_revision != revision:
        raise DriverError(f"{leg} changed revision during collection")
    binary_role = "control" if leg.startswith("A") else "candidate"
    if seen_binary_revisions is not None:
        previous_binary_revision = seen_binary_revisions.get(binary_role)
        if previous_binary_revision is not None and previous_binary_revision != revision:
            raise DriverError(f"{binary_role} revision changed across ABBA legs")
    tool = _require_mapping(report.get("tool"), "child tool")
    tool_revision = _require_string(tool.get("revision"), "child tool.revision")
    if tool_revision != revision:
        raise DriverError(f"{leg}/{selector.key} top-level/tool revision mismatch")

    child_oracle = _require_mapping(report.get("semantic_oracle"), "child semantic_oracle")
    cell_key = selector.key
    frozen_oracle = semantic_oracles.get(cell_key)
    if frozen_oracle is None:
        if leg != "A1":
            raise DriverError(
                f"{leg}/{cell_key} semantic oracle is missing; only A1 may establish it"
            )
        frozen_oracle = dict(child_oracle)
        semantic_oracles[cell_key] = frozen_oracle
    elif canonical_json(dict(child_oracle)) != canonical_json(dict(frozen_oracle)):
        raise DriverError(f"{leg}/{cell_key} semantic oracle changed")
    expected = expected_observation(frozen_oracle, selector)

    records = report.get("records")
    elapsed_samples = report.get("elapsed_samples_ns")
    if not isinstance(records, list) or len(records) != 1:
        raise DriverError(f"{leg}/{selector.key} must contain exactly one record")
    if not isinstance(elapsed_samples, list) or len(elapsed_samples) != 1:
        raise DriverError(f"{leg}/{selector.key} must contain exactly one elapsed sample")
    _require_nonnegative_int(elapsed_samples[0], "elapsed_samples_ns[0]")
    record = _require_mapping(records[0], "child record")
    elapsed_ns = _require_nonnegative_int(record.get("elapsed_ns"), "record.elapsed_ns")
    if elapsed_samples[0] != elapsed_ns:
        raise DriverError(f"{leg}/{selector.key} elapsed_samples_ns disagrees with record")
    if record.get("source_version_stable") is not True:
        raise DriverError(f"{leg}/{selector.key} source version was not stable")
    if record.get("eager_phases") is not None:
        raise DriverError(f"{leg}/{selector.key} unexpectedly reported eager phases")
    observation = validate_observation(
        record.get("observation"), expected, f"{leg}/{selector.key} observation"
    )
    metrics = _require_mapping(record.get("metrics"), "record.metrics")
    for name in metrics:
        _require_string(name, "metric name")
    metric_names = set(metrics)
    if metric_names != METRIC_ALLOWLIST:
        missing = sorted(METRIC_ALLOWLIST - metric_names)
        extra = sorted(metric_names - METRIC_ALLOWLIST)
        raise DriverError(
            f"{leg}/{selector.key} metric allowlist mismatch; missing={missing}, extra={extra}"
        )
    normalized_metrics: dict[str, int] = {}
    for name, value in metrics.items():
        normalized_metrics[_require_string(name, "metric name")] = _require_nonnegative_int(
            value, f"metric {name}"
        )

    normalized_record = {
        "elapsed_ns": elapsed_ns,
        "metrics": dict(sorted(normalized_metrics.items())),
        "observation": dict(observation) if isinstance(observation, dict) else observation,
        "source_version_stable": True,
    }
    return (
        normalized_record,
        semantic_oracles,
        revision,
        tool_revision,
        reported_binary_path,
        reported_input_path,
    )


def _comparison_for_groups(
    left: GroupAccumulator,
    right: GroupAccumulator,
    *,
    kind: str,
) -> dict[str, Any]:
    if not left.complete or not right.complete:
        raise DriverError(
            f"cannot compare incomplete groups {left.leg}/{left.selector.key} and "
            f"{right.leg}/{right.selector.key}"
        )
    left_elapsed = percentile_stats(left.elapsed_ns)
    right_elapsed = percentile_stats(right.elapsed_ns)
    metrics: dict[str, Any] = {}
    if set(left.metrics) != set(right.metrics):
        raise DriverError(f"metric schema differs for {left.selector.key}")
    for name in sorted(left.metrics):
        metrics[name] = compare_stats(
            percentile_stats(left.metrics[name]), percentile_stats(right.metrics[name])
        )
    return {
        "kind": kind,
        "left_leg": left.leg,
        "right_leg": right.leg,
        "selector": left.selector.key,
        "mode": left.selector.mode,
        "operation": left.selector.operation,
        "elapsed_ns": compare_stats(left_elapsed, right_elapsed),
        "metrics": metrics,
        "counters": {
            "left": dict(sorted((left.counters or {}).items())),
            "right": dict(sorted((right.counters or {}).items())),
        },
    }


def _gate_check(
    *,
    checks: list[dict[str, Any]],
    failures: list[str],
    comparison: Mapping[str, Any],
    stat: str,
    requirement: str,
    threshold: float,
) -> None:
    value = comparison["elapsed_ns"][stat]["improvement_percent"]
    passed = value is not None and (
        abs(value) <= threshold if requirement == "same_side" else value >= threshold
    )
    check = {
        "kind": requirement,
        "selector": comparison["selector"],
        "left_leg": comparison["left_leg"],
        "right_leg": comparison["right_leg"],
        "stat": stat,
        "observed_improvement_percent": value,
        "threshold_percent": threshold,
        "passed": passed,
    }
    checks.append(check)
    if not passed:
        failures.append(
            f"{requirement} gate failed for {comparison['selector']} "
            f"{comparison['left_leg']}->{comparison['right_leg']} {stat}: "
            f"observed {value!r}, threshold {threshold:g}%"
        )


def evaluate_gates(
    groups: Mapping[tuple[str, str], GroupAccumulator],
    *,
    samples: int = DEFAULT_SAMPLES,
    same_side_limit_percent: float = 5.0,
    tail_limit_percent: float = 5.0,
    claim_minimum_percent: float = 1.0,
) -> tuple[list[dict[str, Any]], dict[str, Any], list[str]]:
    """Return comparisons, gate details, and fail-closed gate messages."""

    failures: list[str] = []
    checks: list[dict[str, Any]] = []
    comparisons: list[dict[str, Any]] = []
    for selector in SELECTORS:
        for leg in LEGS:
            group = groups.get((leg, selector.key))
            if group is None or len(group.elapsed_ns) != samples:
                failures.append(f"missing complete group {leg}/{selector.key}")
    if failures:
        return comparisons, {"passed": False, "checks": checks}, failures

    for selector in SELECTORS:
        for left_leg, right_leg in (("A1", "A2"), ("B1", "B2")):
            comparison = _comparison_for_groups(
                groups[(left_leg, selector.key)],
                groups[(right_leg, selector.key)],
                kind="same_side",
            )
            comparisons.append(comparison)
            for stat in STAT_NAMES:
                _gate_check(
                    checks=checks,
                    failures=failures,
                    comparison=comparison,
                    stat=stat,
                    requirement="same_side",
                    threshold=same_side_limit_percent,
                )
            for name, left_value in comparison["counters"]["left"].items():
                right_value = comparison["counters"]["right"].get(name)
                passed = right_value == left_value
                checks.append(
                    {
                        "kind": "same_side_counter",
                        "selector": selector.key,
                        "left_leg": left_leg,
                        "right_leg": right_leg,
                        "metric": name,
                        "left": left_value,
                        "right": right_value,
                        "passed": passed,
                    }
                )
                if not passed:
                    failures.append(
                        f"same-side counter changed for {selector.key} {left_leg}->{right_leg}: "
                        f"{name} {left_value!r}->{right_value!r}"
                    )

        for left_leg, right_leg in (("A1", "B1"), ("A2", "B2")):
            comparison = _comparison_for_groups(
                groups[(left_leg, selector.key)],
                groups[(right_leg, selector.key)],
                kind="candidate",
            )
            comparisons.append(comparison)
            claim = selector.key == "file-source/one-cell"
            for stat in ("p50", "mean"):
                _gate_check(
                    checks=checks,
                    failures=failures,
                    comparison=comparison,
                    stat=stat,
                    requirement="claim_direction" if claim else "guard_direction",
                    threshold=claim_minimum_percent if claim else -tail_limit_percent,
                )
            for stat in ("p95", "p99"):
                _gate_check(
                    checks=checks,
                    failures=failures,
                    comparison=comparison,
                    stat=stat,
                    requirement="candidate_tail",
                    threshold=-tail_limit_percent,
                )

    gate_summary = {
        "passed": not failures,
        "checks": checks,
        "limits": {
            "same_side_max_abs_percent": same_side_limit_percent,
            "candidate_tail_max_regression_percent": tail_limit_percent,
            "claim_minimum_p50_and_mean_improvement_percent": claim_minimum_percent,
        },
    }
    return comparisons, gate_summary, failures


def _child_preexec(cpu: int | None, memory_limit_bytes: int):
    def apply_limits() -> None:
        if resource is None or not hasattr(resource, "RLIMIT_AS"):
            raise OSError("POSIX address-space limits are unavailable")
        if cpu is not None:
            if not hasattr(os, "sched_setaffinity"):
                raise OSError("CPU affinity is unavailable on this platform")
            os.sched_setaffinity(0, {cpu})
        resource.setrlimit(resource.RLIMIT_AS, (memory_limit_bytes, memory_limit_bytes))

    return apply_limits


def _kill_reap(process: Any) -> None:
    """Best-effort process-group termination used on cap and timeout paths."""

    try:
        if process.poll() is None:
            try:
                os.killpg(process.pid, signal.SIGKILL)
            except (OSError, AttributeError):
                process.kill()
    except (OSError, subprocess.SubprocessError, AttributeError):
        pass
    try:
        process.wait(timeout=5.0)
    except subprocess.TimeoutExpired:
        try:
            process.kill()
        except (OSError, subprocess.SubprocessError, AttributeError):
            pass
        try:
            process.wait(timeout=5.0)
        except (
            OSError,
            subprocess.SubprocessError,
            subprocess.TimeoutExpired,
            AttributeError,
        ):
            pass
    except (OSError, subprocess.SubprocessError, AttributeError):
        pass


def _bounded_capture(process: Any, timeout_seconds: float) -> tuple[bytes, bytes, int]:
    """Drain both pipes without allowing combined output to exceed the cap."""

    stdout = process.stdout
    stderr = process.stderr
    if stdout is None or stderr is None:
        _kill_reap(process)
        raise DriverError("child did not provide bounded stdout/stderr pipes")
    streams = {stdout: "stdout", stderr: "stderr"}
    buffers = {"stdout": bytearray(), "stderr": bytearray()}
    poller = selector_io.DefaultSelector()
    total = 0
    deadline = time.monotonic() + timeout_seconds
    try:
        for stream in streams:
            os.set_blocking(stream.fileno(), False)
            poller.register(stream, selector_io.EVENT_READ)
        while streams:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                _kill_reap(process)
                raise DriverError(f"child timed out after {timeout_seconds:g}s")
            events = poller.select(remaining)
            if not events:
                continue
            for key, _ in events:
                stream = key.fileobj
                remaining_cap = MAX_CHILD_OUTPUT_BYTES - total
                if remaining_cap <= 0:
                    try:
                        probe = os.read(stream.fileno(), 1)
                    except BlockingIOError:
                        continue
                    except OSError as error:
                        _kill_reap(process)
                        raise DriverError(f"bounded child capture failed: {error}") from error
                    if probe:
                        _kill_reap(process)
                        raise DriverError(
                            f"child stdout/stderr exceeded {MAX_CHILD_OUTPUT_BYTES} bytes"
                        )
                    poller.unregister(stream)
                    del streams[stream]
                    stream.close()
                    continue
                try:
                    chunk = os.read(stream.fileno(), min(65536, remaining_cap))
                except BlockingIOError:
                    continue
                except OSError as error:
                    _kill_reap(process)
                    raise DriverError(f"bounded child capture failed: {error}") from error
                if not chunk:
                    poller.unregister(stream)
                    del streams[stream]
                    stream.close()
                    continue
                total += len(chunk)
                if total > MAX_CHILD_OUTPUT_BYTES:
                    _kill_reap(process)
                    raise DriverError(
                        f"child stdout/stderr exceeded {MAX_CHILD_OUTPUT_BYTES} bytes"
                    )
                buffers[streams[stream]].extend(chunk)
    except DriverError:
        raise
    except (
        OSError,
        ValueError,
        TypeError,
        KeyError,
        AttributeError,
        RecursionError,
        OverflowError,
    ) as error:
        _kill_reap(process)
        raise DriverError(f"bounded child capture failed: {error}") from error
    finally:
        poller.close()
    remaining = max(0.0, deadline - time.monotonic())
    try:
        return_code = process.wait(timeout=remaining)
    except subprocess.TimeoutExpired as error:
        _kill_reap(process)
        raise DriverError(f"child timed out after {timeout_seconds:g}s") from error
    except (OSError, subprocess.SubprocessError) as error:
        _kill_reap(process)
        raise DriverError(f"child wait failed: {error}") from error
    return bytes(buffers["stdout"]), bytes(buffers["stderr"]), return_code


def invoke_child(
    *,
    binary: Path,
    input_path: Path,
    revision: str,
    selector: Selector,
    warmups: int,
    tmpdir: Path,
    cwd: Path,
    cpu: int | None,
    memory_limit_bytes: int,
    timeout_seconds: float,
) -> dict[str, Any]:
    """Invoke exactly one profiler child with bounded pipe capture."""

    command = [
        str(binary),
        "--input",
        str(input_path),
        "--mode",
        selector.mode,
        "--operation",
        selector.operation,
        "--warmups",
        str(warmups),
        "--samples",
        "1",
        "--worksheet-index",
        "1",
        "--row",
        "1",
        "--column",
        "0",
    ]
    environment = os.environ.copy()
    environment.update(
        {
            "TMPDIR": str(tmpdir),
            "TMP": str(tmpdir),
            "TEMP": str(tmpdir),
            "CARGO_BUILD_JOBS": "1",
            "CARGO_INCREMENTAL": "0",
            "RAYON_NUM_THREADS": "1",
            "OMP_NUM_THREADS": "1",
            "LITCHI_REVISION": revision,
        }
    )
    process = None
    try:
        process = subprocess.Popen(
            command,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            cwd=str(cwd),
            env=environment,
            close_fds=True,
            start_new_session=True,
            preexec_fn=(
                _child_preexec(cpu, memory_limit_bytes) if os.name == "posix" else None
            ),
        )
        stdout, stderr, return_code = _bounded_capture(process, timeout_seconds)
    except DriverError:
        if process is not None:
            _kill_reap(process)
        raise
    except (
        OSError,
        subprocess.SubprocessError,
        ValueError,
        TypeError,
        KeyError,
        AttributeError,
        RecursionError,
        OverflowError,
    ) as error:
        if process is not None:
            _kill_reap(process)
        raise DriverError(f"cannot launch {shlex_join(command)}: {error}") from error
    if return_code != 0:
        detail = stderr[:4096].decode("utf-8", errors="replace").strip()
        suffix = f"; stderr={detail!r}" if detail else ""
        raise DriverError(f"child exited {return_code}: {shlex_join(command)}{suffix}")
    return parse_json_bytes(stdout, f"child {selector.key}")


def shlex_join(parts: Iterable[str]) -> str:
    """Small local equivalent for shells unavailable on every target."""

    import shlex

    return shlex.join(list(parts))


def _write_bytes(path: Path, data: bytes) -> None:
    if path.exists():
        raise DriverError(f"refusing to overwrite artifact {path}")
    temporary = path.with_name(f".{path.name}.partial")
    try:
        with temporary.open("xb") as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    except OSError as error:
        try:
            temporary.unlink()
        except OSError:
            pass
        raise DriverError(f"cannot write artifact {path}: {error}") from error


def _write_json(path: Path, value: Any) -> None:
    _write_bytes(path, canonical_json(value) + b"\n")


def publish_samples(partial_path: Path, final_path: Path) -> None:
    """Atomically publish streamed samples after the verdict is final."""

    if final_path.exists():
        raise DriverError(f"refusing to overwrite samples artifact {final_path}")
    if not partial_path.is_file():
        raise DriverError(f"missing partial samples artifact {partial_path}")
    try:
        os.replace(partial_path, final_path)
    except OSError as error:
        raise DriverError(f"cannot publish samples artifact {final_path}: {error}") from error
    if partial_path.exists() or not final_path.is_file():
        raise DriverError(f"samples artifact publication was not atomic: {final_path}")


def open_samples_partial(path: Path) -> BinaryIO:
    """Open the streamed samples artifact as an exclusive binary file."""

    try:
        return path.open("xb")
    except OSError as error:
        raise DriverError(f"cannot create partial samples artifact {path}: {error}") from error


def encode_sample_row(row: Mapping[str, Any], current_bytes: int) -> bytes:
    """Canonicalize one JSONL row under both row and aggregate size caps."""

    if type(current_bytes) is not int or current_bytes < 0:
        raise DriverError("current samples artifact size must be a non-negative integer")
    try:
        encoded = canonical_json(dict(row))
    except (DriverError, RecursionError, TypeError, ValueError, OverflowError) as error:
        if isinstance(error, DriverError):
            raise
        raise DriverError(f"cannot encode normalized sample row: {error}") from error
    row_bytes = len(encoded) + 1
    if len(encoded) > MAX_NORMALIZED_ROW_BYTES:
        raise DriverError(
            f"normalized sample row exceeds {MAX_NORMALIZED_ROW_BYTES} bytes"
        )
    if current_bytes > MAX_SAMPLES_ARTIFACT_BYTES - row_bytes:
        raise DriverError(
            f"samples artifact would exceed {MAX_SAMPLES_ARTIFACT_BYTES} bytes"
        )
    return encoded + b"\n"


def resolve_verdict(
    failures: Sequence[str], *, test_mode: bool
) -> tuple[str, list[str]]:
    """Make explicit test-mode runs ineligible for an accepted claim."""

    final_failures = list(failures)
    if test_mode:
        final_failures.append(
            "test-mode run is non-evidence; accepted performance claims are disabled"
        )
    return ("accepted" if not final_failures else "rejected", final_failures)


def _summary_text(summary: Mapping[str, Any]) -> str:
    lines = [
        f"change: {CHANGE_ID}",
        f"verdict: {summary.get('verdict')}",
        f"samples: {summary.get('samples_per_group')}",
        f"groups: {summary.get('complete_groups')}/{summary.get('expected_groups')}",
    ]
    gates = summary.get("gates")
    if isinstance(gates, dict):
        lines.append(f"gates_passed: {gates.get('passed')}")
    failures = summary.get("failures")
    if isinstance(failures, list):
        lines.append(f"failures: {len(failures)}")
        lines.extend(f"- {failure}" for failure in failures)
    return "\n".join(lines) + "\n"


ARTIFACT_NAMES = (
    "protocol.json",
    "identity.json",
    "oracle.json",
    "samples.jsonl",
    "summary.json",
    "summary.txt",
    "comparisons.json",
    "comparisons.tsv",
    "failures.log",
)


def _manifest(output_dir: Path, verdict: str) -> dict[str, Any]:
    artifacts = []
    for name in ARTIFACT_NAMES:
        path = output_dir / name
        if not path.is_file():
            raise DriverError(f"missing artifact before manifest: {path}")
        digest = hashlib.sha256()
        size = 0
        with path.open("rb") as stream:
            while True:
                chunk = stream.read(1024 * 1024)
                if not chunk:
                    break
                digest.update(chunk)
                size += len(chunk)
        artifacts.append({"path": name, "bytes": size, "sha256": digest.hexdigest()})
    return {
        "schema_version": DRIVER_SCHEMA_VERSION,
        "manifest_kind": "litchi-xls-source-attribution-abba",
        "change": CHANGE_ID,
        "verdict": verdict,
        "artifacts": artifacts,
    }


def _parse_args(argv: Sequence[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--control-binary", type=Path, required=True)
    parser.add_argument("--candidate-binary", type=Path, required=True)
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--tmpdir", type=Path)
    parser.add_argument("--cwd", type=Path)
    parser.add_argument("--cpu", type=int, default=DEFAULT_CPU, help="CPU number; -1 disables affinity")
    parser.add_argument("--memory-limit-bytes", type=int, default=DEFAULT_MEMORY_LIMIT_BYTES)
    parser.add_argument("--timeout-seconds", type=float, default=DEFAULT_TIMEOUT_SECONDS)
    parser.add_argument("--warmups", type=int, default=DEFAULT_WARMUPS)
    parser.add_argument("--samples", type=int, default=DEFAULT_SAMPLES)
    parser.add_argument("--control-revision", required=True)
    parser.add_argument("--candidate-revision", required=True)
    parser.add_argument(
        "--test-mode",
        action="store_true",
        help=f"permit small explicit runs (at most {MAX_TEST_SAMPLES} samples and {MAX_TEST_WARMUPS} warmups)",
    )
    args = parser.parse_args(argv)
    if args.samples <= 0 or args.warmups <= 0:
        parser.error("--samples and --warmups must be greater than zero")
    if not args.test_mode and (args.samples != DEFAULT_SAMPLES or args.warmups != DEFAULT_WARMUPS):
        parser.error("non-default sample counts require --test-mode")
    if args.test_mode and (args.samples > MAX_TEST_SAMPLES or args.warmups > MAX_TEST_WARMUPS):
        parser.error("--test-mode sample/warmup counts are too large")
    if args.memory_limit_bytes <= 0:
        parser.error("--memory-limit-bytes must be positive")
    if args.memory_limit_bytes > HARD_MEMORY_LIMIT_BYTES:
        parser.error(f"--memory-limit-bytes cannot exceed {HARD_MEMORY_LIMIT_BYTES}")
    if args.timeout_seconds <= 0 or not math.isfinite(args.timeout_seconds):
        parser.error("--timeout-seconds must be finite and positive")
    if args.cpu < -1:
        parser.error("--cpu must be -1 or a non-negative CPU number")
    return args


def run_batch(args: argparse.Namespace) -> dict[str, Any]:
    """Collect, validate, summarize, and publish one evidence directory."""

    input_identity = _hash_file(args.input, "input corpus")
    control_identity = _hash_file(args.control_binary, "control binary")
    candidate_identity = _hash_file(args.candidate_binary, "candidate binary")
    require_enforceable_limits(args.memory_limit_bytes)
    input_path = Path(input_identity["path"])
    control_path = Path(control_identity["path"])
    candidate_path = Path(candidate_identity["path"])
    output_dir = args.output_dir.expanduser().resolve()
    if output_dir.exists():
        raise DriverError(f"refusing to use existing output directory {output_dir}")
    try:
        output_dir.parent.mkdir(parents=True, exist_ok=True)
        output_dir.mkdir()
    except OSError as error:
        raise DriverError(f"cannot create output directory {output_dir}: {error}") from error
    tmp_requested = args.tmpdir or output_dir / ".tmp"
    tmpdir, owns_tmpdir = ensure_tmpdir(tmp_requested)
    cwd = (args.cwd or Path.cwd()).expanduser().resolve(strict=True)
    cpu = None if args.cpu == -1 else args.cpu
    protocol = build_protocol(
        corpus=input_identity,
        control_binary=control_identity,
        candidate_binary=candidate_identity,
        tmpdir=str(tmpdir),
        cpu=cpu,
        memory_limit_bytes=args.memory_limit_bytes,
        timeout_seconds=args.timeout_seconds,
        warmups=args.warmups,
        samples=args.samples,
        test_mode=args.test_mode,
        cwd=str(cwd),
    )
    protocol["expected_revisions"] = {
        "control": args.control_revision,
        "candidate": args.candidate_revision,
    }
    _write_json(output_dir / "protocol.json", protocol)
    observed = {
        "revisions": {},
        "binary_revisions": {},
        "tool_revisions": {},
        "binary_paths": {},
        "input_path": None,
        "validated_samples": 0,
    }
    samples_path = output_dir / "samples.jsonl"
    samples_partial_path = output_dir / "samples.jsonl.partial"
    failures: list[str] = []
    groups = {
        (leg, selector.key): GroupAccumulator(selector, leg, args.samples)
        for selector in SELECTORS
        for leg in LEGS
    }
    semantic_oracles: dict[str, Mapping[str, Any]] = {}
    expected_revisions = {
        "A1": args.control_revision,
        "A2": args.control_revision,
        "B1": args.candidate_revision,
        "B2": args.candidate_revision,
    }
    sequence = 0
    samples_bytes = 0
    with open_samples_partial(samples_partial_path) as samples_stream:
        try:
            for selector in SELECTORS:
                for leg in LEGS:
                    binary = control_path if leg.startswith("A") else candidate_path
                    binary_identity = control_identity if leg.startswith("A") else candidate_identity
                    for sample_index in range(1, args.samples + 1):
                        report = invoke_child(
                            binary=binary,
                            input_path=input_path,
                            revision=expected_revisions[leg],
                            selector=selector,
                            warmups=args.warmups,
                            tmpdir=tmpdir,
                            cwd=cwd,
                            cpu=cpu,
                            memory_limit_bytes=args.memory_limit_bytes,
                            timeout_seconds=args.timeout_seconds,
                        )
                        (
                            record,
                            _semantic_oracles,
                            revision,
                            tool_revision,
                            binary_path,
                            input_report_path,
                        ) = validate_child_report(
                            report,
                            selector=selector,
                            leg=leg,
                            input_identity=input_identity,
                            binary_identity=binary_identity,
                            warmups=args.warmups,
                            semantic_oracles=semantic_oracles,
                            seen_revisions=observed["revisions"],
                            seen_binary_revisions=observed["binary_revisions"],
                            expected_revision=expected_revisions[leg],
                        )
                        groups[(leg, selector.key)].add(record["elapsed_ns"], record["metrics"])
                        sequence += 1
                        observed["revisions"].setdefault(leg, revision)
                        binary_role = "control" if leg.startswith("A") else "candidate"
                        observed["binary_revisions"].setdefault(binary_role, revision)
                        observed["tool_revisions"].setdefault(leg, tool_revision)
                        previous_binary_path = observed["binary_paths"].get(leg)
                        if previous_binary_path is not None and previous_binary_path != binary_path:
                            raise DriverError(f"{leg} reported inconsistent binary path")
                        observed["binary_paths"].setdefault(leg, binary_path)
                        previous_input_path = observed["input_path"]
                        if previous_input_path is not None and previous_input_path != input_report_path:
                            raise DriverError("child reports used inconsistent input paths")
                        observed["input_path"] = input_report_path
                        observed["validated_samples"] += 1
                        normalized = {
                            "schema_version": DRIVER_SCHEMA_VERSION,
                            "sequence": sequence,
                            "leg": leg,
                            "selector": selector.key,
                            "mode": selector.mode,
                            "operation": selector.operation,
                            "sample_index": sample_index,
                            "revision": revision,
                            "elapsed_ns": record["elapsed_ns"],
                            "metrics": record["metrics"],
                            "observation": record["observation"],
                            "source_version_stable": True,
                        }
                        encoded_row = encode_sample_row(normalized, samples_bytes)
                        written = samples_stream.write(encoded_row)
                        if written != len(encoded_row):
                            raise DriverError(
                                "partial samples writer reported a short binary write"
                            )
                        samples_stream.flush()
                        samples_bytes += written
        except (DriverError, OSError) as error:
            failures.append(str(error))

    try:
        if samples_partial_path.stat().st_size != samples_bytes:
            failures.append("partial samples size disagrees with bounded stream accounting")
        if samples_bytes > MAX_SAMPLES_ARTIFACT_BYTES:
            failures.append("partial samples artifact exceeds its declared size bound")
    except OSError as error:
        failures.append(f"cannot inspect partial samples artifact: {error}")

    for label, path, expected in (
        ("input corpus", input_path, input_identity),
        ("control binary", control_path, control_identity),
        ("candidate binary", candidate_path, candidate_identity),
    ):
        try:
            actual = _hash_file(path, f"final {label}")
            if actual != expected:
                failures.append(
                    f"final {label} identity changed: expected {expected!r}, got {actual!r}"
                )
        except DriverError as error:
            failures.append(str(error))

    complete_groups = sum(group.complete for group in groups.values())
    comparisons: list[dict[str, Any]] = []
    gate_summary: dict[str, Any] = {"passed": False, "checks": []}
    if not failures:
        try:
            comparisons, gate_summary, gate_failures = evaluate_gates(
                groups, samples=args.samples
            )
            failures.extend(gate_failures)
        except (
            DriverError,
            RecursionError,
            TypeError,
            KeyError,
            AttributeError,
            ValueError,
            OverflowError,
            UnicodeError,
        ) as error:
            failures.append(f"collection/gate computation failed: {error}")
    try:
        group_summaries = {
            f"{leg}:{selector.key}": groups[(leg, selector.key)].summary()
            for selector in SELECTORS
            for leg in LEGS
        }
    except (
        DriverError,
        RecursionError,
        TypeError,
        KeyError,
        AttributeError,
        ValueError,
        OverflowError,
        UnicodeError,
    ) as error:
        failures.append(f"collection summary computation failed: {error}")
        group_summaries = {}
    verdict, failures = resolve_verdict(failures, test_mode=args.test_mode)
    selectors_by_key = {selector.key: selector for selector in SELECTORS}
    ordered_cell_keys = sorted(selectors_by_key)
    source_semantic_oracles = {
        key: semantic_oracles[key]
        for key in ordered_cell_keys
        if key in semantic_oracles
    }
    observations = {
        key: expected_observation(semantic_oracles[key], selectors_by_key[key])
        for key in ordered_cell_keys
        if key in semantic_oracles
    }
    missing_cell_keys = [key for key in ordered_cell_keys if key not in semantic_oracles]
    if missing_cell_keys:
        oracle_artifact = {
            "schema_version": DRIVER_SCHEMA_VERSION,
            "status": "not-established",
            "reason": (
                failures[0]
                if failures
                else "missing semantic oracles: " + ", ".join(missing_cell_keys)
            ),
            "source_semantic_oracles": source_semantic_oracles,
            "observations": observations,
        }
    else:
        oracle_artifact = {
            "schema_version": DRIVER_SCHEMA_VERSION,
            "status": "frozen",
            "source_semantic_oracles": source_semantic_oracles,
            "observations": observations,
        }
    _write_json(output_dir / "oracle.json", oracle_artifact)
    identity_artifact = {
        "schema_version": DRIVER_SCHEMA_VERSION,
        "status": "validated" if not failures else "rejected",
        "corpus": input_identity,
        "binaries": {"control": control_identity, "candidate": candidate_identity},
        "observed": {
            "revisions": dict(sorted(observed["revisions"].items())),
            "binary_revisions": dict(sorted(observed["binary_revisions"].items())),
            "tool_revisions": dict(sorted(observed["tool_revisions"].items())),
            "binary_paths": {
                leg: path for leg, path in sorted(observed["binary_paths"].items())
            },
            "input_path": observed["input_path"],
            "validated_samples": observed["validated_samples"],
        },
    }
    _write_json(output_dir / "identity.json", identity_artifact)
    summary = {
        "schema_version": DRIVER_SCHEMA_VERSION,
        "change": CHANGE_ID,
        "verdict": verdict,
        "samples_per_group": args.samples,
        "expected_groups": len(groups),
        "complete_groups": complete_groups,
        "evidence_eligible": not args.test_mode,
        "groups": group_summaries,
        "gates": gate_summary,
        "failures": failures,
    }
    _write_json(output_dir / "summary.json", summary)
    _write_bytes(output_dir / "summary.txt", _summary_text(summary).encode("utf-8"))
    _write_json(
        output_dir / "comparisons.json",
        {
            "schema_version": DRIVER_SCHEMA_VERSION,
            "status": "evaluated" if comparisons else "not-evaluated",
            "comparisons": comparisons,
            "gates": gate_summary,
        },
    )
    tsv_lines = [
        "kind\tselector\tleft_leg\tright_leg\tp50_pct\tmean_pct\tp95_pct\tp99_pct\tp50_delta_ns\tmean_delta_ns"
    ]
    for comparison in comparisons:
        elapsed = comparison["elapsed_ns"]
        tsv_lines.append(
            "\t".join(
                str(value)
                for value in (
                    comparison["kind"],
                    comparison["selector"],
                    comparison["left_leg"],
                    comparison["right_leg"],
                    elapsed["p50"]["improvement_percent"],
                    elapsed["mean"]["improvement_percent"],
                    elapsed["p95"]["improvement_percent"],
                    elapsed["p99"]["improvement_percent"],
                    elapsed["p50"]["delta"],
                    elapsed["mean"]["delta"],
                )
            )
        )
    _write_bytes(output_dir / "comparisons.tsv", ("\n".join(tsv_lines) + "\n").encode("utf-8"))
    _write_bytes(output_dir / "failures.log", ("\n".join(failures) + ("\n" if failures else "")).encode("utf-8"))
    publish_samples(samples_partial_path, samples_path)
    _write_json(output_dir / "manifest.json", _manifest(output_dir, verdict))
    if owns_tmpdir:
        try:
            tmpdir.rmdir()
        except OSError:
            pass
    return {"output_dir": str(output_dir), "verdict": verdict, "failures": failures}


def main(argv: Sequence[str] | None = None) -> int:
    try:
        args = _parse_args(argv)
        result = run_batch(args)
    except (DriverError, OSError) as error:
        print(f"xls_source_attribution_abba: {error}", file=sys.stderr)
        return 2
    print(json.dumps(result, sort_keys=True))
    return 0 if result["verdict"] == "accepted" else 1


if __name__ == "__main__":  # pragma: no cover - exercised by the CLI
    raise SystemExit(main())
