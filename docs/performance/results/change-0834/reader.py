#!/usr/bin/env python3
"""Offline admission reader for the 0834 filesystem qualification.

This module only reads JSON and source/build receipts.  It never starts a
process, probes a filesystem, builds a binary, compares timings, or treats a
qualification row as a performance result.  The driver may use
``validate_report`` directly, or write the small immutable manifest consumed
by ``main``::

    python3 -B reader.py qualification-v3.json

The manifest is intentionally tolerant about receipt field names (``report``
and ``report_path`` are both accepted) so the capture driver can keep its
custody vocabulary.  The report itself remains strict: a missing field or an
unrecognised value fails closed.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import re
from pathlib import Path
from typing import Any, Iterable


HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
BASE = "41bb4c9670935496a5ca421d86bc76f7f6485971"
ACTIVE_STAGE = "repaired-v3"
ACTIVE_SUFFIX = "v3"
QUALITY_GATES = ("fmt", "check", "test", "clippy", "doc", "boundaries")
EXPECTED_OPC_FAILURE = 'Error: ProofError("cold ZIP proof does not accept data-descriptor member framing")'
SCHEMA = "litchi.0834.filesystem-qualification.v1"
REPORT_SCHEMA_VERSION = 1

CASES = (
    "opc_file_eager_open",
    "opc_file_source_open",
    "opc_file_eager_one_part_atomic_save",
    "opc_file_source_one_part_atomic_save",
    "pptx_file_eager_open_selected_slide_lifecycle",
    "pptx_file_source_open_selected_slide_lifecycle",
)
STATES = ("warm", "cold-verified")
EXPECTED_PAIRS = tuple((case, state) for state in STATES for case in CASES)
SIX_CASES = frozenset(CASES)
SHA256_RE = r"[0-9a-f]{64}"
EMPTY_SHA256 = hashlib.sha256(b"").hexdigest()
U64_MAX = (1 << 64) - 1

CLAIM_SCOPE = (
    "external fincore page-cache residency/dirty/writeback proof plus "
    "positive process read_bytes; no physical-media claim"
)
FINCORE_COMMAND = (
    "fincore --json --bytes --output FILE,SIZE,RES,DIRTY,WRITEBACK --"
)
FINCORE_METHOD = "external_fincore_json_columns"
FINCORE_FALLBACK = "none"
FINCORE_TOOL = "fincore"
FINCORE_SHA256 = "5586d15dedd490ce676adb19f2848ce77ecd46ffa5ac74d768396bcac17e021f"
COLD_ADVICE = "posix_fadvise_dontneed_accepted"

# These dimensions are copied from the deterministic builders.  The source
# hashes and the ordinary OPC output hash are read from the frozen repaired
# source below; a captured report never supplies its own oracle.
OPC_CORPUS = {
    "name": "few-large-incompressible",
    "generator": "litchi-opc-synthetic-v2",
    "package_format": "OPC/ZIP",
    "shape": "few-large",
    "payload_kind": "incompressible",
    "compression": "deflate",
    "entry_count": 4,
    "archive_member_count": 6,
    "entry_bytes": 4 * 1024 * 1024,
    "uncompressed_payload_bytes": 16 * 1024 * 1024,
    "archive_bytes": 16_783_632,
    "archive_sha256": None,
    "target_entry": "benchmark/parts/00002.bin",
    "target_payload_bytes": 4 * 1024 * 1024,
    "target_payload_sha256": "3dbf6225021a99c1da8750a738bde21f57591c0be1a60aa510966c47ee25b098",
    "xlsx": None,
}
PPTX_CORPUS = {
    "name": "pptx-source-backed-media",
    "generator": "litchi-pptx-source-edit-media-v1",
    "package_format": "PPTX/OPC/ZIP",
    "shape": "media-rich",
    "payload_kind": "deterministic-incompressible-media",
    "compression": "deflate",
    "entry_count": 229,
    "archive_member_count": 445,
    "entry_bytes": 2 * 1024 * 1024,
    "uncompressed_payload_bytes": 17_568_429,
    "archive_bytes": None,
    "archive_sha256": None,
    "target_entry": "slide:100/shape:0",
    "target_payload_bytes": 52,
    "target_payload_sha256": "26c9e1dffc347568407eb56d54c87d5fddb8cdb895a867fffb3cd9f302be18e6",
    "xlsx": None,
}
PPTX_TAIL_BYTES = 65_536
PPTX_SELECTED_SEMANTIC_SHA256 = "f5f7db181150c00a4323a48c142721ead73aca3ad7c3b3594e8b1a18a686b257"
_SOURCE_STAGE = ACTIVE_STAGE
_SOURCE_ORACLE_CACHE: tuple[str, dict[str, Any], dict[str, Any], str] | None = None

BUCKETS = (
    "bytes_0",
    "bytes_1_to_512",
    "bytes_513_to_4096",
    "bytes_4097_to_16384",
    "bytes_16385_to_65536",
    "bytes_over_65536",
)
NONNEGATIVE_COLD_STATUSES = {
    "eligible",
    "ineligible_non_linux",
    "ineligible_linux_non64_bit",
    "ineligible_filesystem_unknown",
    "ineligible_filesystem_unsupported",
    "ineligible_fincore_unavailable",
    "ineligible_fincore_failed",
    "ineligible_fincore_invalid_json",
    "ineligible_fincore_multiple_records",
    "ineligible_fincore_path_mismatch",
    "ineligible_fincore_size_mismatch",
    "ineligible_fincore_metadata_unavailable",
    "ineligible_fincore_unrecognized_fallback",
    "ineligible_source_not_regular",
    "ineligible_source_empty",
    "ineligible_source_read_write_unavailable",
    "ineligible_source_hash_failed",
    "ineligible_source_page_size_unavailable",
    "ineligible_source_not_page_aligned",
    "ineligible_source_fsync_failed",
    "ineligible_source_advice_failed",
    "ineligible_source_resident",
    "ineligible_source_dirty",
    "ineligible_source_writeback",
    "ineligible_proc_io_unavailable",
    "ineligible_read_bytes_backwards",
    "ineligible_read_bytes_zero",
    "ineligible_post_fincore",
    "ineligible_prepared_query_control",
    "ineligible_source_alignment_unavailable",
    "ineligible_source_write_failed",
}


class QualificationError(AssertionError):
    """An admission invariant failed."""


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


def _hash(value: Any, path: str) -> str:
    require(isinstance(value, str) and len(value) == 64 and value == value.lower(),
            f"{path}: expected lowercase SHA-256")
    require(all(char in "0123456789abcdef" for char in value), f"{path}: malformed SHA-256")
    return value


def _uint(value: Any, path: str, *, positive: bool = False) -> int:
    require(type(value) is int and value <= U64_MAX and value >= (1 if positive else 0),
            f"{path}: expected {'positive' if positive else 'non-negative'} integer")
    return value


def _same(actual: Any, expected: Any, path: str) -> None:
    require(actual == expected, f"{path}: oracle differs (expected {expected!r}, got {actual!r})")


def _source_text(relative: str) -> str:
    """Read a source file without executing or invoking repository tooling."""

    source_stage = _SOURCE_STAGE
    candidates = (
        HERE / "sources" / source_stage / relative,
        ROOT / relative,
    )
    for candidate in candidates:
        if candidate.is_file() and not candidate.is_symlink():
            try:
                return candidate.read_text(encoding="utf-8")
            except (OSError, UnicodeError) as error:
                raise QualificationError(f"cannot read frozen source {candidate}: {error}") from error
    raise QualificationError(f"missing frozen source oracle: {relative}")


def _const_string(text: str, name: str) -> str:
    match = re.search(
        rf"\bconst\s+{re.escape(name)}\s*:\s*[^=]+?=\s*\"([0-9a-zA-Z_./:-]+)\"",
        text,
    )
    require(match is not None, f"source oracle is missing const {name}")
    return match.group(1)


def _const_usize(text: str, name: str) -> int:
    match = re.search(
        rf"\bconst\s+{re.escape(name)}\s*:\s*[^=]+?=\s*([0-9][0-9_]*)\s*;",
        text,
    )
    require(match is not None, f"source oracle is missing const {name}")
    return int(match.group(1).replace("_", ""))


def _source_oracles() -> tuple[dict[str, Any], dict[str, Any], str]:
    global _SOURCE_ORACLE_CACHE
    if _SOURCE_ORACLE_CACHE is not None and _SOURCE_ORACLE_CACHE[0] == _SOURCE_STAGE:
        return _SOURCE_ORACLE_CACHE[1:]
    filesystem = _source_text("tools/perf-baseline/src/filesystem.rs")
    library = _source_text("tools/perf-baseline/src/lib.rs")
    opc_source = _const_string(filesystem, "OPC_FILE_SOURCE_SHA256")
    opc_output = _const_string(filesystem, "OPC_FILE_EXPECTED_OUTPUT_SHA256")
    pptx_source = _const_string(filesystem, "PPTX_FILE_SOURCE_SHA256")
    pptx_bytes = _const_usize(filesystem, "PPTX_FILE_SOURCE_ARCHIVE_BYTES")
    pptx_generator = _const_string(library, "PPTX_SOURCE_EDIT_CORPUS_GENERATOR")
    # The source builder computes the target text digest rather than retaining
    # it as a literal.  This value is pinned by the current source generator's
    # deterministic manifest and cross-checked in every report below.
    opc = dict(OPC_CORPUS, archive_sha256=opc_source)
    pptx = dict(PPTX_CORPUS, archive_bytes=pptx_bytes, archive_sha256=pptx_source,
                generator=pptx_generator)
    _hash(opc_source, "source.OPC_FILE_SOURCE_SHA256")
    _hash(opc_output, "source.OPC_FILE_EXPECTED_OUTPUT_SHA256")
    _hash(pptx_source, "source.PPTX_FILE_SOURCE_SHA256")
    _SOURCE_ORACLE_CACHE = (_SOURCE_STAGE, opc, pptx, opc_output)
    return _SOURCE_ORACLE_CACHE[1:]


def _corpus(case: str) -> dict[str, Any]:
    opc, pptx, _ = _source_oracles()
    return opc if case.startswith("opc_") else pptx


def _manifest(value: Any, path: str, case: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{path}: corpus manifest is missing")
    expected = _corpus(case)
    # The full manifest is compared.  This catches generator, archive shape,
    # member count, target payload and source identity drift independently of
    # the binary's own correctness checks.
    if expected.get("target_payload_sha256") is None:
        _hash(value.get("target_payload_sha256"), f"{path}.target_payload_sha256")
        expected = dict(expected, target_payload_sha256=value["target_payload_sha256"])
    _same(value, expected, path)
    return value


def _path_from(value: Any, base: Path, label: str) -> Path:
    require(isinstance(value, str) and value, f"{label}: path is missing")
    candidate = Path(value)
    if not candidate.is_absolute():
        candidate = base / candidate
    candidate = candidate.resolve()
    require(candidate.is_file() and not candidate.is_symlink(), f"{label}: file is missing")
    return candidate


def _descriptor(value: Any, base: Path, label: str, *, allow_missing: bool = False) -> Path:
    """Verify a path/length/hash descriptor without requiring a build binary."""

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
    """Find an exact path/length/hash descriptor in a retained cleanup receipt."""

    if isinstance(value, dict):
        if all(value.get(field) == expected.get(field)
               for field in ("path", "bytes", "sha256")):
            return True
        return any(_contains_descriptor(child, expected) for child in value.values())
    if isinstance(value, list):
        return any(_contains_descriptor(child, expected) for child in value)
    return False


def _binary_identity(report: dict[str, Any], expected: dict[str, Any] | None, path: str) -> dict[str, Any]:
    identity = report.get("binary_identity")
    require(isinstance(identity, dict), f"{path}: binary identity is missing")
    _hash(identity.get("binary_sha256"), f"{path}.binary_identity.binary_sha256")
    _uint(identity.get("binary_bytes"), f"{path}.binary_identity.binary_bytes", positive=True)
    _same(identity.get("executable"), True, f"{path}.binary_identity.executable")
    _same(identity.get("profile"), "release", f"{path}.binary_identity.profile")
    if expected is not None:
        require(isinstance(expected.get("path"), str) and expected["path"],
                f"{path}.binary_descriptor.path: missing")
        expected_bytes = expected.get("bytes", expected.get("binary_bytes"))
        expected_sha = expected.get("sha256", expected.get("binary_sha256"))
        _uint(expected_bytes, f"{path}.binary_descriptor.bytes", positive=True)
        _hash(expected_sha, f"{path}.binary_descriptor.sha256")
        aliases = {
            "binary_sha256": ("binary_sha256", "sha256"),
            "binary_bytes": ("binary_bytes", "bytes"),
            "profile": ("profile",),
            "executable": ("executable",),
        }
        for field, names in aliases.items():
            expected_value = next((expected[name] for name in names if name in expected), None)
            if expected_value is not None:
                _same(identity.get(field), expected_value, f"{path}.binary_identity.{field}")
    return identity


def _tool(report: dict[str, Any], case: str, path: str) -> None:
    tool = report.get("tool")
    require(isinstance(tool, dict), f"{path}: tool identity is missing")
    _same(tool.get("name"), "litchi-perf-baseline", f"{path}.tool.name")
    _same(tool.get("version"), "0.1.0", f"{path}.tool.version")
    _same(tool.get("profile"), "release", f"{path}.tool.profile")
    expected_binary = "litchi-perf-baseline"
    _same(tool.get("binary"), expected_binary, f"{path}.tool.binary")
    _same(tool.get("instrumentation"), "none", f"{path}.tool.instrumentation")
    _same(tool.get("target_os"), "linux", f"{path}.tool.target_os")
    _same(tool.get("target_arch"), "x86_64", f"{path}.tool.target_arch")


def _configuration(report: dict[str, Any], case: str, states: tuple[str, ...],
                   samples: int, warmup: int, path: str) -> None:
    configuration = report.get("configuration")
    require(isinstance(configuration, dict), f"{path}: configuration is missing")
    _same(configuration.get("samples_per_case"), samples, f"{path}.configuration.samples_per_case")
    _same(configuration.get("warmup_iterations_per_case"), warmup,
          f"{path}.configuration.warmup_iterations_per_case")
    _same(configuration.get("filesystem_cache_states"), list(states),
          f"{path}.configuration.filesystem_cache_states")
    _same(configuration.get("filesystem_fresh_child_per_sample"), True,
          f"{path}.configuration.filesystem_fresh_child_per_sample")
    _same(configuration.get("filesystem_process_isolated"), True,
          f"{path}.configuration.filesystem_process_isolated")
    _same(configuration.get("filesystem_root_selected"), True,
          f"{path}.configuration.filesystem_root_selected")
    _same(configuration.get("cases"), [case], f"{path}.configuration.cases")


def _environment(report: dict[str, Any], expected: dict[str, Any] | None, path: str) -> dict[str, Any]:
    environment = report.get("environment")
    require(isinstance(environment, dict), f"{path}: environment is missing")
    _same(environment.get("git_revision"), BASE, f"{path}.environment.git_revision")
    require(type(environment.get("git_worktree_dirty")) is bool,
            f"{path}.environment.git_worktree_dirty: expected boolean")
    _same(environment.get("os"), "linux", f"{path}.environment.os")
    _same(environment.get("page_size_bytes"), 4096, f"{path}.environment.page_size_bytes")
    if expected is not None:
        _same(environment, expected, f"{path}.environment")
    return environment


def _stats(value: Any, path: str, expected_samples: int = 1) -> list[int]:
    require(isinstance(value, dict), f"{path}: elapsed statistics are missing")
    _same(value.get("unit"), "ns", f"{path}.unit")
    raw = value.get("samples")
    require(isinstance(raw, list) and len(raw) == expected_samples,
            f"{path}.samples: expected {expected_samples} values")
    values = [_uint(item, f"{path}.samples[{index}]", positive=True) for index, item in enumerate(raw)]
    order = value.get("sample_order")
    require(isinstance(order, list) and len(order) == expected_samples,
            f"{path}.sample_order: missing or wrong length")
    order_int = [_uint(item, f"{path}.sample_order[{index}]") for index, item in enumerate(order)]
    require(sorted(order_int) == list(range(expected_samples)), f"{path}.sample_order: not a permutation")
    ordered_pairs = sorted(enumerate(values), key=lambda item: (item[1], item[0]))
    ordered = [item[1] for item in ordered_pairs]
    expected_order = [item[0] for item in ordered_pairs]
    _same(order_int, expected_order, f"{path}.sample_order")
    _same(values, ordered, f"{path}.samples: values are not sorted with sample_order")
    nearest = lambda rank: ordered[min(expected_samples - 1,
                                       (rank * expected_samples + 99) // 100 - 1)]
    midpoint = lambda left, right: left // 2 + right // 2 + (left % 2 + right % 2) // 2
    _same(value.get("min"), ordered[0], f"{path}.min")
    _same(value.get("p50"), midpoint(ordered[(expected_samples - 1) // 2],
                                      ordered[expected_samples // 2]), f"{path}.p50")
    _same(value.get("p95"), nearest(95), f"{path}.p95")
    _same(value.get("p99"), nearest(99), f"{path}.p99")
    _same(value.get("max"), ordered[-1], f"{path}.max")
    mean = value.get("mean")
    mean_expected = sum(ordered) / expected_samples
    require(type(mean) in (int, float) and math.isfinite(float(mean)) and
            math.isclose(float(mean), mean_expected, rel_tol=0.0, abs_tol=1e-6),
            f"{path}.mean: does not match sample vector")
    running_mean = 0.0
    squared = 0.0
    for index, item in enumerate(ordered):
        count = index + 1
        delta = item - running_mean
        next_mean = running_mean + delta / count
        squared += delta * (item - next_mean)
        running_mean = next_mean
    deviation = math.sqrt(squared / (expected_samples - 1)) if expected_samples > 1 else 0.0
    _number_close(value.get("standard_deviation"), deviation, f"{path}.standard_deviation")
    interval = value.get("confidence_interval_95")
    require(isinstance(interval, dict), f"{path}.confidence_interval_95: missing")
    _same(interval.get("method"), "two-sided Student's t interval for the mean",
          f"{path}.confidence_interval_95.method")
    critical = _student_t_critical_95(expected_samples - 1)
    margin = critical * deviation / math.sqrt(expected_samples) if expected_samples > 1 else 0.0
    _number_close(interval.get("lower"), max(0.0, mean_expected - margin),
                  f"{path}.confidence_interval_95.lower")
    _number_close(interval.get("upper"), mean_expected + margin,
                  f"{path}.confidence_interval_95.upper")
    return values


def _number_close(actual: Any, expected: float, path: str) -> None:
    require(type(actual) in (int, float) and math.isfinite(float(actual)) and
            math.isclose(float(actual), expected, rel_tol=0.0, abs_tol=1e-6),
            f"{path}: differs from sample vector")


def _student_t_critical_95(degrees: int) -> float:
    values = (12.706, 4.303, 3.182, 2.776, 2.571, 2.447, 2.365, 2.306, 2.262, 2.228,
              2.201, 2.179, 2.160, 2.145, 2.131, 2.120, 2.110, 2.101, 2.093, 2.086,
              2.080, 2.074, 2.069, 2.064, 2.060, 2.056, 2.052, 2.048, 2.045, 2.042)
    if degrees <= 0:
        return 0.0
    if degrees <= len(values):
        return values[degrees - 1]
    z = 1.959963984540054
    d = float(degrees)
    z2, z3, z5, z7 = z * z, z * z * z, z * z * z * z * z, z * z * z * z * z * z * z
    return (z + (z3 + z) / (4 * d)
            + (5 * z5 + 16 * z3 + 3 * z) / (96 * d * d)
            + (3 * z7 + 19 * z5 + 17 * z3 - 15 * z) / (384 * d * d * d))


def _buckets(value: Any, path: str) -> dict[str, int]:
    require(isinstance(value, dict), f"{path}: request-size buckets are missing")
    result: dict[str, int] = {}
    for bucket in BUCKETS:
        result[bucket] = _uint(value.get(bucket), f"{path}.{bucket}")
    return result


def _logical_reads(sample: dict[str, Any], case: str, path: str) -> None:
    scope = sample.get("logical_read_counter_scope")
    if case == "opc_file_source_open" or case == "opc_file_source_one_part_atomic_save":
        _same(scope, "timed_read_at", f"{path}.logical_read_counter_scope")
        calls = _uint(sample.get("logical_read_calls"), f"{path}.logical_read_calls")
        requested = _uint(sample.get("logical_read_requested_bytes"),
                          f"{path}.logical_read_requested_bytes")
        returned = _uint(sample.get("logical_read_bytes"), f"{path}.logical_read_bytes")
        require(returned <= requested, f"{path}: returned logical bytes exceed requested bytes")
        sizes = sample.get("logical_read_request_sizes")
        require(isinstance(sizes, list), f"{path}.logical_read_request_sizes: missing")
        require(len(sizes) == calls, f"{path}: logical call count differs from size histogram")
        sizes_int = [_uint(size, f"{path}.logical_read_request_sizes[{index}]")
                     for index, size in enumerate(sizes)]
        _same(sum(sizes_int), requested, f"{path}: request-size sum differs from requested bytes")
        buckets = _buckets(sample.get("logical_read_request_size_buckets"),
                           f"{path}.logical_read_request_size_buckets")
        _same(sum(buckets.values()), calls, f"{path}: request buckets do not conserve calls")
        _same(buckets, _count_buckets(sizes_int), f"{path}: request buckets differ from sizes")
        largest_requested = max(sizes_int, default=0)
        _same(sample.get("logical_read_largest_requested_bytes"), largest_requested,
              f"{path}.logical_read_largest_requested_bytes")
        # The wrapper records only the largest returned value, not every return
        # length.  Its conservation invariant is therefore bounded by the
        # requested total and the per-call largest request.
        largest_returned = _uint(sample.get("logical_read_largest_returned_bytes"),
                                  f"{path}.logical_read_largest_returned_bytes")
        require(largest_returned <= largest_requested,
                f"{path}: largest returned range exceeds largest request")
        _uint(sample.get("max_concurrent_reads"), f"{path}.max_concurrent_reads")
        pattern = sample.get("logical_read_pattern")
        if pattern is not None:
            require(pattern in {"sequential", "random", "unknown"},
                    f"{path}.logical_read_pattern: unknown value")
    elif case.startswith("opc_file_eager_"):
        _same(scope, "not_applicable_eager_opc", f"{path}.logical_read_counter_scope")
        _zero_logical_reads(sample, path)
    elif case.startswith("pptx_file_eager_"):
        _same(scope, "not_applicable_eager_pptx", f"{path}.logical_read_counter_scope")
        _zero_logical_reads(sample, path)
    else:
        _same(scope, "untimed_source_replay_only", f"{path}.logical_read_counter_scope")
        _zero_logical_reads(sample, path)


def _count_buckets(sizes: Iterable[int]) -> dict[str, int]:
    result = dict.fromkeys(BUCKETS, 0)
    for size in sizes:
        if size == 0:
            bucket = "bytes_0"
        elif size <= 512:
            bucket = "bytes_1_to_512"
        elif size <= 4096:
            bucket = "bytes_513_to_4096"
        elif size <= 16384:
            bucket = "bytes_4097_to_16384"
        elif size <= 65536:
            bucket = "bytes_16385_to_65536"
        else:
            bucket = "bytes_over_65536"
        result[bucket] += 1
    return result


def _zero_logical_reads(sample: dict[str, Any], path: str) -> None:
    for field in (
        "logical_read_calls",
        "logical_read_requested_bytes",
        "logical_read_bytes",
        "logical_read_largest_requested_bytes",
        "logical_read_largest_returned_bytes",
        "max_concurrent_reads",
    ):
        _same(sample.get(field), 0, f"{path}.{field}")
    _same(sample.get("logical_read_request_sizes"), [], f"{path}.logical_read_request_sizes")
    _same(_buckets(sample.get("logical_read_request_size_buckets"),
                   f"{path}.logical_read_request_size_buckets"),
          dict.fromkeys(BUCKETS, 0), f"{path}: empty request buckets differ")
    pattern = sample.get("logical_read_pattern")
    require(pattern is None, f"{path}.logical_read_pattern: N/A route exposed a pattern")


def _process_metrics(sample: dict[str, Any], path: str) -> dict[str, Any]:
    metrics = sample.get("process_metrics")
    require(isinstance(metrics, dict), f"{path}: process metrics are missing")
    for field in (
        "rchar", "wchar", "read_bytes", "write_bytes", "cancelled_write_bytes",
        "syscr", "syscw", "minor_faults", "major_faults", "user_cpu_ticks",
        "system_cpu_ticks", "clock_ticks_per_second", "voluntary_context_switches",
        "nonvoluntary_context_switches", "rss_bytes", "peak_rss_bytes",
    ):
        _uint(metrics.get(field), f"{path}.process_metrics.{field}")
    require(metrics["clock_ticks_per_second"] > 0,
            f"{path}.process_metrics.clock_ticks_per_second: zero")
    return metrics


def _cold_proof(proof: Any, source: dict[str, Any], path: str) -> dict[str, Any]:
    require(isinstance(proof, dict), f"{path}: cold proof is missing")
    _same(proof.get("status"), "eligible", f"{path}.status")
    require(proof.get("status") in NONNEGATIVE_COLD_STATUSES, f"{path}.status: unknown status")
    for field in (
        "filesystem_magic", "page_size_bytes", "source_bytes", "source_pages",
        "aligned_source_bytes", "fincore_size_bytes", "resident_bytes", "dirty_bytes",
        "writeback_bytes", "fincore_stderr_bytes", "fincore_version_stderr_bytes",
    ):
        _uint(proof.get(field), f"{path}.{field}")
    _same(proof.get("fsync_completed"), True, f"{path}.fsync_completed")
    _same(proof.get("advice"), COLD_ADVICE, f"{path}.advice")
    _same(proof.get("source_bytes"), proof.get("aligned_source_bytes"), f"{path}: source sizes differ")
    _same(proof.get("aligned_source_bytes"), proof.get("fincore_size_bytes"),
          f"{path}: fincore size differs")
    require(proof["source_bytes"] > 0 and proof["page_size_bytes"] > 0,
            f"{path}: source/page size must be positive")
    require(proof["source_bytes"] % proof["page_size_bytes"] == 0,
            f"{path}: aligned source is not page aligned")
    _same(proof.get("page_size_bytes"), 4096, f"{path}.page_size_bytes")
    _same(proof.get("filesystem_magic"), 0xEF53, f"{path}.filesystem_magic")
    aligned_source_bytes = ((source["archive_bytes"] + 4095) // 4096) * 4096
    _same(proof.get("source_bytes"), aligned_source_bytes, f"{path}.source_bytes")
    _same(proof.get("source_pages"), proof["source_bytes"] // proof["page_size_bytes"],
          f"{path}.source_pages")
    _same(proof.get("resident_bytes"), 0, f"{path}.resident_bytes")
    _same(proof.get("dirty_bytes"), 0, f"{path}.dirty_bytes")
    _same(proof.get("writeback_bytes"), 0, f"{path}.writeback_bytes")
    _same(proof.get("fincore_tool"), FINCORE_TOOL, f"{path}.fincore_tool")
    _same(proof.get("fincore_sha256"), FINCORE_SHA256, f"{path}.fincore_sha256")
    _same(proof.get("fincore_method"), FINCORE_METHOD, f"{path}.fincore_method")
    _same(proof.get("fincore_fallback"), FINCORE_FALLBACK, f"{path}.fincore_fallback")
    for field in ("fincore_sha256", "fincore_stderr_sha256", "fincore_version_stderr_sha256",
                  "aligned_source_sha256"):
        _hash(proof.get(field), f"{path}.{field}")
    _same(proof.get("fincore_stderr_sha256"), EMPTY_SHA256, f"{path}.fincore_stderr_sha256")
    _same(proof.get("fincore_stderr_bytes"), 0, f"{path}.fincore_stderr_bytes")
    _same(proof.get("fincore_version_stderr_sha256"), EMPTY_SHA256,
          f"{path}.fincore_version_stderr_sha256")
    _same(proof.get("fincore_version_stderr_bytes"), 0,
          f"{path}.fincore_version_stderr_bytes")
    require(isinstance(proof.get("fincore_version"), str) and proof["fincore_version"],
            f"{path}.fincore_version: missing")
    _uint(proof.get("read_bytes_before"), f"{path}.read_bytes_before")
    _uint(proof.get("read_bytes_after"), f"{path}.read_bytes_after")
    delta = _uint(proof.get("read_bytes_delta"), f"{path}.read_bytes_delta", positive=True)
    _same(proof["read_bytes_after"] - proof["read_bytes_before"], delta,
          f"{path}: read_bytes delta is inconsistent")
    post = proof.get("fincore_post")
    require(isinstance(post, dict), f"{path}.fincore_post: post observation is missing")
    _same(post.get("status"), "eligible", f"{path}.fincore_post.status")
    for field in ("size_bytes", "resident_bytes", "dirty_bytes", "writeback_bytes",
                  "fincore_stderr_bytes", "fincore_version_stderr_bytes"):
        _uint(post.get(field), f"{path}.fincore_post.{field}")
    _same(post.get("size_bytes"), proof["fincore_size_bytes"], f"{path}: pre/post sizes differ")
    require(0 < post["resident_bytes"] <= post["size_bytes"],
            f"{path}.fincore_post.resident_bytes: expected post-operation residency")
    _same(post.get("dirty_bytes"), 0, f"{path}.fincore_post.dirty_bytes")
    _same(post.get("writeback_bytes"), 0, f"{path}.fincore_post.writeback_bytes")
    for field in ("fincore_tool", "fincore_sha256", "fincore_version", "fincore_method",
                  "fincore_fallback", "fincore_stderr_sha256", "fincore_version_stderr_sha256",
                  "fincore_stderr_bytes", "fincore_version_stderr_bytes"):
        _same(post.get(field), proof.get(field), f"{path}.fincore_post.{field}")
    return proof


def _ranges(value: Any, path: str, source_bytes: int) -> list[tuple[int, int]]:
    require(isinstance(value, list), f"{path}: range vector is missing")
    result: list[tuple[int, int]] = []
    for index, item in enumerate(value):
        require(isinstance(item, dict), f"{path}[{index}]: range is malformed")
        start = _uint(item.get("start"), f"{path}[{index}].start")
        end = _uint(item.get("end"), f"{path}[{index}].end")
        require(start < end <= source_bytes, f"{path}[{index}]: range is outside source")
        result.append((start, end))
    # Slide/media vectors retain presentation order.  Their byte offsets are
    # therefore not required to be monotone; check disjointness on a sorted
    # view while returning the source order for semantic partition checks.
    ordered = sorted(result)
    for left, right in zip(ordered, ordered[1:]):
        require(left[1] <= right[0], f"{path}: payload ranges overlap or are unordered")
    return result


def _one_range(value: Any, path: str, source_bytes: int) -> tuple[int, int]:
    require(isinstance(value, dict), f"{path}: selected range is missing")
    return _ranges([value], path, source_bytes)[0]


def _overlap(start: int, end: int, ranges: Iterable[tuple[int, int]]) -> int:
    total = 0
    for left, right in ranges:
        total += max(0, min(end, right) - max(start, left))
    return total


def _covered_bytes(ranges: list[tuple[int, int]],
                   observations: list[tuple[int, int]]) -> int:
    intersections: list[tuple[int, int]] = []
    for start, end in observations:
        for left, right in ranges:
            overlap_start = max(start, left)
            overlap_end = min(end, right)
            if overlap_start < overlap_end:
                intersections.append((overlap_start, overlap_end))
    intersections.sort()
    total = 0
    current: tuple[int, int] | None = None
    for start, end in intersections:
        if current is None:
            current = (start, end)
        elif start <= current[1]:
            current = (current[0], max(current[1], end))
        else:
            total += current[1] - current[0]
            current = (start, end)
    if current is not None:
        total += current[1] - current[0]
    return total


def _replay(replay: Any, case: str, source: dict[str, Any], state: str,
            proof: dict[str, Any] | None, path: str) -> None:
    require(isinstance(replay, dict), f"{path}: PPTX replay evidence is missing")
    if state == "cold-verified":
        require(proof is not None, f"{path}: cold replay has no cold proof")
        expected_bytes = proof["aligned_source_bytes"]
        expected_sha = proof["aligned_source_sha256"]
    else:
        expected_bytes = source["archive_bytes"]
        expected_sha = source["archive_sha256"]
    _same(replay.get("source_bytes"), expected_bytes, f"{path}.source_bytes")
    _same(replay.get("source_sha256"), expected_sha, f"{path}.source_sha256")
    _hash(replay.get("source_sha256"), f"{path}.source_sha256")
    _same(replay.get("implementation"), "litchi_pptx::SourceBackedPresentation",
          f"{path}.implementation")
    _same(replay.get("slide_count"), 200, f"{path}.slide_count")
    _same(replay.get("selected_position"), 100, f"{path}.selected_position")
    _same(replay.get("semantic_sha256"), PPTX_SELECTED_SEMANTIC_SHA256,
          f"{path}.semantic_sha256")
    _hash(replay.get("semantic_sha256"), f"{path}.semantic_sha256")
    # The semantic digest is an untimed source-replay oracle.  The timed child
    # does not return a slide digest, so this reader never promotes it to one.
    operation = replay.get("operation")
    _same(operation, "open_selected_slide_lifecycle", f"{path}.operation")
    classification = replay.get("classification")
    expected_classification = (
        "selected-slide-only:aligned-eocd-tail-metadata-probe;"
        "target-slide-no-unselected-or-media-overlap"
        if state == "cold-verified" else
        "selected-slide-only:target-slide-no-unselected-or-media-overlap"
    )
    _same(classification, expected_classification, f"{path}.classification")

    sizes = replay.get("read_return_sizes")
    require(isinstance(sizes, list), f"{path}.read_return_sizes: missing")
    values = [_uint(value, f"{path}.read_return_sizes[{index}]")
              for index, value in enumerate(sizes)]
    _same(values, sorted(values), f"{path}.read_return_sizes: source vector is not sorted")
    _same(replay.get("read_calls"), len(values), f"{path}.read_calls")
    _same(replay.get("read_bytes"), sum(values), f"{path}.read_bytes")
    fields = (
        "slide_payload_read_calls", "slide_payload_read_bytes", "slide_payload_covered_bytes",
        "slide_payload_ranges_fully_covered", "selected_slide_payload_read_calls",
        "selected_slide_payload_read_bytes", "selected_slide_payload_covered_bytes",
        "unselected_slide_payload_read_calls", "unselected_slide_payload_read_bytes",
        "unselected_slide_payload_covered_bytes", "media_payload_read_calls",
        "media_payload_read_bytes", "media_payload_covered_bytes",
    )
    for field in fields:
        _uint(replay.get(field), f"{path}.{field}")
    _same(replay.get("selected_slide_payload_fully_covered"), True,
          f"{path}.selected_slide_payload_fully_covered")
    _same(replay.get("selected_slide_payload_read_bytes"),
          replay.get("selected_slide_payload_covered_bytes"),
          f"{path}: selected payload read/coverage differs")
    _same(replay.get("selected_slide_payload_covered_bytes"), 522,
          f"{path}.selected_slide_payload_covered_bytes")

    probe = replay.get("aligned_eocd_tail_probe")
    if state == "warm":
        _same(replay.get("unselected_slide_payload_read_bytes"), 0,
              f"{path}.unselected_slide_payload_read_bytes")
        _same(replay.get("media_payload_read_bytes"), 0,
              f"{path}.media_payload_read_bytes")
        require(probe is None, f"{path}: unaligned warm replay exposed aligned proof")
        return
    require(isinstance(probe, dict), f"{path}: cold aligned replay proof is missing")
    alignment = probe.get("alignment_proof")
    require(isinstance(alignment, dict), f"{path}.aligned_eocd_tail_probe.alignment_proof: missing")
    for field in ("page_size_bytes", "eocd_offset", "base_bytes", "aligned_bytes",
                  "padding_bytes", "base_comment_bytes", "aligned_comment_bytes"):
        _uint(alignment.get(field), f"{path}.alignment_proof.{field}")
    _same(alignment.get("page_size_bytes"), 4096, f"{path}.alignment_proof.page_size_bytes")
    _same(alignment.get("base_bytes"), source["archive_bytes"], f"{path}.alignment_proof.base_bytes")
    _same(alignment.get("aligned_bytes"), expected_bytes, f"{path}.alignment_proof.aligned_bytes")
    _same(alignment.get("base_sha256"), source["archive_sha256"],
          f"{path}.alignment_proof.base_sha256")
    _same(alignment.get("aligned_sha256"), expected_sha,
          f"{path}.alignment_proof.aligned_sha256")
    _hash(alignment.get("base_sha256"), f"{path}.alignment_proof.base_sha256")
    _hash(alignment.get("aligned_sha256"), f"{path}.alignment_proof.aligned_sha256")
    padding = (source["archive_bytes"] + 4095) // 4096 * 4096 - source["archive_bytes"]
    _same(alignment.get("padding_bytes"), padding, f"{path}.alignment_proof.padding_bytes")
    _same(alignment.get("aligned_bytes"), alignment["base_bytes"] + alignment["padding_bytes"],
          f"{path}.alignment_proof: size/padding conservation")
    _same(alignment["eocd_offset"] + 22 + alignment["base_comment_bytes"],
          alignment["base_bytes"], f"{path}.alignment_proof: base EOCD boundary differs")
    _same(alignment["eocd_offset"] + 22 + alignment["aligned_comment_bytes"],
          alignment["aligned_bytes"], f"{path}.alignment_proof: aligned EOCD boundary differs")
    _same(alignment.get("aligned_comment_bytes"),
          alignment["base_comment_bytes"] + alignment["padding_bytes"],
          f"{path}.alignment_proof: EOCD comment transform differs")
    require(alignment["eocd_offset"] < alignment["base_bytes"],
            f"{path}.alignment_proof.eocd_offset: outside base archive")
    for field in ("outside_eocd_bytes_equal", "outside_comment_bytes_equal",
                  "zero_suffix", "suffix_zero", "eocd_offset_unchanged"):
        if field in alignment:
            _same(alignment[field], True, f"{path}.alignment_proof.{field}")

    _uint(probe.get("eocd_tail_probe_offset"), f"{path}.tail.offset")
    _same(probe.get("eocd_tail_probe_bytes"), PPTX_TAIL_BYTES, f"{path}.tail.bytes")
    _same(probe.get("eocd_tail_probe_read_count"), 1, f"{path}.tail.read_count")
    tail_offset = expected_bytes - PPTX_TAIL_BYTES
    _same(probe.get("eocd_tail_probe_offset"), tail_offset, f"{path}.tail.offset")
    _uint(probe.get("open_read_count"), f"{path}.open_read_count")
    raw = probe.get("raw_reads")
    require(isinstance(raw, list), f"{path}.raw_reads: missing")
    boundary = probe["open_read_count"]
    require(boundary <= len(raw), f"{path}.open_read_count: boundary exceeds raw vector")
    payload = probe.get("payload_ranges")
    require(isinstance(payload, dict), f"{path}.payload_ranges: missing")
    slides = _ranges(payload.get("slides"), f"{path}.payload_ranges.slides", expected_bytes)
    selected = _one_range(payload.get("selected_slide"),
                          f"{path}.payload_ranges.selected_slide", expected_bytes)
    unselected = _ranges(payload.get("unselected_slides"),
                          f"{path}.payload_ranges.unselected_slides", expected_bytes)
    media = _ranges(payload.get("media"), f"{path}.payload_ranges.media", expected_bytes)
    require(len(slides) == 200 and len(unselected) == 199 and len(media) == 8,
            f"{path}.payload_ranges: corpus counts differ")
    require(selected in slides, f"{path}.payload_ranges: selected range is not a slide range")
    _same(selected, slides[100],
          f"{path}.payload_ranges.selected_slide: selected range is not presentation slide 100")
    _same(unselected, [slide for index, slide in enumerate(slides) if index != 100],
          f"{path}.payload_ranges.unselected_slides: not the slide partition")

    parsed: list[tuple[int, int, int]] = []
    for index, item in enumerate(raw):
        require(isinstance(item, dict), f"{path}.raw_reads[{index}]: malformed")
        offset = _uint(item.get("offset"), f"{path}.raw_reads[{index}].offset")
        requested = _uint(item.get("requested_length"), f"{path}.raw_reads[{index}].requested_length")
        returned = _uint(item.get("returned_length"), f"{path}.raw_reads[{index}].returned_length")
        require(returned <= requested, f"{path}.raw_reads[{index}]: returned exceeds requested")
        require(offset <= U64_MAX - requested and offset <= U64_MAX - returned,
                f"{path}.raw_reads[{index}]: range overflows")
        require(returned == 0 or offset + returned <= expected_bytes,
                f"{path}.raw_reads[{index}]: returned range exceeds source")
        parsed.append((offset, requested, returned))
    _same(len(parsed), replay["read_calls"], f"{path}: raw read call conservation")
    _same(sorted(item[2] for item in parsed), values,
          f"{path}: raw/read-return vector differs")
    _same(sum(item[2] for item in parsed), replay["read_bytes"],
          f"{path}: raw returned-byte conservation")
    exact = [index for index, (offset, requested, returned) in enumerate(parsed)
             if offset == tail_offset and requested == PPTX_TAIL_BYTES and returned == PPTX_TAIL_BYTES]
    tail_starts = [index for index, (offset, _requested, _returned) in enumerate(parsed)
                   if offset == tail_offset]
    require(len(exact) == 1 and len(tail_starts) == 1 and exact[0] < boundary,
            f"{path}: exact aligned tail must occur once in open phase")

    totals = {"slides": 0, "selected": 0, "unselected": 0, "media": 0}
    call_totals = {"slides": 0, "selected": 0, "unselected": 0, "media": 0}
    open_totals = {"slides": 0, "selected": 0, "unselected": 0, "media": 0}
    returned_ranges: list[tuple[int, int]] = []
    query_returned_ranges: list[tuple[int, int]] = []
    for index, (offset, _requested, returned) in enumerate(parsed):
        end = offset + returned
        if returned:
            returned_ranges.append((offset, end))
            if index >= boundary:
                query_returned_ranges.append((offset, end))
        current = {
            "slides": _overlap(offset, end, slides),
            "selected": _overlap(offset, end, [selected]),
            "media": _overlap(offset, end, media),
        }
        current["unselected"] = current["slides"] - current["selected"]
        require(current["unselected"] >= 0, f"{path}.raw_reads[{index}]: overlap partition is invalid")
        for key in totals:
            totals[key] += current[key]
            if current[key]:
                call_totals[key] += 1
            if index < boundary:
                open_totals[key] += current[key]
    _same(totals["slides"], replay["slide_payload_read_bytes"], f"{path}: slide overlap conservation")
    _same(call_totals["slides"], replay["slide_payload_read_calls"],
          f"{path}: slide call conservation")
    _same(totals["selected"], replay["selected_slide_payload_read_bytes"], f"{path}: selected overlap conservation")
    _same(call_totals["selected"], replay["selected_slide_payload_read_calls"],
          f"{path}: selected call conservation")
    _same(totals["unselected"], replay["unselected_slide_payload_read_bytes"], f"{path}: unselected overlap conservation")
    _same(call_totals["unselected"], replay["unselected_slide_payload_read_calls"],
          f"{path}: unselected call conservation")
    _same(totals["media"], replay["media_payload_read_bytes"], f"{path}: media overlap conservation")
    _same(call_totals["media"], replay["media_payload_read_calls"],
          f"{path}: media call conservation")
    _same(_covered_bytes(slides, returned_ranges), replay["slide_payload_covered_bytes"],
          f"{path}: slide coverage conservation")
    _same(_covered_bytes([selected], returned_ranges), replay["selected_slide_payload_covered_bytes"],
          f"{path}: selected coverage conservation")
    _same(_covered_bytes(unselected, returned_ranges), replay["unselected_slide_payload_covered_bytes"],
          f"{path}: unselected coverage conservation")
    _same(_covered_bytes(media, returned_ranges), replay["media_payload_covered_bytes"],
          f"{path}: media coverage conservation")
    full_slide_count = sum(
        _covered_bytes([slide], returned_ranges) == slide[1] - slide[0] for slide in slides
    )
    _same(full_slide_count, replay["slide_payload_ranges_fully_covered"],
          f"{path}: full-slide coverage conservation")
    _same(_covered_bytes([selected], returned_ranges) == selected[1] - selected[0],
          replay["selected_slide_payload_fully_covered"],
          f"{path}: selected coverage boolean differs")
    _same(open_totals["slides"], probe.get("open_slide_payload_overlap_bytes"), f"{path}: open slide overlap")
    _same(open_totals["selected"], probe.get("open_selected_slide_payload_overlap_bytes"), f"{path}: open selected overlap")
    _same(open_totals["unselected"], probe.get("open_unselected_slide_payload_overlap_bytes"), f"{path}: open unselected overlap")
    _same(open_totals["media"], probe.get("open_media_payload_overlap_bytes"), f"{path}: open media overlap")
    tail_overlaps = {
        "slides": _overlap(tail_offset, expected_bytes, slides),
        "selected": _overlap(tail_offset, expected_bytes, [selected]),
        "media": _overlap(tail_offset, expected_bytes, media),
    }
    tail_overlaps["unselected"] = tail_overlaps["slides"] - tail_overlaps["selected"]
    _same(open_totals, tail_overlaps, f"{path}: open payload overlap is not the exact tail proof")
    query_totals = {key: totals[key] - open_totals[key] for key in totals}
    require(all(value >= 0 for value in query_totals.values()),
            f"{path}: open/query payload overlap conservation is invalid")
    require(query_totals["unselected"] == 0 and query_totals["media"] == 0,
            f"{path}: semantic query read an unselected slide or media payload")
    _same(_covered_bytes([selected], query_returned_ranges), 522,
          f"{path}: semantic query did not cover the selected slide")
    _same(_covered_bytes(unselected, query_returned_ranges), 0,
          f"{path}: semantic query unselected coverage is nonzero")
    _same(_covered_bytes(media, query_returned_ranges), 0,
          f"{path}: semantic query media coverage is nonzero")
    require(replay["selected_slide_payload_fully_covered"],
            f"{path}: selected slide is not fully covered")


def _sample(sample: Any, case: str, state: str, source: dict[str, Any], expected_output: str | None,
            path: str, seen_pids: set[int]) -> dict[str, Any] | None:
    require(isinstance(sample, dict), f"{path}: sample is malformed")
    _same(sample.get("cache_state"), state, f"{path}.cache_state")
    _uint(sample.get("elapsed_ns"), f"{path}.elapsed_ns", positive=True)
    _uint(sample.get("parent_wall_ns"), f"{path}.parent_wall_ns", positive=True)
    require(sample["parent_wall_ns"] >= sample["elapsed_ns"],
            f"{path}: parent wall clock is shorter than child elapsed clock")
    pid = _uint(sample.get("child_process_id"), f"{path}.child_process_id", positive=True)
    require(pid not in seen_pids, f"{path}: fresh-child PID was reused")
    seen_pids.add(pid)
    _logical_reads(sample, case, path)
    metrics = _process_metrics(sample, path)
    if case.startswith("opc_file_") and "source_open" not in case and "source_one_part" not in case:
        _same(sample.get("opc_materialized_parts"), 4, f"{path}.opc_materialized_parts")
    if case == "opc_file_source_open":
        _same(sample.get("opc_materialized_parts"), 0, f"{path}.opc_materialized_parts")
    if case == "opc_file_source_one_part_atomic_save":
        _same(sample.get("opc_materialized_parts"), 0, f"{path}.opc_materialized_parts")
    is_save = case in {"opc_file_eager_one_part_atomic_save", "opc_file_source_one_part_atomic_save"}
    observed_output = sample.get("output_sha256")
    observed_bytes: int | None = None
    if is_save:
        _hash(observed_output, f"{path}.output_sha256")
        require(observed_output == expected_output,
                f"{path}.output_sha256: differs from the report's state-specific oracle")
        output_bytes = _uint(sample.get("output_bytes"), f"{path}.output_bytes", positive=True)
        observed_bytes = output_bytes
        aligned_output = ((source["archive_bytes"] + 4095) // 4096) * 4096
        expected_bytes = aligned_output if (
            state == "cold-verified" and case == "opc_file_source_one_part_atomic_save"
        ) else source["archive_bytes"]
        _same(output_bytes, expected_bytes, f"{path}.output_bytes")
        _, _, ordinary_output = _source_oracles()
        if state == "warm" or case == "opc_file_eager_one_part_atomic_save":
            _same(observed_output, ordinary_output, f"{path}.output_sha256")
    else:
        require(observed_output is None and sample.get("output_bytes") is None,
                f"{path}: open operation emitted output evidence")
    proof: dict[str, Any] | None = None
    if state == "warm":
        _same(sample.get("cold_advice"), "not_requested", f"{path}.cold_advice")
        require(sample.get("cold_verified") is None, f"{path}: warm row has cold proof")
    else:
        _same(sample.get("cold_advice"), "not_requested", f"{path}.cold_advice")
        proof = sample.get("cold_verified")
        proof = _cold_proof(proof, source, f"{path}.cold_verified")
        _same(proof.get("read_bytes_delta"), metrics["read_bytes"],
              f"{path}: process read_bytes does not match cold proof")
        require(proof["read_bytes_delta"] > 0, f"{path}: cold read_bytes is not positive")
    if case == "pptx_file_source_open_selected_slide_lifecycle":
        _replay(sample.get("pptx_source_replay"), case, source, state, proof,
                f"{path}.pptx_source_replay")
    else:
        require(sample.get("pptx_source_replay") is None,
                f"{path}: eager PPTX or OPC row emitted PPTX replay evidence")
    return proof


def _result_oracle(result: Any, case: str, state: str, source: dict[str, Any],
                   samples: int, path: str) -> str | None:
    require(isinstance(result, dict), f"{path}: timed result is malformed")
    _same(result.get("case"), case, f"{path}.case")
    _same(result.get("cache_state"), state, f"{path}.cache_state")
    result_corpus = _manifest(result.get("corpus"), f"{path}.corpus", case)
    _same(result_corpus, source, f"{path}.corpus: differs from evidence corpus")
    _stats(result.get("elapsed_ns"), f"{path}.elapsed_ns", samples)
    output = result.get("output_sha256")
    is_save = case in {"opc_file_eager_one_part_atomic_save", "opc_file_source_one_part_atomic_save"}
    if is_save:
        _hash(output, f"{path}.output_sha256")
        _, _, ordinary_output = _source_oracles()
        if state == "warm" or case == "opc_file_eager_one_part_atomic_save":
            _same(output, ordinary_output, f"{path}.output_sha256")
        return output
    require(output is None, f"{path}.output_sha256: open operation emitted output")
    require(result.get("output_bytes") is None,
            f"{path}.output_bytes: open operation emitted output size")
    return None


def _opc_evidence_proof(evidence: dict[str, Any], case: str, source: dict[str, Any],
                        cold_samples: list[dict[str, Any]], path: str) -> None:
    if case not in {"opc_file_eager_one_part_atomic_save",
                    "opc_file_source_one_part_atomic_save"}:
        require(evidence.get("opc_cold_alignment_proof") is None,
                f"{path}: non-save evidence exposed OPC alignment proof")
        require(evidence.get("opc_cold_output_proof") is None,
                f"{path}: non-save evidence exposed OPC output proof")
        return
    alignment = evidence.get("opc_cold_alignment_proof")
    output = evidence.get("opc_cold_output_proof")
    require(isinstance(alignment, dict), f"{path}.opc_cold_alignment_proof: missing")
    require(isinstance(output, dict), f"{path}.opc_cold_output_proof: missing")
    for field in ("page_size_bytes", "eocd_offset", "base_bytes", "aligned_bytes",
                  "padding_bytes", "base_comment_bytes", "aligned_comment_bytes"):
        _uint(alignment.get(field), f"{path}.opc_cold_alignment_proof.{field}")
    _same(alignment.get("page_size_bytes"), 4096, f"{path}.opc_cold_alignment_proof.page_size_bytes")
    _same(alignment.get("base_bytes"), source["archive_bytes"], f"{path}.opc_cold_alignment_proof.base_bytes")
    _same(alignment.get("base_sha256"), source["archive_sha256"], f"{path}.opc_cold_alignment_proof.base_sha256")
    _hash(alignment.get("base_sha256"), f"{path}.opc_cold_alignment_proof.base_sha256")
    aligned_bytes = ((source["archive_bytes"] + 4095) // 4096) * 4096
    _same(alignment.get("aligned_bytes"), aligned_bytes, f"{path}.opc_cold_alignment_proof.aligned_bytes")
    _same(alignment.get("padding_bytes"), aligned_bytes - source["archive_bytes"],
          f"{path}.opc_cold_alignment_proof.padding_bytes")
    _same(alignment.get("aligned_comment_bytes"),
          alignment["base_comment_bytes"] + alignment["padding_bytes"],
          f"{path}.opc_cold_alignment_proof.comment_transform")
    _hash(alignment.get("aligned_sha256"), f"{path}.opc_cold_alignment_proof.aligned_sha256")
    _same(alignment.get("aligned_bytes"), alignment["base_bytes"] + alignment["padding_bytes"],
          f"{path}.opc_cold_alignment_proof.size_conservation")
    _same(alignment["eocd_offset"] + 22 + alignment["base_comment_bytes"],
          alignment["base_bytes"], f"{path}.opc_cold_alignment_proof.base_eocd_boundary")
    _same(alignment["eocd_offset"] + 22 + alignment["aligned_comment_bytes"],
          alignment["aligned_bytes"], f"{path}.opc_cold_alignment_proof.aligned_eocd_boundary")
    require(alignment["eocd_offset"] < alignment["base_bytes"],
            f"{path}.opc_cold_alignment_proof.eocd_offset: outside source")
    for field in ("outside_eocd_bytes_equal", "outside_comment_bytes_equal", "zero_suffix",
                  "suffix_zero", "eocd_offset_unchanged"):
        if field in alignment:
            _same(alignment[field], True, f"{path}.opc_cold_alignment_proof.{field}")

    _same(output.get("route"), "eager" if case.startswith("opc_file_eager_") else "source-backed",
          f"{path}.opc_cold_output_proof.route")
    for field in ("eocd_offset", "aligned_source_bytes", "output_bytes",
                  "output_comment_bytes", "unchanged_member_count"):
        _uint(output.get(field), f"{path}.opc_cold_output_proof.{field}")
    _same(output.get("eocd_offset"), alignment["eocd_offset"], f"{path}.opc_cold_output_proof.eocd_offset")
    _same(output.get("aligned_source_bytes"), alignment["aligned_bytes"],
          f"{path}.opc_cold_output_proof.aligned_source_bytes")
    _same(output.get("aligned_source_sha256"), alignment["aligned_sha256"],
          f"{path}.opc_cold_output_proof.aligned_source_sha256")
    _hash(output.get("aligned_source_sha256"), f"{path}.opc_cold_output_proof.aligned_source_sha256")
    _hash(output.get("output_sha256"), f"{path}.opc_cold_output_proof.output_sha256")
    _hash(output.get("canonical_sha256"), f"{path}.opc_cold_output_proof.canonical_sha256")
    _, _, ordinary_output = _source_oracles()
    _same(output.get("canonical_sha256"), ordinary_output,
          f"{path}.opc_cold_output_proof.canonical_sha256")
    _same(output.get("unchanged_member_count"), source["archive_member_count"] - 1,
          f"{path}.opc_cold_output_proof.unchanged_member_count")
    route = output["route"]
    expected_output_bytes = source["archive_bytes"] if route == "eager" else alignment["aligned_bytes"]
    expected_comment_bytes = 0 if route == "eager" else alignment["aligned_comment_bytes"]
    _same(output.get("output_bytes"), expected_output_bytes, f"{path}.opc_cold_output_proof.output_bytes")
    _same(output.get("output_comment_bytes"), expected_comment_bytes,
          f"{path}.opc_cold_output_proof.output_comment_bytes")
    for sample in cold_samples:
        cold = sample.get("cold_verified")
        require(isinstance(cold, dict), f"{path}: cold sample omitted cold verifier proof")
        _same(cold.get("aligned_source_sha256"), alignment.get("aligned_sha256"),
              f"{path}: cold sample/alignment source hash differs")
        _same(cold.get("aligned_source_bytes"), alignment.get("aligned_bytes"),
              f"{path}: cold sample/alignment source length differs")
        _same(sample.get("output_sha256"), output.get("output_sha256"),
              f"{path}: cold sample/output proof hash differs")
        _same(sample.get("output_bytes"), output.get("output_bytes"),
              f"{path}: cold sample/output proof length differs")


def validate_report(report: Path | dict[str, Any], case: str,
                    states: Iterable[str] = STATES, samples: int = 1,
                    warmup: int = 0,
                    binary_descriptor: dict[str, Any] | None = None, *,
                    expected_environment: dict[str, Any] | None = None,
                    seen_pids: set[int] | None = None) -> dict[str, Any]:
    """Validate one report containing one or more cache states.

    ``report`` may be a parsed object for callers that already verified the
    content-addressed receipt.  The positional prefix is intentionally small
    so a later measurement driver can reuse this gate with larger sample
    vectors without changing the report schema.
    """

    require(case in SIX_CASES, f"unknown qualification case: {case}")
    state_tuple = tuple(states)
    require(state_tuple and all(state in STATES for state in state_tuple),
            "report cache-state selection is invalid")
    require(len(set(state_tuple)) == len(state_tuple),
            "report cache-state selection is duplicated")
    require(type(samples) is int and samples > 0,
            "report sample count must be positive")
    require(type(warmup) is int and warmup >= 0,
            "report warmup count must be non-negative")
    if isinstance(report, Path):
        value = read(report)
        label = str(report)
    else:
        value = report
        label = "<report>"
    _finite(value, label)
    require(isinstance(value, dict), f"{label}: report is not an object")
    _same(value.get("schema_version"), REPORT_SCHEMA_VERSION, f"{label}.schema_version")
    _tool(value, case, label)
    require(isinstance(binary_descriptor, dict),
            f"{label}: retained binary descriptor is missing")
    identity = _binary_identity(value, binary_descriptor, label)
    environment = _environment(value, expected_environment, label)
    _configuration(value, case, state_tuple, samples, warmup, label)

    evidence_list = value.get("filesystem_evidence")
    require(isinstance(evidence_list, list) and len(evidence_list) == 1,
            f"{label}: expected one filesystem evidence record")
    evidence = evidence_list[0]
    require(isinstance(evidence, dict), f"{label}: filesystem evidence is malformed")
    _same(evidence.get("case"), case, f"{label}.filesystem_evidence[0].case")
    source = _manifest(evidence.get("corpus"), f"{label}.filesystem_evidence[0].corpus", case)
    _same(evidence.get("warmup_iterations"), warmup, f"{label}: warmup contract differs")
    _same(evidence.get("sample_count"), samples, f"{label}: sample contract differs")
    _same(evidence.get("cache_states"), list(state_tuple), f"{label}: cache-state contract differs")
    _same(evidence.get("fresh_child_per_sample"), True,
          f"{label}: fresh-child contract absent")
    evidence_samples = evidence.get("samples")
    expected_evidence_count = samples * len(state_tuple)
    require(isinstance(evidence_samples, list) and len(evidence_samples) == expected_evidence_count,
            f"{label}: expected {expected_evidence_count} evidence samples")

    results = value.get("results")
    require(isinstance(results, list) and len(results) == len(state_tuple),
            f"{label}: timed result count differs from cache-state selection")
    result_by_state: dict[str, dict[str, Any]] = {}
    output_by_state: dict[str, str | None] = {}
    for index, result in enumerate(results):
        require(isinstance(result, dict), f"{label}.results[{index}]: malformed")
        state = result.get("cache_state")
        require(state in state_tuple,
                f"{label}.results[{index}].cache_state: unexpected state")
        require(state not in result_by_state,
                f"{label}: duplicate timed result state {state}")
        result_by_state[state] = result
        output_by_state[state] = _result_oracle(
            result, case, state, source, samples, f"{label}.results[{index}]"
        )
    require(set(result_by_state) == set(state_tuple),
            f"{label}: timed results do not cover selected states")

    pids = set() if seen_pids is None else seen_pids
    grouped: dict[str, list[dict[str, Any]]] = {state: [] for state in state_tuple}
    pptx_semantic_digests: list[str] = []
    for index, sample in enumerate(evidence_samples):
        require(isinstance(sample, dict),
                f"{label}.filesystem_evidence.samples[{index}]: malformed")
        state = sample.get("cache_state")
        require(state in grouped,
                f"{label}.filesystem_evidence.samples[{index}]: unexpected state")
        grouped[state].append(sample)
    for state in state_tuple:
        selected = grouped[state]
        require(len(selected) == samples,
                f"{label}: state {state} has {len(selected)} samples, expected {samples}")
        expected_output = output_by_state[state]
        for index, sample in enumerate(selected):
            _same(sample.get("sample_index"), index,
                  f"{label}.filesystem_evidence.{state}[{index}].sample_index")
            _sample(sample, case, state, source, expected_output,
                    f"{label}.filesystem_evidence.{state}[{index}]", pids)
            if case == "pptx_file_source_open_selected_slide_lifecycle":
                replay = sample.get("pptx_source_replay")
                require(isinstance(replay, dict),
                        f"{label}: PPTX semantic replay is missing")
                pptx_semantic_digests.append(replay["semantic_sha256"])

    if pptx_semantic_digests:
        _same(len(set(pptx_semantic_digests)), 1,
              f"{label}: PPTX semantic oracle differs by sample or cache state")

    if "cold-verified" in state_tuple:
        _same(evidence.get("cold_verified_status"), "eligible",
              f"{label}.cold_verified_status")
        _same(evidence.get("cold_verified_claim_scope"), CLAIM_SCOPE,
              f"{label}.cold_verified_claim_scope")
        _same(evidence.get("cold_verified_fincore_command"), FINCORE_COMMAND,
              f"{label}.cold_verified_fincore_command")
        proofs = evidence.get("cold_verified_samples")
        require(isinstance(proofs, list) and len(proofs) == samples,
                f"{label}: evidence-level cold proof is missing")
        for index, (proof, sample) in enumerate(zip(proofs, grouped["cold-verified"])):
            _cold_proof(proof, source, f"{label}.cold_verified_samples[{index}]")
            _same(proof, sample.get("cold_verified"),
                  f"{label}: sample/evidence cold proofs differ at {index}")
        _opc_evidence_proof(evidence, case, source, grouped["cold-verified"], label)
    else:
        require(evidence.get("cold_verified_status") is None,
                f"{label}: report omitted state but exposed cold status")
        require(evidence.get("cold_verified_samples") in (None, []),
                f"{label}: report omitted state but exposed cold proof samples")
        require(evidence.get("cold_verified_claim_scope") is None,
                f"{label}: report omitted state but exposed cold claim scope")
        require(evidence.get("cold_verified_fincore_command") is None,
                f"{label}: report omitted state but exposed fincore command")
        require(evidence.get("opc_cold_alignment_proof") is None,
                f"{label}: report omitted state but exposed cold alignment proof")
        require(evidence.get("opc_cold_output_proof") is None,
                f"{label}: report omitted state but exposed cold output proof")
    return {"report": value, "environment": environment, "binary": identity,
            "corpus": source, "outputs": output_by_state,
            "output_sha256": output_by_state.get("warm")}

def _receipt_file(receipt_path: Path, label: str, expected_exit: int) -> dict[str, Any]:
    require(receipt_path.is_file() and not receipt_path.is_symlink(),
            f"{label}: receipt is missing")
    receipt = read(receipt_path)
    require(isinstance(receipt, dict), f"{label}: receipt is not an object")
    require(type(receipt.get("exit_code")) is int,
            f"{label}.exit_code: expected integer")
    _same(receipt.get("exit_code"), expected_exit, f"{label}.exit_code")
    _same(receipt.get("error"), None, f"{label}.error")
    _descriptor(receipt.get("log"), receipt_path.parent, f"{label}.log")
    return receipt


def _receipt(path: Path, descriptor: Any, label: str, expected_exit: int,
             expected_parent: str | None = None) -> dict[str, Any]:
    receipt_path = _descriptor(descriptor, path.parent, f"{label}.descriptor")
    _same(receipt_path.name, "receipt.json", f"{label}: receipt filename")
    if expected_parent is not None:
        _same(receipt_path.parent.name, expected_parent, f"{label}: receipt label")
    return _receipt_file(receipt_path, label, expected_exit)


def _receipt_freeze(receipt: dict[str, Any], binding: dict[str, Any], label: str) -> None:
    freeze_path = binding["freeze_path"]
    freeze_descriptor = receipt.get("freeze")
    require(isinstance(freeze_descriptor, dict), f"{label}.freeze: missing")
    _same(freeze_descriptor.get("path"), str(freeze_path), f"{label}.freeze.path")
    _same(freeze_descriptor.get("bytes"), freeze_path.stat().st_size,
          f"{label}.freeze.bytes")
    _same(freeze_descriptor.get("sha256"), sha(freeze_path),
          f"{label}.freeze.sha256")


def _freeze_binding(packet: Path, stage_hint: str | None = None) -> dict[str, Any]:
    global _SOURCE_STAGE, _SOURCE_ORACLE_CACHE
    packet = packet.resolve()
    origin_path = packet / "origin.json"
    origin = read(origin_path)
    require(isinstance(origin, dict), f"{origin_path}: origin receipt is malformed")
    _same(origin.get("base"), BASE, f"{origin_path}.base")
    if stage_hint is not None:
        require(re.fullmatch(r"repaired-v[0-9]+", stage_hint) is not None,
                f"{packet}/active-stage: invalid requested stage")
        requested_suffix = stage_hint.rsplit("-", 1)[-1]
        active_candidates = (packet / f"active-stage-{requested_suffix}.json",
                             packet / "active-stage.json")
    else:
        active_candidates = (packet / "active-stage.json",)
    active_path = next((candidate for candidate in active_candidates
                        if candidate.is_file() and not candidate.is_symlink()), None)
    require(active_path is not None, f"{packet}: active stage receipt is missing")
    active = read(active_path)
    require(isinstance(active, dict), f"{active_path}: active stage is malformed")
    stage = active.get("stage")
    require(isinstance(stage, str) and re.fullmatch(r"repaired-v[0-9]+", stage) is not None,
            f"{active_path}.stage: invalid repaired stage")
    if stage_hint is not None:
        _same(stage, stage_hint, f"{active_path}.stage")
    suffix = stage.rsplit("-", 1)[-1]
    version = int(suffix[1:])
    require(version >= 2, f"{active_path}.stage: unsupported repaired version")
    expected_previous = "repaired" if version == 2 else f"repaired-v{version - 1}"
    _same(active.get("previous_stage"), expected_previous,
          f"{active_path}.previous_stage")
    active_driver = active.get("driver")
    active_driver_path = _descriptor(active_driver, packet, f"{active_path}.driver")
    _same(active_driver_path.parent, packet, f"{active_path}.driver.parent")
    require(active_driver_path.name.startswith("driver_") and
            active_driver_path.suffix == ".py",
            f"{active_path}.driver.path: unexpected driver name")
    freeze_path = packet / f"freeze-{stage}.json"
    freeze = read(freeze_path)
    require(isinstance(freeze, dict), f"{freeze_path}: freeze receipt is malformed")
    _same(freeze.get("stage"), stage, f"{freeze_path}.stage")
    _same(freeze.get("driver"), active_driver, f"{freeze_path}.driver")
    source = freeze.get("source")
    require(isinstance(source, dict) and source, f"{freeze_path}.source: missing")
    for relative, digest in source.items():
        relative_path = Path(relative) if isinstance(relative, str) else Path("/")
        require(isinstance(relative, str) and relative and
                not relative_path.is_absolute() and ".." not in relative_path.parts,
                f"{freeze_path}.source: invalid path")
        _hash(digest, f"{freeze_path}.source.{relative}")
        frozen = packet / "sources" / stage / relative
        if frozen.exists():
            require(frozen.is_file() and not frozen.is_symlink(),
                    f"{freeze_path}.source: frozen path is not regular {relative}")
            _same(sha(frozen), digest, f"{freeze_path}.frozen.{relative}")
        else:
            raw_candidate = ROOT / relative
            require(not raw_candidate.is_symlink(),
                    f"{freeze_path}.source: current source is a symlink {relative}")
            candidate = raw_candidate.resolve()
            require(ROOT in candidate.parents and candidate.is_file() and not candidate.is_symlink(),
                    f"{freeze_path}.source: missing current source {relative}")
            _same(sha(candidate), digest, f"{freeze_path}.source.{relative}")
    _descriptor(freeze.get("driver"), packet, f"{freeze_path}.driver")
    _descriptor(freeze.get("origin"), packet, f"{freeze_path}.origin")
    _SOURCE_STAGE = stage
    _SOURCE_ORACLE_CACHE = None
    return {"origin": origin, "active": active, "active_path": active_path,
            "freeze": freeze, "freeze_path": freeze_path,
            "stage": stage, "suffix": suffix}


def _prepare_binding(packet: Path, stage_hint: str | None = None) -> dict[str, Any]:
    # Keep the historical helper name for callers, but 0834 binds the repaired
    # freeze and build receipt directly.  No live executable is required.
    binding = _freeze_binding(packet, stage_hint)
    build_path = packet / f"build-{binding['stage']}.json"
    build = read(build_path)
    require(isinstance(build, dict), f"{build_path}: build receipt is malformed")
    binary = build.get("binary")
    binary_path = _descriptor(binary, packet, f"{build_path}.binary", allow_missing=True)
    _same(binary_path.name, "litchi-perf-baseline", f"{build_path}.binary.path")
    _same(binary_path.parent.name, binding["stage"], f"{build_path}.binary.stage")
    if not binary_path.exists():
        cleanup_path = packet / "cleanup.json"
        cleanup = read(cleanup_path)
        require(isinstance(cleanup, dict), f"{cleanup_path}: cleanup receipt is malformed")
        _same(cleanup.get("status"), "pass", f"{cleanup_path}.status")
        require(_contains_descriptor(cleanup, binary),
                f"{cleanup_path}: retained binary descriptor is missing")
    receipt = _receipt(packet, build.get("receipt"), f"{build_path}.receipt", 0)
    freeze_descriptor = binding["freeze_path"]
    receipt_freeze = receipt.get("freeze")
    require(isinstance(receipt_freeze, dict), f"{build_path}.receipt.freeze: missing")
    _same(receipt_freeze.get("path"), str(freeze_descriptor),
          f"{build_path}.receipt.freeze.path")
    _same(receipt_freeze.get("bytes"), freeze_descriptor.stat().st_size,
          f"{build_path}.receipt.freeze.bytes")
    _same(receipt_freeze.get("sha256"), sha(freeze_descriptor),
          f"{build_path}.receipt.freeze.sha256")
    quality_path = packet / f"quality-{binding['suffix']}.json"
    quality = read(quality_path)
    require(isinstance(quality, dict), f"{quality_path}: quality receipt is malformed")
    _same(quality.get("status"), "pass", f"{quality_path}.status")
    _same(quality.get("gates"), list(QUALITY_GATES), f"{quality_path}.gates")
    _same(quality.get("reused"), False, f"{quality_path}.reused")
    _same(quality.get("test_scope"), "filesystem::aligned_zip::tests",
          f"{quality_path}.test_scope")
    active_descriptor = {
        "path": str(binding["active_path"]),
        "bytes": binding["active_path"].stat().st_size,
        "sha256": sha(binding["active_path"]),
    }
    _same(quality.get("amendment"), active_descriptor, f"{quality_path}.amendment")
    full_suite = binding["active"].get("full_suite")
    require(isinstance(full_suite, dict), f"{quality_path}.full_suite: missing")
    _same(quality.get("full_suite"), full_suite, f"{quality_path}.full_suite")
    _receipt(packet, full_suite, f"{quality_path}.full_suite", 0,
             expected_parent="quality-test")
    for gate in QUALITY_GATES:
        gate_path = packet / "commands" / f"quality-{binding['suffix']}-{gate}" / "receipt.json"
        gate_receipt = _receipt_file(gate_path, f"{quality_path}.quality-{gate}", 0)
        _receipt_freeze(gate_receipt, binding, f"{quality_path}.quality-{gate}")
    return {**binding, "build": build, "build_path": build_path, "binary": binary,
            "quality": quality}


def _report_descriptor(row: dict[str, Any], packet: Path, label: str,
                       expected_name: str | None = None) -> tuple[Path, str]:
    descriptor = row.get("report")
    require(isinstance(descriptor, dict), f"{label}.report: successful row omitted report")
    path = _descriptor(descriptor, packet, f"{label}.report")
    if expected_name is not None:
        _same(path.name, expected_name, f"{label}.report.filename")
    digest = descriptor.get("sha256")
    _same(sha(path), digest, f"{label}.report.sha256")
    return path, digest


def _receipt_contract(receipt: dict[str, Any], binding: dict[str, Any], label: str,
                      case_arg: str, samples: int = 1, warmup: int = 0,
                      cache_states: str = "warm,cold-verified") -> None:
    _receipt_freeze(receipt, binding, label)
    argv = receipt.get("argv")
    require(isinstance(argv, list), f"{label}.argv: missing")
    binary = binding.get("binary")
    require(isinstance(binary, dict), f"{label}: build binary descriptor is missing")
    require(isinstance(binary.get("path"), str) and binary["path"],
            f"{label}: build binary path is missing")
    require(binary["path"] in argv, f"{label}.argv: retained binary is not bound")
    require(case_arg in argv, f"{label}.argv: case binding is missing")
    for option, expected in (("--samples", str(samples)), ("--warmup", str(warmup)),
                             ("--filesystem-cache", cache_states)):
        require(option in argv, f"{label}.argv: missing {option}")
        index = argv.index(option)
        require(index + 1 < len(argv), f"{label}.argv: {option} has no value")
        _same(argv[index + 1], expected, f"{label}.argv.{option}")
    require("--filesystem-root" in argv and "--json" in argv,
            f"{label}.argv: selected root/report binding is missing")


def _validate_bundle(path: Path, cases: tuple[str, ...], states: tuple[str, ...],
                     samples: int, warmup: int, binary: dict[str, Any],
                     expected_environment: dict[str, Any] | None,
                     seen_pids: set[int]) -> dict[str, Any]:
    value = read(path)
    _finite(value, str(path))
    require(isinstance(value, dict), f"{path}: report is not an object")
    configuration = value.get("configuration")
    require(isinstance(configuration, dict), f"{path}.configuration: missing")
    _same(configuration.get("cases"), list(cases), f"{path}.configuration.cases")
    evidence = value.get("filesystem_evidence")
    require(isinstance(evidence, list) and len(evidence) == len(cases),
            f"{path}: evidence/case count differs")
    result_list = value.get("results")
    require(isinstance(result_list, list) and len(result_list) == len(cases) * len(states),
            f"{path}: result/case-state count differs")
    expected_environment_local = expected_environment
    outputs: dict[str, dict[str, str | None]] = {}
    checked: list[dict[str, Any]] = []
    for case in cases:
        single = dict(value)
        single_configuration = dict(configuration)
        single_configuration["cases"] = [case]
        single["configuration"] = single_configuration
        single["filesystem_evidence"] = [item for item in evidence
                                           if isinstance(item, dict) and item.get("case") == case]
        single["results"] = [item for item in result_list
                              if isinstance(item, dict) and item.get("case") == case]
        require(len(single["filesystem_evidence"]) == 1,
                f"{path}: duplicate/missing evidence for {case}")
        require(len(single["results"]) == len(states),
                f"{path}: duplicate/missing results for {case}")
        checked_report = validate_report(
            single, case, states, samples, warmup, binary,
            expected_environment=expected_environment_local, seen_pids=seen_pids,
        )
        expected_environment_local = expected_environment_local or checked_report["environment"]
        outputs[case] = checked_report["outputs"]
        checked.append({"case": case, "outputs": checked_report["outputs"]})
    if {item.get("case") for item in evidence if isinstance(item, dict)} != set(cases):
        raise QualificationError(f"{path}: evidence cases are not exactly the requested pair")
    if {item.get("case") for item in result_list if isinstance(item, dict)} != set(cases):
        raise QualificationError(f"{path}: result cases are not exactly the requested pair")
    return {"environment": expected_environment_local, "outputs": outputs,
            "report": value, "reports": checked}


def _pair_output_oracle(bundle: dict[str, Any], cases: tuple[str, ...], states: tuple[str, ...],
                        path: str) -> None:
    if set(cases) != {"opc_file_eager_one_part_atomic_save",
                      "opc_file_source_one_part_atomic_save"}:
        return
    outputs = bundle["outputs"]
    for state in states:
        eager = outputs[cases[0]].get(state)
        source = outputs[cases[1]].get(state)
        _hash(eager, f"{path}.{cases[0]}.{state}")
        _hash(source, f"{path}.{cases[1]}.{state}")
        # Warm output is the ordinary source oracle.  Cold source-backed output
        # is retained as a route-specific oracle: alignment may change only
        # the private EOCD framing, so it is never silently normalized or
        # compared as if it were a timed semantic result.
        if state == "warm":
            _same(eager, source, f"{path}.{state}: warm eager/source output differs")


def _stage_hint_for_packet(path: Path) -> str | None:
    match = re.fullmatch(r"(?:qualification|capture)-(v[0-9]+)", path.stem)
    return f"repaired-{match.group(1)}" if match else None


def validate_manifest(path: Path) -> dict[str, Any]:
    """Validate qualification receipts and retained reports without writes."""

    value = read(path)
    _finite(value, str(path))
    require(isinstance(value, dict), f"{path}: qualification manifest is not an object")
    rows = value.get("rows")
    require(isinstance(rows, list) and len(rows) == len(CASES) + 1,
            f"{path}: expected six cases plus the paired OPC save row")
    status = value.get("status")
    require(status in {"commands_pass", "failed"}, f"{path}.status: unknown status")
    packet = path.parent
    binding = _prepare_binding(packet, _stage_hint_for_packet(path))
    binary = binding["binary"]
    expected_environment: dict[str, Any] | None = None
    seen_pids: set[int] = set()
    corpus_oracles: dict[str, tuple[int, str]] = {}
    checked: list[dict[str, Any]] = []
    failed_rows: list[dict[str, Any]] = []
    validated_sample_count = 0
    expected_cases: list[tuple[str, ...]] = [(case,) for case in CASES]
    expected_cases.append((CASES[2], CASES[3]))
    for index, (row, row_cases) in enumerate(zip(rows, expected_cases)):
        label = f"{path}.rows[{index}]"
        row_label = (f"qualification-{binding['suffix']}-{index:02}"
                     if index < len(CASES) else f"qualification-{binding['suffix']}-opc-pair")
        require(isinstance(row, dict), f"{label}: row is malformed")
        case_arg = row_cases[0] if len(row_cases) == 1 else ",".join(row_cases)
        _same(row.get("case"), case_arg, f"{label}.case")
        _same(row.get("state"), "warm,cold-verified", f"{label}.state")
        require(type(row.get("exit_code")) is int and row["exit_code"] >= 0,
                f"{label}.exit_code: invalid value")
        receipt = _receipt(path, row.get("receipt"), f"{label}.receipt", row["exit_code"],
                           expected_parent=row_label)
        _receipt_contract(receipt, binding, f"{label}.receipt", case_arg)
        report_descriptor = row.get("report")
        if row["exit_code"] != 0:
            require(status == "failed", f"{label}: commands_pass contains a failed command")
            require(index in {2, 3, len(rows) - 1},
                    f"{label}: unexpected failed qualification row")
            _same(row["exit_code"], 1, f"{label}.exit_code")
            require(report_descriptor is None, f"{label}: failed command emitted a report")
            log_path = _descriptor(receipt.get("log"), packet, f"{label}.log")
            try:
                text = log_path.read_text(encoding="utf-8", errors="replace")
            except OSError as error:
                raise QualificationError(f"{label}: cannot read retained error log: {error}") from error
            _same(text.strip(), EXPECTED_OPC_FAILURE, f"{label}: retained failure text")
            failed_rows.append({"index": index, "case": case_arg,
                                "exit_code": row["exit_code"],
                                "diagnostic_sha256": sha(log_path),
                                "diagnostic": EXPECTED_OPC_FAILURE})
            continue
        report_path, report_digest = _report_descriptor(row, packet, label,
                                                        f"{row_label}.json")
        if len(row_cases) == 1:
            checked_report = validate_report(
                report_path, row_cases[0], STATES, 1, 0, binary,
                expected_environment=expected_environment, seen_pids=seen_pids,
            )
            expected_environment = expected_environment or checked_report["environment"]
            outputs = {row_cases[0]: checked_report["outputs"]}
            report_record = {"case": row_cases[0], "report": str(report_path),
                             "report_sha256": report_digest, "outputs": checked_report["outputs"]}
            validated_sample_count += len(STATES)
        else:
            bundle = _validate_bundle(report_path, row_cases, STATES, 1, 0, binary,
                                      expected_environment, seen_pids)
            expected_environment = expected_environment or bundle["environment"]
            _pair_output_oracle(bundle, row_cases, STATES, label)
            outputs = bundle["outputs"]
            report_record = {"case": case_arg, "report": str(report_path),
                             "report_sha256": report_digest, "outputs": outputs}
            validated_sample_count += len(row_cases) * len(STATES)
        for case, case_outputs in outputs.items():
            corpus = _corpus(case)
            key = corpus["generator"]
            oracle = (corpus["archive_bytes"], corpus["archive_sha256"])
            if key in corpus_oracles:
                _same(oracle, corpus_oracles[key], f"{label}.{case}.corpus: source identity differs")
            else:
                corpus_oracles[key] = oracle
        checked.append(report_record)
    if status == "commands_pass":
        require(not failed_rows and len(checked) == len(rows),
                f"{path}: commands_pass manifest has failed rows")
    else:
        require([row["index"] for row in failed_rows] == [2, 3, len(rows) - 1],
                f"{path}: failed rows do not match the retained OPC proof failure set")
        require(len(checked) == 4 and len(failed_rows) == 3 and
                len(checked) + len(failed_rows) == len(rows),
                f"{path}: failed manifest does not retain every row")
    return {"status": "pass" if status == "commands_pass" else "failed",
            "qualification_valid": status == "commands_pass" and not failed_rows,
            "reports": checked, "validated_sample_count": validated_sample_count,
            "failed_rows": failed_rows,
            "claim_authorized": False, "performance_claim": "none"}


def validate_capture(path: Path) -> dict[str, Any]:
    """Validate the later 72-report capture using the same report gate.

    This function checks custody and evidence for each native report.  It does
    not calculate ratios, quantiles across cases, or any before/after claim.
    """

    value = read(path)
    _finite(value, str(path))
    require(isinstance(value, dict), f"{path}: capture manifest is not an object")
    _same(value.get("status"), "commands_pass", f"{path}.status")
    rows = value.get("rows")
    require(isinstance(rows, list) and len(rows) == 72,
            f"{path}: expected 72 native capture rows")
    require(value.get("report_count") == 72, f"{path}.report_count: mismatch")
    require(value.get("sample_count") == 2160, f"{path}.sample_count: mismatch")
    packet = path.parent
    binding = _prepare_binding(packet, _stage_hint_for_packet(path))
    binary = binding["binary"]
    expected_environment: dict[str, Any] | None = None
    seen_pids: set[int] = set()
    checked: list[dict[str, Any]] = []
    for index, row in enumerate(rows):
        label = f"{path}.rows[{index}]"
        require(isinstance(row, dict), f"{label}: row is malformed")
        plan = row.get("plan")
        result = row.get("result")
        require(isinstance(plan, dict) and isinstance(result, dict), f"{label}: plan/result missing")
        case = plan.get("case")
        state = plan.get("cache_state")
        require(case in SIX_CASES and state in STATES, f"{label}: unknown case/state")
        _same(plan.get("samples"), 30, f"{label}.plan.samples")
        _same(plan.get("warmup"), 3, f"{label}.plan.warmup")
        _same(result.get("case"), case, f"{label}.result.case")
        _same(result.get("state"), state, f"{label}.result.state")
        require(type(result.get("exit_code")) is int and result["exit_code"] == 0,
                f"{label}.result.exit_code: expected zero")
        row_label = f"native-{index:03}"
        receipt = _receipt(path, result.get("receipt"), f"{label}.receipt", 0,
                           expected_parent=row_label)
        _receipt_contract(receipt, binding, f"{label}.receipt", case, 30, 3, state)
        report_path, report_digest = _report_descriptor(result, packet, label,
                                                        f"{row_label}.json")
        checked_report = validate_report(
            report_path, case, (state,), 30, 3, binary,
            expected_environment=expected_environment, seen_pids=seen_pids,
        )
        expected_environment = expected_environment or checked_report["environment"]
        checked.append({"index": index, "case": case, "state": state,
                        "report": str(report_path), "report_sha256": report_digest})
    return {"status": "pass", "capture_valid": True, "reports": checked,
            "claim_authorized": False, "performance_claim": "none"}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("manifest", nargs="?", type=Path, default=None)
    parser.add_argument("--capture", type=Path, default=None,
                        help="validate a retained formal capture.json instead of qualification.json")
    args = parser.parse_args()
    path = args.capture or args.manifest or (HERE / f"qualification-{ACTIVE_SUFFIX}.json")
    capture_mode = args.capture is not None or (
        args.manifest is not None and path.name == "capture.json"
    )
    try:
        result = validate_capture(path.resolve()) if capture_mode else validate_manifest(path.resolve())
    except Exception as error:
        result = {"status": "invalid", "qualification_valid": False,
                  "claim_authorized": False, "performance_claim": "none",
                  "error": str(error)}
        print(json.dumps(result, sort_keys=True))
        return 1
    print(json.dumps(result, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
