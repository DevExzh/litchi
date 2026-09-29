#!/usr/bin/env python3
"""Offline admission reader for the 0833 filesystem qualification.

This module only reads JSON and source/build receipts.  It never starts a
process, probes a filesystem, builds a binary, compares timings, or treats a
qualification row as a performance result.  The driver may use
``validate_report`` directly, or write the small immutable manifest consumed
by ``main``::

    python3 -B reader.py qualification.json

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
from pathlib import Path
from typing import Any, Iterable


HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
BASE = "eefaca16e39ace3c1e6219aedbac4fe1c0cd060a"
SCHEMA = "litchi.0833.filesystem-qualification.v1"
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

# These values are copied from the literals and deterministic builders in
# tools/perf-baseline/src/filesystem.rs and lib.rs.  They are deliberately
# independent of a captured report: a report cannot redefine its own oracle.
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
    "archive_sha256": "a0c1af9e2c7a19148b44fc2a8c594c7a274131d74f9f042d55b487d5337cd1e6",
    "target_entry": "benchmark/parts/00002.bin",
    "target_payload_bytes": 4 * 1024 * 1024,
    "target_payload_sha256": "3dbf6225021a99c1da8750a738bde21f57591c0be1a60aa510966c47ee25b098",
    "xlsx": None,
}
# The PPTX builder intentionally computes its archive bytes and digest from the
# current source generator.  Keep the stable generator/shape oracles here, but
# bind the archive identity to the fresh qualification reports below instead of
# importing a digest from an older capture.
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
OPC_OUTPUT_SHA256 = "f4bbe4de18853444cc6cd093cf561249decaa81f776afcf5de122667f5dd7009"
PPTX_REPLAY_SHA256 = "f5f7db181150c00a4323a48c142721ead73aca3ad7c3b3594e8b1a18a686b257"
PPTX_REPLAY_CLASSIFICATION = "selected-slide-only:target-slide-no-unselected-or-media-overlap"

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
    require(type(value) is int and value >= (1 if positive else 0),
            f"{path}: expected {'positive' if positive else 'non-negative'} integer")
    return value


def _same(actual: Any, expected: Any, path: str) -> None:
    require(actual == expected, f"{path}: oracle differs (expected {expected!r}, got {actual!r})")


def _corpus(case: str) -> dict[str, Any]:
    return OPC_CORPUS if case.startswith("opc_") else PPTX_CORPUS


def _manifest(value: Any, path: str, case: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{path}: corpus manifest is missing")
    expected = _corpus(case)
    if case.startswith("opc_"):
        _same(value, expected, path)
    else:
        # These fields are literals or fixed source-generator dimensions.  The
        # archive length/hash are generated by the current PPTX source builder
        # and are checked for identity across this qualification set below.
        fixed = {key: item for key, item in expected.items()
                 if key not in {"archive_bytes", "archive_sha256"}}
        for key, item in fixed.items():
            _same(value.get(key), item, f"{path}.{key}")
        _uint(value.get("archive_bytes"), f"{path}.archive_bytes", positive=True)
        _hash(value.get("archive_sha256"), f"{path}.archive_sha256")
        require(set(value) == set(expected), f"{path}: corpus manifest fields differ")
    return value


def _source_map(value: Any, label: str) -> dict[str, str]:
    """Verify a driver source map against the current repository bytes."""

    require(isinstance(value, dict) and value, f"{label}: source map is missing")
    result: dict[str, str] = {}
    for relative, digest in value.items():
        require(isinstance(relative, str) and relative and not Path(relative).is_absolute(),
                f"{label}: source path is invalid")
        path = (ROOT / relative).resolve()
        require(path.is_file() and not path.is_symlink(), f"{label}: source is missing: {relative}")
        require(ROOT in path.parents, f"{label}: source escaped repository: {relative}")
        _hash(digest, f"{label}.{relative}")
        _same(sha(path), digest, f"{label}.{relative}")
        result[relative] = digest
    return result


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
    candidate = candidate.resolve()
    if not candidate.exists():
        require(allow_missing, f"{label}: file is missing: {candidate}")
        return candidate
    require(candidate.is_file() and not candidate.is_symlink(),
            f"{label}: descriptor is not a regular file")
    _same(candidate.stat().st_size, value["bytes"], f"{label}.bytes")
    _same(sha(candidate), value["sha256"], f"{label}.sha256")
    return candidate


def _receipt_binding(value: Any, base: Path, label: str) -> None:
    """Check optional content-addressed build/command receipt bindings."""

    if value is None:
        return
    require(isinstance(value, dict), f"{label}: receipt binding is malformed")
    path_value = value.get("path", value.get("receipt"))
    digest = value.get("sha256", value.get("receipt_sha256"))
    if path_value is None:
        require(digest is None, f"{label}: digest has no receipt path")
        return
    path = _path_from(path_value, base, f"{label}.path")
    _hash(digest, f"{label}.sha256")
    _same(sha(path), digest, f"{label}.sha256")


def _binary_identity(report: dict[str, Any], expected: dict[str, Any] | None, path: str) -> dict[str, Any]:
    identity = report.get("binary_identity")
    require(isinstance(identity, dict), f"{path}: binary identity is missing")
    _hash(identity.get("binary_sha256"), f"{path}.binary_identity.binary_sha256")
    _uint(identity.get("binary_bytes"), f"{path}.binary_identity.binary_bytes", positive=True)
    _same(identity.get("executable"), True, f"{path}.binary_identity.executable")
    _same(identity.get("profile"), "release", f"{path}.binary_identity.profile")
    if expected:
        aliases = {
            "binary_sha256": ("binary_sha256", "sha256"),
            "binary_bytes": ("binary_bytes", "bytes"),
            "profile": ("profile",),
            "executable": ("executable",),
            "path": ("path",),
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
    values = [_uint(item, f"{path}.samples[{index}]", positive=True)
              for index, item in enumerate(raw)]
    _same(value.get("sample_order"), list(range(expected_samples)), f"{path}.sample_order")
    ordered = sorted(values)
    percentile = lambda rank: ordered[min(expected_samples - 1, math.ceil(rank * expected_samples) - 1)]
    _same(value.get("min"), min(values), f"{path}.min")
    _same(value.get("p50"), percentile(0.50), f"{path}.p50")
    _same(value.get("p95"), percentile(0.95), f"{path}.p95")
    _same(value.get("p99"), percentile(0.99), f"{path}.p99")
    _same(value.get("max"), max(values), f"{path}.max")
    mean = value.get("mean")
    require(type(mean) in (int, float) and math.isfinite(float(mean)) and
            math.isclose(float(mean), sum(values) / expected_samples, rel_tol=0.0, abs_tol=1e-9),
            f"{path}.mean: does not match sample vector")
    return values


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
    _same(proof.get("fincore_version_stderr_sha256"), EMPTY_SHA256,
          f"{path}.fincore_version_stderr_sha256")
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
    _same(replay.get("slide_count"), 200, f"{path}.slide_count")
    _same(replay.get("selected_position"), 100, f"{path}.selected_position")
    _same(replay.get("semantic_sha256"), PPTX_REPLAY_SHA256, f"{path}.semantic_sha256")
    _hash(replay.get("semantic_sha256"), f"{path}.semantic_sha256")
    _same(replay.get("classification"), PPTX_REPLAY_CLASSIFICATION, f"{path}.classification")
    operation = replay.get("operation")
    _same(operation, "open_selected_slide_lifecycle", f"{path}.operation")
    sizes = replay.get("read_return_sizes")
    require(isinstance(sizes, list), f"{path}.read_return_sizes: missing")
    values = [_uint(value, f"{path}.read_return_sizes[{index}]")
              for index, value in enumerate(sizes)]
    _same(replay.get("read_calls"), len(values), f"{path}.read_calls")
    _same(replay.get("read_bytes"), sum(values), f"{path}.read_bytes")
    for field in (
        "slide_payload_read_calls", "slide_payload_read_bytes", "slide_payload_covered_bytes",
        "slide_payload_ranges_fully_covered", "selected_slide_payload_read_calls",
        "selected_slide_payload_read_bytes", "selected_slide_payload_covered_bytes",
        "unselected_slide_payload_read_calls", "unselected_slide_payload_read_bytes",
        "unselected_slide_payload_covered_bytes", "media_payload_read_calls",
        "media_payload_read_bytes", "media_payload_covered_bytes",
    ):
        _uint(replay.get(field), f"{path}.{field}")
    _same(replay.get("selected_slide_payload_fully_covered"), True,
          f"{path}.selected_slide_payload_fully_covered")
    _same(replay.get("selected_slide_payload_read_bytes"),
          replay.get("selected_slide_payload_covered_bytes"),
          f"{path}: selected payload read/coverage differs")
    _same(replay.get("selected_slide_payload_covered_bytes"), 522,
          f"{path}.selected_slide_payload_covered_bytes")
    _same(replay.get("unselected_slide_payload_read_bytes"), 0,
          f"{path}.unselected_slide_payload_read_bytes")
    _same(replay.get("media_payload_read_bytes"), 0, f"{path}.media_payload_read_bytes")


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
    if is_save:
        _hash(observed_output, f"{path}.output_sha256")
        require(observed_output == expected_output,
                f"{path}.output_sha256: differs from the report's state-specific oracle")
        output_bytes = _uint(sample.get("output_bytes"), f"{path}.output_bytes", positive=True)
        aligned_output = ((source["archive_bytes"] + 4095) // 4096) * 4096
        expected_bytes = aligned_output if (
            state == "cold-verified" and case == "opc_file_source_one_part_atomic_save"
        ) else source["archive_bytes"]
        _same(output_bytes, expected_bytes, f"{path}.output_bytes")
        if state == "warm" or case == "opc_file_eager_one_part_atomic_save":
            _same(observed_output, OPC_OUTPUT_SHA256, f"{path}.output_sha256")
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
        if state == "warm" or case == "opc_file_eager_one_part_atomic_save":
            _same(output, OPC_OUTPUT_SHA256, f"{path}.output_sha256")
        return output
    require(output is None, f"{path}.output_sha256: open operation emitted output")
    require(result.get("output_bytes") is None,
            f"{path}.output_bytes: open operation emitted output size")
    return None


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
    else:
        require(evidence.get("cold_verified_status") is None,
                f"{label}: report omitted state but exposed cold status")
        require(evidence.get("cold_verified_samples") in (None, []),
                f"{label}: report omitted state but exposed cold proof samples")
    return {"report": value, "environment": environment, "binary": identity,
            "corpus": source, "outputs": output_by_state,
            "output_sha256": output_by_state.get("warm")}

def _run_report_path(run: dict[str, Any], base: Path, label: str) -> Path:
    require(isinstance(run, dict), f"{label}: run is malformed")
    value = run.get("report", run.get("report_path", run.get("path")))
    return _path_from(value, base, f"{label}.report")


def _receipt(path: Path, descriptor: Any, label: str, expected_exit: int) -> dict[str, Any]:
    receipt_path = _descriptor(descriptor, path.parent, f"{label}.descriptor")
    receipt = read(receipt_path)
    require(isinstance(receipt, dict), f"{label}: receipt is not an object")
    _same(receipt.get("exit_code"), expected_exit, f"{label}.exit_code")
    _same(receipt.get("error"), None, f"{label}.error")
    log = receipt.get("log")
    _descriptor(log, path.parent, f"{label}.log")
    return receipt


def _prepare_binding(packet: Path) -> dict[str, Any]:
    prepare_path = packet / "prepare.json"
    prepare = read(prepare_path)
    require(isinstance(prepare, dict), f"{prepare_path}: prepare receipt is malformed")
    _same(prepare.get("base"), BASE, f"{prepare_path}.base")
    _source_map(prepare.get("source"), f"{prepare_path}.source")
    _source_map(prepare.get("normative"), f"{prepare_path}.normative")
    unrelated = prepare.get("unrelated")
    require(isinstance(unrelated, dict), f"{prepare_path}.unrelated: missing")
    for relative, digest in unrelated.items():
        candidate = (ROOT / relative).resolve()
        require(ROOT in candidate.parents, f"{prepare_path}.unrelated: path escaped root")
        require(candidate.is_file() and not candidate.is_symlink(),
                f"{prepare_path}.unrelated: file is missing: {relative}")
        _hash(digest, f"{prepare_path}.unrelated.{relative}")
        _same(sha(candidate), digest, f"{prepare_path}.unrelated.{relative}")
    frozen = prepare.get("frozen")
    require(isinstance(frozen, dict) and frozen, f"{prepare_path}.frozen: missing")
    for relative, digest in frozen.items():
        candidate = packet / relative
        require(candidate.is_file() and not candidate.is_symlink(),
                f"{prepare_path}.frozen: file is missing: {relative}")
        _hash(digest, f"{prepare_path}.frozen.{relative}")
        _same(sha(candidate), digest, f"{prepare_path}.frozen.{relative}")
    fincore = prepare.get("fincore")
    _descriptor(fincore, packet, f"{prepare_path}.fincore", allow_missing=True)
    return prepare


def validate_manifest(path: Path) -> dict[str, Any]:
    """Validate the six-row driver's receipts and any retained reports.

    A failed terminal row is admissible only as a recorded qualification
    failure: its receipt and diagnostic log are checked, its report is absent,
    and no formal result or performance claim is produced.
    """

    value = read(path)
    _finite(value, str(path))
    require(isinstance(value, dict), f"{path}: qualification manifest is not an object")
    rows = value.get("rows")
    require(isinstance(rows, list) and len(rows) == len(CASES),
            f"{path}: expected six qualification rows")
    status = value.get("status")
    require(status in {"commands_pass", "failed"}, f"{path}.status: unknown status")
    packet = path.parent
    prepare = _prepare_binding(packet)
    prepare_hash = sha(packet / "prepare.json")
    build_path = packet / "build.json"
    build = read(build_path)
    require(isinstance(build, dict), f"{build_path}: build receipt is malformed")
    binary = build.get("binary")
    require(isinstance(binary, dict), f"{build_path}.binary: missing")
    _descriptor(binary, packet, f"{build_path}.binary", allow_missing=True)
    build_receipt_descriptor = build.get("receipt")
    build_receipt = _receipt(packet, build_receipt_descriptor,
                             f"{build_path}.receipt", 0)
    _same(build_receipt.get("prepare_sha256"), prepare_hash,
          f"{build_path}.receipt.prepare_sha256")

    expected_environment: dict[str, Any] | None = None
    seen_pids: set[int] = set()
    corpus_oracles: dict[str, tuple[int, str]] = {}
    checked: list[dict[str, Any]] = []
    failed_rows: list[dict[str, Any]] = []
    for index, (row, case) in enumerate(zip(rows, CASES)):
        label = f"{path}.rows[{index}]"
        require(isinstance(row, dict), f"{label}: row is malformed")
        _same(row.get("case"), case, f"{label}.case")
        exit_code = row.get("exit_code")
        require(type(exit_code) is int and exit_code >= 0,
                f"{label}.exit_code: invalid value")
        receipt = _receipt(path, row.get("receipt"), f"{label}.receipt", exit_code)
        _same(receipt.get("prepare_sha256"), prepare_hash,
              f"{label}.receipt.prepare_sha256")
        argv = receipt.get("argv")
        require(isinstance(argv, list), f"{label}.receipt.argv: missing")
        require(case in argv and "--warmup" in argv and "--samples" in argv,
                f"{label}.receipt.argv: case/sample contract missing")
        _same(argv[argv.index("--warmup") + 1], "0", f"{label}.receipt.argv.warmup")
        _same(argv[argv.index("--samples") + 1], "1", f"{label}.receipt.argv.samples")
        require("--filesystem-cache" in argv,
                f"{label}.receipt.argv: cache contract missing")
        _same(argv[argv.index("--filesystem-cache") + 1], "warm,cold-verified",
              f"{label}.receipt.argv.cache_states")

        report_descriptor = row.get("report")
        if exit_code == 0:
            require(report_descriptor is not None,
                    f"{label}: successful command omitted report descriptor")
            report_path = _descriptor(report_descriptor, packet, f"{label}.report")
            report_digest = row["report"]["sha256"]
            require(sha(report_path) == report_digest,
                    f"{label}.report.sha256: receipt and report differ")
            checked_report = validate_report(
                report_path, case, STATES, 1, 0, binary,
                expected_environment=expected_environment, seen_pids=seen_pids,
            )
            expected_environment = expected_environment or checked_report["environment"]
            corpus = checked_report["corpus"]
            key = corpus["generator"]
            oracle = (corpus["archive_bytes"], corpus["archive_sha256"])
            if key in corpus_oracles:
                _same(oracle, corpus_oracles[key],
                      f"{label}.report.corpus: source identity differs")
            else:
                corpus_oracles[key] = oracle
            checked.append({"case": case, "report": str(report_path),
                            "report_sha256": report_digest,
                            "outputs": checked_report["outputs"]})
        else:
            require(status == "failed", f"{label}: nonzero command in commands_pass manifest")
            require(index == len(CASES) - 1 and case == CASES[-1],
                    f"{label}: only the terminal PPTX replay row may fail qualification")
            require(report_descriptor is None,
                    f"{label}: failed command emitted a report")
            log_path = Path(receipt["log"]["path"])
            diagnostic = log_path.read_text(encoding="utf-8", errors="replace")
            require("PPTX source replay violated" in diagnostic,
                    f"{label}: expected replay-classification terminal error")
            failed_rows.append({"case": case, "exit_code": exit_code,
                                "diagnostic": "pptx_source_replay_classification"})

    if status == "commands_pass":
        require(not failed_rows and len(checked) == len(CASES),
                f"{path}: commands_pass manifest has failed rows")
    else:
        require(failed_rows and len(checked) == len(CASES) - 1,
                f"{path}: failed manifest does not contain the expected terminal failure")
    return {
        "status": "pass" if status == "commands_pass" else "failed",
        "qualification_valid": status == "commands_pass",
        "reports": checked,
        "failed_rows": failed_rows,
        "claim_authorized": False,
        "performance_claim": "none",
    }

def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("manifest", nargs="?", type=Path, default=HERE / "qualification.json")
    args = parser.parse_args()
    result = validate_manifest(args.manifest.resolve())
    print(json.dumps(result, sort_keys=True))


if __name__ == "__main__":
    main()
