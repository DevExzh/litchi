#!/usr/bin/env python3
"""Validate and summarize the normal 0467 XLSX ABBA capture.

The capture driver owns process isolation and raw report creation.  This module
is deliberately read-only: it validates the formal normal lanes, delegates
the report/schema/statistic checks to :mod:`tools.perf_abba_summary`, parses
the process-total RSS receipts, and emits a descriptive comparison containing
one row for each case/corpus pair.

The nested ``abba_summary`` is the canonical summary produced by the existing
ABBA tool.  It keeps the repository claim-registry contract intact (including
its p50/mean/p95/p99 drift policy).  The separate ``review_triggers`` section
applies the 0467 protocol's 5% review threshold to every statistic and to
process RSS.  A review trigger is not a speedup or resource claim.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import math
from pathlib import Path
import re
import sys
from typing import Any, Iterable, Mapping, Sequence


ROOT = Path(__file__).resolve().parent
REPO_ROOT = next(
    (candidate for candidate in ROOT.parents if (candidate / "tools").is_dir()),
    ROOT,
)
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from tools import perf_abba_summary, perf_compare  # noqa: E402


SCHEMA = "litchi-0467-xlsx-abba-analysis-v1"
PROTOCOL_SCHEMA = "litchi-0467-protocol-v2"
LANES = ("a1", "b1", "b2", "a2")
LANE_NAMES = ("A1-clean", "B1-clean", "B2-clean", "A2-clean")
LANE_ROLES = {
    "A1-clean": "control",
    "A2-clean": "control",
    "B1-clean": "candidate",
    "B2-clean": "candidate",
}


def _lane_key(name: str) -> str:
    """Return the canonical lower-case ABBA key for a formal lane name."""

    return name.split("-", 1)[0].lower()
CASES = ("xlsx_one_cell_commit_save", "xlsx_one_percent_commit_save")
SHAPES = ("tiny", "medium", "dense-wide")
STATISTICS = ("mean", "p50", "p95", "p99")
SAMPLES = 100
WARMUPS = 5
CPU = "2"
WORKERS = 1
REVIEW_THRESHOLD_PERCENT = 5.0
FULL_GUARD_SAMPLES = 15
FULL_GUARD_WARMUPS = 3
HEAP_GUARD_SAMPLES = 5
HEAP_GUARD_WARMUPS = 1
FULL_GUARD_LANES = ("A-full-clean", "B-full-clean")
HEAP_GUARD_LANES = ("A-heap-clean", "B-heap-clean")
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
REVISION_RE = re.compile(r"^[0-9a-f]{40}$")
EXPECTED_TOOL = {
    "name": "litchi-perf-baseline",
    "version": "0.1.0",
    "binary": "litchi-perf-baseline",
    "profile": "release",
    "target_os": "linux",
    "target_arch": "x86_64",
    "instrumentation": "none",
}


class AnalysisError(ValueError):
    """Raised when a formal capture is unsafe or does not match the protocol."""


def _reject_constant(value: str) -> Any:
    raise ValueError(f"non-finite JSON constant {value!r}")


def _reject_duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def load_json(path: Path) -> Any:
    try:
        with path.open(encoding="utf-8") as stream:
            return json.load(
                stream,
                object_pairs_hook=_reject_duplicate_keys,
                parse_constant=_reject_constant,
            )
    except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as error:
        raise AnalysisError(f"cannot read {path}: {error}") from error


def canonical_bytes(value: Any) -> bytes:
    try:
        return json.dumps(
            value, sort_keys=True, separators=(",", ":"), allow_nan=False
        ).encode("utf-8")
    except (TypeError, ValueError, OverflowError) as error:
        raise AnalysisError(f"cannot canonicalize JSON: {error}") from error


def canonical_sha256(value: Any) -> str:
    return hashlib.sha256(canonical_bytes(value)).hexdigest()


def raw_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError as error:
        raise AnalysisError(f"cannot hash {path}: {error}") from error
    return digest.hexdigest()


def _portable_repo_path(path: Path) -> str:
    """Render retained in-repository paths independently of the checkout root."""

    try:
        return path.resolve().relative_to(REPO_ROOT.resolve()).as_posix()
    except ValueError:
        return str(path)


def _object(value: Any, location: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise AnalysisError(f"{location} must be an object")
    return value


def _string(value: Any, location: str) -> str:
    if not isinstance(value, str) or not value:
        raise AnalysisError(f"{location} must be a non-empty string")
    return value


def _sha(value: Any, location: str) -> str:
    value = _string(value, location)
    if SHA256_RE.fullmatch(value) is None:
        raise AnalysisError(f"{location} must be a lowercase SHA-256")
    return value


def _revision(value: Any, location: str) -> str:
    value = _string(value, location)
    if REVISION_RE.fullmatch(value) is None:
        raise AnalysisError(f"{location} must be a 40-character revision")
    return value


def _positive_int(value: Any, location: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
        raise AnalysisError(f"{location} must be a positive integer")
    return value


def _exact_list(value: Any, expected: Sequence[Any], location: str) -> None:
    if value != list(expected):
        raise AnalysisError(f"{location} does not match {list(expected)!r}: {value!r}")


def _report_profile(report: Mapping[str, Any], label: str) -> None:
    tool = _object(report.get("tool"), f"{label}.tool")
    if tool != EXPECTED_TOOL:
        raise AnalysisError(f"{label}.tool identity differs from the formal harness")
    if report.get("schema_version") != 1:
        raise AnalysisError(f"{label}.schema_version must be 1")


def _protocol(path: Path) -> dict[str, Any]:
    protocol = _object(load_json(path), str(path))
    if protocol.get("schema") != PROTOCOL_SCHEMA:
        raise AnalysisError(f"{path} has unsupported protocol schema")
    if protocol.get("status") != "frozen":
        raise AnalysisError(f"{path}.status must be frozen")
    _exact_list(protocol.get("order"), LANE_NAMES, f"{path}.order")
    if protocol.get("samples") != SAMPLES or protocol.get("warmups") != WARMUPS:
        raise AnalysisError(f"{path} sample/warmup counts do not match 0467 normal protocol")
    if protocol.get("cpu") != 2 or protocol.get("workers") != WORKERS:
        raise AnalysisError(f"{path} CPU/worker configuration does not match 0467")
    _exact_list(protocol.get("cases"), CASES, f"{path}.cases")
    _exact_list(protocol.get("shapes"), SHAPES, f"{path}.shapes")
    triggers = _object(protocol.get("review_triggers"), f"{path}.review_triggers")
    if triggers.get("latency_percent") != REVIEW_THRESHOLD_PERCENT:
        raise AnalysisError(f"{path}.review_triggers.latency_percent must be 5")
    if triggers.get("process_rss_percent") != REVIEW_THRESHOLD_PERCENT:
        raise AnalysisError(f"{path}.review_triggers.process_rss_percent must be 5")
    return protocol


def _formal_lane_paths(root: Path) -> dict[str, Path]:
    """Find clean formal lane directories, never silently using diagnostics."""

    clean = {name: root / name for name in LANE_NAMES}
    if all(path.is_dir() for path in clean.values()):
        return clean
    missing = [str(path) for path in clean.values() if not path.is_dir()]
    raise AnalysisError(
        "formal normal lanes are incomplete; expected A1-clean/B1-clean/B2-clean/A2-clean "
        f"(missing {missing!r})"
    )


def _binding(root: Path, role: str) -> tuple[Path, dict[str, Any]]:
    path = root / f"{role}-binding.json"
    if not path.is_file():
        raise AnalysisError(f"missing {role} binding {path}")
    binding = _object(load_json(path), str(path))
    revision = _revision(binding.get("revision"), f"{path}.revision")
    binary = _sha(binding.get("binary_sha256"), f"{path}.binary_sha256")
    binary_bytes = _positive_int(binding.get("bytes"), f"{path}.bytes")
    build_receipt_name = _string(binding.get("build_receipt"), f"{path}.build_receipt")
    build_receipt_path = Path(build_receipt_name)
    if build_receipt_path.is_absolute() or ".." in build_receipt_path.parts:
        raise AnalysisError(f"{path}.build_receipt must be a safe bundle-relative path")
    build_receipt_path = path.parent / build_receipt_path
    build_receipt = _object(load_json(build_receipt_path), str(build_receipt_path))
    if raw_sha256(build_receipt_path) != _sha(
        binding.get("build_receipt_sha256"), f"{path}.build_receipt_sha256"
    ):
        raise AnalysisError(f"{path}.build_receipt_sha256 does not match the retained build receipt")
    if build_receipt.get("role") != role:
        raise AnalysisError(f"{build_receipt_path}.role does not match {role}")
    if _revision(build_receipt.get("revision"), f"{build_receipt_path}.revision") != revision:
        raise AnalysisError(f"{build_receipt_path}.revision does not match its binding")
    if build_receipt.get("exit_code") != 0:
        raise AnalysisError(f"{build_receipt_path}.exit_code must be zero")
    if build_receipt.get("clean_before") is not True or build_receipt.get("clean_after") is not True:
        raise AnalysisError(f"{build_receipt_path} does not prove a clean build tree")
    if binding.get("clean_build") is not True:
        raise AnalysisError(f"{path}.clean_build must be true")
    build_argv = build_receipt.get("argv")
    if not isinstance(build_argv, list) or build_argv[:2] != ["cargo", "build"]:
        raise AnalysisError(f"{build_receipt_path}.argv is not the authenticated cargo build")

    # Some future bundles may add an explicit source-manifest record.  It is
    # additive and verified when present, while the frozen 0467 binding uses
    # the clean build receipt plus the exact revision as its source binding.
    source_binding: dict[str, Any] = {
        "kind": "clean_build_receipt_revision",
        "revision": revision,
        "build_receipt_path": build_receipt_name,
        "build_receipt_sha256": raw_sha256(build_receipt_path),
        "build_receipt_canonical_sha256": canonical_sha256(build_receipt),
    }
    optional_source = binding.get("source")
    if optional_source is not None:
        source_object = _object(optional_source, f"{path}.source")
        source_name = _string(source_object.get("path"), f"{path}.source.path")
        source_path = Path(source_name)
        if source_path.is_absolute() or ".." in source_path.parts:
            raise AnalysisError(f"{path}.source.path must be bundle-relative")
        source_path = path.parent / source_path
        source_digest = _sha(source_object.get("sha256"), f"{path}.source.sha256")
        if raw_sha256(source_path) != source_digest:
            raise AnalysisError(f"{path}.source.sha256 does not match the retained source manifest")
        source_binding.update(
            {
                "kind": "explicit_source_manifest",
                "path": source_name,
                "sha256": source_digest,
                "files": _positive_int(source_object.get("files"), f"{path}.source.files"),
            }
        )
    optional_source_digest = binding.get("source_manifest_sha256")
    if optional_source_digest is not None:
        source_binding["source_manifest_sha256"] = _sha(
            optional_source_digest, f"{path}.source_manifest_sha256"
        )
    return path, {
        "revision": revision,
        "binary_sha256": binary,
        "bytes": binary_bytes,
        "build_receipt": build_receipt_name,
        "build_receipt_sha256": raw_sha256(build_receipt_path),
        "source_binding": source_binding,
        "binding_sha256": raw_sha256(path),
    }


def _validate_receipt(
    path: Path,
    lane: str,
    binding_path: Path,
    binding: Mapping[str, Any],
    *,
    role: str | None = None,
    samples: int = SAMPLES,
    warmups: int = WARMUPS,
    heaptrack: bool = False,
) -> dict[str, Any]:
    receipt = _object(load_json(path), str(path))
    if receipt.get("schema") != "litchi-0467-capture-v1":
        raise AnalysisError(f"{path}.schema is not the 0467 capture schema")
    expected_role = role if role is not None else LANE_ROLES[lane]
    if receipt.get("lane") != lane or receipt.get("role") != expected_role:
        raise AnalysisError(f"{path} lane/role identity mismatch")
    if receipt.get("revision") != binding["revision"]:
        raise AnalysisError(f"{path}.revision does not match its binding")
    if receipt.get("binary_sha256") != binding["binary_sha256"]:
        raise AnalysisError(f"{path}.binary_sha256 does not match its binding")
    if receipt.get("binding_sha256") != raw_sha256(binding_path):
        raise AnalysisError(f"{path}.binding_sha256 does not match its binding file")
    if not (
        receipt.get("clean_before") is True
        and receipt.get("clean_after") is True
        and receipt.get("binary_unchanged") is True
    ):
        raise AnalysisError(f"{path} does not prove clean before/after and binary immutability")
    if receipt.get("samples") != samples or receipt.get("warmups") != warmups:
        raise AnalysisError(f"{path} sample/warmup counts do not match the normal protocol")
    if receipt.get("report_metadata_matches_clean_role") is not True:
        raise AnalysisError(f"{path} does not prove clean report metadata")
    argv = receipt.get("argv")
    if not isinstance(argv, list) or argv[:3] != ["taskset", "-c", CPU]:
        raise AnalysisError(f"{path}.argv does not pin the formal run to CPU 2")
    has_heaptrack = "heaptrack" in argv
    if has_heaptrack != heaptrack:
        expected = "with" if heaptrack else "without"
        raise AnalysisError(f"{path}.argv must run {expected} Heaptrack")
    if "--workers" not in argv or argv[argv.index("--workers") + 1] != str(WORKERS):
        raise AnalysisError(f"{path}.argv does not use one worker")
    if not isinstance(receipt.get("exit_code"), int) or receipt["exit_code"] != 0:
        raise AnalysisError(f"{path} does not record a successful capture")
    return receipt


def _resource_rss(path: Path) -> dict[str, Any]:
    """Parse GNU ``time -v`` process RSS and bind the raw receipt."""

    try:
        text = path.read_text(encoding="utf-8")
    except OSError as error:
        raise AnalysisError(f"cannot read {path}: {error}") from error
    match = re.search(r"^\s*Maximum resident set size \(kbytes\):\s*(\d+)\s*$", text, re.MULTILINE)
    if match is None:
        raise AnalysisError(f"{path} has no maximum resident set size")
    value = int(match.group(1))
    if value <= 0:
        raise AnalysisError(f"{path} maximum resident set size must be positive")
    return {"path": path.name, "sha256": raw_sha256(path), "bytes": path.stat().st_size, "max_rss_kib": value}


def _catalog_file_identity(
    root: Path,
    lane_dir: Path,
    report: Mapping[str, Any],
    label: str,
) -> dict[str, Any]:
    """Validate a catalog and return custody plus corpus-only identities.

    ``catalog_sha256`` covers the catalog build metadata, so clean control and
    candidate catalogs necessarily differ when their revisions differ.  The
    cross-leg identity therefore excludes that build metadata and the two
    derived hashes while retaining the manifest identity, case bindings, and
    complete corpus items.  The full file/canonical hashes remain in the
    returned record for custody.
    """

    catalog_path = lane_dir / "corpus-catalog.json"
    catalog = _object(load_json(catalog_path), str(catalog_path))
    if catalog.get("manifest_kind") != "corpus-catalog" or catalog.get("manifest_version") != 2:
        raise AnalysisError(f"{catalog_path} is not corpus-catalog version 2")
    if catalog.get("catalog_id") != "litchi-perf-corpus-v2":
        raise AnalysisError(f"{catalog_path}.catalog_id is not litchi-perf-corpus-v2")
    canonicalization = _object(catalog.get("canonicalization"), f"{catalog_path}.canonicalization")
    if canonicalization != {"algorithm": "sorted-json-utf8-compact-v1", "hash": "sha256"}:
        raise AnalysisError(f"{catalog_path}.canonicalization is unsupported")
    catalog_sha = _sha(catalog.get("catalog_sha256"), f"{catalog_path}.catalog_sha256")
    content_set_sha = _sha(
        catalog.get("content_set_sha256"), f"{catalog_path}.content_set_sha256"
    )
    build = _object(catalog.get("build"), f"{catalog_path}.build")
    environment = _object(report.get("environment"), f"{label}.environment")
    report_revision = _revision(environment.get("git_revision"), f"{label}.environment.git_revision")
    catalog_revision = _revision(build.get("git_revision"), f"{catalog_path}.build.git_revision")
    if catalog_revision != report_revision:
        raise AnalysisError(f"{catalog_path}.build.git_revision does not match {label}")
    if build.get("git_worktree_dirty") is not False:
        raise AnalysisError(f"{catalog_path}.build.git_worktree_dirty must be false")
    if build.get("tool") != EXPECTED_TOOL["name"] or build.get("tool_version") != EXPECTED_TOOL["version"]:
        raise AnalysisError(f"{catalog_path}.build tool identity differs from the harness")
    corpora = catalog.get("corpora")
    case_bindings = catalog.get("case_bindings")
    if not isinstance(corpora, list) or not isinstance(case_bindings, list):
        raise AnalysisError(f"{catalog_path} corpus items/bindings must be arrays")

    # Recompute both producer-side derived values.  The content-set value is
    # the stable corpus identity; catalog_sha256 remains a per-build custody
    # identity because it includes the build object.
    content_corpora: list[dict[str, Any]] = []
    for corpus_index, raw_corpus in enumerate(corpora):
        corpus = _object(raw_corpus, f"{catalog_path}.corpora[{corpus_index}]")
        bytes_object = _object(
            corpus.get("bytes"), f"{catalog_path}.corpora[{corpus_index}].bytes"
        )
        members_object = _object(
            corpus.get("members"), f"{catalog_path}.corpora[{corpus_index}].members"
        )
        members = members_object.get("items")
        if not isinstance(members, list):
            raise AnalysisError(f"{catalog_path}.corpora[{corpus_index}].members.items must be an array")
        member_projection: list[dict[str, Any]] = []
        for member_index, raw_member in enumerate(members):
            member = _object(
                raw_member,
                f"{catalog_path}.corpora[{corpus_index}].members.items[{member_index}]",
            )
            ordinal = member.get("ordinal")
            if isinstance(ordinal, bool) or not isinstance(ordinal, int) or ordinal < 0:
                raise AnalysisError(f"{catalog_path}.corpora[{corpus_index}] member ordinal is invalid")
            member_projection.append(
                {
                    "ordinal": ordinal,
                    "name": _string(member.get("name"), f"{catalog_path}.member.name"),
                    "sha256": _sha(member.get("sha256"), f"{catalog_path}.member.sha256"),
                }
            )
        content_corpora.append(
            {
                "id": _string(corpus.get("id"), f"{catalog_path}.corpora[{corpus_index}].id"),
                "archive_sha256": _sha(
                    bytes_object.get("archive_sha256"),
                    f"{catalog_path}.corpora[{corpus_index}].bytes.archive_sha256",
                ),
                "members": member_projection,
            }
        )
    content_bindings: list[dict[str, Any]] = []
    for binding_index, raw_binding in enumerate(case_bindings):
        binding = _object(raw_binding, f"{catalog_path}.case_bindings[{binding_index}]")
        content_bindings.append(
            {
                "case": _string(binding.get("case"), f"{catalog_path}.case_bindings.case"),
                "corpus_id": _string(
                    binding.get("corpus_id"), f"{catalog_path}.case_bindings.corpus_id"
                ),
                "role": _string(binding.get("role"), f"{catalog_path}.case_bindings.role"),
            }
        )
    content_projection = {
        "corpora": content_corpora,
        "case_bindings": content_bindings,
    }
    if canonical_sha256(content_projection) != content_set_sha:
        raise AnalysisError(f"{catalog_path}.content_set_sha256 does not match its corpus items")
    catalog_without_hash = {key: value for key, value in catalog.items() if key != "catalog_sha256"}
    if canonical_sha256(catalog_without_hash) != catalog_sha:
        raise AnalysisError(f"{catalog_path}.catalog_sha256 does not match its contents")
    corpus_identity = {
        "manifest_kind": catalog["manifest_kind"],
        "manifest_version": catalog["manifest_version"],
        "catalog_id": catalog["catalog_id"],
        "canonicalization": canonicalization,
        "content_set_sha256": content_set_sha,
        "case_bindings": case_bindings,
        "corpora": corpora,
    }
    report_catalog = _object(report.get("corpus_catalog"), f"{label}.corpus_catalog")
    report_projection = {
        field: report_catalog.get(field)
        for field in ("catalog_id", "catalog_sha256", "content_set_sha256", "manifest_version")
    }
    expected_report_projection = {
        "catalog_id": catalog["catalog_id"],
        "catalog_sha256": catalog_sha,
        "content_set_sha256": content_set_sha,
        "manifest_version": catalog["manifest_version"],
    }
    if report_projection != expected_report_projection:
        raise AnalysisError(f"{label}.corpus_catalog does not match its retained catalog")
    return {
        "path": str(catalog_path.relative_to(root)),
        "sha256": raw_sha256(catalog_path),
        "canonical_sha256": canonical_sha256(catalog),
        "bytes": catalog_path.stat().st_size,
        "catalog_id": catalog["catalog_id"],
        "catalog_sha256": catalog_sha,
        "content_set_sha256": content_set_sha,
        "manifest_version": catalog["manifest_version"],
        "corpus_identity_sha256": canonical_sha256(corpus_identity),
    }


def _guard_statistics(elapsed: Mapping[str, Any], location: str, sample_count: int) -> dict[str, Any]:
    """Recompute guard percentiles, including the five-sample Heaptrack lane."""

    if sample_count >= perf_abba_summary.MIN_RETAINED_SAMPLES:
        return perf_abba_summary.recompute_statistics(elapsed, location)
    if elapsed.get("unit") != "ns" or not isinstance(elapsed.get("samples"), list):
        raise AnalysisError(f"{location} must contain ns samples")
    samples = elapsed["samples"]
    if len(samples) != sample_count or not samples:
        raise AnalysisError(f"{location}.samples cardinality is not {sample_count}")
    if any(isinstance(value, bool) or not isinstance(value, int) or value < 0 for value in samples):
        raise AnalysisError(f"{location}.samples must contain non-negative integers")
    if samples != sorted(samples):
        raise AnalysisError(f"{location}.samples must be sorted ascending")
    p50 = samples[(len(samples) - 1) // 2] // 2 + samples[len(samples) // 2] // 2
    p50 += (samples[(len(samples) - 1) // 2] % 2 + samples[len(samples) // 2] % 2) // 2
    p95 = samples[max(1, (95 * len(samples) + 99) // 100) - 1]
    p99 = samples[max(1, (99 * len(samples) + 99) // 100) - 1]
    computed = {"sample_count": sample_count, "min": samples[0], "p50": p50, "p95": p95, "p99": p99, "max": samples[-1], "mean": sum(samples) / len(samples)}
    for field in ("min", "p50", "p95", "p99", "max"):
        if elapsed.get(field) != computed[field]:
            raise AnalysisError(f"{location}.{field} disagrees with samples")
    reported_mean = elapsed.get("mean")
    if not isinstance(reported_mean, (int, float)) or not math.isfinite(float(reported_mean)):
        raise AnalysisError(f"{location}.mean must be finite")
    if abs(float(reported_mean) - computed["mean"]) > max(1e-12, abs(computed["mean"]) * 1e-12):
        raise AnalysisError(f"{location}.mean disagrees with samples")
    return computed


def _validate_report(report: Mapping[str, Any], label: str, binding: Mapping[str, Any]) -> dict[tuple[str, str], dict[str, Any]]:
    _report_profile(report, label)
    environment = _object(report.get("environment"), f"{label}.environment")
    revision = _revision(environment.get("git_revision"), f"{label}.environment.git_revision")
    if revision != binding["revision"]:
        raise AnalysisError(f"{label}.environment.git_revision does not match its binding")
    if environment.get("git_worktree_dirty") is not False:
        raise AnalysisError(f"{label}.environment.git_worktree_dirty must be false for formal ABBA")
    if environment.get("cpu_affinity") != CPU:
        raise AnalysisError(f"{label}.environment.cpu_affinity must be CPU 2")
    configuration = _object(report.get("configuration"), f"{label}.configuration")
    if configuration.get("samples_per_case") != SAMPLES:
        raise AnalysisError(f"{label}.configuration.samples_per_case must be 100")
    if configuration.get("warmup_iterations_per_case") != WARMUPS:
        raise AnalysisError(f"{label}.configuration.warmup_iterations_per_case must be 5")
    _exact_list(configuration.get("cases"), CASES, f"{label}.configuration.cases")
    _exact_list(configuration.get("xlsx_shapes"), SHAPES, f"{label}.configuration.xlsx_shapes")
    workers = configuration.get("execution_workers")
    if workers != [WORKERS]:
        raise AnalysisError(f"{label}.configuration.execution_workers must be [1]")
    binary = _object(report.get("binary_identity"), f"{label}.binary_identity")
    if _sha(binary.get("binary_sha256"), f"{label}.binary_identity.binary_sha256") != binding["binary_sha256"]:
        raise AnalysisError(f"{label}.binary_identity.binary_sha256 does not match its binding")
    if binary.get("binary_bytes") != binding["bytes"]:
        raise AnalysisError(f"{label}.binary_identity.binary_bytes does not match its binding")
    results = report.get("results")
    if not isinstance(results, list) or len(results) != len(CASES) * len(SHAPES):
        raise AnalysisError(f"{label}.results must contain exactly six rows")
    indexed: dict[tuple[str, str], dict[str, Any]] = {}
    for index, row_value in enumerate(results):
        row = _object(row_value, f"{label}.results[{index}]")
        case = _string(row.get("case"), f"{label}.results[{index}].case")
        corpus = _object(row.get("corpus"), f"{label}.results[{index}].corpus")
        shape = _string(corpus.get("shape"), f"{label}.results[{index}].corpus.shape")
        if case not in CASES or shape not in SHAPES:
            raise AnalysisError(f"{label}.results[{index}] is outside the formal case/shape matrix")
        key = (case, canonical_sha256(corpus))
        if key in indexed:
            raise AnalysisError(f"{label} contains duplicate {case}/{shape} corpus identity")
        elapsed = _object(row.get("elapsed_ns"), f"{label}.{case}[{shape}].elapsed_ns")
        if not isinstance(elapsed.get("samples"), list) or len(elapsed["samples"]) != SAMPLES:
            raise AnalysisError(f"{label}.{case}[{shape}] must contain 100 elapsed samples")
        try:
            checked = perf_abba_summary.recompute_statistics(
                elapsed, f"{label}.{case}[{shape}].elapsed_ns"
            )
        except Exception as error:
            raise AnalysisError(f"{label}.{case}[{shape}] statistic validation failed: {error}") from error
        indexed[key] = {"row": row, "corpus": corpus, "statistics": checked}
    if {(case, _string(row["corpus"].get("shape"), "shape")) for (case, _identity), row in indexed.items()} != set((case, shape) for case in CASES for shape in SHAPES):
        raise AnalysisError(f"{label} does not cover every case/shape exactly once")
    return indexed


def _validate_guard_report(
    report: Mapping[str, Any],
    label: str,
    binding: Mapping[str, Any],
    *,
    samples: int,
    warmups: int,
    expected_rows: int,
    expected_cases: Sequence[str] | None = None,
    expected_shapes: Sequence[str] | None = None,
) -> dict[tuple[str, str], dict[str, Any]]:
    """Validate one full/default or instrumented guard report.

    This intentionally does not compare latency.  Full normal guards are
    compared by the existing default-matrix comparator; Heaptrack guards are
    identity/semantics evidence only because instrumentation changes timing.
    """

    _report_profile(report, label)
    environment = _object(report.get("environment"), f"{label}.environment")
    revision = _revision(environment.get("git_revision"), f"{label}.environment.git_revision")
    if revision != binding["revision"]:
        raise AnalysisError(f"{label}.environment.git_revision does not match its binding")
    if environment.get("git_worktree_dirty") is not False:
        raise AnalysisError(f"{label}.environment.git_worktree_dirty must be false")
    if environment.get("cpu_affinity") != CPU:
        raise AnalysisError(f"{label}.environment.cpu_affinity must be CPU 2")
    configuration = _object(report.get("configuration"), f"{label}.configuration")
    if configuration.get("samples_per_case") != samples:
        raise AnalysisError(f"{label}.configuration.samples_per_case must be {samples}")
    if configuration.get("warmup_iterations_per_case") != warmups:
        raise AnalysisError(f"{label}.configuration.warmup_iterations_per_case must be {warmups}")
    if expected_cases is not None:
        _exact_list(configuration.get("cases"), expected_cases, f"{label}.configuration.cases")
    if expected_shapes is not None:
        _exact_list(configuration.get("xlsx_shapes"), expected_shapes, f"{label}.configuration.xlsx_shapes")
    workers = configuration.get("execution_workers")
    if workers != [WORKERS]:
        raise AnalysisError(f"{label}.configuration.execution_workers must be [1]")
    binary = _object(report.get("binary_identity"), f"{label}.binary_identity")
    if _sha(binary.get("binary_sha256"), f"{label}.binary_identity.binary_sha256") != binding["binary_sha256"]:
        raise AnalysisError(f"{label}.binary_identity.binary_sha256 does not match its binding")
    if binary.get("binary_bytes") != binding["bytes"]:
        raise AnalysisError(f"{label}.binary_identity.binary_bytes does not match its binding")
    results = report.get("results")
    if not isinstance(results, list) or len(results) != expected_rows:
        raise AnalysisError(f"{label}.results must contain exactly {expected_rows} rows")
    indexed: dict[tuple[str, str], dict[str, Any]] = {}
    for index, raw_row in enumerate(results):
        row = _object(raw_row, f"{label}.results[{index}]")
        case = _string(row.get("case"), f"{label}.results[{index}].case")
        corpus = _object(row.get("corpus"), f"{label}.results[{index}].corpus")
        shape = _string(corpus.get("shape"), f"{label}.results[{index}].corpus.shape")
        if expected_cases is not None and case not in expected_cases:
            raise AnalysisError(f"{label}.results[{index}] has an unexpected case")
        if expected_shapes is not None and shape not in expected_shapes:
            raise AnalysisError(f"{label}.results[{index}] has an unexpected shape")
        key = (case, canonical_sha256(corpus))
        if key in indexed:
            raise AnalysisError(f"{label}.results contains duplicate case/corpus identity")
        elapsed = _object(row.get("elapsed_ns"), f"{label}.results[{index}].elapsed_ns")
        if not isinstance(elapsed.get("samples"), list) or len(elapsed["samples"]) != samples:
            raise AnalysisError(f"{label}.{case}[{shape}] elapsed sample count is not {samples}")
        try:
            statistics = _guard_statistics(
                elapsed, f"{label}.{case}[{shape}].elapsed_ns", samples
            )
        except Exception as error:
            raise AnalysisError(f"{label}.{case}[{shape}] statistic validation failed: {error}") from error
        indexed[key] = {"row": row, "corpus": corpus, "statistics": statistics}
    return indexed


def _ratio_percent(baseline: float, current: float, location: str) -> dict[str, float | bool | None]:
    if not math.isfinite(baseline) or not math.isfinite(current) or baseline <= 0 or current <= 0:
        raise AnalysisError(f"{location} requires positive finite values")
    ratio = current / baseline
    delta = (ratio - 1.0) * 100.0
    reduction = -delta
    return {
        "baseline": baseline,
        "current": current,
        "ratio_current_over_baseline": ratio,
        "delta_percent": delta,
        "reduction_percent": reduction,
        "regression_over_5_percent": delta > REVIEW_THRESHOLD_PERCENT,
        "improvement_over_5_percent": reduction > REVIEW_THRESHOLD_PERCENT,
    }


def _drift(first: float, second: float, location: str) -> dict[str, float | bool]:
    comparison = _ratio_percent(first, second, location)
    return {
        "first": first,
        "second": second,
        "delta_percent": comparison["delta_percent"],
        "absolute_delta_percent": abs(comparison["delta_percent"]),
        "review_triggered": abs(comparison["delta_percent"]) > REVIEW_THRESHOLD_PERCENT,
    }


def _row_result(
    key: tuple[str, str],
    indexed: Mapping[str, Mapping[tuple[str, str], Mapping[str, Any]]],
    rss: Mapping[str, Mapping[str, Any]],
) -> dict[str, Any]:
    case, identity = key
    rows = {lane: indexed[lane][key] for lane in LANES}
    corpus = rows["a1"]["corpus"]
    if any(canonical_sha256(item["corpus"]) != identity for item in rows.values()):
        raise AnalysisError(f"{case} corpus identity differs between ABBA lanes")
    stats = {
        lane: {name: rows[lane]["statistics"][name] for name in STATISTICS}
        for lane in LANES
    }
    pairings = {
        "a1_control_to_b1_candidate": {
            name: _ratio_percent(stats["a1"][name], stats["b1"][name], f"{case}.{name}.a1-b1")
            for name in STATISTICS
        },
        "a2_control_to_b2_candidate": {
            name: _ratio_percent(stats["a2"][name], stats["b2"][name], f"{case}.{name}.a2-b2")
            for name in STATISTICS
        },
    }
    drift = {
        "control_a1_to_a2": {
            name: _drift(stats["a1"][name], stats["a2"][name], f"{case}.{name}.a1-a2")
            for name in STATISTICS
        },
        "candidate_b1_to_b2": {
            name: _drift(stats["b1"][name], stats["b2"][name], f"{case}.{name}.b1-b2")
            for name in STATISTICS
        },
    }
    rss_pairings = {
        "a1_control_to_b1_candidate": _ratio_percent(
            rss["a1"]["max_rss_kib"], rss["b1"]["max_rss_kib"], f"{case}.rss.a1-b1"
        ),
        "a2_control_to_b2_candidate": _ratio_percent(
            rss["a2"]["max_rss_kib"], rss["b2"]["max_rss_kib"], f"{case}.rss.a2-b2"
        ),
    }
    rss_drift = {
        "control_a1_to_a2": _drift(rss["a1"]["max_rss_kib"], rss["a2"]["max_rss_kib"], f"{case}.rss.a1-a2"),
        "candidate_b1_to_b2": _drift(rss["b1"]["max_rss_kib"], rss["b2"]["max_rss_kib"], f"{case}.rss.b1-b2"),
    }
    latency_review = {
        "pair_regression_triggered": any(
            value["regression_over_5_percent"]
            for pair in pairings.values()
            for value in pair.values()
        ),
        "same_role_drift_triggered": any(
            value["review_triggered"]
            for pair in drift.values()
            for value in pair.values()
        ),
    }
    rss_review = {
        "pair_regression_triggered": any(
            value["regression_over_5_percent"] for value in rss_pairings.values()
        ),
        "same_role_drift_triggered": any(value["review_triggered"] for value in rss_drift.values()),
    }
    return {
        "case": case,
        "shape": corpus.get("shape"),
        "corpus": corpus,
        "statistics_ns": stats,
        "latency_pairings": pairings,
        "latency_same_role_drift": drift,
        "process_rss_kib": {lane: rss[lane]["max_rss_kib"] for lane in LANES},
        "rss_scope": "whole process per lane; repeated for row context, not row-local RSS",
        "rss_pairings": rss_pairings,
        "rss_same_role_drift": rss_drift,
        "review_triggers": {
            "threshold_percent": REVIEW_THRESHOLD_PERCENT,
            "latency": latency_review,
            "process_rss": rss_review,
        },
    }


def _guard_lane_role(lane: str) -> str:
    if lane.startswith("A-"):
        return "control"
    if lane.startswith("B-"):
        return "candidate"
    raise AnalysisError(f"unsupported guard lane {lane!r}")


def _load_guard_lane(
    root: Path,
    lane: str,
    *,
    samples: int,
    warmups: int,
    expected_rows: int,
    expected_cases: Sequence[str] | None = None,
    expected_shapes: Sequence[str] | None = None,
    heaptrack: bool = False,
) -> dict[str, Any]:
    lane_dir = root / lane
    if not lane_dir.is_dir():
        raise AnalysisError(f"missing guard lane directory {lane_dir}")
    role = _guard_lane_role(lane)
    binding_path, binding = _binding(root, role)
    receipt = _validate_receipt(
        lane_dir / "receipt.json",
        lane,
        binding_path,
        binding,
        role=role,
        samples=samples,
        warmups=warmups,
        heaptrack=heaptrack,
    )
    report_path = lane_dir / "report.json"
    report = _object(load_json(report_path), str(report_path))
    indexed = _validate_guard_report(
        report,
        lane,
        binding,
        samples=samples,
        warmups=warmups,
        expected_rows=expected_rows,
        expected_cases=expected_cases,
        expected_shapes=expected_shapes,
    )
    resource = _resource_rss(lane_dir / "resource.log")
    catalog = _catalog_file_identity(root, lane_dir, report, lane)
    artifacts = {
        "report": {
            "path": report_path.name,
            "sha256": raw_sha256(report_path),
            "canonical_sha256": canonical_sha256(report),
            "bytes": report_path.stat().st_size,
        },
        "catalog": catalog,
        "resource": resource,
    }
    if heaptrack:
        heap_files = [
            item
            for item in sorted(lane_dir.rglob("*"))
            if item.is_file() and item.name.startswith("heaptrack")
        ]
        if not heap_files:
            raise AnalysisError(f"{lane_dir} has no retained Heaptrack artifact")
        artifacts["heaptrack"] = [
            {
                "path": str(item.relative_to(lane_dir)),
                "sha256": raw_sha256(item),
                "bytes": item.stat().st_size,
            }
            for item in heap_files
        ]
    return {
        "lane": lane,
        "role": role,
        "binding": binding,
        "binding_path": str(binding_path.relative_to(root)),
        "receipt": receipt,
        "report": report,
        "indexed": indexed,
        "artifacts": artifacts,
    }


def _catalog_identity(guard: Mapping[str, Any]) -> str:
    # The full catalog digest includes the clean build revision.  Control and
    # candidate therefore have distinct custody digests even when they use the
    # same corpus.  Guard comparisons need the stable corpus projection.
    return guard["artifacts"]["catalog"]["corpus_identity_sha256"]


def _full_guard_latency_review(
    control: Mapping[str, Any], candidate: Mapping[str, Any]
) -> dict[str, Any]:
    """Compare every full-guard latency statistic with the 5% review rule.

    ``perf_compare`` is the canonical policy result and intentionally compares
    only the policy's p50/p95/p99 cells at their configured thresholds.  The
    protocol's review threshold is a separate descriptive layer: it includes
    the mean and records every row so an adverse result cannot disappear into
    the default regression list.
    """

    control_index = control["indexed"]
    candidate_index = candidate["indexed"]
    if set(control_index) != set(candidate_index):
        raise AnalysisError("full guard latency row identities differ")
    rows: list[dict[str, Any]] = []
    for key in sorted(control_index):
        control_row = control_index[key]
        candidate_row = candidate_index[key]
        if canonical_sha256(control_row["corpus"]) != canonical_sha256(candidate_row["corpus"]):
            raise AnalysisError(f"full guard corpus differs for {key[0]}")
        case, _identity = key
        corpus = control_row["corpus"]
        for statistic in STATISTICS:
            values = _ratio_percent(
                control_row["statistics"][statistic],
                candidate_row["statistics"][statistic],
                f"full_guard.{case}.{statistic}",
            )
            rows.append(
                {
                    "case": case,
                    "corpus": corpus,
                    "metric": f"elapsed_ns.{statistic}",
                    "metric_class": "latency",
                    "statistic": statistic,
                    "baseline": values["baseline"],
                    "current": values["current"],
                    "ratio_current_over_baseline": values["ratio_current_over_baseline"],
                    "delta_percent": values["delta_percent"],
                    "reduction_percent": values["reduction_percent"],
                    "review_threshold_percent": REVIEW_THRESHOLD_PERCENT,
                    "review_triggered": values["regression_over_5_percent"],
                    "regression_over_5_percent": values["regression_over_5_percent"],
                    "improvement_over_5_percent": values["improvement_over_5_percent"],
                }
            )
    triggered = [row for row in rows if row["review_triggered"]]
    return {
        "threshold_percent": REVIEW_THRESHOLD_PERCENT,
        "statistics": list(STATISTICS),
        "compared_rows": len(rows),
        "triggered_rows": triggered,
        "triggered_count": len(triggered),
        "triggered": bool(triggered),
        "rows": rows,
    }


def _comparison_report(report: Mapping[str, Any]) -> dict[str, Any]:
    """Normalize absent optional work vectors for the shared comparator.

    Some full-matrix cases legitimately emit an empty top-level ``source``
    vector (for example ``zip_index`` has no source reads).  The policy marks
    those metrics optional; treating an empty vector as an absent optional
    field preserves the report and lets the comparator enforce that both roles
    have the same presence.  No samples or measured values are changed.
    """

    normalized = copy.deepcopy(dict(report))
    results = normalized.get("results")
    if not isinstance(results, list):
        raise AnalysisError("full guard reports must contain a results list")

    def remove_empty_vectors(value: Any) -> Any:
        if isinstance(value, dict):
            return {
                key: remove_empty_vectors(item)
                for key, item in value.items()
                if not (isinstance(item, list) and not item)
            }
        if isinstance(value, list):
            return [remove_empty_vectors(item) for item in value]
        return value

    for index, row in enumerate(results):
        if not isinstance(row, dict):
            raise AnalysisError(f"full guard result {index} is not an object")
        if "source" in row:
            row["source"] = remove_empty_vectors(row["source"])
    return normalized


def analyze_full_guard(
    root: Path = ROOT,
    *,
    policy_path: Path | None = None,
) -> dict[str, Any]:
    """Validate and compare the two normal 15/3 full default-matrix guards."""

    root = root.resolve()
    if policy_path is None:
        policy_path = REPO_ROOT / "docs/performance/perf-regression-policy-v1.json"
    policy = _object(load_json(policy_path), str(policy_path))
    control = _load_guard_lane(
        root,
        FULL_GUARD_LANES[0],
        samples=FULL_GUARD_SAMPLES,
        warmups=FULL_GUARD_WARMUPS,
        expected_rows=201,
    )
    candidate = _load_guard_lane(
        root,
        FULL_GUARD_LANES[1],
        samples=FULL_GUARD_SAMPLES,
        warmups=FULL_GUARD_WARMUPS,
        expected_rows=201,
    )
    if _catalog_identity(control) != _catalog_identity(candidate):
        raise AnalysisError("full guard corpus catalogs differ")
    try:
        comparison = perf_compare.compare_reports(
            _comparison_report(control["report"]),
            _comparison_report(candidate["report"]),
            policy,
        )
    except Exception as error:
        raise AnalysisError(f"full default guard comparison failed: {error}") from error
    comparison_rows = comparison.get("comparisons")
    if not isinstance(comparison_rows, list):
        raise AnalysisError("full default guard comparison did not return comparison rows")
    latency_review = _full_guard_latency_review(control, candidate)
    rss_pair = _ratio_percent(
        control["artifacts"]["resource"]["max_rss_kib"],
        candidate["artifacts"]["resource"]["max_rss_kib"],
        "full_guard.process_rss",
    )
    latency_regressions = [
        item
        for item in comparison.get("comparisons", [])
        if item.get("metric_class") == "latency" and item.get("regression") is True
    ]
    return {
        "schema": "litchi-0467-full-default-guard-v1",
        "scope": "normal full default 201-row matrix, 15 samples/3 warmups per row",
        "protocol": {
            "samples": FULL_GUARD_SAMPLES,
            "warmups": FULL_GUARD_WARMUPS,
            "cpu": 2,
            "workers": WORKERS,
            "policy_path": _portable_repo_path(policy_path),
            "policy_sha256": raw_sha256(policy_path),
        },
        "control": {
            "lane": control["lane"],
            "binding": control["binding"],
            "binding_path": control["binding_path"],
            "artifacts": control["artifacts"],
        },
        "candidate": {
            "lane": candidate["lane"],
            "binding": candidate["binding"],
            "binding_path": candidate["binding_path"],
            "artifacts": candidate["artifacts"],
        },
        # Preserve the complete canonical comparator result, including its
        # policy thresholds and every comparison row.  The separate review
        # layer below must not rewrite this policy result.
        "comparison": comparison,
        "process_rss": {
            "scope": "whole process per full guard lane",
            "control_kib": control["artifacts"]["resource"]["max_rss_kib"],
            "candidate_kib": candidate["artifacts"]["resource"]["max_rss_kib"],
            "pair": rss_pair,
            "review_triggered": rss_pair["regression_over_5_percent"],
        },
        "review_triggers": {
            "latency_threshold_percent": REVIEW_THRESHOLD_PERCENT,
            "process_rss_threshold_percent": REVIEW_THRESHOLD_PERCENT,
            "latency_regression_count": len(latency_regressions),
            "latency_regression_triggered": bool(latency_regressions),
            "latency_review_5_percent": latency_review,
            "latency_review_5_percent_count": latency_review["triggered_count"],
            "latency_review_5_percent_triggered": latency_review["triggered"],
            "process_rss_regression_triggered": rss_pair["regression_over_5_percent"],
        },
        "verification": {
            "clean_worktrees_verified": True,
            "distinct_revisions_verified": control["binding"]["revision"] != candidate["binding"]["revision"],
            "distinct_binaries_verified": control["binding"]["binary_sha256"] != candidate["binding"]["binary_sha256"],
            "build_receipt_bindings_verified": True,
            "catalog_identity_verified": True,
            "default_policy_comparison_verified": True,
            "instrumented_latency_not_compared": True,
        },
    }


def analyze_heap_guard(root: Path = ROOT) -> dict[str, Any]:
    """Validate the separate 5/1 Heaptrack lanes without comparing latency."""

    root = root.resolve()
    guards = [
        _load_guard_lane(
            root,
            lane,
            samples=HEAP_GUARD_SAMPLES,
            warmups=HEAP_GUARD_WARMUPS,
            expected_rows=1,
            expected_cases=("xlsx_one_percent_commit_save",),
            expected_shapes=("dense-wide",),
            heaptrack=True,
        )
        for lane in HEAP_GUARD_LANES
    ]
    if _catalog_identity(guards[0]) != _catalog_identity(guards[1]):
        raise AnalysisError("Heaptrack guard corpus catalogs differ")
    return {
        "schema": "litchi-0467-heaptrack-guard-v1",
        "scope": "whole-process Heaptrack dense-wide one-percent commit/save, 5 samples/1 warmup",
        "lanes": [
            {
                "lane": guard["lane"],
                "role": guard["role"],
                "binding": guard["binding"],
                "binding_path": guard["binding_path"],
                "report": guard["artifacts"]["report"],
                "catalog": guard["artifacts"]["catalog"],
                "resource": guard["artifacts"]["resource"],
                "heaptrack": guard["artifacts"]["heaptrack"],
            }
            for guard in guards
        ],
        "verification": {
            "clean_worktrees_verified": True,
            "distinct_revisions_verified": guards[0]["binding"]["revision"] != guards[1]["binding"]["revision"],
            "distinct_binaries_verified": guards[0]["binding"]["binary_sha256"] != guards[1]["binding"]["binary_sha256"],
            "build_receipt_bindings_verified": True,
            "catalog_identity_verified": True,
            "heaptrack_artifacts_verified": True,
            "latency_comparison_forbidden": True,
        },
    }


def analyze(
    root: Path = ROOT,
    *,
    protocol_path: Path | None = None,
    lane_paths: Mapping[str, Path] | None = None,
) -> dict[str, Any]:
    root = root.resolve()
    if protocol_path is None:
        candidate = root / "protocol-r1.json"
        protocol_path = candidate if candidate.is_file() else root / "protocol.json"
    protocol = _protocol(protocol_path)
    lanes = dict(lane_paths or _formal_lane_paths(root))
    if set(lanes) != set(LANE_NAMES):
        raise AnalysisError(f"lane set must be {list(LANE_NAMES)!r}")
    bindings = {role: _binding(root, role) for role in ("control", "candidate")}
    binding_values = {role: item[1] for role, item in bindings.items()}
    if binding_values["control"]["revision"] == binding_values["candidate"]["revision"]:
        raise AnalysisError("control and candidate revisions must be distinct")
    if binding_values["control"]["binary_sha256"] == binding_values["candidate"]["binary_sha256"]:
        raise AnalysisError("control and candidate binaries must be distinct")
    loaded_reports: dict[str, dict[str, Any]] = {}
    indexed: dict[str, dict[tuple[str, str], dict[str, Any]]] = {}
    receipts: dict[str, dict[str, Any]] = {}
    resources: dict[str, dict[str, Any]] = {}
    for lane in LANE_NAMES:
        lane_dir = lanes[lane]
        if not lane_dir.is_dir():
            raise AnalysisError(f"missing formal lane directory {lane_dir}")
        role = LANE_ROLES[lane]
        binding_path, binding = bindings[role]
        receipt_path = lane_dir / "receipt.json"
        report_path = lane_dir / "report.json"
        resource_path = lane_dir / "resource.log"
        lane_key = _lane_key(lane)
        receipts[lane_key] = _validate_receipt(receipt_path, lane, binding_path, binding)
        report = _object(load_json(report_path), str(report_path))
        loaded_reports[lane_key] = report
        indexed[lane_key] = _validate_report(report, lane, binding)
        resources[lane_key] = _resource_rss(resource_path)
    first_keys = set(indexed["a1"])
    if len(first_keys) != len(CASES) * len(SHAPES):
        raise AnalysisError("A1 does not contain exactly six formal case/corpus rows")
    if any(set(indexed[lane]) != first_keys for lane in LANES):
        raise AnalysisError("case/corpus identities differ between ABBA lanes")
    corpora = [indexed["a1"][key]["corpus"] for key in sorted(first_keys)]
    catalogs: dict[str, dict[str, Any]] = {}
    for lane in LANE_NAMES:
        lane_key = _lane_key(lane)
        catalogs[lane_key] = _catalog_file_identity(
            root, lanes[lane], loaded_reports[lane_key], lane
        )
    catalog_hashes = {item["corpus_identity_sha256"] for item in catalogs.values()}
    if len(catalog_hashes) != 1:
        raise AnalysisError("corpus catalog content differs between ABBA lanes")
    report_values = [loaded_reports[lane] for lane in LANES]
    try:
        canonical_summary = perf_abba_summary.summarize_reports(
            report_values,
            cases=CASES,
            shapes=SHAPES,
        )
    except Exception as error:
        raise AnalysisError(f"canonical ABBA summary validation failed: {error}") from error
    rows = [
        _row_result(key, indexed, resources)
        for key in sorted(first_keys, key=lambda value: (value[0], indexed["a1"][value]["corpus"].get("shape", "")))
    ]
    lane_records = {}
    for lane in LANE_NAMES:
        lane_key = _lane_key(lane)
        lane_records[lane] = {
            "role": LANE_ROLES[lane],
            "path": str(lanes[lane].relative_to(root)),
            "report_sha256": raw_sha256(lanes[lane] / "report.json"),
            "report_canonical_sha256": canonical_sha256(loaded_reports[lane_key]),
            "receipt_sha256": raw_sha256(lanes[lane] / "receipt.json"),
            "receipt": receipts[lane_key],
            "resource": resources[lane_key],
            "catalog": catalogs[lane_key],
        }
    return {
        "schema": SCHEMA,
        "purpose": "descriptive clean normal XLSX ABBA comparison; no automatic speedup or resource claim",
        "protocol": {
            "path": protocol_path.name,
            "sha256": raw_sha256(protocol_path),
            "canonical_sha256": canonical_sha256(protocol),
            "schema": protocol.get("schema"),
            "order": list(LANE_NAMES),
            "samples": SAMPLES,
            "warmups": WARMUPS,
            "cpu": 2,
            "workers": WORKERS,
            "review_trigger_percent": REVIEW_THRESHOLD_PERCENT,
        },
        "bindings": {
            role: {
                **binding,
                "path": str(path.relative_to(root)),
            }
            for role, (path, binding) in bindings.items()
        },
        "lanes": lane_records,
        "corpora": corpora,
        "results": rows,
        "abba_summary": canonical_summary,
        "claim_registration": {
            "registry_path": "docs/performance/claim-registry-v1.json",
            "latency_policy": "latency-abba-v1",
            "registry_minimum_samples_per_case": 500,
            "formal_samples_per_case": SAMPLES,
            "strict_latency_claim_eligible": False,
            "reason_codes": [
                "below_registry_minimum_samples",
                "resource_guardrail_is_separate_evidence",
            ],
            "required_for_strict_registration": [
                "four clean ABBA reports at or above the registry sample minimum",
                "canonical summary recomputation from retained raw reports",
                "scoped selectors and exact corpus identities",
                "separate resource guardrail evidence when the landed claim requires it",
            ],
            "minimal_strict500_followup": {
                "lane_order": ["A1-500-clean", "B1-500-clean", "B2-500-clean", "A2-500-clean"],
                "samples_per_case": 500,
                "warmups_per_case": WARMUPS,
                "cpu": 2,
                "workers": WORKERS,
                "scope": {
                    "case": "xlsx_one_percent_commit_save",
                    "shape": "dense-wide",
                },
                "retain_initial_six_descriptive_rows": True,
                "qualification_is_conditional_on_accepted_cells_and_guards": True,
            },
            "current_disposition": "descriptive evidence only; do not register as a strict landed claim",
        },
        "verification": {
            "clean_worktree_verified": True,
            "distinct_revisions_verified": True,
            "distinct_binaries_verified": True,
            "binding_identity_verified": True,
            "source_binding_verified_via_build_receipts": True,
            "explicit_source_manifest_present": all(
                item[1]["source_binding"]["kind"] == "explicit_source_manifest"
                for item in bindings.values()
            ),
            "configuration_identity_verified": True,
            "corpus_identity_verified": True,
            "catalog_identity_verified": True,
            "statistics_recomputed_from_100_samples": True,
            "canonical_abba_summary_verified": True,
            "rss_process_total_only": True,
            "review_trigger_is_not_claim": True,
        },
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--protocol", type=Path)
    parser.add_argument("--json-out", type=Path)
    parser.add_argument("--full-guard-out", type=Path)
    parser.add_argument("--heap-guard-out", type=Path)
    for lane in LANE_NAMES:
        parser.add_argument(
            f"--{lane}",
            dest=f"{_lane_key(lane)}_path",
            type=Path,
            help=f"{lane} formal lane directory",
        )
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    args = build_parser().parse_args(list(argv) if argv is not None else None)
    try:
        root = args.root.resolve()
        supplied = {
            lane: getattr(args, f"{_lane_key(lane)}_path")
            for lane in LANE_NAMES
            if getattr(args, f"{_lane_key(lane)}_path") is not None
        }
        if supplied and len(supplied) != len(LANE_NAMES):
            raise AnalysisError("supply all four lane paths or none")
        result = analyze(
            root,
            protocol_path=args.protocol.resolve() if args.protocol else None,
            lane_paths={lane: path.resolve() for lane, path in supplied.items()} or None,
        )
        if args.json_out is not None:
            args.json_out.parent.mkdir(parents=True, exist_ok=True)
            with args.json_out.open("x", encoding="utf-8") as stream:
                json.dump(result, stream, indent=2, sort_keys=True, allow_nan=False)
                stream.write("\n")
        if args.full_guard_out is not None:
            full_guard = analyze_full_guard(root)
            args.full_guard_out.parent.mkdir(parents=True, exist_ok=True)
            with args.full_guard_out.open("x", encoding="utf-8") as stream:
                json.dump(full_guard, stream, indent=2, sort_keys=True, allow_nan=False)
                stream.write("\n")
        if args.heap_guard_out is not None:
            heap_guard = analyze_heap_guard(root)
            args.heap_guard_out.parent.mkdir(parents=True, exist_ok=True)
            with args.heap_guard_out.open("x", encoding="utf-8") as stream:
                json.dump(heap_guard, stream, indent=2, sort_keys=True, allow_nan=False)
                stream.write("\n")
        json.dump(result, sys.stdout, indent=2, sort_keys=True, allow_nan=False)
        sys.stdout.write("\n")
        return 0
    except (AnalysisError, OSError, ValueError) as error:
        print(f"{SCHEMA}: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
