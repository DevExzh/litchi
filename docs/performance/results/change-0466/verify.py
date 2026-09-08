#!/usr/bin/env python3
"""Portable, read-only verifier for the 0466 XLSX profiling bundle.

This verifier authenticates the capture receipts and their raw artifacts.  It
also checks the one timed case and corpus identity, and recomputes the
reported latency statistics.  It does not build, run, or post-process a
profile.  The prior 0465 binding/source archive is used only as the recorded
build provenance for the reused normal binary.
"""

from __future__ import annotations

import argparse
import gzip
import hashlib
import importlib.util
import json
import math
from pathlib import Path
import re
import statistics
import sys
from typing import Any, Mapping


ROOT = Path(__file__).resolve().parent
CHANGE = 466
SCHEMA = "litchi-0466-verification-v1"
SHA256_RE = re.compile(r"[0-9a-f]{64}\Z")
CASE = "xlsx_one_percent_commit_save"
CORPUS_SHA256 = "5dd3ad701eb686f6d2d14e9f177a4e9433445728b57b484d53f663b2f87a7714"
CATALOG_SHA256 = "93de3286ae2c598752f38ec235392276d31af5b16895e9e075199eae5ac5c6f2"
CONTENT_SET_SHA256 = "66889d1124b43cd8ec4580ba0675509d97419df2d10e0d8a6130a3a509076936"
LANES = {
    "normal-r1": (30, 3),
    "normal-r2": (30, 3),
    "normal-r3": (30, 3),
    "normal-r4": (30, 3),
    "counters": (30, 3),
    "samples": (50, 3),
    "heaptrack": (5, 1),
}
COMMON_ARTIFACTS = {"corpus-catalog.json", "report.json", "resource.log", "stderr.log", "stdout.log"}
DRIVER_FILES = ("capture.py", "capture-r1.py.txt", "capture-initial.py.txt")
REQUIRED_COUNTERS = ("cycles:u", "instructions:u", "branches:u", "branch-misses:u", "cache-misses:u", "page-faults")
PROFILE_ENVIRONMENT = {
    "RUSTUP_TOOLCHAIN": "1.98.1",
    "RUSTFLAGS": "-C force-frame-pointers=yes -C force-unwind-tables=yes",
    "CARGO_PROFILE_RELEASE_DEBUG": "1",
    "CARGO_BUILD_JOBS": "4",
    "CARGO_INCREMENTAL": "0",
    "DEBUGINFOD_URLS": "",
}
PROFILE_BUILD_ARGV = [
    "cargo",
    "build",
    "--release",
    "--locked",
    "--manifest-path",
    "tools/perf-baseline/Cargo.toml",
    "--bin",
    "litchi-perf-baseline",
]
FP_STACK_BLOCKS = 15483
FP_WEIGHTED_PERIOD = 135131742466
FP_CONTEXTS = {
    "attribute_lookup_under_commit": (1317, 11458403980),
    "commit": (8610, 74497523419),
    "counting_sink_writer": (2384, 21137002598),
    "worksheet_parser_under_commit": (4100, 35501949212),
}


class VerificationError(ValueError):
    """A fail-closed bundle validation error."""


def fail(message: str) -> None:
    raise VerificationError(message)


def regular(path: Path, label: str) -> Path:
    if path.is_symlink() or not path.is_file():
        fail(f"{label}: missing, symlinked, or non-regular file")
    return path


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_file(path: Path, *, gzip_stream: bool = False) -> tuple[str, int]:
    digest = hashlib.sha256()
    size = 0
    try:
        stream = gzip.open(path, "rb") if gzip_stream else path.open("rb")
        with stream:
            while True:
                block = stream.read(1024 * 1024)
                if not block:
                    break
                digest.update(block)
                size += len(block)
    except (OSError, EOFError) as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest(), size


def read_json(path: Path, label: str) -> Any:
    regular(path, label)
    try:
        if path.suffix == ".gz":
            with gzip.open(path, "rt", encoding="utf-8") as stream:
                return json.load(stream)
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"{label}: invalid JSON ({error})")


def mapping(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label}: expected object")
    return value


def nonempty_text(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        fail(f"{label}: expected non-empty string")
    return value


def sha_text(value: Any, label: str) -> str:
    value = nonempty_text(value, label).lower()
    if SHA256_RE.fullmatch(value) is None:
        fail(f"{label}: expected lowercase SHA-256")
    return value


def integer(value: Any, label: str, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        fail(f"{label}: expected integer >= {minimum}")
    return value


def safe_relative(value: Any, label: str) -> Path:
    raw = nonempty_text(value, label)
    path = Path(raw)
    if path.is_absolute() or ".." in path.parts or path.as_posix() != raw:
        fail(f"{label}: expected a relative traversal-free path")
    return path


def bundle_path(value: Any, label: str, base: Path = ROOT) -> Path:
    relative = safe_relative(value, label)
    path = base / relative
    try:
        resolved = path.resolve(strict=True)
        if not resolved.is_relative_to(base.resolve()):
            fail(f"{label}: path escapes its bundle")
    except OSError as error:
        fail(f"{label}: cannot resolve ({error})")
    return regular(path, label)


def canonical_json(value: Any) -> bytes:
    try:
        return json.dumps(
            value,
            ensure_ascii=False,
            allow_nan=False,
            sort_keys=True,
            separators=(",", ":"),
        ).encode("utf-8")
    except (TypeError, ValueError) as error:
        fail(f"cannot canonicalize JSON: {error}")
    raise AssertionError("unreachable")


def json_sha(value: Any) -> str:
    return sha256_bytes(canonical_json(value))


def verify_sha256sums() -> int:
    sums = ROOT / "SHA256SUMS"
    if not sums.exists():
        return 0
    regular(sums, "SHA256SUMS")
    rows: dict[str, str] = {}
    try:
        lines = sums.read_text(encoding="utf-8").splitlines()
    except (OSError, UnicodeError) as error:
        fail(f"SHA256SUMS: cannot read ({error})")
    for number, line in enumerate(lines, 1):
        fields = line.split("  ", 1)
        if len(fields) != 2 or SHA256_RE.fullmatch(fields[0]) is None:
            fail(f"SHA256SUMS: malformed line {number}")
        name = fields[1]
        relative = safe_relative(name, f"SHA256SUMS line {number}")
        if name == "SHA256SUMS" or name in rows:
            fail(f"SHA256SUMS: duplicate or self entry {name}")
        member = bundle_path(relative.as_posix(), f"SHA256SUMS.{name}")
        if sha256_file(member)[0] != fields[0]:
            fail(f"SHA256SUMS: hash differs for {name}")
        rows[name] = fields[0]
    actual: set[str] = set()
    for member in ROOT.rglob("*"):
        if member.is_symlink():
            fail(f"bundle contains symlink {member.relative_to(ROOT)}")
        if member.is_file() and member != sums:
            actual.add(member.relative_to(ROOT).as_posix())
    if set(rows) != actual:
        missing = sorted(actual - set(rows))[:3]
        extra = sorted(set(rows) - actual)[:3]
        fail(f"SHA256SUMS: exact coverage differs (missing={missing}, extra={extra})")
    return len(rows)


def _prior_binding_path(receipt: Mapping[str, Any], label: str) -> Path:
    raw = nonempty_text(receipt.get("original_build_binding"), f"{label}.original_build_binding")
    relative = Path(raw)
    if relative.is_absolute() or relative.as_posix() != raw:
        fail(f"{label}.original_build_binding: expected a relative path")
    path = ROOT / relative
    try:
        resolved = path.resolve(strict=True)
        if not resolved.is_relative_to(ROOT.parent.resolve()):
            fail(f"{label}.original_build_binding: path escapes adjacent evidence")
    except OSError as error:
        fail(f"{label}.original_build_binding: cannot resolve ({error})")
    path = regular(path, f"{label}.original_build_binding")
    expected = sha_text(receipt.get("original_build_binding_sha256"), f"{label}.original_build_binding_sha256")
    if sha256_file(path)[0] != expected:
        fail(f"{label}: original binding hash differs")
    return path


def verify_provenance(receipts: Mapping[str, Mapping[str, Any]]) -> dict[str, Any]:
    first = next(iter(receipts.values()))
    binding_path = _prior_binding_path(first, "receipt")
    binding = mapping(read_json(binding_path, "prior binding"), "prior binding")
    if binding.get("schema") != "litchi-0465-binaries-v1":
        fail("prior binding schema differs")
    binaries = mapping(binding.get("binaries"), "prior binding.binaries")
    normal = mapping(binaries.get("normal"), "prior binding.binaries.normal")
    binary_sha = sha_text(normal.get("sha256"), "prior binding.normal.sha256")
    binary_bytes = integer(normal.get("bytes"), "prior binding.normal.bytes", 1)
    source = mapping(normal.get("source"), "prior binding.normal.source")
    source_path = bundle_path(source.get("path"), "prior source manifest", binding_path.parent)
    source_sha = sha_text(source.get("sha256"), "prior source manifest.sha256")
    source_files = integer(source.get("files"), "prior source manifest.files", 1)
    if sha256_file(source_path)[0] != source_sha:
        fail("prior source manifest hash differs")
    source_map = mapping(read_json(source_path, "prior source manifest"), "prior source manifest")
    if len(source_map) != source_files:
        fail("prior source manifest file count differs")
    for name, value in source_map.items():
        if not isinstance(name, str) or Path(name).is_absolute() or ".." in Path(name).parts:
            fail(f"prior source manifest has unsafe path {name!r}")
        sha_text(value, f"prior source manifest[{name}]")
    build_path = bundle_path(normal.get("build_receipt"), "prior normal build receipt", binding_path.parent)
    build_sha = sha_text(normal.get("build_receipt_sha256"), "prior normal build receipt.sha256")
    if sha256_file(build_path)[0] != build_sha:
        fail("prior normal build receipt hash differs")
    build = mapping(read_json(build_path, "prior normal build receipt"), "prior normal build receipt")
    if build.get("status") != "pass" or build.get("exit_code") != 0 or build.get("source_unchanged") is not True:
        fail("prior normal build receipt is not passing")
    if build.get("revision") != binding.get("revision"):
        fail("prior build revision differs from binding")
    if build.get("source_before") != source or build.get("source_after") != source:
        fail("prior build receipt source custody differs")
    for label, receipt in receipts.items():
        if sha_text(receipt.get("binary_sha256"), f"{label}.binary_sha256") != binary_sha:
            fail(f"{label}: binary hash differs from prior binding")
        if integer(receipt.get("samples"), f"{label}.samples") != LANES[label][0]:
            fail(f"{label}: sample count differs")
        if sha_text(receipt.get("source_manifest_sha256"), f"{label}.source_manifest_sha256") != source_sha:
            fail(f"{label}: source manifest hash differs from prior binding")
    return {
        "binding_path": str(binding_path.relative_to(ROOT.parent.parent.parent)),
        "binding_sha256": sha256_file(binding_path)[0],
        "binary_sha256": binary_sha,
        "binary_bytes": binary_bytes,
        "source_sha256": source_sha,
        "source_files": source_files,
        "build_revision": binding.get("revision"),
        "build_receipt_sha256": build_sha,
    }


def _artifact_rows(receipt: Mapping[str, Any], label: str) -> dict[str, Mapping[str, Any]]:
    rows = receipt.get("artifacts")
    if isinstance(rows, dict):
        result = {}
        for name, value in rows.items():
            if not isinstance(name, str) or not name:
                fail(f"{label}.artifacts: invalid name")
            result[name] = mapping(value, f"{label}.artifacts.{name}")
        return result
    if isinstance(rows, list):
        result = {}
        for index, value in enumerate(rows):
            row = mapping(value, f"{label}.artifacts[{index}]")
            path = nonempty_text(row.get("path"), f"{label}.artifacts[{index}].path")
            name = Path(path).name
            if name in result:
                fail(f"{label}.artifacts: duplicate {name}")
            result[name] = row
        return result
    fail(f"{label}.artifacts: expected object or array")
    raise AssertionError("unreachable")


def _artifact_target(lane_dir: Path, name: str, row: Mapping[str, Any], label: str) -> tuple[Path, bool]:
    raw = row.get("path", name)
    path = Path(nonempty_text(raw, f"{label}.path"))
    if path.is_absolute() or ".." in path.parts or path.as_posix() != str(path):
        fail(f"{label}.path: unsafe path")
    # A receipt may use either a lane-relative name or a bundle-relative path.
    candidate = ROOT / path if path.parts and path.parts[0] == lane_dir.name else lane_dir / path
    try:
        resolved = candidate.resolve(strict=False)
        if not resolved.is_relative_to(lane_dir.resolve()):
            fail(f"{label}.path: escapes lane")
    except OSError as error:
        fail(f"{label}.path: cannot resolve ({error})")
    if candidate.exists():
        return candidate, False
    # Finalization may gzip a raw log while retaining its original receipt
    # identity.  Verify the decompressed bytes against the recorded hash.
    if candidate.suffix != ".gz":
        compressed = Path(str(candidate) + ".gz")
        if compressed.exists():
            return compressed, True
    fail(f"{label}: artifact is missing")
    raise AssertionError("unreachable")


def _artifact_bytes(path: Path) -> bytes:
    try:
        if path.suffix == ".gz":
            with gzip.open(path, "rb") as stream:
                return stream.read()
        return path.read_bytes()
    except (OSError, EOFError) as error:
        fail(f"cannot read artifact {path}: {error}")
    raise AssertionError("unreachable")


def verify_artifacts(lane: str, receipt: Mapping[str, Any]) -> dict[str, Path]:
    lane_dir = ROOT / lane
    if not lane_dir.is_dir() or lane_dir.is_symlink():
        fail(f"{lane}: missing or unsafe lane directory")
    rows = _artifact_rows(receipt, f"{lane}.receipt")
    expected = set(COMMON_ARTIFACTS)
    expected.add("started.json") if lane in {"normal-r1", "normal-r2", "normal-r3", "normal-r4", "counters", "heaptrack"} else None
    if lane == "counters":
        expected.add("counters.csv")
    if lane == "samples":
        expected.add("perf.data")
    if lane == "heaptrack":
        expected.add("heaptrack.zst")
    if not expected.issubset(rows):
        fail(f"{lane}: receipt artifact set misses {sorted(expected - set(rows))}")
    paths: dict[str, Path] = {}
    for name, row in rows.items():
        expected_bytes = integer(row.get("bytes"), f"{lane}.artifacts.{name}.bytes")
        expected_sha = sha_text(row.get("sha256"), f"{lane}.artifacts.{name}.sha256")
        target, compressed_fallback = _artifact_target(lane_dir, name, row, f"{lane}.artifacts.{name}")
        if compressed_fallback or target.suffix == ".gz":
            actual_sha, actual_bytes = sha256_file(target, gzip_stream=True)
            stored_sha, stored_bytes = sha256_file(target)
            if actual_sha != expected_sha or actual_bytes != expected_bytes:
                fail(f"{lane}.artifacts.{name}: decompressed identity differs")
            if "compressed_sha256" in row and sha_text(row["compressed_sha256"], f"{lane}.artifacts.{name}.compressed_sha256") != stored_sha:
                fail(f"{lane}.artifacts.{name}: compressed hash differs")
            if "compressed_bytes" in row and integer(row["compressed_bytes"], f"{lane}.artifacts.{name}.compressed_bytes") != stored_bytes:
                fail(f"{lane}.artifacts.{name}: compressed size differs")
        else:
            actual_sha, actual_bytes = sha256_file(target)
            if actual_sha != expected_sha or actual_bytes != expected_bytes:
                fail(f"{lane}.artifacts.{name}: identity differs")
        paths[name] = target
    return paths


def _close_number(actual: Any, expected: float, label: str) -> None:
    if isinstance(actual, bool) or not isinstance(actual, (int, float)) or not math.isfinite(float(actual)):
        fail(f"{label}: expected finite number")
    if abs(float(actual) - expected) > 0.5:
        fail(f"{label}: differs from recomputed statistic ({actual} != {expected})")


def _student_t_critical_95(degrees: int) -> float:
    values = [12.706, 4.303, 3.182, 2.776, 2.571, 2.447, 2.365, 2.306, 2.262, 2.228, 2.201, 2.179, 2.160, 2.145, 2.131, 2.120, 2.110, 2.101, 2.093, 2.086, 2.080, 2.074, 2.069, 2.064, 2.060, 2.056, 2.052, 2.048, 2.045, 2.042]
    if degrees <= 0:
        return 0.0
    if degrees <= len(values):
        return values[degrees - 1]
    z = 1.959963984540054
    d = float(degrees)
    z2 = z * z
    z3 = z2 * z
    z5 = z3 * z2
    z7 = z5 * z2
    return z + (z3 + z) / (4 * d) + (5 * z5 + 16 * z3 + 3 * z) / (96 * d * d) + (3 * z7 + 19 * z5 + 17 * z3 - 15 * z) / (384 * d * d * d)


def verify_elapsed(result: Mapping[str, Any], expected_samples: int, label: str) -> None:
    elapsed = mapping(result.get("elapsed_ns"), f"{label}.elapsed_ns")
    if elapsed.get("unit") != "ns":
        fail(f"{label}.elapsed_ns.unit differs")
    samples = elapsed.get("samples")
    order = elapsed.get("sample_order")
    if not isinstance(samples, list) or len(samples) != expected_samples:
        fail(f"{label}: expected {expected_samples} elapsed samples")
    if any(isinstance(x, bool) or not isinstance(x, int) or x <= 0 for x in samples):
        fail(f"{label}: elapsed samples must be positive integers")
    if samples != sorted(samples):
        fail(f"{label}: elapsed samples are not sorted")
    if not isinstance(order, list) or sorted(order) != list(range(expected_samples)):
        fail(f"{label}: sample_order is not a complete permutation")
    if any(samples[i] == samples[i + 1] and order[i] > order[i + 1] for i in range(expected_samples - 1)):
        fail(f"{label}: sample_order does not preserve ties")
    midpoint = samples[(expected_samples - 1) // 2] // 2 + samples[expected_samples // 2] // 2 + ((samples[(expected_samples - 1) // 2] % 2 + samples[expected_samples // 2] % 2) // 2)
    nearest = lambda percentile: samples[min((percentile * expected_samples + 99) // 100 - 1, expected_samples - 1)]
    if elapsed.get("min") != samples[0] or elapsed.get("max") != samples[-1] or elapsed.get("p50") != midpoint or elapsed.get("p95") != nearest(95) or elapsed.get("p99") != nearest(99):
        fail(f"{label}: min/percentile/max statistics do not recompute")
    mean = 0.0
    squared = 0.0
    for index, value in enumerate(samples):
        value_float = float(value)
        count = float(index + 1)
        delta = value_float - mean
        mean += delta / count
        squared += delta * (value_float - mean)
    deviation = math.sqrt(squared / (expected_samples - 1)) if expected_samples > 1 else 0.0
    margin = _student_t_critical_95(expected_samples - 1) * deviation / math.sqrt(expected_samples) if expected_samples > 1 else 0.0
    interval = mapping(elapsed.get("confidence_interval_95"), f"{label}.confidence_interval_95")
    if interval.get("method") != "two-sided Student's t interval for the mean":
        fail(f"{label}: confidence interval method differs")
    _close_number(elapsed.get("mean"), mean, f"{label}.mean")
    _close_number(elapsed.get("standard_deviation"), deviation, f"{label}.standard_deviation")
    _close_number(interval.get("lower"), max(mean - margin, 0.0), f"{label}.confidence_interval_95.lower")
    _close_number(interval.get("upper"), mean + margin, f"{label}.confidence_interval_95.upper")


def verify_report(
    lane: str,
    report_path: Path,
    catalog_path: Path,
    receipt: Mapping[str, Any],
    provenance: Mapping[str, Any],
    expected_samples: int,
    expected_warmup: int,
    *,
    binary_sha256: str | None = None,
    binary_bytes: int | None = None,
) -> tuple[str, str]:
    report = mapping(read_json(report_path, f"{lane}.report.json"), f"{lane}.report.json")
    catalog = mapping(read_json(catalog_path, f"{lane}.corpus-catalog.json"), f"{lane}.corpus-catalog.json")
    if report.get("schema_version") != 1 or report.get("tool", {}).get("name") != "litchi-perf-baseline":
        fail(f"{lane}: report schema/tool differs")
    configuration = mapping(report.get("configuration"), f"{lane}.configuration")
    if configuration.get("cases") != [CASE] or configuration.get("samples_per_case") != expected_samples or configuration.get("warmup_iterations_per_case") != expected_warmup:
        fail(f"{lane}: report configuration differs from receipt")
    if configuration.get("execution_workers") != [1]:
        fail(f"{lane}: report worker configuration differs")
    binary = mapping(report.get("binary_identity"), f"{lane}.binary_identity")
    expected_binary_sha = binary_sha256 if binary_sha256 is not None else provenance["binary_sha256"]
    expected_binary_bytes = binary_bytes if binary_bytes is not None else provenance["binary_bytes"]
    if binary.get("binary_sha256") != expected_binary_sha or binary.get("binary_bytes") != expected_binary_bytes:
        fail(f"{lane}: report binary identity differs")
    reference = {key: catalog.get(key) for key in ("manifest_version", "catalog_id", "catalog_sha256", "content_set_sha256")}
    if reference != report.get("corpus_catalog"):
        fail(f"{lane}: report/catalog reference differs")
    if catalog.get("manifest_version") != 2 or catalog.get("catalog_id") != "litchi-perf-corpus-v2" or catalog.get("catalog_sha256") != CATALOG_SHA256 or catalog.get("content_set_sha256") != CONTENT_SET_SHA256:
        fail(f"{lane}: catalog identity differs")
    catalog_without_hash = dict(catalog)
    catalog_without_hash.pop("catalog_sha256", None)
    if json_sha(catalog_without_hash) != CATALOG_SHA256:
        fail(f"{lane}: catalog_sha256 does not recompute")
    rows = report.get("results")
    if not isinstance(rows, list) or len(rows) != 1:
        fail(f"{lane}: report must contain exactly one result")
    result = mapping(rows[0], f"{lane}.results[0]")
    if result.get("case") != CASE:
        fail(f"{lane}: result case differs")
    corpus = mapping(result.get("corpus"), f"{lane}.results[0].corpus")
    if corpus.get("archive_sha256") != CORPUS_SHA256:
        fail(f"{lane}: result corpus archive identity differs")
    if corpus.get("name") != "xlsx-dense-wide" or corpus.get("generator") != "litchi-xlsx-synthetic-v1":
        fail(f"{lane}: result corpus metadata differs")
    expected_id = f"xlsx-opc-zip:sha256:{CORPUS_SHA256}"
    corpus_records = catalog.get("corpora")
    if not isinstance(corpus_records, list) or not any(isinstance(x, dict) and x.get("id") == expected_id for x in corpus_records):
        fail(f"{lane}: catalog lacks exact timed corpus")
    bindings = catalog.get("case_bindings")
    if not isinstance(bindings, list) or len(bindings) != 1 or bindings[0].get("case") != CASE or bindings[0].get("corpus_id") != expected_id:
        fail(f"{lane}: catalog case binding differs")
    verify_elapsed(result, expected_samples, f"{lane}.results[0]")
    operation = mapping(result.get("operation_metrics"), f"{lane}.results[0].operation_metrics")
    if operation.get("sample_count") != expected_samples or operation.get("sample_indices") != list(range(expected_samples)):
        fail(f"{lane}: operation sample identity differs")
    return CASE, CORPUS_SHA256


def verify_lane(lane: str, receipt: Mapping[str, Any], provenance: Mapping[str, Any]) -> tuple[str, str]:
    samples, warmup = LANES[lane]
    if receipt.get("schema") != "litchi-0466-capture-v1" or receipt.get("lane") != lane:
        fail(f"{lane}: receipt schema/lane differs")
    if receipt.get("exit_code") != 0 or receipt.get("source_and_binary_unchanged") is not True:
        fail(f"{lane}: capture did not pass custody/exit checks")
    if receipt.get("cpu") != 2 or receipt.get("warmups") != warmup:
        fail(f"{lane}: receipt warmup/cpu identity differs")
    driver_hashes = {sha256_file(ROOT / name)[0] for name in DRIVER_FILES if (ROOT / name).is_file()}
    driver = sha_text(receipt.get("driver_sha256"), f"{lane}.driver_sha256")
    if driver not in driver_hashes:
        fail(f"{lane}: driver hash does not bind a retained capture helper")
    if receipt.get("original_driver_sha256") is not None:
        if sha_text(receipt.get("original_driver_sha256"), f"{lane}.original_driver_sha256") != sha256_file(ROOT / "capture-initial.py.txt")[0]:
            fail(f"{lane}: original driver hash differs")
        if not isinstance(receipt.get("recovery"), str) or not receipt["recovery"]:
            fail(f"{lane}: recovered receipt lacks recovery explanation")
    artifacts = verify_artifacts(lane, receipt)
    if b"Exit status: 0" not in _artifact_bytes(artifacts["resource.log"]):
        fail(f"{lane}: resource log lacks successful exit marker")
    if lane == "counters":
        lines = _artifact_bytes(artifacts["counters.csv"]).decode("utf-8", errors="replace").splitlines()
        seen: set[str] = set()
        for line in lines:
            fields = line.split(";")
            if len(fields) < 4 or not fields[2]:
                continue
            event = fields[2]
            if event in REQUIRED_COUNTERS:
                if not fields[0].strip().isdigit():
                    fail("counters.csv contains a non-numeric event count")
                seen.add(event)
        if seen != set(REQUIRED_COUNTERS):
            fail(f"counters.csv missing events: {sorted(set(REQUIRED_COUNTERS) - seen)}")
    if lane == "samples":
        if b"Captured and wrote" not in _artifact_bytes(artifacts["stderr.log"]) or len(_artifact_bytes(artifacts["perf.data"])) == 0:
            fail("samples: perf capture output is incomplete")
    if lane == "heaptrack":
        heaptrack_output = _artifact_bytes(artifacts["stderr.log"]) + _artifact_bytes(artifacts["stdout.log"])
        if b"Heaptrack finished!" not in heaptrack_output or artifacts["heaptrack.zst"].stat().st_size == 0:
            fail("heaptrack: capture output is incomplete")
    return verify_report(lane, artifacts["report.json"], artifacts["corpus-catalog.json"], receipt, provenance, samples, warmup)


def _bundle_artifact(relative: str, label: str) -> Path:
    """Resolve a retained artifact, accepting finalization's ``.gz`` form."""

    path = ROOT / safe_relative(relative, f"{label}.path")
    if path.exists() or path.is_symlink():
        return regular(path, label)
    compressed = Path(str(path) + ".gz")
    if compressed.exists() or compressed.is_symlink():
        return regular(compressed, label)
    fail(f"{label}: artifact is missing")
    raise AssertionError("unreachable")


def _artifact_sha256(path: Path) -> tuple[str, int]:
    if path.suffix == ".gz":
        return sha256_file(path, gzip_stream=True)
    return sha256_file(path)


def _verify_process_receipt(path: Path, label: str, expected_environment: Mapping[str, str] | None = None) -> dict[str, Any]:
    receipt = mapping(read_json(path, label), label)
    argv = receipt.get("argv")
    if not isinstance(argv, list) or not argv or any(not isinstance(item, str) or not item for item in argv):
        fail(f"{label}: argv is missing or malformed")
    if receipt.get("exit_code") != 0:
        fail(f"{label}: command did not exit successfully")
    for field in ("started_utc", "finished_utc"):
        nonempty_text(receipt.get(field), f"{label}.{field}")
    environment = mapping(receipt.get("environment"), f"{label}.environment")
    if expected_environment is not None:
        if environment != dict(expected_environment):
            fail(f"{label}: environment differs from the profile build")
    return receipt


def _require_argv(argv: list[str], required: list[str], label: str) -> None:
    position = 0
    for item in required:
        try:
            position = argv.index(item, position) + 1
        except ValueError:
            fail(f"{label}: argv misses {item!r}")


def verify_compression() -> int:
    """Verify packed large artifacts against their recorded raw identities."""

    path = ROOT / "compression.json"
    if not path.exists():
        fail("compression.json: missing")
    rows = mapping(read_json(path, "compression.json"), "compression.json")
    expected_names = {
        "samples/perf.data",
        "samples/perf-script.txt",
        "samples-fp/perf.data",
        "samples-fp/perf-script.stdout",
        "heaptrack/print.txt",
    }
    if set(rows) != expected_names:
        fail(f"compression.json: expected exactly {sorted(expected_names)}")
    for name, value in rows.items():
        original = mapping(value, f"compression.json.{name}").get("original")
        compressed = mapping(mapping(value, f"compression.json.{name}").get("compressed"), f"compression.json.{name}.compressed")
        original = mapping(original, f"compression.json.{name}.original")
        expected_original_sha = sha_text(original.get("sha256"), f"compression.json.{name}.original.sha256")
        expected_original_bytes = integer(original.get("bytes"), f"compression.json.{name}.original.bytes", 1)
        compressed_name = nonempty_text(compressed.get("path"), f"compression.json.{name}.compressed.path")
        if not compressed_name.endswith(".gz"):
            fail(f"compression.json.{name}: compressed path is not gzip")
        compressed_path = _bundle_artifact(compressed_name, f"compression.json.{name}.compressed")
        compressed_sha = sha_text(compressed.get("sha256"), f"compression.json.{name}.compressed.sha256")
        compressed_bytes = integer(compressed.get("bytes"), f"compression.json.{name}.compressed.bytes", 1)
        actual_compressed_sha, actual_compressed_bytes = sha256_file(compressed_path)
        actual_original_sha, actual_original_bytes = sha256_file(compressed_path, gzip_stream=True)
        if (actual_compressed_sha, actual_compressed_bytes) != (compressed_sha, compressed_bytes):
            fail(f"compression.json.{name}: compressed identity differs")
        if (actual_original_sha, actual_original_bytes) != (expected_original_sha, expected_original_bytes):
            fail(f"compression.json.{name}: decompressed identity differs")
    return len(rows)


def verify_postprocess() -> int:
    """Authenticate the retained standard perf and heaptrack post-processors."""

    receipt = mapping(read_json(ROOT / "postprocess.json", "postprocess.json"), "postprocess.json")
    commands = receipt.get("commands")
    if not isinstance(commands, list) or len(commands) != 3:
        fail("postprocess.json: expected three successful commands")
    expected_artifacts = [
        {"samples/perf-script.txt", "samples/perf-script.stderr"},
        {"samples/top-symbols.txt", "samples/top-symbols.stderr"},
        {"heaptrack/print.txt", "heaptrack/print.stderr"},
    ]
    expected_tools = ("perf", "perf", "heaptrack_print")
    for index, command_value in enumerate(commands):
        command = mapping(command_value, f"postprocess.json.commands[{index}]")
        if command.get("exit_code") != 0:
            fail(f"postprocess.json.commands[{index}]: command did not exit successfully")
        argv = command.get("argv")
        if not isinstance(argv, list) or not argv or any(not isinstance(item, str) or not item for item in argv):
            fail(f"postprocess.json.commands[{index}]: malformed argv")
        if argv[0] != expected_tools[index]:
            fail(f"postprocess.json.commands[{index}]: tool differs")
        if index == 0:
            _require_argv(argv, ["script", "--no-inline", "-i"], f"postprocess.json.commands[{index}]")
        elif index == 1:
            _require_argv(argv, ["report", "--stdio", "--no-children", "--no-inline", "-i", "--sort", "symbol"], f"postprocess.json.commands[{index}]")
        else:
            _require_argv(argv, ["-f", "-n", "30"], f"postprocess.json.commands[{index}]")
        artifacts = mapping(command.get("artifacts"), f"postprocess.json.commands[{index}].artifacts")
        if set(artifacts) != expected_artifacts[index]:
            fail(f"postprocess.json.commands[{index}]: artifact set differs")
        for name, expected in artifacts.items():
            expected_sha = sha_text(expected, f"postprocess.json.commands[{index}].artifacts.{name}")
            target = _bundle_artifact(name, f"postprocess.json.commands[{index}].artifacts.{name}")
            if _artifact_sha256(target)[0] != expected_sha:
                fail(f"postprocess.json.commands[{index}].artifacts.{name}: identity differs")
    return len(commands)


def verify_profile(provenance: Mapping[str, Any]) -> dict[str, Any]:
    """Verify the separate frame-pointer build, capture, and export receipts."""

    binding = mapping(read_json(ROOT / "profile-binding.json", "profile-binding.json"), "profile-binding.json")
    profile_sha = sha_text(binding.get("binary_sha256"), "profile-binding.json.binary_sha256")
    profile_bytes = integer(binding.get("bytes"), "profile-binding.json.bytes", 1)
    profile_source_sha = sha_text(binding.get("source_manifest_sha256"), "profile-binding.json.source_manifest_sha256")
    profile_revision = nonempty_text(binding.get("revision"), "profile-binding.json.revision")
    if not re.fullmatch(r"[0-9a-f]{40}", profile_revision):
        fail("profile-binding.json.revision: expected a commit id")
    if profile_source_sha != provenance["source_sha256"]:
        fail("profile-binding.json: source manifest differs from normal provenance")
    if profile_sha == provenance["binary_sha256"]:
        fail("profile-binding.json: profile binary hash must differ from normal binary")
    if binding.get("scope") != "same source as normal, different profiling build flags; not normal latency evidence":
        fail("profile-binding.json.scope: diagnostic scope differs")

    build = _verify_process_receipt(ROOT / "profile-build.json", "profile-build.json", PROFILE_ENVIRONMENT)
    if build.get("argv") != PROFILE_BUILD_ARGV:
        fail("profile-build.json: build argv differs")

    capture = _verify_process_receipt(ROOT / "samples-fp" / "capture.json", "samples-fp/capture.json", PROFILE_ENVIRONMENT)
    capture_argv = capture["argv"]
    _require_argv(capture_argv, ["perf", "record", "-F", "499", "-e", "cycles:u", "--call-graph", "fp", "-o", "--", "--warmup", "3", "--samples", "50"], "samples-fp/capture.json")
    script_receipt = _verify_process_receipt(ROOT / "samples-fp" / "perf-script.json", "samples-fp/perf-script.json", PROFILE_ENVIRONMENT)
    _require_argv(script_receipt["argv"], ["perf", "script", "--no-inline", "-i", "-F"], "samples-fp/perf-script.json")
    report_receipt = _verify_process_receipt(ROOT / "samples-fp" / "top-symbols.json", "samples-fp/top-symbols.json", PROFILE_ENVIRONMENT)
    _require_argv(report_receipt["argv"], ["perf", "report", "--stdio", "--no-children", "--no-inline", "-g", "none", "-i", "--sort", "symbol", "--percent-limit", "0.5"], "samples-fp/top-symbols.json")

    resource = _bundle_artifact("samples-fp/resource.log", "samples-fp/resource.log")
    if b"Exit status: 0" not in _artifact_bytes(resource):
        fail("samples-fp/resource.log: capture lacks successful exit marker")
    perf_data = _bundle_artifact("samples-fp/perf.data", "samples-fp/perf.data")
    if not _artifact_bytes(perf_data):
        fail("samples-fp/perf.data: capture is empty")
    for name in ("capture.stdout", "capture.stderr", "perf-script.stderr", "top-symbols.stdout", "top-symbols.stderr", "report.json", "corpus-catalog.json"):
        _bundle_artifact(f"samples-fp/{name}", f"samples-fp/{name}")
    top_symbols = _bundle_artifact("samples-fp/top-symbols.stdout", "samples-fp/top-symbols.stdout")
    if not _artifact_bytes(top_symbols):
        fail("samples-fp/top-symbols.stdout: report is empty")

    report = mapping(read_json(ROOT / "samples-fp" / "report.json", "samples-fp.report.json"), "samples-fp.report.json")
    binary = mapping(report.get("binary_identity"), "samples-fp.report.binary_identity")
    if binary.get("binary_sha256") != profile_sha or binary.get("binary_bytes") != profile_bytes:
        fail("samples-fp.report: binary identity differs from profile binding")
    environment = mapping(report.get("environment"), "samples-fp.report.environment")
    if environment.get("git_revision") != profile_revision or environment.get("rustflags") != PROFILE_ENVIRONMENT["RUSTFLAGS"]:
        fail("samples-fp.report: build environment differs from profile binding")
    verify_report(
        "samples-fp",
        ROOT / "samples-fp" / "report.json",
        ROOT / "samples-fp" / "corpus-catalog.json",
        capture,
        provenance,
        50,
        3,
        binary_sha256=profile_sha,
        binary_bytes=profile_bytes,
    )
    return {
        "binary_sha256": profile_sha,
        "binary_bytes": profile_bytes,
        "source_sha256": profile_source_sha,
        "revision": profile_revision,
    }


def _load_analyzer() -> Any:
    path = regular(ROOT / "analyze.py", "analyze.py")
    spec = importlib.util.spec_from_file_location("litchi_0466_analyze_verify", path)
    if spec is None or spec.loader is None:
        fail("analyze.py: cannot load parser")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    try:
        spec.loader.exec_module(module)
    except (OSError, SyntaxError, TypeError, ValueError) as error:
        fail(f"analyze.py: cannot load parser ({error})")
    return module


def _verify_fp_summary_values(summary: Mapping[str, Any]) -> None:
    whole = mapping(summary.get("whole_process"), "summary-fp.whole_process")
    if whole.get("raw_stack_blocks") != FP_STACK_BLOCKS or whole.get("weighted_event_period") != FP_WEIGHTED_PERIOD:
        fail("summary-fp: whole-process FP totals differ")
    parser = mapping(summary.get("sample_parser"), "summary-fp.sample_parser")
    if parser.get("accepted_sample_blocks") != FP_STACK_BLOCKS or parser.get("accepted_weighted_event_period") != FP_WEIGHTED_PERIOD:
        fail("summary-fp: parser FP totals differ")
    diagnostics = mapping(parser.get("parser_diagnostics"), "summary-fp.sample_parser.parser_diagnostics")
    for key in ("sample_blocks", "sample_blocks_seen"):
        if diagnostics.get(key) != FP_STACK_BLOCKS:
            fail(f"summary-fp: parser diagnostic {key} differs")
    for key in ("accepted_cycle_period", "total_weighted_event_period"):
        if diagnostics.get(key) != FP_WEIGHTED_PERIOD:
            fail(f"summary-fp: parser diagnostic {key} differs")
    partition = mapping(mapping(summary.get("disjoint_scopes"), "summary-fp.disjoint_scopes").get("partition"), "summary-fp.disjoint_scopes.partition")
    for key in ("blocks_partition_exact", "periods_partition_exact"):
        if partition.get(key) is not True:
            fail(f"summary-fp: partition flag {key} is false")
    for key in ("classified_raw_stack_blocks", "whole_process_raw_stack_blocks"):
        if partition.get(key) != FP_STACK_BLOCKS:
            fail(f"summary-fp: partition block count {key} differs")
    for key in ("classified_weighted_event_period", "whole_process_weighted_event_period"):
        if partition.get(key) != FP_WEIGHTED_PERIOD:
            fail(f"summary-fp: partition period {key} differs")
    contexts = mapping(summary.get("exact_contexts"), "summary-fp.exact_contexts")
    if contexts.get("scope") != "whole-process exact ancestor contexts; inclusive rows overlap; not elapsed phases":
        fail("summary-fp: exact-context scope differs")
    rows = mapping(contexts.get("rows"), "summary-fp.exact_contexts.rows")
    for name, (blocks, period) in FP_CONTEXTS.items():
        row = mapping(rows.get(name), f"summary-fp.exact_contexts.rows.{name}")
        if row.get("raw_stack_blocks") != blocks or row.get("weighted_event_period") != period:
            fail(f"summary-fp: exact context {name} differs")


def verify_summaries() -> dict[str, Any]:
    """Recompute both retained summaries from the actual (possibly gzipped) scripts."""

    analyzer = _load_analyzer()
    reports = [
        _bundle_artifact("normal-r3/report.json", "normal-r3/report.json"),
        _bundle_artifact("normal-r4/report.json", "normal-r4/report.json"),
    ]
    verified: dict[str, Any] = {}
    for summary_name, script_name, expected_fp in (
        ("summary-dwarf.json", "samples/perf-script.txt", False),
        ("summary-fp.json", "samples-fp/perf-script.stdout", True),
    ):
        summary_path = _bundle_artifact(summary_name, summary_name)
        script_path = _bundle_artifact(script_name, script_name)
        summary = mapping(read_json(summary_path, summary_name), summary_name)
        try:
            recomputed = analyzer.analyze_script(script_path, reports, top=40, bundle_root=ROOT)
        except (OSError, UnicodeError, ValueError, TypeError, KeyError) as error:
            fail(f"{summary_name}: analyze.py recomputation failed ({error})")
        if summary != recomputed:
            fail(f"{summary_name}: derived summary differs from analyze.py recomputation")
        if expected_fp:
            _verify_fp_summary_values(summary)
        verified[summary_name] = {
            "script": str(script_path.relative_to(ROOT)),
            "raw_stack_blocks": summary["whole_process"]["raw_stack_blocks"],
            "weighted_event_period": summary["whole_process"]["weighted_event_period"],
        }
    return verified


def verify(*, expected_root: Path = ROOT) -> dict[str, Any]:
    if expected_root != ROOT:
        fail("the verifier's bundle root is immutable")
    receipts: dict[str, Mapping[str, Any]] = {}
    for lane in LANES:
        receipt = mapping(read_json(ROOT / lane / "receipt.json", f"{lane}.receipt.json"), f"{lane}.receipt.json")
        receipts[lane] = receipt
    provenance = verify_provenance(receipts)
    identities = {lane: verify_lane(lane, receipt, provenance) for lane, receipt in receipts.items()}
    if len(set(identities.values())) != 1:
        fail("lanes do not share one exact case/corpus identity")
    profile = verify_profile(provenance)
    postprocess_commands = verify_postprocess()
    compressed_files = verify_compression()
    summaries = verify_summaries()
    sealed = verify_sha256sums()
    return {
        "schema": SCHEMA,
        "change": CHANGE,
        "status": "pass",
        "sealed_files": sealed,
        "lanes": list(LANES),
        "case": CASE,
        "corpus_archive_sha256": CORPUS_SHA256,
        "normal_samples": 30,
        "profile_samples": 50,
        "heaptrack_samples": 5,
        "binary_sha256": provenance["binary_sha256"],
        "profile_binary_sha256": profile["binary_sha256"],
        "profile_revision": profile["revision"],
        "postprocess_commands": postprocess_commands,
        "compressed_files": compressed_files,
        "summaries": summaries,
        "build_revision": provenance["build_revision"],
        "source_files": provenance["source_files"],
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.parse_args(sys.argv[1:] if argv is None else argv)
    try:
        result = verify()
    except (OSError, KeyError, TypeError, ValueError, VerificationError) as error:
        result = {"schema": SCHEMA, "change": CHANGE, "status": "fail", "error": str(error)}
    print(json.dumps(result, ensure_ascii=False, sort_keys=True, indent=2))
    return 0 if result.get("status") == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
