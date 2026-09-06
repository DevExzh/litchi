#!/usr/bin/env python3
"""Replay the portable change-0430 frame-pointer evidence bundle.

The bundle contains two profiles captured from the unchanged change-0429
binary.  This verifier intentionally reads the copied verifier and copied
custody records from the bundle.  It does not import a producer, inspect the
working tree, or require the executable or the source tree to be present.
"""

from __future__ import annotations

import argparse
import copy
import gzip
import hashlib
import json
import math
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
from typing import Any
import zlib


ROOT = Path(__file__).resolve().parent
U64_MAX = (1 << 64) - 1
HEX64 = set("0123456789abcdef")
HEX40 = set("0123456789abcdef")

EXPECTED_PROTOCOL = {
    "change": 430,
    "baseline_revision": "e86d46b8317448575db34137cdafdd7f8f309e3a",
    "source_revision": "56ec912e5b387e8f4f55f7044a73bb4cdd11e324",
    "adr_tree": "c950b6c8be822561b498d7bbe87c460873dcbf49",
    "binary_sha256": "9b3fca778443b6c9bd49a8522f56a251f1df27fc775a76af6124186398f78969",
    "binary_bytes": 451782312,
    "capture_binary": "/tmp/litchi-goal-0430-binaries/litchi-perf-baseline",
    "original_binary": "tools/perf-baseline/target/release/litchi-perf-baseline",
    "cpu": 2,
    "event": "cycles:u",
    "frequency_hz": 499,
    "call_graph": "fp",
    "corpus": "media-rich",
    "providers": ["bytes", "file"],
    "samples": 100,
    "warmup": 3,
    "scope": (
        "CPU attribution controls on unchanged 0429 binary, existing public API "
        "and exact output gates. Not normal timing, not an optimization or causal "
        "provider comparison. Roots serialize CPU workloads. Compare ancestry "
        "completeness with retained historical DWARF16384 profiles without treating "
        "counts as wall-clock fractions."
    ),
}

PROTOCOL_FIELDS = tuple(EXPECTED_PROTOCOL)
BUILD_REQUIRED_FIELDS = (
    "revision", "baseline_revision", "build_receipt", "build_receipt_sha256",
    "source_manifest", "original_binary", "capture_binary", "binary_sha256",
    "binary_bytes", "protocol_sha256", "planned_checks_sha256", "machine_sha256",
    "capture_driver_sha256", "verifier_sha256", "replay_verifier_sha256",
    "probe_sha256", "summary_driver_sha256", "strict_comparison_driver_sha256",
    "additional_artifact_sha256", "rust_toolchain", "profile", "scope",
    "validation_amendment",
)

RECEIPT_REQUIRED_FIELDS = (
    "provider", "argv", "source", "binary_sha256", "protocol_sha256",
    "driver_sha256", "verifier_sha256", "started_utc", "status",
    "record_exit_code", "finished_utc", "artifacts",
)

REPORT_SEMANTIC_FIELDS = (
    "schema", "corpus", "source_revision", "binary_sha256", "binary_bytes",
    "source_archive_sha256", "source_archive_bytes", "destination_archive_sha256",
    "destination_archive_bytes", "expected_output_sha256", "expected_output_bytes",
    "corpus_manifest", "gates", "configured_limits", "destination_configured_limits",
    "destination_editor_consumed_during_publish", "final_memory_objects_depth_zero_checked",
    "phases",
)

PROFILE_ARTIFACT_SUFFIXES = (
    "-record.log", "-report.log", "-script.log", ".data", ".json",
)


class VerificationError(ValueError):
    """A serialized evidence contract violation."""


def fail(path: str, message: str) -> None:
    raise VerificationError(f"{path}: {message}")


def reject_duplicate_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise VerificationError(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def reject_nonfinite(value: str) -> None:
    raise VerificationError(f"non-finite JSON number {value}")


def load_json(path: Path) -> Any:
    try:
        return json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=reject_duplicate_pairs,
            parse_constant=reject_nonfinite,
        )
    except VerificationError:
        raise
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise VerificationError(f"{path}: cannot read JSON: {exc}") from exc


def sha(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def digest(path: Path) -> str:
    try:
        return sha(path.read_bytes())
    except OSError as exc:
        raise VerificationError(f"{path}: cannot hash artifact: {exc}") from exc


def uint(value: Any, path: str) -> int:
    if type(value) is not int:
        fail(path, "expected an unsigned integer, not a boolean or other type")
    if value < 0 or value > U64_MAX:
        fail(path, "outside u64 range")
    return value


def text(value: Any, path: str) -> str:
    if type(value) is not str:
        fail(path, "expected a string")
    return value


def boolean(value: Any, path: str) -> bool:
    if type(value) is not bool:
        fail(path, "expected a boolean")
    return value


def finite_number(value: Any, path: str) -> float:
    if type(value) not in (int, float) or type(value) is bool:
        fail(path, "expected a number")
    if not math.isfinite(float(value)):
        fail(path, "non-finite number")
    return float(value)


def validate_numbers(value: Any, path: str) -> None:
    """Reject oversized integers and non-finite floats.

    Signed values are allowed in descriptive analysis deltas.  Report and
    receipt fields that are required to be u64 use :func:`uint` at their
    contract boundary, which also rejects bool-as-int values.
    """
    if type(value) is bool or value is None or isinstance(value, str):
        return
    if type(value) is int:
        if value < -U64_MAX or value > U64_MAX:
            fail(path, "outside signed u64 range")
        return
    if type(value) is float:
        finite_number(value, path)
        return
    if isinstance(value, list):
        for index, item in enumerate(value):
            validate_numbers(item, f"{path}[{index}]")
        return
    if isinstance(value, dict):
        for key, item in value.items():
            validate_numbers(item, f"{path}.{key}")


def exact_keys(value: Any, required: tuple[str, ...] | set[str], path: str,
               optional: tuple[str, ...] = ()) -> None:
    if not isinstance(value, dict):
        fail(path, "expected an object")
    expected = set(required) | set(optional)
    actual = set(value)
    missing = sorted(set(required) - actual)
    extra = sorted(actual - expected)
    if missing:
        fail(path, f"missing fields {missing}")
    if extra:
        fail(path, f"unexpected fields {extra}")


def valid_hash(value: Any, path: str) -> str:
    value = text(value, path)
    if len(value) != 64 or any(char not in HEX64 for char in value):
        fail(path, "expected a lowercase SHA-256 digest")
    return value


def valid_revision(value: Any, path: str) -> str:
    value = text(value, path)
    if len(value) != 40 or any(char not in HEX40 for char in value):
        fail(path, "expected a lowercase Git revision")
    return value


def safe_path(root: Path, name: str, path: str = "path") -> Path:
    if type(name) is not str or not name or Path(name).is_absolute():
        fail(path, "expected a relative bundle path")
    candidate = (root / name).resolve()
    root_resolved = root.resolve()
    try:
        candidate.relative_to(root_resolved)
    except ValueError:
        fail(path, "path escapes the bundle")
    return candidate


def read_payload(root: Path, name: str, path: str = "path") -> tuple[bytes, Path, bytes]:
    """Read a raw artifact, or its lossless ``.gz`` companion.

    The first return value is the logical/original payload.  The second is the
    physical path used.  The third is its stored bytes, which permits checking
    both receipt hashes and the compression inventory.
    """
    candidate = safe_path(root, name, path)
    physical = candidate
    if not physical.is_file():
        compressed = candidate.with_name(candidate.name + ".gz")
        if compressed.is_file():
            physical = compressed
        else:
            fail(path, f"missing artifact {name!r} (raw or .gz)")
    try:
        stored = physical.read_bytes()
        logical = gzip.decompress(stored) if physical.suffix == ".gz" else stored
    except (OSError, EOFError, gzip.BadGzipFile, zlib.error) as exc:
        fail(path, f"cannot read lossless artifact: {exc}")
    return logical, physical, stored


def check_custody_artifact(root: Path, name: str, row: Any, path: str) -> bytes:
    exact_keys(row, ("sha256", "bytes"), path, optional=("path",))
    if "path" in row:
        compare(row["path"], name, f"{path}.path")
    expected_hash = valid_hash(row["sha256"], f"{path}.sha256")
    expected_bytes = uint(row["bytes"], f"{path}.bytes")
    logical, _physical, _stored = read_payload(root, name, path)
    logical_hash = sha(logical)
    if expected_hash == logical_hash and expected_bytes == len(logical):
        return logical
    fail(path, "artifact digest or byte count does not match raw/gzip payload")


def require_regular(root: Path, name: str, path: str | None = None) -> Path:
    target = safe_path(root, name, path or name)
    if not target.is_file():
        fail(path or name, "missing regular file")
    return target


def compare(value: Any, expected: Any, path: str) -> None:
    if value != expected:
        fail(path, f"expected {expected!r}, got {value!r}")


def verify_protocol(root: Path, build: dict[str, Any]) -> dict[str, Any]:
    path = root / "profile-protocol.json"
    protocol = load_json(path)
    validate_numbers(protocol, "profile-protocol")
    exact_keys(protocol, PROTOCOL_FIELDS, "profile-protocol")
    for field, expected in EXPECTED_PROTOCOL.items():
        compare(protocol[field], expected, f"profile-protocol.{field}")
    compare(protocol["source_revision"], build["revision"], "profile-protocol.source_revision")
    compare(protocol["binary_sha256"], build["binary_sha256"], "profile-protocol.binary_sha256")
    compare(protocol["binary_bytes"], build["binary_bytes"], "profile-protocol.binary_bytes")
    valid_revision(protocol["baseline_revision"], "profile-protocol.baseline_revision")
    valid_revision(protocol["source_revision"], "profile-protocol.source_revision")
    valid_revision(protocol["adr_tree"], "profile-protocol.adr_tree")
    valid_hash(protocol["binary_sha256"], "profile-protocol.binary_sha256")
    return protocol


def verify_source_manifest(root: Path, build: dict[str, Any]) -> dict[str, str]:
    source = build["source_manifest"]
    exact_keys(source, ("path", "sha256", "files"), "input-build-0429.source_manifest")
    source_hash = valid_hash(source["sha256"], "input-build-0429.source_manifest.sha256")
    files = uint(source["files"], "input-build-0429.source_manifest.files")
    copied_path = safe_path(root, source["path"], "input-build-0429.source_manifest.path")
    raw = copied_path.read_bytes() if copied_path.is_file() else b""
    if not copied_path.is_file():
        fail("input-build-0429.source_manifest.path", "copied source manifest is missing")
    if sha(raw) != source_hash:
        fail("input-build-0429.source_manifest", "copied manifest digest differs from input build")
    parsed = load_json(copied_path)
    validate_numbers(parsed, "source-manifest")
    if not isinstance(parsed, dict):
        fail("source-manifest", "expected path-to-SHA-256 object")
    if len(parsed) != files:
        fail("source-manifest", "file count differs from input build")
    for name, value in parsed.items():
        if type(name) is not str:
            fail("source-manifest", "manifest path is not a string")
        valid_hash(value, f"source-manifest.{name}")
    duplicate = root / "input-source-manifest.json"
    if not duplicate.is_file():
        fail("input-source-manifest.json", "copied input manifest is missing")
    if duplicate.read_bytes() != raw:
        fail("input-source-manifest.json", "does not equal the bound source manifest")
    return parsed


def build_receipt_path(root: Path, build: dict[str, Any]) -> tuple[Path, str]:
    declared = text(build["build_receipt"], "input-build-0429.build_receipt")
    candidate = safe_path(root, declared, "input-build-0429.build_receipt")
    if candidate.is_file():
        return candidate, declared
    # The retained 0429 receipt is deliberately renamed in the 0430 bundle so
    # that the old checks directory need not be copied wholesale.  Only this
    # exact historical alias is accepted; arbitrary fallback paths are not.
    if declared == "checks/release-build.json":
        alias = root / "input-build-receipt-0429.json"
        if alias.is_file():
            return alias, "input-build-receipt-0429.json"
    fail("input-build-0429.build_receipt", "declared build receipt is missing")


def verify_build(root: Path, build: dict[str, Any], protocol: dict[str, Any],
                 source: dict[str, str]) -> dict[str, Any]:
    exact_keys(build, BUILD_REQUIRED_FIELDS, "input-build-0429")
    validate_numbers(build, "input-build-0429")
    for field in ("revision", "baseline_revision"):
        valid_revision(build[field], f"input-build-0429.{field}")
    for field in (
        "build_receipt_sha256", "protocol_sha256", "planned_checks_sha256",
        "machine_sha256", "capture_driver_sha256", "verifier_sha256",
        "replay_verifier_sha256", "probe_sha256", "summary_driver_sha256",
        "strict_comparison_driver_sha256",
    ):
        valid_hash(build[field], f"input-build-0429.{field}")
    additional = build["additional_artifact_sha256"]
    if not isinstance(additional, dict) or not additional:
        fail("input-build-0429.additional_artifact_sha256", "expected nonempty artifact hash map")
    for name, value in additional.items():
        text(name, "input-build-0429.additional_artifact_sha256.path")
        if Path(name).is_absolute() or ".." in Path(name).parts:
            fail("input-build-0429.additional_artifact_sha256.path", "artifact path escapes source bundle")
        valid_hash(value, f"input-build-0429.additional_artifact_sha256.{name}")
    amendment = build["validation_amendment"]
    exact_keys(amendment, ("path", "sha256"), "input-build-0429.validation_amendment")
    text(amendment["path"], "input-build-0429.validation_amendment.path")
    valid_hash(amendment["sha256"], "input-build-0429.validation_amendment.sha256")
    valid_hash(build["binary_sha256"], "input-build-0429.binary_sha256")
    compare(build["revision"], protocol["source_revision"], "input-build-0429.revision")
    compare(build["binary_sha256"], protocol["binary_sha256"], "input-build-0429.binary_sha256")
    compare(build["binary_bytes"], protocol["binary_bytes"], "input-build-0429.binary_bytes")
    compare(build["original_binary"], protocol["original_binary"],
            "input-build-0429.original_binary")
    compare(build["capture_binary"], "/tmp/litchi-goal-0429-binaries/litchi-perf-baseline",
            "input-build-0429.capture_binary")
    compare(digest(root / "verify-report.py"), build["verifier_sha256"],
            "input-build-0429.verifier_sha256")
    compare(build["baseline_revision"], "995e217ba5624d6e9926a7bf2677191067f2bd6c",
            "input-build-0429.baseline_revision")
    compare(build["rust_toolchain"], "1.98.1", "input-build-0429.rust_toolchain")
    compare(build["profile"], "release", "input-build-0429.profile")
    compare(build["scope"],
            "one frozen provider baseline build with ZIP short-read correctness enabler; no causal before/after speedup",
            "input-build-0429.scope")
    compare(build["source_manifest"], {
        "path": "sources/1f85ca448f9fd4d970576870e10d699ddfc6b12ad38f0e70ed2be733fe43f122.json",
        "sha256": "1f85ca448f9fd4d970576870e10d699ddfc6b12ad38f0e70ed2be733fe43f122",
        "files": 6634,
    }, "input-build-0429.source_manifest")
    compare(source, load_json(safe_path(root, build["source_manifest"]["path"])),
            "input-build-0429.source_manifest.contents")

    build_receipt, receipt_name = build_receipt_path(root, build)
    if digest(build_receipt) != build["build_receipt_sha256"]:
        fail("input-build-0429.build_receipt_sha256", "copied build receipt digest differs")
    receipt = load_json(build_receipt)
    validate_numbers(receipt, "input-build-receipt-0429")
    required = (
        "change", "argv", "cwd", "revision", "driver_sha256", "environment",
        "source_scope", "started_utc", "source_before", "status", "exit_code",
        "finished_utc", "source_after", "source_unchanged", "log",
    )
    exact_keys(receipt, required, "input-build-receipt-0429")
    compare(receipt["change"], 429, "input-build-receipt-0429.change")
    compare(receipt["revision"], build["revision"], "input-build-receipt-0429.revision")
    compare(receipt["status"], "pass", "input-build-receipt-0429.status")
    compare(receipt["exit_code"], 0, "input-build-receipt-0429.exit_code")
    compare(receipt["source_before"], source_ref(build), "input-build-receipt-0429.source_before")
    compare(receipt["source_after"], source_ref(build), "input-build-receipt-0429.source_after")
    compare(receipt["source_unchanged"], True, "input-build-receipt-0429.source_unchanged")
    argv = receipt["argv"]
    if type(argv) is not list or argv[:2] != ["env", "RUSTFLAGS=-Cforce-frame-pointers=yes"]:
        fail("input-build-receipt-0429.argv", "release frame-pointer build command is not bound")
    required_args = {
        "CARGO_PROFILE_RELEASE_DEBUG=1", "cargo", "build", "--locked", "--release",
        "--manifest-path", "tools/perf-baseline/Cargo.toml", "--bin",
        "litchi-perf-baseline",
    }
    if not required_args.issubset(argv):
        fail("input-build-receipt-0429.argv", "missing pinned release build argument")
    compare(receipt["environment"], {
        "RUSTUP_TOOLCHAIN": "1.98.1", "CARGO_BUILD_JOBS": "4",
        "CARGO_INCREMENTAL": "0", "PYTHONDONTWRITEBYTECODE": "1",
    }, "input-build-receipt-0429.environment")
    expected_driver = build["additional_artifact_sha256"].get("check.py")
    if expected_driver is not None:
        compare(receipt["driver_sha256"], expected_driver,
                "input-build-receipt-0429.driver_sha256")
    log = receipt["log"]
    exact_keys(log, ("path", "bytes", "sha256"), "input-build-receipt-0429.log")
    log_name = text(log["path"], "input-build-receipt-0429.log.path")
    log_bytes = uint(log["bytes"], "input-build-receipt-0429.log.bytes")
    log_hash = valid_hash(log["sha256"], "input-build-receipt-0429.log.sha256")
    log_candidate = safe_path(root, log_name)
    if not log_candidate.is_file() and not log_candidate.with_name(log_candidate.name + ".gz").is_file():
        # This alias is pinned to the receipt above and is retained as a
        # compressed historical build log after the 0429 cleanup.
        if log_name != "checks/release-build.log":
            fail("input-build-receipt-0429.log.path", "build log is missing")
        log_name = "input-build-log-0429.log.gz"
    raw_log, _physical, _stored = read_payload(root, log_name, "input-build-receipt-0429.log")
    if sha(raw_log) != log_hash or len(raw_log) != log_bytes:
        fail("input-build-receipt-0429.log", "lossless build log differs from receipt")
    machine = root / "machine-prior.json"
    if digest(machine) != build["machine_sha256"]:
        fail("machine-prior.json", "machine record is not the input-build record")
    current = require_regular(root, "machine-current.json")
    current_row = load_json(current)
    validate_numbers(current_row, "machine-current")
    if current_row.get("prior_machine_sha256") != build["machine_sha256"]:
        fail("machine-current.prior_machine_sha256", "does not bind prior machine record")
    return {"receipt": receipt_name, "log": log_name}


def source_ref(build: dict[str, Any]) -> dict[str, Any]:
    return copy.deepcopy(build["source_manifest"])


def check_report_path(root: Path, provider: str) -> Path:
    raw = root / "profiles" / f"{provider}.json"
    if raw.is_file():
        return raw
    compressed = root / "profiles" / f"{provider}.json.gz"
    if compressed.is_file():
        # Report JSON is deliberately retained raw for direct verifier
        # invocation.  A compressed-only report cannot be passed by pathname
        # to the bound verifier without materializing an untrusted temporary
        # file, so reject it explicitly.
        fail(f"profiles/{provider}.json", "report JSON must remain directly readable")
    fail(f"profiles/{provider}.json", "report is missing")


def run_bound_report_verifier(root: Path, report_path: Path) -> dict[str, Any]:
    verifier = require_regular(root, "verify-report.py")
    if verifier.resolve().parent != root.resolve():
        fail("verify-report.py", "bound verifier is outside the copied bundle")
    try:
        result = subprocess.run(
            [sys.executable, "-B", str(verifier), str(report_path)],
            cwd=root,
            env={**os.environ, "PYTHONDONTWRITEBYTECODE": "1"},
            capture_output=True,
            timeout=300,
            check=False,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        fail("verify-report.py", f"bound verifier could not run: {exc}")
    if result.returncode != 0:
        stderr = result.stderr.decode("utf-8", "replace").strip()
        fail(str(report_path.relative_to(root)), f"bound verifier rejected report: {stderr[:400]}")
    try:
        parsed = json.loads(result.stdout.decode("utf-8"), object_pairs_hook=reject_duplicate_pairs,
                            parse_constant=reject_nonfinite)
    except (UnicodeError, json.JSONDecodeError, VerificationError) as exc:
        fail("verify-report.py.stdout", f"bound verifier returned invalid JSON: {exc}")
    if not isinstance(parsed, dict) or parsed.get("status") != "valid":
        fail("verify-report.py.stdout", "bound verifier did not return status valid")
    return parsed


def report_identity(report: dict[str, Any], provider: str, build: dict[str, Any],
                    protocol: dict[str, Any]) -> None:
    path = f"profiles/{provider}.json"
    validate_numbers(report, path)
    if report.get("schema") != "pptx_provider_lifecycle_v1":
        fail(f"{path}.schema", "expected provider lifecycle schema")
    compare(report.get("provider"), provider, f"{path}.provider")
    compare(report.get("corpus"), protocol["corpus"], f"{path}.corpus")
    compare(report.get("samples"), protocol["samples"], f"{path}.samples")
    compare(report.get("warmup"), protocol["warmup"], f"{path}.warmup")
    compare(report.get("checked_iteration_count"), protocol["samples"] + protocol["warmup"],
            f"{path}.checked_iteration_count")
    compare(report.get("source_revision"), build["revision"], f"{path}.source_revision")
    compare(report.get("binary_sha256"), build["binary_sha256"], f"{path}.binary_sha256")
    compare(report.get("binary_bytes"), build["binary_bytes"], f"{path}.binary_bytes")
    valid_hash(report.get("binary_sha256"), f"{path}.binary_sha256")
    for field in ("source_archive_sha256", "destination_archive_sha256", "expected_output_sha256"):
        valid_hash(report.get(field), f"{path}.{field}")
    for field in ("source_archive_bytes", "destination_archive_bytes", "expected_output_bytes"):
        uint(report.get(field), f"{path}.{field}")
    if not isinstance(report.get("corpus_manifest"), dict):
        fail(f"{path}.corpus_manifest", "expected object")
    gates = report.get("gates")
    if not isinstance(gates, dict) or not gates:
        fail(f"{path}.gates", "expected nonempty gate object")
    for name, value in gates.items():
        if type(value) is not bool or value is not True:
            fail(f"{path}.gates.{name}", "all retained output gates must be true")
    rows = report.get("samples_raw")
    if not isinstance(rows, list) or len(rows) != protocol["samples"]:
        fail(f"{path}.samples_raw", "row count differs from the pinned sample count")
    for index, row in enumerate(rows):
        if not isinstance(row, dict):
            fail(f"{path}.samples_raw[{index}]", "expected object")
        if row.get("sample_index") != index:
            fail(f"{path}.samples_raw[{index}].sample_index", "rows are not ordered")
        if row.get("exact_output_verified") is not True:
            fail(f"{path}.samples_raw[{index}].exact_output_verified", "output gate is false")
        valid_hash(row.get("output_sha256"), f"{path}.samples_raw[{index}].output_sha256")
        uint(row.get("output_bytes"), f"{path}.samples_raw[{index}].output_bytes")
        if row["output_sha256"] != report["expected_output_sha256"]:
            fail(f"{path}.samples_raw[{index}].output_sha256", "differs from expected output")
        if row["output_bytes"] != report["expected_output_bytes"]:
            fail(f"{path}.samples_raw[{index}].output_bytes", "differs from expected output")


def verify_profile_receipt(root: Path, receipt_name: str, provider: str,
                           build: dict[str, Any], protocol: dict[str, Any],
                           source: dict[str, Any]) -> tuple[dict[str, Any], Path, dict[str, Any]]:
    receipt_path = safe_path(root, receipt_name, f"profile-index[{provider}]")
    receipt = load_json(receipt_path)
    validate_numbers(receipt, f"profiles/{provider}-receipt")
    exact_keys(receipt, RECEIPT_REQUIRED_FIELDS, f"profiles/{provider}-receipt")
    compare(receipt["provider"], provider, f"profiles/{provider}-receipt.provider")
    compare(receipt["source"], source, f"profiles/{provider}-receipt.source")
    compare(receipt["binary_sha256"], protocol["binary_sha256"], f"profiles/{provider}-receipt.binary_sha256")
    compare(receipt["protocol_sha256"], digest(root / "profile-protocol.json"),
            f"profiles/{provider}-receipt.protocol_sha256")
    compare(receipt["driver_sha256"], digest(root / "capture-profiles.py"),
            f"profiles/{provider}-receipt.driver_sha256")
    compare(receipt["verifier_sha256"], digest(root / "verify-report.py"),
            f"profiles/{provider}-receipt.verifier_sha256")
    compare(receipt["status"], "pass", f"profiles/{provider}-receipt.status")
    compare(receipt["record_exit_code"], 0, f"profiles/{provider}-receipt.record_exit_code")
    if not isinstance(receipt["argv"], list) or not all(type(item) is str for item in receipt["argv"]):
        fail(f"profiles/{provider}-receipt.argv", "expected string argument vector")
    argv = receipt["argv"]
    required_args = [
        "taskset", "-c", str(protocol["cpu"]), "perf", "record", "-e", protocol["event"],
        "-F", str(protocol["frequency_hz"]), "--call-graph", protocol["call_graph"],
        "provider-lifecycle", "--corpus", protocol["corpus"], "--provider", provider,
        "--samples", str(protocol["samples"]), "--warmup", str(protocol["warmup"]),
        "--source-revision", protocol["source_revision"], "--output",
    ]
    if not all(item in argv for item in required_args):
        fail(f"profiles/{provider}-receipt.argv", "capture command does not match frozen protocol")
    output_index = argv.index("--output")
    if output_index + 1 >= len(argv) or Path(argv[output_index + 1]).name != f"{provider}.json":
        fail(f"profiles/{provider}-receipt.argv", "output argument is not the retained report")
    if protocol["capture_binary"] not in argv:
        fail(f"profiles/{provider}-receipt.argv", "capture executable does not match protocol")
    artifacts = receipt["artifacts"]
    if not isinstance(artifacts, dict):
        fail(f"profiles/{provider}-receipt.artifacts", "expected object")
    expected_names = {f"profiles/{provider}{suffix}" for suffix in PROFILE_ARTIFACT_SUFFIXES}
    if set(artifacts) != expected_names:
        fail(f"profiles/{provider}-receipt.artifacts", f"expected exactly {sorted(expected_names)}")
    for name, metadata in artifacts.items():
        check_custody_artifact(root, name, metadata, f"profiles/{provider}-receipt.artifacts.{name}")
    report_path = check_report_path(root, provider)
    report = load_json(report_path)
    report_identity(report, provider, build, protocol)
    verifier_result = run_bound_report_verifier(root, report_path)
    return report, report_path, verifier_result


def cross_provider_invariants(reports: dict[str, dict[str, Any]]) -> None:
    if set(reports) != {"bytes", "file"}:
        fail("reports", "exactly bytes and file controls are required")
    first = reports["bytes"]
    for provider, report in reports.items():
        for field in REPORT_SEMANTIC_FIELDS:
            if field == "schema":
                continue
            if report.get(field) != first.get(field):
                fail(f"reports.{provider}.{field}", "cross-provider corpus/output semantics differ")
        rows = report["samples_raw"]
        first_rows = first["samples_raw"]
        for index, row in enumerate(rows):
            reference = first_rows[index]
            for field in ("sample_index", "exact_output_verified", "output_sha256", "output_bytes"):
                if row.get(field) != reference.get(field):
                    fail(f"reports.{provider}.samples_raw[{index}].{field}",
                         "cross-provider output identity differs")
            labels = [item.get("label") for item in row.get("phases", [])]
            reference_labels = [item.get("label") for item in reference.get("phases", [])]
            if labels != reference_labels:
                fail(f"reports.{provider}.samples_raw[{index}].phases", "phase labels differ")
    for report in reports.values():
        if report["expected_output_sha256"] != first["expected_output_sha256"]:
            fail("reports.expected_output_sha256", "provider outputs differ")
        if report["expected_output_bytes"] != first["expected_output_bytes"]:
            fail("reports.expected_output_bytes", "provider output sizes differ")


def verify_checks(root: Path, build: dict[str, Any], protocol: dict[str, Any],
                  source: dict[str, Any], allow_pending: bool) -> int:
    expected_path = root / "expected-checks.json"
    expected: dict[str, Any] = {}
    if expected_path.is_file():
        expected = load_json(expected_path)
        validate_numbers(expected, "expected-checks")
        if not isinstance(expected, dict):
            fail("expected-checks", "expected object")
        for name, status in expected.items():
            text(name, "expected-checks.name")
            if status not in ("pass", "failed"):
                fail(f"expected-checks.{name}", "status must be pass or failed")
    receipt_paths = sorted((root / "checks").glob("*.json")) if (root / "checks").is_dir() else []
    actual: dict[str, dict[str, Any]] = {}
    for path in receipt_paths:
        row = load_json(path)
        if isinstance(row, dict) and "source_before" in row:
            actual[path.stem] = row
    names = set(expected) if expected else set(actual)
    for name in sorted(names):
        path = root / "checks" / f"{name}.json"
        if not path.is_file():
            if allow_pending:
                continue
            fail(f"checks/{name}.json", "planned command receipt is missing")
        row = actual.get(name)
        if row is None:
            fail(f"checks/{name}.json", "not a source-custody command receipt")
        status = expected.get(name, "pass")
        validate_numbers(row, f"checks/{name}")
        required = ("change", "argv", "cwd", "revision", "driver_sha256", "environment",
                    "source_scope", "started_utc", "source_before", "status", "exit_code",
                    "finished_utc", "source_after", "source_unchanged", "log")
        exact_keys(row, required, f"checks/{name}",
                   optional=("passed_tests", "failed_tests", "ignored_tests"))
        compare(row["change"], 430, f"checks/{name}.change")
        compare(row["driver_sha256"], digest(root / "check.py"), f"checks/{name}.driver_sha256")
        compare(row["source_before"], source, f"checks/{name}.source_before")
        compare(row["source_after"], source, f"checks/{name}.source_after")
        compare(row["source_unchanged"], True, f"checks/{name}.source_unchanged")
        compare(row["status"], status, f"checks/{name}.status")
        if status == "pass":
            compare(row["exit_code"], 0, f"checks/{name}.exit_code")
        log = row["log"]
        exact_keys(log, ("path", "bytes", "sha256"), f"checks/{name}.log")
        check_custody_artifact(root, log["path"], log, f"checks/{name}.log")
        if row["revision"] not in (protocol["baseline_revision"], build["revision"]):
            fail(f"checks/{name}.revision", "receipt revision is outside the bound evidence pair")
    if expected:
        unexpected = set(actual) - set(expected)
        if unexpected:
            fail("checks", f"unplanned command receipts present: {sorted(unexpected)}")
    return len(names)


def validate_profile_summary(root: Path, protocol: dict[str, Any], allow_pending: bool) -> None:
    script = root / "profile-analysis.py"
    summary_path = root / "profile-summary.json"
    if not script.is_file() or not summary_path.is_file():
        if allow_pending:
            return
        fail("profile-summary", "profile-analysis.py and profile-summary.json are required")
    script_text = script.read_text(encoding="utf-8")
    for marker in ("cycles:u", "profile-summary.json", "parse_perf_script", "bytes", "file"):
        if marker not in script_text:
            fail("profile-analysis.py", f"decoded-stack analysis lacks marker {marker!r}")
    summary = load_json(summary_path)
    validate_numbers(summary, "profile-summary")
    if not isinstance(summary, dict):
        fail("profile-summary", "expected object")
    exact_keys(summary, (
        "status", "schema", "change", "protocol", "artifact_binding", "source_audit",
        "profiles", "scope", "historical_0429_comparison", "performance_claim",
        "analysis_script_sha256",
    ), "profile-summary")
    compare(summary["status"], "pass", "profile-summary.status")
    compare(summary["schema"], "pptx_frame_pointer_profile_analysis_v1", "profile-summary.schema")
    compare(summary["change"], 430, "profile-summary.change")
    compare(summary["performance_claim"], None, "profile-summary.performance_claim")
    compare(summary["analysis_script_sha256"], digest(script),
            "profile-summary.analysis_script_sha256")
    scope = text(summary["scope"], "profile-summary.scope")
    if "cycles:u" not in scope or "sample" not in scope.lower() or "warmup" not in scope.lower():
        fail("profile-summary.scope", "does not state sampled cycles scope")

    summary_protocol = summary["protocol"]
    exact_keys(summary_protocol, (
        "path", "sha256", "binary_sha256", "source_revision", "event", "call_graph",
        "frequency_hz", "corpus", "providers", "samples", "warmup",
    ), "profile-summary.protocol")
    compare(summary_protocol["path"], "profile-protocol.json", "profile-summary.protocol.path")
    compare(summary_protocol["sha256"], digest(root / "profile-protocol.json"),
            "profile-summary.protocol.sha256")
    compare(summary_protocol["binary_sha256"], EXPECTED_PROTOCOL["binary_sha256"],
            "profile-summary.protocol.binary_sha256")
    compare(summary_protocol["source_revision"], EXPECTED_PROTOCOL["source_revision"],
            "profile-summary.protocol.source_revision")
    compare(summary_protocol["event"], EXPECTED_PROTOCOL["event"], "profile-summary.protocol.event")
    compare(summary_protocol["call_graph"], EXPECTED_PROTOCOL["call_graph"],
            "profile-summary.protocol.call_graph")
    compare(summary_protocol["frequency_hz"], EXPECTED_PROTOCOL["frequency_hz"],
            "profile-summary.protocol.frequency_hz")
    compare(summary_protocol["corpus"], EXPECTED_PROTOCOL["corpus"], "profile-summary.protocol.corpus")
    compare(summary_protocol["providers"], EXPECTED_PROTOCOL["providers"],
            "profile-summary.protocol.providers")
    compare(summary_protocol["samples"], EXPECTED_PROTOCOL["samples"], "profile-summary.protocol.samples")
    compare(summary_protocol["warmup"], EXPECTED_PROTOCOL["warmup"], "profile-summary.protocol.warmup")

    binding = summary["artifact_binding"]
    exact_keys(binding, ("protocol", "profile_index", "receipts"), "profile-summary.artifact_binding")
    for field, expected_name in (("protocol", "profile-protocol.json"),
                                 ("profile_index", "profile-index.json")):
        item = binding[field]
        exact_keys(item, ("path", "sha256", "bytes"), f"profile-summary.artifact_binding.{field}")
        compare(item["path"], expected_name, f"profile-summary.artifact_binding.{field}.path")
        target = root / expected_name
        compare(item["sha256"], digest(target), f"profile-summary.artifact_binding.{field}.sha256")
        compare(item["bytes"], target.stat().st_size, f"profile-summary.artifact_binding.{field}.bytes")
    if not isinstance(binding["receipts"], list) or len(binding["receipts"]) != 2:
        fail("profile-summary.artifact_binding.receipts", "expected two provider receipts")
    for index, item in enumerate(binding["receipts"]):
        path = f"profile-summary.artifact_binding.receipts[{index}]"
        exact_keys(item, ("provider", "path", "sha256", "bytes", "artifacts", "report"), path)
        provider = text(item["provider"], f"{path}.provider")
        if provider not in ("bytes", "file"):
            fail(f"{path}.provider", "unknown provider")
        receipt_path = safe_path(root, item["path"], f"{path}.path")
        compare(item["sha256"], digest(receipt_path), f"{path}.sha256")
        compare(item["bytes"], receipt_path.stat().st_size, f"{path}.bytes")
        if not isinstance(item["artifacts"], dict):
            fail(f"{path}.artifacts", "expected artifact map")
        for artifact_name, artifact_row in item["artifacts"].items():
            check_custody_artifact(root, artifact_name, artifact_row, f"{path}.artifacts.{artifact_name}")
        report = item["report"]
        exact_keys(report, (
            "schema", "corpus", "samples", "warmup", "checked_iteration_count",
            "source_revision", "binary_sha256",
        ), f"{path}.report")
        compare(report["schema"], "pptx_provider_lifecycle_v1", f"{path}.report.schema")
        compare(report["corpus"], EXPECTED_PROTOCOL["corpus"], f"{path}.report.corpus")
        compare(report["samples"], EXPECTED_PROTOCOL["samples"], f"{path}.report.samples")
        compare(report["warmup"], EXPECTED_PROTOCOL["warmup"], f"{path}.report.warmup")
        compare(report["checked_iteration_count"], 103, f"{path}.report.checked_iteration_count")
        compare(report["source_revision"], EXPECTED_PROTOCOL["source_revision"],
                f"{path}.report.source_revision")
        compare(report["binary_sha256"], EXPECTED_PROTOCOL["binary_sha256"],
                f"{path}.report.binary_sha256")

    source_audit = summary["source_audit"]
    if not isinstance(source_audit, dict):
        fail("profile-summary.source_audit", "expected object")
    source_manifest = source_audit.get("source_manifest")
    if not isinstance(source_manifest, dict):
        fail("profile-summary.source_audit.source_manifest", "expected object")
    for field in ("path", "sha256", "files"):
        if field not in source_manifest:
            fail("profile-summary.source_audit.source_manifest", f"missing {field}")
    compare(source_manifest.get("path"),
            "sources/1f85ca448f9fd4d970576870e10d699ddfc6b12ad38f0e70ed2be733fe43f122.json",
            "profile-summary.source_audit.source_manifest.path")
    compare(source_manifest.get("sha256"),
            "1f85ca448f9fd4d970576870e10d699ddfc6b12ad38f0e70ed2be733fe43f122",
            "profile-summary.source_audit.source_manifest.sha256")
    compare(source_manifest.get("files"), 6634,
            "profile-summary.source_audit.source_manifest.files")
    compare(source_manifest.get("bytes"),
            (root / "sources/1f85ca448f9fd4d970576870e10d699ddfc6b12ad38f0e70ed2be733fe43f122.json").stat().st_size,
            "profile-summary.source_audit.source_manifest.bytes")
    input_source = source_audit.get("input_source_manifest")
    if not isinstance(input_source, dict) or input_source.get("path") != "input-source-manifest.json":
        fail("profile-summary.source_audit.input_source_manifest", "input source custody is not bound")
    compare(input_source.get("sha256"), digest(root / "input-source-manifest.json"),
            "profile-summary.source_audit.input_source_manifest.sha256")
    compare(input_source.get("bytes"), (root / "input-source-manifest.json").stat().st_size,
            "profile-summary.source_audit.input_source_manifest.bytes")
    input_build = source_audit.get("input_build")
    if not isinstance(input_build, dict):
        fail("profile-summary.source_audit.input_build", "expected object")
    compare(input_build.get("path"), "input-build-0429.json", "profile-summary.source_audit.input_build.path")
    compare(input_build.get("sha256"), digest(root / "input-build-0429.json"),
            "profile-summary.source_audit.input_build.sha256")
    compare(input_build.get("revision"), EXPECTED_PROTOCOL["source_revision"],
            "profile-summary.source_audit.input_build.revision")
    compare(input_build.get("binary_sha256"), EXPECTED_PROTOCOL["binary_sha256"],
            "profile-summary.source_audit.input_build.binary_sha256")
    for check_name in ("capture_check", "symbol_identity_check"):
        check = source_audit.get(check_name)
        if not isinstance(check, dict) or not isinstance(check.get("log"), dict):
            fail(f"profile-summary.source_audit.{check_name}", "check receipt/log binding is missing")
        check_path = text(check.get("path"), f"profile-summary.source_audit.{check_name}.path")
        check_file = safe_path(root, check_path, f"profile-summary.source_audit.{check_name}.path")
        compare(check.get("sha256"), digest(check_file), f"profile-summary.source_audit.{check_name}.sha256")
        check_log = check["log"]
        check_custody_artifact(root, check_log["path"], check_log,
                               f"profile-summary.source_audit.{check_name}.log")

    profiles = summary["profiles"]
    if type(profiles) is not list or len(profiles) != 2:
        fail("profile-summary.profiles", "expected one decoded analysis row per provider")
    names = set()
    for index, row in enumerate(profiles):
        path = f"profile-summary.profiles[{index}]"
        if not isinstance(row, dict):
            fail(path, "expected object")
        exact_keys(row, (
            "provider", "receipt", "receipt_sha256", "artifact_binding", "samples",
            "recorded_samples", "exclusive_ancestry_counts", "exclusive_ancestry_percent",
            "ancestry_coverage", "iteration_api_ancestry_counts_nonexclusive",
            "deflate_intersections", "top_leaf_symbols_by_ancestry", "parse_quality",
            "stack_input", "scope_note",
        ), path)
        provider = text(row["provider"], f"{path}.provider")
        if provider not in ("bytes", "file"):
            fail(f"{path}.provider", "unknown provider")
        names.add(provider)
        samples = uint(row["samples"], f"{path}.samples")
        if samples == 0:
            fail(f"{path}.samples", "decoded stack sample count is zero")
        compare(row["recorded_samples"], samples, f"{path}.recorded_samples")
        receipt_name = f"profiles/{provider}-receipt.json"
        compare(row["receipt"], receipt_name, f"{path}.receipt")
        receipt_path = root / receipt_name
        compare(row["receipt_sha256"], digest(receipt_path), f"{path}.receipt_sha256")
        artifacts = row["artifact_binding"]
        if not isinstance(artifacts, dict):
            fail(f"{path}.artifact_binding", "expected artifact map")
        for artifact_name, artifact_row in artifacts.items():
            check_custody_artifact(root, artifact_name, artifact_row,
                                   f"{path}.artifact_binding.{artifact_name}")
        counts = row["exclusive_ancestry_counts"]
        percents = row["exclusive_ancestry_percent"]
        if not isinstance(counts, dict) or not isinstance(percents, dict):
            fail(path, "ancestry counts and percentages must be objects")
        if set(counts) != {"iteration", "corpus_setup", "other_or_unresolved"} or set(counts) != set(percents):
            fail(path, "ancestry count and percentage categories differ")
        total = 0
        for category, count in counts.items():
            total += uint(count, f"{path}.exclusive_ancestry_counts.{category}")
            percentage = finite_number(percents[category],
                                       f"{path}.exclusive_ancestry_percent.{category}")
            expected = count * 100.0 / samples
            if abs(percentage - expected) > 1e-9:
                fail(f"{path}.exclusive_ancestry_percent.{category}",
                     "does not derive from decoded stack count")
        if total != samples:
            fail(f"{path}.exclusive_ancestry_counts", "counts do not sum to samples")
        coverage = row["ancestry_coverage"]
        exact_keys(coverage, (
            "resolved_iteration_or_setup_samples", "resolved_iteration_or_setup_percent",
            "unresolved_or_ambiguous_samples", "unresolved_or_ambiguous_percent",
        ), f"{path}.ancestry_coverage")
        resolved = counts["iteration"] + counts["corpus_setup"]
        compare(coverage["resolved_iteration_or_setup_samples"], resolved,
                f"{path}.ancestry_coverage.resolved_iteration_or_setup_samples")
        compare(coverage["unresolved_or_ambiguous_samples"], counts["other_or_unresolved"],
                f"{path}.ancestry_coverage.unresolved_or_ambiguous_samples")
        if abs(finite_number(coverage["resolved_iteration_or_setup_percent"],
                             f"{path}.ancestry_coverage.resolved_iteration_or_setup_percent")
               - resolved * 100.0 / samples) > 1e-9:
            fail(f"{path}.ancestry_coverage.resolved_iteration_or_setup_percent", "does not derive from counts")
        if abs(finite_number(coverage["unresolved_or_ambiguous_percent"],
                             f"{path}.ancestry_coverage.unresolved_or_ambiguous_percent")
               - counts["other_or_unresolved"] * 100.0 / samples) > 1e-9:
            fail(f"{path}.ancestry_coverage.unresolved_or_ambiguous_percent", "does not derive from counts")
        nonexclusive = row["iteration_api_ancestry_counts_nonexclusive"]
        if not isinstance(nonexclusive, dict):
            fail(f"{path}.iteration_api_ancestry_counts_nonexclusive", "expected object")
        for category, count in nonexclusive.items():
            uint(count, f"{path}.iteration_api_ancestry_counts_nonexclusive.{category}")
        leaves = row["top_leaf_symbols_by_ancestry"]
        if not isinstance(leaves, dict) or set(leaves) != set(counts):
            fail(f"{path}.top_leaf_symbols_by_ancestry", "leaf categories differ from counts")
        for category, entries in leaves.items():
            if not isinstance(entries, list):
                fail(f"{path}.top_leaf_symbols_by_ancestry.{category}", "expected list")
            previous = None
            for entry_index, entry in enumerate(entries):
                if not isinstance(entry, list) or len(entry) != 2:
                    fail(f"{path}.top_leaf_symbols_by_ancestry.{category}[{entry_index}]",
                         "expected [symbol, count]")
                text(entry[0], f"{path}.top_leaf_symbols_by_ancestry.{category}[{entry_index}][0]")
                current = uint(entry[1],
                               f"{path}.top_leaf_symbols_by_ancestry.{category}[{entry_index}][1]")
                if previous is not None and current > previous:
                    fail(f"{path}.top_leaf_symbols_by_ancestry.{category}",
                         "leaf counts are not descending")
                previous = current
        intersections = row["deflate_intersections"]
        if not isinstance(intersections, dict):
            fail(f"{path}.deflate_intersections", "expected object")
        for key, value in intersections.items():
            if key in ("marker",):
                text(value, f"{path}.deflate_intersections.{key}")
            elif key == "deflate_medium_by_exclusive_ancestry":
                if not isinstance(value, dict) or set(value) != set(counts):
                    fail(f"{path}.deflate_intersections.{key}", "categories differ from ancestry counts")
                for category, count in value.items():
                    uint(count, f"{path}.deflate_intersections.{key}.{category}")
            else:
                uint(value, f"{path}.deflate_intersections.{key}")
        parse_quality = row["parse_quality"]
        if not isinstance(parse_quality, dict):
            fail(f"{path}.parse_quality", "expected object")
        for key in ("sample_blocks", "blocks_with_decoded_frames", "blocks_without_decoded_frames",
                    "malformed_frame_lines", "ambiguous_ancestry_blocks", "diagnostic_lines"):
            uint(parse_quality.get(key), f"{path}.parse_quality.{key}")
        compare(parse_quality["sample_blocks"], samples, f"{path}.parse_quality.sample_blocks")
        if parse_quality["blocks_with_decoded_frames"] + parse_quality["blocks_without_decoded_frames"] != samples:
            fail(f"{path}.parse_quality", "decoded and unresolved blocks do not sum to samples")
        stack_input = row["stack_input"]
        exact_keys(stack_input, ("path", "accepted_storage"), f"{path}.stack_input")
        compare(stack_input["path"], f"profiles/{provider}-script.log", f"{path}.stack_input.path")
        text(stack_input["accepted_storage"], f"{path}.stack_input.accepted_storage")
        text(row["scope_note"], f"{path}.scope_note")
    if names != {"bytes", "file"}:
        fail("profile-summary.profiles", "provider rows are not bytes and file")

    historical = summary["historical_0429_comparison"]
    if not isinstance(historical, dict):
        fail("profile-summary.historical_0429_comparison", "expected object")
    if historical.get("available") is True:
        compare(historical.get("path"), "input-profile-summary-0429.json",
                "profile-summary.historical_0429_comparison.path")
        historical_file = root / "input-profile-summary-0429.json"
        compare(historical.get("sha256"), digest(historical_file),
                "profile-summary.historical_0429_comparison.sha256")
        rows = historical.get("profiles")
        if not isinstance(rows, list) or len(rows) != 2:
            fail("profile-summary.historical_0429_comparison.profiles", "expected two provider comparisons")
        for index, row in enumerate(rows):
            path = f"profile-summary.historical_0429_comparison.profiles[{index}]"
            if not isinstance(row, dict) or row.get("provider") not in ("bytes", "file"):
                fail(path, "unknown historical provider")
            for field in ("current_samples", "historical_samples"):
                uint(row.get(field), f"{path}.{field}")
            for field in ("current_exclusive_ancestry_counts", "historical_exclusive_ancestry_counts"):
                values = row.get(field)
                if not isinstance(values, dict):
                    fail(f"{path}.{field}", "expected ancestry count object")
                for category, value in values.items():
                    uint(value, f"{path}.{field}.{category}")
            delta = row.get("delta_exclusive_ancestry_counts")
            if delta is not None:
                if not isinstance(delta, dict):
                    fail(f"{path}.delta_exclusive_ancestry_counts", "expected signed count object")
                for category, value in delta.items():
                    if type(value) is not int or abs(value) > U64_MAX:
                        fail(f"{path}.delta_exclusive_ancestry_counts.{category}",
                             "invalid signed count delta")
    elif historical.get("available") is not False:
        fail("profile-summary.historical_0429_comparison.available", "expected boolean")


def run_bound_analysis(root: Path, allow_pending: bool) -> dict[str, Any] | None:
    analysis = root / "profile-analysis.py"
    summary = root / "profile-summary.json"
    if not analysis.is_file() or not summary.is_file():
        if allow_pending:
            return None
        fail("profile-analysis.py", "bound analysis driver is missing")
    if analysis.resolve().parent != root.resolve():
        fail("profile-analysis.py", "bound analysis driver is outside the bundle")
    try:
        result = subprocess.run(
            [sys.executable, "-B", str(analysis), "--check"],
            cwd=root,
            env={**os.environ, "PYTHONDONTWRITEBYTECODE": "1"},
            capture_output=True,
            timeout=300,
            check=False,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        fail("profile-analysis.py", f"bound analysis could not run: {exc}")
    if result.returncode != 0:
        stderr = result.stderr.decode("utf-8", "replace").strip()
        fail("profile-analysis.py", f"--check rejected the decoded stacks: {stderr[:400]}")
    try:
        parsed = json.loads(result.stdout.decode("utf-8"), object_pairs_hook=reject_duplicate_pairs,
                            parse_constant=reject_nonfinite)
    except (UnicodeError, json.JSONDecodeError, VerificationError) as exc:
        fail("profile-analysis.py.stdout", f"bound analysis returned invalid JSON: {exc}")
    if not isinstance(parsed, dict) or parsed.get("status") != "pass":
        fail("profile-analysis.py.stdout", "bound analysis did not return status pass")
    return parsed


def verify_compression(root: Path, allow_pending: bool) -> int:
    path = root / "compression.json"
    if not path.is_file():
        if allow_pending:
            return 0
        fail("compression.json", "lossless compression inventory is missing")
    records = load_json(path)
    validate_numbers(records, "compression")
    if not isinstance(records, dict):
        fail("compression", "expected path-to-record object")
    listed = set()
    for name, row in records.items():
        text(name, "compression.path")
        if not name.endswith(".gz"):
            fail(f"compression.{name}", "record key is not a gzip artifact")
        exact_keys(row, ("original_path", "original_sha256", "original_bytes",
                         "stored_sha256", "stored_bytes"), f"compression.{name}")
        compressed = require_regular(root, name, f"compression.{name}")
        stored = compressed.read_bytes()
        compare(row["stored_sha256"], sha(stored), f"compression.{name}.stored_sha256")
        compare(row["stored_bytes"], len(stored), f"compression.{name}.stored_bytes")
        try:
            original = gzip.decompress(stored)
        except (EOFError, gzip.BadGzipFile, zlib.error) as exc:
            fail(f"compression.{name}", f"invalid gzip payload: {exc}")
        compare(row["original_sha256"], sha(original), f"compression.{name}.original_sha256")
        compare(row["original_bytes"], len(original), f"compression.{name}.original_bytes")
        original_name = text(row["original_path"], f"compression.{name}.original_path")
        original_path = safe_path(root, original_name, f"compression.{name}.original_path")
        if original_path.is_file() and original_path.read_bytes() != original:
            fail(f"compression.{name}.original_path", "raw file differs from decompressed bytes")
        listed.add(name)
    physical = {str(path.relative_to(root)) for path in root.rglob("*.gz") if path.is_file()}
    if physical != listed:
        fail("compression", f"gzip inventory differs from physical files: {sorted(physical ^ listed)}")
    return len(listed)


def verify_inventory(root: Path, allow_pending: bool) -> int:
    path = root / "SHA256SUMS"
    if not path.is_file():
        if allow_pending:
            return 0
        fail("SHA256SUMS", "complete bundle inventory is missing")
    entries: dict[str, str] = {}
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except (OSError, UnicodeError) as exc:
        fail("SHA256SUMS", f"cannot read inventory: {exc}")
    for index, line in enumerate(lines):
        if not line:
            continue
        if "  " not in line:
            fail(f"SHA256SUMS[{index}]", "expected two-space digest separator")
        value, name = line.split("  ", 1)
        valid_hash(value, f"SHA256SUMS[{index}].sha256")
        if not name or name in entries or name == "SHA256SUMS":
            fail(f"SHA256SUMS[{index}]", "duplicate or invalid inventory path")
        target = safe_path(root, name, f"SHA256SUMS[{index}].path")
        if not target.is_file() or digest(target) != value:
            fail(f"SHA256SUMS[{index}]", "digest does not match physical artifact")
        entries[name] = value
    physical = {str(path.relative_to(root)) for path in root.rglob("*")
                if path.is_file() and path.name != "SHA256SUMS" and "__pycache__" not in path.parts}
    if set(entries) != physical:
        fail("SHA256SUMS", f"inventory set differs from bundle files: {sorted(physical ^ set(entries))}")
    return len(entries)


def verify_bundle(root: Path, allow_pending: bool = False) -> dict[str, Any]:
    build = load_json(root / "input-build-0429.json")
    validate_numbers(build, "input-build-0429")
    exact_keys(build, BUILD_REQUIRED_FIELDS, "input-build-0429")
    source = verify_source_manifest(root, build)
    protocol = verify_protocol(root, build)
    build_result = verify_build(root, build, protocol, source)
    index_path = root / "profile-index.json"
    profile_index = load_json(index_path)
    validate_numbers(profile_index, "profile-index")
    if profile_index != ["profiles/bytes-receipt.json", "profiles/file-receipt.json"]:
        fail("profile-index", "must retain bytes then file in frozen order")
    reports: dict[str, dict[str, Any]] = {}
    verifier_results: dict[str, dict[str, Any]] = {}
    for provider, receipt_name in zip(("bytes", "file"), profile_index, strict=True):
        report, _path, verifier_result = verify_profile_receipt(
            root, receipt_name, provider, build, protocol, build["source_manifest"])
        reports[provider] = report
        verifier_results[provider] = verifier_result
    cross_provider_invariants(reports)
    checks = verify_checks(root, build, protocol, build["source_manifest"], allow_pending)
    validate_profile_summary(root, protocol, allow_pending)
    analysis_result = run_bound_analysis(root, allow_pending)
    compressed = verify_compression(root, allow_pending)
    inventory = verify_inventory(root, allow_pending)
    return {
        "status": "pass",
        "change": 430,
        "providers": ["bytes", "file"],
        "reports": 2,
        "samples": sum(report["samples"] for report in reports.values()),
        "source_manifest_files": len(source),
        "command_receipts": checks,
        "analysis_replay": analysis_result,
        "compressed_artifacts": compressed,
        "inventory_files": inventory,
        "bound_report_verifier": verifier_results,
        "build_receipt": build_result,
        "performance_claim": None,
    }


def mutate_report(root: Path, provider: str, mutation: str) -> None:
    path = root / "profiles" / f"{provider}.json"
    report = load_json(path)
    if mutation == "numeric-delta":
        report["samples_raw"][0]["timings"]["api_sum_ns"] += 1
    elif mutation == "output":
        report["samples_raw"][0]["output_bytes"] += 1
    else:
        raise AssertionError(mutation)
    path.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")


def expect_rejected(root: Path, label: str) -> None:
    try:
        verify_bundle(root, allow_pending=False)
    except (VerificationError, OSError, subprocess.SubprocessError):
        return
    raise VerificationError(f"portable mutation {label!r} was accepted")


def expect_report_rejected(root: Path, provider: str, label: str) -> None:
    """Exercise the bound report verifier before bundle custody rejects a mutation."""
    report_path = check_report_path(root, provider)
    try:
        run_bound_report_verifier(root, report_path)
    except VerificationError:
        return
    raise VerificationError(f"portable report mutation {label!r} was accepted by bound verifier")


def portable_check(root: Path, baseline: dict[str, Any]) -> dict[str, Any]:
    with tempfile.TemporaryDirectory(prefix="litchi-0430-portable-") as temporary:
        exported = Path(temporary) / "bundle"
        shutil.copytree(root, exported)
        # A portable replay deliberately removes any accidental copy of the
        # executable or source tree.  The source manifest remains evidence and
        # is required; source files and binary bytes are never read.
        for candidate in (
            exported / "tools/perf-baseline/target/release/litchi-perf-baseline",
            exported / "target/release/litchi-perf-baseline",
        ):
            if candidate.is_file():
                candidate.unlink()
        copied = verify_bundle(exported, allow_pending=False)
        if copied["status"] != baseline["status"] or copied["reports"] != baseline["reports"]:
            fail("portable-check", "copied bundle replay differs from source replay")

        for mutation in ("numeric-delta", "output"):
            mutated = Path(temporary) / mutation
            shutil.copytree(exported, mutated)
            mutate_report(mutated, "bytes", mutation)
            expect_report_rejected(mutated, "bytes", mutation)
            expect_rejected(mutated, mutation)

        bound = Path(temporary) / "bound-verifier"
        shutil.copytree(exported, bound)
        verifier = bound / "verify-report.py"
        verifier.write_text(verifier.read_text(encoding="utf-8") + "\n# custody mutation\n",
                            encoding="utf-8")
        expect_rejected(bound, "bound verifier mutation")

        stack = Path(temporary) / "decoded-stack"
        shutil.copytree(exported, stack)
        stack_path = stack / "profiles/bytes-script.log"
        if stack_path.is_file():
            raw = bytearray(stack_path.read_bytes())
            if raw:
                raw[len(raw) // 2] ^= 1
                stack_path.write_bytes(bytes(raw))
                expect_rejected(stack, "decoded stack corruption")
            else:
                fail("portable-check", "bytes script log is empty")
        else:
            compressed = stack / "profiles/bytes-script.log.gz"
            if not compressed.is_file():
                fail("portable-check", "bytes decoded stack log is missing")
            raw = bytearray(gzip.decompress(compressed.read_bytes()))
            if not raw:
                fail("portable-check", "bytes decoded stack log is empty")
            raw[len(raw) // 2] ^= 1
            compressed.write_bytes(gzip.compress(bytes(raw), compresslevel=9, mtime=0))
            expect_rejected(stack, "decoded stack corruption")

        summary = Path(temporary) / "summary"
        shutil.copytree(exported, summary)
        summary_path = summary / "profile-summary.json"
        if summary_path.is_file():
            value = load_json(summary_path)
            value["profiles"][0]["samples"] += 1
            summary_path.write_text(json.dumps(value, indent=2) + "\n", encoding="utf-8")
            expect_rejected(summary, "profile summary mutation")

    return {
        "portable_export": "pass; copied bundle replay without executable/source tree",
        "mutations_rejected": [
            "numeric-delta", "output", "bound verifier mutation",
            "decoded stack corruption", "profile summary mutation",
        ],
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--portable-check", action="store_true",
                        help="copy the bundle, replay it without executable/source, and reject mutations")
    parser.add_argument("--allow-pending", action="store_true",
                        help="permit analysis/final inventory artifacts that root has not sealed yet")
    parser.add_argument("--bundle", type=Path, default=None,
                        help="verify another copied bundle directory")
    args = parser.parse_args(argv)
    root = (args.bundle or ROOT).resolve()
    try:
        result = verify_bundle(root, allow_pending=args.allow_pending)
        if args.portable_check:
            result.update(portable_check(root, result))
    except Exception as exc:
        print(f"INVALID: {exc}", file=sys.stderr)
        return 1
    print(json.dumps(result, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
