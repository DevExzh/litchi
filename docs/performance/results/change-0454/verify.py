#!/usr/bin/env python3
"""Verify the portable evidence retained for change 0454.

Receipts, inventories, and reports are checked from their recorded bytes.
Failed exploratory commands remain useful history, but only completed passing
receipts are included in the final summary.
"""

from __future__ import annotations

import argparse
import collections
import contextlib
import datetime
import hashlib
import importlib.util
import json
import re
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Any, Iterator


ROOT = Path(__file__).resolve().parent
HEX40 = re.compile(r"^[0-9a-f]{40}$")
HEX64 = re.compile(r"^[0-9a-f]{64}$")
FIXTURE_SHA256 = "88a4755fa90815802c8f439c9e0488772e5e7d8db63cfd0326e4d3f35fdeaa44"
FIXTURE_BYTES = 29956
ORIGINAL_DERIVE_SHA256 = "316841a5e46e248c486bb2f7b4c16ae17eabc051539bdd2a67332ea2f4da4c3f"
CORRECTED_DERIVE_SHA256 = "2be514cfb04c4e21d80881a605040943da608347555611c70e9b213ec1999797"
ORIGINAL_PROTOCOL_SHA256 = "6c3ab253d2e6680d1beffe975fc0c8fce472d10f04304eff986264ddfdaaaa3b"
INTERMEDIATE_PROTOCOL_SHA256 = "932d4ccb85820a460fe40ff5a43dd3eb0e1466182fe2cc4aaa0a84df8c868363"
FUZZ_TARGET = "crates/litchi-opc/fuzz/fuzz_targets/parse_opc.rs"
FUZZ_GATE_TAGS = {
    "fuzz-lock",
    "fuzz-build",
    "fuzz-smoke",
    "final-fuzz-strict",
    "final-fuzz-format",
}
CAPTURE_RECEIPT_ROOTS = {
    "provider-runs",
    "external-runs",
    "provider-pilots",
    "external-pilots",
    "external-pilot-attempts",
}
HISTORICAL_RECEIPT_ROOTS = {"external-pilot-attempts"}
FUZZ_RUSTFLAGS = (
    "-C passes=sancov-module -C llvm-args=-sanitizer-coverage-level=4 "
    "-C llvm-args=-sanitizer-coverage-inline-8bit-counters "
    "-C llvm-args=-sanitizer-coverage-pc-table "
    "-C llvm-args=-sanitizer-coverage-trace-compares "
    "-Z sanitizer=address --cfg fuzzing"
)


def final_gate_commands() -> dict[str, list[str]]:
    """Return the exact command capability required for final evidence.

    ``run-checks.py`` and ``run-fuzz.py`` are capture drivers, but accepting a
    receipt by tag alone would let a weaker command stand in for a required
    gate.  Keep this small, explicit table in the verifier so the final
    decision remains portable after those drivers or their working directory
    disappear.
    """
    source_files = load(ROOT / "source-files.json")
    require(isinstance(source_files, list) and all(isinstance(item, str) for item in source_files), "source-files.json: invalid source list for final format gate")
    fuzz_manifest = "/tmp/litchi-goal-0454-opc-fuzz/Cargo.toml"
    return {
        "final-strict": [
            "cargo", "clippy", "--locked", "--release", "-p", "litchi-pptx",
            "-p", "litchi-opc", "--all-targets", "--all-features", "--", "-D",
            "warnings",
        ],
        "final-harness-strict": [
            "cargo", "clippy", "--locked", "--release", "--manifest-path",
            "tools/perf-baseline/Cargo.toml", "--all-targets", "--all-features",
            "--", "-D", "warnings",
        ],
        "final-pptx": [
            "cargo", "test", "--locked", "--release", "-p", "litchi-pptx",
            "--all-features", "--", "--test-threads=1",
        ],
        "final-opc": [
            "cargo", "test", "--locked", "--release", "-p", "litchi-opc",
            "--all-features", "--", "--test-threads=1",
        ],
        "final-harness": [
            "cargo", "test", "--locked", "--release", "--manifest-path",
            "tools/perf-baseline/Cargo.toml", "--all-features", "--",
            "--test-threads=1",
        ],
        "final-doc": [
            "cargo", "doc", "--locked", "--release", "-p", "litchi-pptx",
            "-p", "litchi-opc", "--no-deps", "--all-features",
        ],
        "final-workspace": [
            "cargo", "check", "--locked", "--release", "--workspace",
            "--all-targets", "--no-default-features", "--exclude", "litchi-iwa*",
            "--exclude", "litchi-keynote", "--exclude", "litchi-numbers*",
            "--exclude", "litchi-pages", "--features", "litchi/odf",
        ],
        "final-format": [
            "rustfmt", "+1.98.1", "--edition", "2024", "--check", "--config",
            "skip_children=true", *source_files,
        ],
        "final-boundaries": ["python3", "-B", "tools/check_crate_boundaries.py"],
        "fuzz-lock": [
            "cargo", "generate-lockfile", "--offline", "--manifest-path", fuzz_manifest,
        ],
        "fuzz-build": [
            "env", "RUSTC_BOOTSTRAP=1", f"RUSTFLAGS={FUZZ_RUSTFLAGS}", "cargo",
            "build", "--release", "--locked", "--manifest-path", fuzz_manifest,
            "--target", "x86_64-unknown-linux-gnu", "--bin", "parse_opc",
        ],
        "fuzz-smoke": [
            "/tmp/litchi-goal-0454-opc-fuzz/target/x86_64-unknown-linux-gnu/release/parse_opc",
            "/tmp/litchi-goal-0454-opc-fuzz/corpus", "-runs=1000", "-seed=454",
            "-max_len=1048576", "-timeout=10",
            "-artifact_prefix=/tmp/litchi-goal-0454-opc-fuzz/",
        ],
        "final-fuzz-strict": [
            "cargo", "clippy", "--locked", "--release", "--manifest-path",
            fuzz_manifest, "--", "-D", "warnings",
        ],
        "final-fuzz-format": [
            "rustfmt", "+1.98.1", "--edition", "2024", "--check", "--config",
            "skip_children=true", "crates/litchi-opc/fuzz/fuzz_targets/parse_opc.rs",
        ],
    }


def latest_gate_receipt(tag: str) -> tuple[Path, int]:
    """Select the newest immutable attempt for a required gate tag."""
    checks = ROOT / "checks"
    candidates: list[tuple[int, Path]] = []
    exact = checks / f"{tag}.json"
    if exact.is_file():
        candidates.append((0, exact))
    pattern = re.compile(re.escape(tag) + r"-r([0-9]+)\.json")
    for path in checks.glob(f"{tag}-r*.json"):
        match = pattern.fullmatch(path.name)
        if match is not None:
            candidates.append((int(match.group(1)), path))
    require(candidates, f"required final gate missing: {tag}")
    attempt, path = max(candidates, key=lambda item: (item[0], item[1].name))
    return path, attempt


def check_final_release_gates(candidate_build: dict[str, Any]) -> dict[str, Any]:
    """Require every release/check/fuzz gate to pass for this source epoch."""
    expected_commands = final_gate_commands()
    checked: list[str] = []
    candidate_source = candidate_build.get("source_manifest")
    verify_source_manifest(candidate_source, "candidate-build.source_manifest")
    fuzz_amendment = check_fuzz_source_amendment(candidate_source)
    fuzz_source = fuzz_amendment["after"]
    candidate_revision = text(candidate_build.get("revision"), "candidate-build.revision")
    for tag, expected_argv in expected_commands.items():
        path, _attempt = latest_gate_receipt(tag)
        label = str(path.relative_to(ROOT))
        receipt = passing_receipt(path, label)
        require(receipt.get("argv") == expected_argv, f"{label}: command does not provide {tag} capability")
        require(receipt.get("revision") == candidate_revision, f"{label}: revision differs from candidate build")
        expected_source = fuzz_source if tag in FUZZ_GATE_TAGS else candidate_source
        require(receipt.get("source_before") == expected_source, f"{label}: source differs from its permitted source epoch")
        require(receipt.get("source_after") == expected_source, f"{label}: source differs from its permitted source epoch")
        require(receipt.get("source_unchanged") is True, f"{label}: source epoch is not immutable")
        checked.append(path.name)
    return {
        "required": len(expected_commands),
        "passed": len(checked),
        "receipts": checked,
        "fuzz_source_before": candidate_source,
        "fuzz_source_after": fuzz_source,
    }


class VerificationError(ValueError):
    """Evidence does not satisfy the portable contract."""


def fail(message: str) -> None:
    raise VerificationError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def sha(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def git_blob(raw: bytes) -> str:
    header = f"blob {len(raw)}\0".encode()
    return hashlib.sha1(header + raw).hexdigest()


def no_duplicate_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            fail(f"duplicate JSON object key: {key}")
        result[key] = value
    return result


def load(path: Path) -> Any:
    try:
        return json.loads(
            path.read_bytes(),
            object_pairs_hook=no_duplicate_pairs,
            parse_constant=lambda value: fail(f"non-finite JSON constant: {value}"),
        )
    except VerificationError:
        raise
    except (OSError, json.JSONDecodeError) as error:
        fail(f"cannot read JSON {path.relative_to(ROOT)}: {error}")


def text(value: Any, label: str, *, empty: bool = False) -> str:
    require(isinstance(value, str) and (empty or value), f"{label}: expected text")
    return value


def uint(value: Any, label: str) -> int:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0, f"{label}: expected unsigned integer")
    return value


def sint(value: Any, label: str) -> int:
    """Validate a subprocess return code, including signal termination."""
    require(isinstance(value, int) and not isinstance(value, bool), f"{label}: expected integer")
    return value


def digest(value: Any, label: str) -> str:
    value = text(value, label)
    require(HEX64.fullmatch(value) is not None, f"{label}: expected SHA-256")
    return value


def safe_member(name: Any, label: str) -> Path:
    name = text(name, label)
    path = Path(name)
    require(not path.is_absolute() and ".." not in path.parts, f"{label}: unsafe relative path")
    candidate = ROOT / path
    current = ROOT
    for part in path.parts:
        current /= part
        require(not current.is_symlink(), f"{label}: symlink evidence is not accepted")
    resolved = candidate.resolve()
    require(resolved.is_relative_to(ROOT.resolve()), f"{label}: path escapes evidence bundle")
    return resolved


def historical_attempt(path: Path) -> str | None:
    """Return an archived oracle attempt name for a retained receipt path."""
    try:
        parts = path.relative_to(ROOT).parts
    except ValueError:
        return None
    if len(parts) >= 2 and parts[0] in HISTORICAL_RECEIPT_ROOTS:
        return parts[1]
    return None


def is_custody_correction_receipt(path: Path) -> bool:
    """Identify archived capture copies handled by restoration custody checks."""
    try:
        parts = path.relative_to(ROOT).parts
    except ValueError:
        return False
    return (
        len(parts) >= 4
        and parts[0] == "custody-corrections"
        and parts[1] == "renderer-r1"
        and parts[-1] == "receipt.json"
    )


def historical_member_path(receipt_path: Path, name: Any, label: str, expected: str | None = None) -> Path:
    """Resolve an immutable archived receipt reference without rewriting it."""
    member_name = text(name, f"{label}.path")
    member = Path(member_name)
    require(not member.is_absolute() and ".." not in member.parts, f"{label}.path: unsafe relative path")
    attempt = historical_attempt(receipt_path)
    if attempt is not None:
        # Capture records retain paths from the original external-pilots lane.
        # The archived copy keeps the same bytes under its attempt directory;
        # map only that known lane shape and preserve the JSON verbatim.
        lane = receipt_path.parent.name
        if len(member.parts) >= 3 and member.parts[:2] == ("external-pilots", lane):
            mapped = receipt_path.parent.joinpath(*member.parts[2:])
            return safe_member(str(mapped.relative_to(ROOT)), label)
        if member == Path("candidate-build.json"):
            archived = ROOT / "candidate-attempts" / attempt / "candidate-build.json"
            if archived.is_file() and (expected is None or sha(archived.read_bytes()) == expected):
                return safe_member(str(archived.relative_to(ROOT)), label)
    return safe_member(member_name, label)


def historical_row(row: Any, label: str, receipt_path: Path | None) -> Any:
    """Return a path-remapped copy of an archived artifact/reference row."""
    if receipt_path is None or historical_attempt(receipt_path) is None:
        return row
    require(isinstance(row, dict), f"{label}: expected file reference")
    expected = digest(row.get("sha256"), f"{label}.sha256")
    replacement = dict(row)
    mapped = historical_member_path(receipt_path, row.get("path"), label, expected)
    replacement["path"] = str(mapped.relative_to(ROOT))
    return replacement


def historical_driver_path(receipt_path: Path, relative: str) -> Path:
    """Use the archived driver snapshot when a historical receipt has one."""
    attempt = historical_attempt(receipt_path)
    if attempt is not None:
        snapshot_name = {
            "protocol.json": "protocol.json",
            "external-verifier.py": "external-verifier.py.txt",
        }.get(relative)
        if snapshot_name is not None:
            snapshot = ROOT / "candidate-attempts" / attempt / snapshot_name
            if snapshot.is_file():
                return snapshot
    return ROOT / relative


def artifact(row: Any, label: str, *, allow_missing: bool = False) -> bytes | None:
    require(isinstance(row, dict), f"{label}: expected artifact object")
    path = safe_member(row.get("path"), f"{label}.path")
    size = uint(row.get("bytes"), f"{label}.bytes")
    expected = digest(row.get("sha256"), f"{label}.sha256")
    if not path.exists():
        require(allow_missing, f"{label}: artifact is missing")
        return None
    require(path.is_file(), f"{label}: artifact is not a file")
    raw = path.read_bytes()
    require(len(raw) == size and sha(raw) == expected, f"{label}: artifact identity mismatch")
    return raw


def absolute_artifact(
    row: Any,
    label: str,
    *,
    allow_missing: bool = True,
    path_override: Path | None = None,
) -> None:
    """Check a binary identity whose producer intentionally stores an absolute path."""
    require(isinstance(row, dict), f"{label}: expected binary object")
    path = path_override or Path(text(row.get("path"), f"{label}.path"))
    require(path.is_absolute(), f"{label}.path: expected absolute path")
    size = uint(row.get("bytes"), f"{label}.bytes")
    expected = digest(row.get("sha256"), f"{label}.sha256")
    if not path.exists():
        require(allow_missing, f"{label}: binary is missing")
        return
    require(path.is_file() and not path.is_symlink(), f"{label}: binary is not a regular file")
    raw = path.read_bytes()
    require(len(raw) == size and sha(raw) == expected, f"{label}: binary identity mismatch")


def verify_source_manifest(row: Any, label: str) -> dict[str, str]:
    require(isinstance(row, dict), f"{label}: expected source manifest reference")
    path = safe_member(row.get("path"), f"{label}.path")
    raw = path.read_bytes()
    require(sha(raw) == digest(row.get("sha256"), f"{label}.sha256"), f"{label}: manifest hash mismatch")
    value = load(path)
    require(isinstance(value, dict), f"{label}: manifest is not an object")
    require(uint(row.get("files"), f"{label}.files") == len(value), f"{label}: file count mismatch")
    for name, file_hash in value.items():
        safe_member(name, f"{label}.{name}")
        digest(file_hash, f"{label}.{name}")
    return value


def check_derivation_amendment(protocol: dict[str, Any]) -> dict[str, Any]:
    """Validate the display-only renderer correction used for measurements."""
    path = ROOT / "derivation-amendment.json"
    require(path.is_file() and not path.is_symlink(), "derivation-amendment.json is required")
    amendment = load(path)
    require(isinstance(amendment, dict), "derivation-amendment.json: expected object")
    require(
        set(amendment)
        == {
            "status",
            "original_driver",
            "original_sha256",
            "corrected_driver",
            "corrected_sha256",
            "captured_protocol_sha256",
            "reason",
            "recorded_utc",
        },
        "derivation-amendment.json: unexpected fields",
    )
    require(amendment.get("status") == "pass", "derivation-amendment.json: amendment did not pass")
    require(amendment.get("original_driver") == "derive.py", "derivation-amendment.json: wrong original driver")
    require(amendment.get("corrected_driver") == "derive-final.py", "derivation-amendment.json: wrong corrected driver")
    text(amendment.get("reason"), "derivation-amendment.json.reason")
    text(amendment.get("recorded_utc"), "derivation-amendment.json.recorded_utc")

    original_path = ROOT / "derive.py"
    corrected_path = ROOT / "derive-final.py"
    require(original_path.is_file() and not original_path.is_symlink(), "derive.py: original driver is missing")
    require(corrected_path.is_file() and not corrected_path.is_symlink(), "derive-final.py: corrected driver is missing")
    original_raw = original_path.read_bytes()
    corrected_raw = corrected_path.read_bytes()
    original_hash = sha(original_raw)
    corrected_hash = sha(corrected_raw)
    require(original_hash == ORIGINAL_DERIVE_SHA256, "derive.py: original driver hash changed")
    require(corrected_hash == CORRECTED_DERIVE_SHA256, "derive-final.py: corrected driver hash changed")
    require(digest(amendment.get("original_sha256"), "derivation-amendment.json.original_sha256") == original_hash, "derivation-amendment.json: original hash mismatch")
    require(digest(amendment.get("corrected_sha256"), "derivation-amendment.json.corrected_sha256") == corrected_hash, "derivation-amendment.json: corrected hash mismatch")
    require(digest(amendment.get("captured_protocol_sha256"), "derivation-amendment.json.captured_protocol_sha256") == ORIGINAL_PROTOCOL_SHA256, "derivation-amendment.json: protocol hash mismatch")
    require(protocol.get("bound_files", {}).get("derive.py") == ORIGINAL_DERIVE_SHA256, "protocol: derive.py is not the original driver")

    old_block = (
        b'        source_calls = work["source_reads"]["logical_calls"] if work["source_reads"] else 0\n'
        b'        destination_calls = work["destination_reads"]["logical_calls"] if work["destination_reads"] else 0'
    )
    new_block = (
        b'        source_reads = work.get("source_reads", work.get("source"))\n'
        b'        destination_reads = work.get("destination_reads", work.get("destination"))\n'
        b'        source_calls = source_reads["logical_calls"] if source_reads else 0\n'
        b'        destination_calls = destination_reads["logical_calls"] if destination_reads else 0'
    )
    require(original_raw.count(old_block) == 1, "derive.py: expected renderer block is absent or duplicated")
    require(corrected_raw == original_raw.replace(old_block, new_block, 1), "derive-final.py: amendment changes more than the exact renderer lookup")
    return {
        "original_sha256": original_hash,
        "corrected_sha256": corrected_hash,
        "protocol_sha256": ORIGINAL_PROTOCOL_SHA256,
    }


def check_fuzz_source_amendment(candidate_source: dict[str, Any]) -> dict[str, Any]:
    """Validate the isolated one-file source epoch used only by fuzz gates."""
    path = ROOT / "fuzz-source-amendment.json"
    require(path.is_file() and not path.is_symlink(), "fuzz-source-amendment.json is required")
    amendment = load(path)
    require(isinstance(amendment, dict), "fuzz-source-amendment.json: expected object")
    require(
        set(amendment)
        == {
            "status",
            "before",
            "after",
            "allowed_changed_path",
            "changed_paths",
            "before_sha256",
            "after_sha256",
            "snapshot",
            "reason",
            "recorded_utc",
        },
        "fuzz-source-amendment.json: unexpected fields",
    )
    require(amendment.get("status") == "pass", "fuzz-source-amendment.json: amendment did not pass")
    require(amendment.get("allowed_changed_path") == FUZZ_TARGET, "fuzz-source-amendment.json: unexpected allowed path")
    changed_paths = amendment.get("changed_paths")
    require(changed_paths == [FUZZ_TARGET], "fuzz-source-amendment.json: source amendment changes more than the fuzz target")
    text(amendment.get("reason"), "fuzz-source-amendment.json.reason")
    text(amendment.get("recorded_utc"), "fuzz-source-amendment.json.recorded_utc")

    before_reference = amendment.get("before")
    after_reference = amendment.get("after")
    before = verify_source_manifest(before_reference, "fuzz-source-amendment.before")
    after = verify_source_manifest(after_reference, "fuzz-source-amendment.after")
    require(before_reference == candidate_source, "fuzz-source-amendment.before: candidate source differs")
    require(isinstance(before_reference, dict) and isinstance(after_reference, dict), "fuzz-source-amendment: source references are incomplete")
    require(set(before) == set(after), "fuzz-source-amendment: source file set changed")
    deltas = [name for name in sorted(set(before) | set(after)) if before.get(name) != after.get(name)]
    require(deltas == [FUZZ_TARGET], "fuzz-source-amendment: source delta is not exactly the fuzz target")
    require(digest(amendment.get("before_sha256"), "fuzz-source-amendment.before_sha256") == before[FUZZ_TARGET], "fuzz-source-amendment: before file hash mismatch")
    require(digest(amendment.get("after_sha256"), "fuzz-source-amendment.after_sha256") == after[FUZZ_TARGET], "fuzz-source-amendment: after file hash mismatch")
    snapshot_reference = text(amendment.get("snapshot"), "fuzz-source-amendment.snapshot")
    require(snapshot_reference == "candidate/fuzz-amended-parse_opc.rs.txt", "fuzz-source-amendment: unexpected retained snapshot")
    snapshot = safe_member(snapshot_reference, "fuzz-source-amendment.snapshot")
    require(snapshot.is_file() and not snapshot.is_symlink(), "fuzz-source-amendment.snapshot: missing snapshot")
    require(sha(snapshot.read_bytes()) == after[FUZZ_TARGET], "fuzz-source-amendment.snapshot: hash mismatch")
    return {"before": before_reference, "after": after_reference}


def check_custody_restoration(*, required: bool) -> dict[str, Any]:
    """Validate the byte-level restoration of the 18 captured receipt fields."""
    archive = ROOT / "custody-corrections" / "renderer-r1"
    if not archive.exists():
        require(not required, "custody-corrections/renderer-r1 is required")
        return {"present": False}
    require(archive.is_dir() and not archive.is_symlink(), "custody-corrections/renderer-r1: invalid archive directory")
    restoration_path = archive / "restoration.json"
    require(restoration_path.is_file() and not restoration_path.is_symlink(), "custody restoration receipt is missing")
    restoration = load(restoration_path)
    require(isinstance(restoration, dict), "custody restoration receipt: expected object")
    require(
        set(restoration)
        == {
            "status",
            "intermediate_protocol_sha256",
            "restored_protocol_sha256",
            "receipts",
            "basis",
            "recorded_utc",
        },
        "custody restoration receipt: unexpected fields",
    )
    require(restoration.get("status") == "pass", "custody restoration receipt: restoration did not pass")
    require(restoration.get("intermediate_protocol_sha256") == INTERMEDIATE_PROTOCOL_SHA256, "custody restoration: intermediate protocol hash mismatch")
    require(restoration.get("restored_protocol_sha256") == ORIGINAL_PROTOCOL_SHA256, "custody restoration: restored protocol hash mismatch")
    text(restoration.get("basis"), "custody restoration.basis")
    text(restoration.get("recorded_utc"), "custody restoration.recorded_utc")

    archive_protocol = archive / "protocol.json"
    archive_derive = archive / "derive.py.txt"
    archive_measurements = archive / "measurements.json"
    archive_markdown = archive / "measurements.md"
    for member, label in (
        (archive_protocol, "custody archive protocol"),
        (archive_derive, "custody archive derive"),
        (archive_measurements, "custody archive measurements"),
        (archive_markdown, "custody archive markdown"),
    ):
        require(member.is_file() and not member.is_symlink(), f"{label}: missing")
    require(sha(archive_protocol.read_bytes()) == INTERMEDIATE_PROTOCOL_SHA256, "custody archive protocol hash mismatch")
    require(sha(archive_derive.read_bytes()) == CORRECTED_DERIVE_SHA256, "custody archive derive hash mismatch")
    archived_measurements = load(archive_measurements)
    require(isinstance(archived_measurements, dict), "custody archive measurements: invalid JSON")
    archived_protocol = archived_measurements.get("protocol")
    require(isinstance(archived_protocol, dict) and archived_protocol.get("sha256") == INTERMEDIATE_PROTOCOL_SHA256, "custody archive measurements: protocol identity mismatch")

    expected_paths = [
        *(f"provider-runs/{index}/receipt.json" for index in range(16)),
        *(f"external-runs/{index}/receipt.json" for index in range(2)),
    ]
    rows = restoration.get("receipts")
    require(isinstance(rows, list) and [row.get("path") for row in rows] == expected_paths, "custody restoration: receipt set/order mismatch")
    for index, row in enumerate(rows):
        prefix = f"custody restoration.receipts[{index}]"
        require(isinstance(row, dict) and set(row) == {"path", "intermediate_sha256", "restored_sha256"}, f"{prefix}: malformed row")
        relative = text(row.get("path"), f"{prefix}.path")
        archived_path = archive / relative
        restored_path = ROOT / relative
        require(archived_path.is_file() and not archived_path.is_symlink(), f"{prefix}: archived receipt is missing")
        require(restored_path.is_file() and not restored_path.is_symlink(), f"{prefix}: restored receipt is missing")
        archived_raw = archived_path.read_bytes()
        restored_raw = restored_path.read_bytes()
        require(digest(row.get("intermediate_sha256"), f"{prefix}.intermediate_sha256") == sha(archived_raw), f"{prefix}: archived hash mismatch")
        require(digest(row.get("restored_sha256"), f"{prefix}.restored_sha256") == sha(restored_raw), f"{prefix}: restored hash mismatch")
        archived_value = load(archived_path)
        restored_value = load(restored_path)
        require(archived_value.get("protocol_sha256") == INTERMEDIATE_PROTOCOL_SHA256, f"{prefix}: archived protocol field mismatch")
        require(restored_value.get("protocol_sha256") == ORIGINAL_PROTOCOL_SHA256, f"{prefix}: restored protocol field mismatch")
        archived_without_protocol = dict(archived_value)
        restored_without_protocol = dict(restored_value)
        archived_without_protocol.pop("protocol_sha256", None)
        restored_without_protocol.pop("protocol_sha256", None)
        require(archived_without_protocol == restored_without_protocol, f"{prefix}: restoration changed fields besides protocol_sha256")

    expected_members = {
        "derive.py.txt",
        "protocol.json",
        "measurements.json",
        "measurements.md",
        "restoration.json",
        *expected_paths,
    }
    actual_members = {str(member.relative_to(archive)) for member in archive.rglob("*") if member.is_file()}
    require(actual_members == expected_members, "custody archive: unexpected or missing retained member")
    return {"present": True, "receipts": len(rows), "intermediate_protocol_sha256": INTERMEDIATE_PROTOCOL_SHA256}


def check_log(row: Any, label: str) -> bytes:
    raw = artifact(row, label)
    require(raw is not None, f"{label}: missing log")
    return raw


def check_file_reference(row: Any, label: str) -> bytes:
    """Validate a path/hash reference used by capture receipts."""
    require(isinstance(row, dict), f"{label}: expected file reference")
    path = safe_member(row.get("path"), f"{label}.path")
    expected = digest(row.get("sha256"), f"{label}.sha256")
    require(path.exists() and path.is_file(), f"{label}: referenced file is missing")
    raw = path.read_bytes()
    if "bytes" in row:
        require(uint(row["bytes"], f"{label}.bytes") == len(raw), f"{label}: referenced file size mismatch")
    require(sha(raw) == expected, f"{label}: referenced file hash mismatch")
    return raw


def json_bytes(raw: bytes, label: str) -> dict[str, Any]:
    """Decode an object artifact with the same strict JSON policy as files."""
    try:
        value = json.loads(
            raw,
            object_pairs_hook=no_duplicate_pairs,
            parse_constant=lambda constant: fail(f"{label}: non-finite JSON constant: {constant}"),
        )
    except VerificationError:
        raise
    except (TypeError, json.JSONDecodeError) as error:
        fail(f"{label}: invalid JSON: {error}")
    require(isinstance(value, dict), f"{label}: expected object")
    return value


def check_report_identity(
    report: dict[str, Any],
    receipt: dict[str, Any],
    build: dict[str, Any] | None,
    label: str,
) -> None:
    """Bind a passing report to the executable and build named by its receipt."""
    require(isinstance(receipt.get("binary"), dict), f"{label}: passing report lacks binary identity")
    binary = receipt["binary"]
    require(report.get("binary_sha256") == binary.get("sha256"), f"{label}: report binary hash differs from receipt")
    require(report.get("binary_bytes") == binary.get("bytes"), f"{label}: report binary size differs from receipt")
    require(report.get("current_exe") == binary.get("path"), f"{label}: report executable differs from receipt")
    if build is not None and "revision" in build:
        require(report.get("source_revision") == build["revision"], f"{label}: report source revision differs from build")

    lane = receipt.get("lane_definition")
    if not isinstance(lane, dict):
        return
    if report.get("schema") == "pptx_provider_lifecycle_v1":
        require(report.get("provider") == lane.get("provider"), f"{label}: report provider differs from lane")
        require(report.get("corpus") == lane.get("corpus"), f"{label}: report corpus differs from lane")
    elif report.get("schema") == "pptx-external-cross-copy-v1":
        require(report.get("provider") == lane.get("provider"), f"{label}: report provider differs from lane")
        fixture_metadata = load(ROOT / "external-fixture.json")
        require(isinstance(fixture_metadata, dict), f"{label}: external fixture metadata is invalid")
        require(report.get("fixture_path") == fixture_metadata.get("path"), f"{label}: report fixture path differs from pinned fixture")
        require(report.get("fixture_sha256") == FIXTURE_SHA256, f"{label}: report fixture hash differs from pinned fixture")
        require(report.get("fixture_bytes") == FIXTURE_BYTES, f"{label}: report fixture size differs from pinned fixture")


TEST_RESULT = re.compile(rb"test result: (?:ok|FAILED)\. (\d+) passed; (\d+) failed; (\d+) ignored;")


def check_command_receipt(
    path: Path,
    row: dict[str, Any],
    *,
    strict_running: bool,
    allow_binary_missing: bool = False,
    allow_output_missing: bool = False,
) -> dict[str, Any] | None:
    """Validate one check/capture receipt and return its status summary."""
    if "source_before" not in row:
        return None
    label = str(path.relative_to(ROOT))
    require(not path.is_symlink(), f"{label}: receipt is a symlink")
    root = path.relative_to(ROOT).parts[0]
    historical = historical_attempt(path) is not None
    status = text(row.get("status"), f"{label}.status")
    require(status in {"pass", "failed", "running"}, f"{label}: unknown receipt status")
    if "change" in row:
        require(uint(row["change"], f"{label}.change") == 454, f"{label}: wrong change number")
    before = verify_source_manifest(row["source_before"], f"{label}.source_before")
    if status == "running":
        require("source_after" not in row and "finished_utc" not in row, f"{label}: running receipt has completion fields")
        # A check receipt or capture-lane receipt is part of the candidate
        # evidence set.  Leaving one at its initial running state must never
        # make a partial batch look complete; failed attempts remain valid
        # history after their receipt is finalized.
        if strict_running:
            fail(f"{label}: running receipt cannot be final evidence")
        return {"path": label, "status": status, "source": before}

    after = verify_source_manifest(row.get("source_after"), f"{label}.source_after")
    require(row.get("source_unchanged") is True, f"{label}: completed receipt changed its source epoch")
    require(before == after, f"{label}: source_before differs from source_after")
    if "revision" in row:
        revision = text(row.get("revision"), f"{label}.revision")
        require(HEX40.fullmatch(revision) is not None, f"{label}: invalid revision")
    else:
        # Capture receipts bind their executable revision through the build
        # manifest and retain the checkout epoch under this explicit name.
        verify_source_manifest(row.get("capture_source_manifest"), f"{label}.capture_source_manifest")
    driver = digest(row.get("driver_sha256"), f"{label}.driver_sha256")
    if root == "checks":
        require(driver == sha((ROOT / "check.py").read_bytes()), f"{label}: check driver hash mismatch")
    elif root in {"provider-runs", "external-runs", "provider-pilots", "external-pilots", *HISTORICAL_RECEIPT_ROOTS}:
        require(driver == sha((ROOT / "capture.py").read_bytes()), f"{label}: capture driver hash mismatch")
    argv = row.get("argv")
    require(isinstance(argv, list) and argv and all(isinstance(item, str) for item in argv), f"{label}: invalid argv")
    sint(row.get("exit_code"), f"{label}.exit_code")
    require("finished_utc" in row, f"{label}: completed receipt lacks finish time")

    if root == "checks":
        require("log" in row, f"{label}: check receipt lacks log identity")
        raw_log = check_log(row["log"], f"{label}.log")
        if "test" in argv:
            matches = TEST_RESULT.findall(raw_log)
            counts = {
                "passed_tests": sum(int(match[0]) for match in matches),
                "failed_tests": sum(int(match[1]) for match in matches),
                "ignored_tests": sum(int(match[2]) for match in matches),
            }
            for key, value in counts.items():
                require(uint(row.get(key), f"{label}.{key}") == value, f"{label}: test count mismatch")

    report_value: dict[str, Any] | None = None
    artifacts = row.get("artifacts")
    if artifacts is not None:
        require(isinstance(artifacts, dict), f"{label}.artifacts: expected object")
        for name, item in artifacts.items():
            item = historical_row(item, f"{label}.artifacts.{name}", path)
            raw_artifact = artifact(
                item,
                f"{label}.artifacts.{name}",
                allow_missing=allow_output_missing and name == "output_artifact",
            )
            if status == "pass" and name == "report":
                report_value = json_bytes(raw_artifact or b"", f"{label}.artifacts.report")
                require(isinstance(report_value, dict) and report_value.get("schema") in {"pptx-external-cross-copy-v1", "pptx_provider_lifecycle_v1"}, f"{label}.artifacts.report: unknown benchmark schema")
    build_value: dict[str, Any] | None = None
    if "build_manifest" in row:
        build_reference = historical_row(row["build_manifest"], f"{label}.build_manifest", path)
        build_raw = check_file_reference(build_reference, f"{label}.build_manifest")
        build_value = json_bytes(build_raw, f"{label}.build_manifest")
    for field, relative in (
        ("protocol_sha256", "protocol.json"),
        ("capture_sha256", "capture.py"),
        ("provider_oracle_sha256", "verify-report.py"),
        ("external_oracle_sha256", "external-verifier.py"),
    ):
        if field in row:
            require(
                digest(row[field], f"{label}.{field}") == sha(historical_driver_path(path, relative).read_bytes()),
                f"{label}: {field} does not bind its driver",
            )
    if "binary" in row:
        binary_path = None
        if historical:
            require(isinstance(row["binary"], dict), f"{label}.binary: expected binary object")
            recorded_binary_path = Path(text(row["binary"].get("path"), f"{label}.binary.path"))
            attempt = historical_attempt(path)
            if attempt is not None and recorded_binary_path.is_absolute():
                historical_alias = recorded_binary_path.with_name(f"{attempt}-{recorded_binary_path.name}")
                # The first archived candidate was moved to an explicitly
                # named custody file. Later attempts have the same identity
                # as the current candidate binary and retain the original
                # path. Keep the r1 alias even after cleanup so its absence
                # is accepted under --cleanup rather than being confused
                # with a different current binary.
                if historical_alias.exists() or attempt == "oracle-r1":
                    binary_path = historical_alias
        absolute_artifact(
            row["binary"],
            f"{label}.binary",
            allow_missing=allow_binary_missing,
            path_override=binary_path,
        )
    if status == "pass" and report_value is not None:
        check_report_identity(report_value, row, build_value, f"{label}.artifacts.report")
    if "oracle_exit_code" in row:
        oracle_exit = sint(row["oracle_exit_code"], f"{label}.oracle_exit_code")
        if status == "pass":
            require(oracle_exit == 0, f"{label}: oracle failed")

    exit_code = row["exit_code"]
    if status == "pass":
        require(exit_code == 0, f"{label}: passing receipt has nonzero exit code")
    else:
        # A retained failed attempt is valid history, but it must be visibly
        # failed so it cannot accidentally be counted as a final pass.
        oracle_failed = "oracle_exit_code" in row and row["oracle_exit_code"] != 0
        require(exit_code != 0 or oracle_failed or row.get("passed_tests", 1) == 0, f"{label}: failed receipt looks successful")
    return {"path": label, "status": status, "source": before}


def check_receipts(
    *,
    strict_running: bool,
    allow_binary_missing: bool = False,
    allow_output_missing: bool = False,
) -> dict[str, Any]:
    statuses: collections.Counter[str] = collections.Counter()
    failed: list[str] = []
    running: list[str] = []
    checked = 0
    for path in sorted(ROOT.rglob("*.json")):
        if "sources" in path.relative_to(ROOT).parts:
            continue
        value = load(path)
        if not isinstance(value, dict):
            continue
        relative = path.relative_to(ROOT)
        if is_custody_correction_receipt(path):
            continue
        if relative.parts[0] == "checks" and path.stem.startswith("final-"):
            require("source_before" in value, f"{relative}: final receipt lacks source_before")
        if "source_before" not in value:
            continue
        result = check_command_receipt(
            path,
            value,
            strict_running=strict_running,
            allow_binary_missing=allow_binary_missing,
            allow_output_missing=allow_output_missing,
        )
        if result is None:
            continue
        checked += 1
        statuses[result["status"]] += 1
        if result["status"] == "failed":
            failed.append(result["path"])
        elif result["status"] == "running":
            running.append(result["path"])
    return {"receipts": checked, "passed": statuses["pass"], "failed": failed, "running": running}


def inventory(path: Path) -> tuple[dict[str, Any], dict[tuple[Any, ...], dict[str, Any]]]:
    label = str(path.relative_to(ROOT))
    value = load(path)
    require(isinstance(value, dict), f"{label}: expected inventory object")
    text(value.get("scope"), f"{label}.scope")
    files_inspected = uint(value.get("files_inspected"), f"{label}.files_inspected")
    require(files_inspected > 0, f"{label}: empty corpus inventory")
    rows = value.get("rows")
    require(isinstance(rows, list) and rows, f"{label}.rows: expected non-empty array")
    errors = value.get("scan_errors", [])
    require(isinstance(errors, list), f"{label}.scan_errors: expected array")
    for index, error in enumerate(errors):
        require(isinstance(error, dict), f"{label}.scan_errors[{index}]: expected object")
        text(error.get("path"), f"{label}.scan_errors[{index}].path")
        text(error.get("error"), f"{label}.scan_errors[{index}].error")
    seen: dict[tuple[Any, ...], dict[str, Any]] = {}
    outcomes: collections.Counter[str] = collections.Counter()
    for index, row in enumerate(rows):
        prefix = f"{label}.rows[{index}]"
        require(isinstance(row, dict), f"{prefix}: expected object")
        fixture = text(row.get("path"), f"{prefix}.path")
        require(not Path(fixture).is_absolute() and ".." not in Path(fixture).parts, f"{prefix}.path: unsafe fixture path")
        file_hash = digest(row.get("sha256"), f"{prefix}.sha256")
        size = uint(row.get("bytes"), f"{prefix}.bytes")
        slide = uint(row.get("slide"), f"{prefix}.slide")
        slide_member = text(row.get("slide_member"), f"{prefix}.slide_member")
        pictures = uint(row.get("pictures"), f"{prefix}.pictures")
        require(size > 0 and pictures > 0, f"{prefix}: invalid fixture dimensions")
        argv = row.get("argv")
        require(isinstance(argv, list) and argv and all(isinstance(item, str) for item in argv), f"{prefix}.argv: invalid command")
        uint(row.get("exit_code"), f"{prefix}.exit_code")
        stdout = text(row.get("stdout"), f"{prefix}.stdout", empty=True)
        stderr = text(row.get("stderr"), f"{prefix}.stderr", empty=True)
        identity = (fixture, file_hash, size, slide, slide_member, pictures)
        require(identity not in seen, f"{prefix}: duplicate fixture identity")
        seen[identity] = row
        outcomes[(stdout or stderr).strip()] += 1
    recorded = value.get("outcomes")
    require(isinstance(recorded, dict), f"{label}.outcomes: expected object")
    require({str(key): uint(item, f"{label}.outcomes.{key}") for key, item in recorded.items()} == dict(outcomes), f"{label}: outcomes do not match rows")
    return value, seen


def comparison(path: Path, before: dict[tuple[Any, ...], dict[str, Any]], after: dict[tuple[Any, ...], dict[str, Any]], label: str) -> set[tuple[str, int]]:
    value = load(path)
    require(isinstance(value, dict), f"{label}: expected comparison object")
    text(value.get("scope"), f"{label}.scope")
    require(uint(value.get("rows_compared"), f"{label}.rows_compared") == len(before), f"{label}: row count mismatch")
    changed = value.get("changed")
    require(isinstance(changed, list), f"{label}.changed: expected array")
    by_fixture_before = {(key[0], key[3]): row for key, row in before.items()}
    by_fixture_after = {(key[0], key[3]): row for key, row in after.items()}
    require(len(by_fixture_before) == len(before) and len(by_fixture_after) == len(after), f"{label}: fixture path/slide selectors are ambiguous")
    actual = {
        (key[0], key[3])
        for key in before
        if (before[key].get("exit_code"), before[key].get("stdout"), before[key].get("stderr"))
        != (after[key].get("exit_code"), after[key].get("stdout"), after[key].get("stderr"))
    }
    listed: set[tuple[str, int]] = set()
    for index, row in enumerate(changed):
        prefix = f"{label}.changed[{index}]"
        require(isinstance(row, dict), f"{prefix}: expected object")
        fixture = text(row.get("path"), f"{prefix}.path")
        slide = uint(row.get("slide"), f"{prefix}.slide")
        key = (fixture, slide)
        require(key not in listed, f"{prefix}: duplicate fixture")
        require(key in by_fixture_before and key in by_fixture_after, f"{prefix}: unknown fixture identity")
        identity = next(identity for identity in before if identity[0] == fixture and identity[3] == slide)
        if "sha256" in row:
            require(digest(row["sha256"], f"{prefix}.sha256") == identity[1], f"{prefix}: fixture hash mismatch")
        if "before" in row or "after" in row:
            old = row.get("before")
            new = row.get("after")
            require(isinstance(old, dict) and isinstance(new, dict), f"{prefix}: nested outcome objects are incomplete")
            for field in ("exit_code", "stdout", "stderr"):
                require(old.get(field) == by_fixture_before[key][field], f"{prefix}: baseline {field} mismatch")
                require(new.get(field) == by_fixture_after[key][field], f"{prefix}: candidate {field} mismatch")
        else:
            require(text(row.get("before_stdout"), f"{prefix}.before_stdout", empty=True) == by_fixture_before[key]["stdout"], f"{prefix}: baseline output mismatch")
            require(text(row.get("after_stderr"), f"{prefix}.after_stderr", empty=True) == by_fixture_after[key]["stderr"], f"{prefix}: candidate output mismatch")
        listed.add(key)
    require(actual == listed, f"{label}: changed fixture set does not match rows")
    if "unchanged" in value:
        require(uint(value["unchanged"], f"{label}.unchanged") == len(before) - len(listed), f"{label}: unchanged count mismatch")
    return listed


def check_inventories(*, require_final: bool = False) -> dict[str, Any]:
    baseline_path = ROOT / "baseline-native-inventory.json"
    candidate_path = ROOT / "candidate-native-inventory.json"
    require(baseline_path.exists() and candidate_path.exists(), "baseline/name-only inventories are required")
    baseline, baseline_rows = inventory(baseline_path)
    candidate, candidate_rows = inventory(candidate_path)
    require(set(baseline_rows) == set(candidate_rows), "baseline/name-only fixture identities differ")
    changed = comparison(ROOT / "name-only-outcome-comparison.json", baseline_rows, candidate_rows, "name-only-outcome-comparison.json")
    integrated_counts: dict[str, int] = {}
    for label, inventory_name, comparison_name in (
        ("integrated", "integrated-native-inventory.json", "integrated-outcome-comparison.json"),
        ("final", "final-native-inventory.json", "final-outcome-comparison.json"),
    ):
        inventory_path = ROOT / inventory_name
        if label == "final" and require_final:
            require(inventory_path.is_file(), f"{inventory_name}: final inventory is required")
        if not inventory_path.exists():
            continue
        _, candidate_rows = inventory(inventory_path)
        require(set(baseline_rows) == set(candidate_rows), f"{label} fixture identities differ from baseline")
        comparison_path = ROOT / comparison_name
        if label == "final" and require_final:
            require(comparison_path.is_file(), f"{comparison_name}: final outcome comparison is required")
        if comparison_path.exists():
            changed_rows = comparison(comparison_path, baseline_rows, candidate_rows, comparison_name)
        else:
            changed_rows = {
                (key[0], key[3])
                for key in baseline_rows
                if (baseline_rows[key].get("exit_code"), baseline_rows[key].get("stdout"), baseline_rows[key].get("stderr"))
                != (candidate_rows[key].get("exit_code"), candidate_rows[key].get("stdout"), candidate_rows[key].get("stderr"))
            }
            require(changed_rows == changed, f"{label} changed fixture set differs from name-only admission set")
        integrated_counts[label] = len(changed_rows)
    return {"baseline_rows": len(baseline["rows"]), "name_only_rows": len(candidate["rows"]), "name_only_changed": len(changed), **{f"{key}_changed": value for key, value in integrated_counts.items()}}


def check_external_fixture(*, precleanup: bool) -> dict[str, Any]:
    metadata_path = ROOT / "external-fixture.json"
    require(metadata_path.exists(), "external-fixture.json is required")
    metadata = load(metadata_path)
    require(isinstance(metadata, dict) and metadata.get("status") == "pass", "external fixture metadata is not passing")
    url = text(metadata.get("url"), "external-fixture.url")
    require(url.startswith("https://"), "external-fixture.url: expected HTTPS")
    require(HEX40.fullmatch(text(metadata.get("commit"), "external-fixture.commit")) is not None, "external fixture commit is invalid")
    require(HEX40.fullmatch(text(metadata.get("git_blob"), "external-fixture.git_blob")) is not None, "external fixture blob is invalid")
    expected_path = Path(text(metadata.get("path"), "external-fixture.path"))
    require(expected_path.is_absolute() and str(expected_path).startswith("/tmp/litchi-goal-0454-"), "external fixture path is outside owned scratch")
    require(not expected_path.is_symlink(), "external fixture path is a symlink")
    require(uint(metadata.get("bytes"), "external-fixture.bytes") == FIXTURE_BYTES, "external fixture size changed")
    require(digest(metadata.get("sha256"), "external-fixture.sha256") == FIXTURE_SHA256, "external fixture digest changed")
    require(digest(metadata.get("driver_sha256"), "external-fixture.driver_sha256") == sha((ROOT / "fetch-fixture.py").read_bytes()), "fixture fetch driver hash mismatch")
    scope = text(metadata.get("scope"), "external-fixture.scope")
    require("native application" in scope and "original producer" in scope, "external fixture scope overclaims provenance")

    members_path = ROOT / "external-members.json"
    require(members_path.exists(), "external-members.json is required")
    members = load(members_path)
    require(isinstance(members, list) and members, "external-members.json: expected members")
    names: list[str] = []
    for index, member in enumerate(members):
        prefix = f"external-members[{index}]"
        require(isinstance(member, dict), f"{prefix}: expected object")
        name = text(member.get("name"), f"{prefix}.name")
        require(name not in names, f"{prefix}: duplicate member")
        names.append(name)
        for field in ("bytes", "compressed_bytes", "crc32", "method"):
            uint(member.get(field), f"{prefix}.{field}")
        digest(member.get("sha256"), f"{prefix}.sha256")

    if expected_path.exists():
        raw = expected_path.read_bytes()
        require(len(raw) == FIXTURE_BYTES and sha(raw) == FIXTURE_SHA256 and git_blob(raw) == metadata["git_blob"], "external fixture file identity mismatch")
        with zipfile.ZipFile(expected_path) as archive:
            infos = archive.infolist()
            require([info.filename for info in infos] == names, "external member ordering changed")
            for index, (info, member) in enumerate(zip(infos, members)):
                prefix = f"external-members[{index}]"
                payload = archive.read(info.filename)
                require(info.file_size == member["bytes"] and info.compress_size == member["compressed_bytes"], f"{prefix}: ZIP size mismatch")
                require(info.CRC == member["crc32"] and info.compress_type == member["method"], f"{prefix}: ZIP metadata mismatch")
                require(sha(payload) == member["sha256"], f"{prefix}: payload digest mismatch")
    elif precleanup:
        fail("external fixture is missing before cleanup")
    return {"fixture_members": len(members), "fixture_present": expected_path.exists()}


def module(name: str, path: Path) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    require(spec is not None and spec.loader is not None, f"cannot import {path.name}")
    value = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(value)
    return value


@contextlib.contextmanager
def relocated_external_evidence(
    report_path: Path,
    receipt_path: Path | None = None,
    oracle_log_path: Path | None = None,
    *,
    copy_output: bool,
) -> Iterator[tuple[Path, Path | None, Path | None]]:
    """Relocate absolute report/output paths into this evidence bundle.

    Captures record absolute paths from their original checkout.  The retained
    JSON is never rewritten; temporary bundle-local copies let the independent
    oracle enforce sibling/path identity when a bundle is verified from a
    clone.  Cleanup replay omits the deleted output bytes and relocates only
    the retained report, receipt, and oracle-log attestations.
    """
    report = load(report_path)
    require(isinstance(report, dict), f"{report_path}: expected report object")
    output = report.get("output_artifact")
    require(isinstance(output, dict), f"{report_path}: output artifact is missing")
    original_output = report_path.with_suffix(".pptx")
    recorded_output = Path(text(output.get("path"), f"{report_path}.output_artifact.path")).resolve()
    if recorded_output != original_output.resolve():
        require(recorded_output.name == original_output.name and recorded_output.parts[-2:] == original_output.parts[-2:], f"{report_path}: output artifact is not the report sibling")
    if copy_output:
        require(original_output.is_file() and not original_output.is_symlink(), f"{report_path}: output artifact is missing")

    metadata = load(ROOT / "external-fixture.json")
    require(isinstance(metadata, dict), "external-fixture.json: expected object")
    with tempfile.TemporaryDirectory(prefix=".verify-external-", dir=ROOT) as directory:
        temporary = Path(directory)
        relocated_report = temporary / "report.json"
        relocated_output = temporary / "report.pptx"
        relocated_report_value = dict(report)
        relocated_output_value = dict(output)
        relocated_output_value["path"] = str(relocated_output.resolve())
        relocated_report_value["output_artifact"] = relocated_output_value
        relocated_report.write_text(json.dumps(relocated_report_value) + "\n", encoding="utf-8")
        if copy_output:
            relocated_output.write_bytes(original_output.read_bytes())

        relocated_receipt: Path | None = None
        relocated_oracle: Path | None = None
        if receipt_path is not None or oracle_log_path is not None:
            require(receipt_path is not None and oracle_log_path is not None, "external replay paths are incomplete")
            receipt = load(receipt_path)
            oracle_log = load(oracle_log_path)
            require(isinstance(receipt, dict), f"{receipt_path}: expected receipt object")
            require(isinstance(oracle_log, dict), f"{oracle_log_path}: expected oracle object")
            relocated_oracle = temporary / "oracle.log"
            relocated_oracle_value = dict(oracle_log)
            relocated_oracle_value["output_artifact"] = str(relocated_output.resolve())
            relocated_oracle.write_text(json.dumps(relocated_oracle_value) + "\n", encoding="utf-8")

            artifacts = receipt.get("artifacts")
            require(isinstance(artifacts, dict), f"{receipt_path}: artifacts are missing")
            normalized = dict(receipt)
            normalized.setdefault("revision", report.get("source_revision"))
            normalized.setdefault(
                "fixture",
                {
                    "path": metadata.get("path"),
                    "bytes": metadata.get("bytes"),
                    "sha256": metadata.get("sha256"),
                },
            )
            normalized_artifacts = dict(artifacts)
            normalized_artifacts["report"] = {
                "path": str(relocated_report.relative_to(ROOT)),
                "bytes": relocated_report.stat().st_size,
                "sha256": sha(relocated_report.read_bytes()),
            }
            normalized_artifacts["oracle"] = {
                "path": str(relocated_oracle.relative_to(ROOT)),
                "bytes": relocated_oracle.stat().st_size,
                "sha256": sha(relocated_oracle.read_bytes()),
            }
            normalized_artifacts["output_artifact"] = {
                # external-verifier.py resolves this retained output path
                # against the process cwd, unlike its bundle-relative
                # report/oracle artifact helper.  Keep the original receipt
                # untouched and bind only this temporary replay path
                # absolutely.
                "path": str(relocated_output.resolve()),
                "bytes": uint(output.get("bytes"), f"{report_path}.output_artifact.bytes"),
                "sha256": digest(output.get("sha256"), f"{report_path}.output_artifact.sha256"),
            }
            normalized["artifacts"] = normalized_artifacts
            relocated_receipt = temporary / "receipt.json"
            relocated_receipt.write_text(json.dumps(normalized) + "\n", encoding="utf-8")
        yield relocated_report, relocated_receipt, relocated_oracle


def check_reports(*, allow_external_fixture_missing: bool = False) -> dict[str, int]:
    external_oracle = module("external_oracle_0454", ROOT / "external-verifier.py")
    report_oracle = module("report_oracle_0454", ROOT / "verify-report.py")
    external = 0
    lifecycle = 0
    for path in sorted(ROOT.rglob("*.json")):
        if "sources" in path.relative_to(ROOT).parts:
            continue
        value = load(path)
        if not isinstance(value, dict):
            continue
        label = str(path.relative_to(ROOT))
        if value.get("schema") == "pptx-external-cross-copy-v1":
            if historical_attempt(path) is not None:
                receipt_path = path.parent / "receipt.json"
                require(receipt_path.is_file(), f"{label}: archived external attempt receipt is missing")
                receipt = load(receipt_path)
                require(isinstance(receipt, dict) and receipt.get("status") == "failed", f"{label}: archived external attempt is not a failed historical capture")
                # check_receipts has already bound every retained report,
                # output, and oracle artifact to this failed receipt.  The
                # failed attempt is custody history and must not be promoted
                # to a passing oracle result by replaying it here.
                continue
            fixture_path = value.get("fixture_path")
            output_path = path.with_suffix(".pptx")
            replay_needed = (
                allow_external_fixture_missing
                and isinstance(fixture_path, str)
                and (not Path(fixture_path).exists() or not output_path.exists())
            )
            if replay_needed:
                receipt_path = path.parent / "receipt.json"
                oracle_log_path = path.parent / "oracle.log"
                require(receipt_path.is_file(), f"{label}: external capture receipt is missing after cleanup")
                require(oracle_log_path.is_file(), f"{label}: external oracle log is missing after cleanup")
                replay = getattr(external_oracle, "verify_replay_attestation", None)
                require(callable(replay), f"{label}: external oracle lacks replay attestation")
                with relocated_external_evidence(path, receipt_path, oracle_log_path, copy_output=False) as (
                    relocated_report,
                    relocated_receipt,
                    relocated_oracle,
                ):
                    try:
                        result = replay(relocated_report, relocated_receipt, relocated_oracle)
                    except Exception as error:  # oracle errors need a portable path
                        fail(f"{label}: external replay attestation rejected report: {error}")
                require(result.get("status") == "pass" and result.get("verification_mode") == "replay-attestation", f"{label}: external replay attestation did not pass")
                external += 1
                continue
            try:
                with relocated_external_evidence(path, copy_output=True) as (relocated_report, _, _):
                    result = external_oracle.verify(relocated_report)
            except Exception as error:  # oracle errors need a portable path
                fail(f"{label}: external oracle rejected report: {error}")
            require(result.get("status") == "pass", f"{label}: external oracle did not pass")
            external += 1
        elif value.get("schema") == "pptx_provider_lifecycle_v1":
            try:
                result = report_oracle.check_report(value)
            except Exception as error:
                fail(f"{label}: lifecycle oracle rejected report: {error}")
            require(result.get("status") == "pass", f"{label}: lifecycle oracle did not pass")
            lifecycle += 1
    return {"external_reports": external, "lifecycle_reports": lifecycle}


def check_build_identities(*, allow_missing: bool) -> int:
    checked = 0
    paths = set(ROOT.glob("*-build.json"))
    paths.update(ROOT / name for name in ("baseline-binary.json", "name-only-binary.json", "baseline-lifecycle-binary.json", "integrated-probe-binary.json", "final-probe-binary.json") if (ROOT / name).exists())
    for path in sorted(paths):
        value = load(path)
        require(isinstance(value, dict), f"{path.name}: expected build object")
        if "revision" in value:
            revision = text(value.get("revision"), f"{path.name}.revision")
            require(HEX40.fullmatch(revision) is not None, f"{path.name}: invalid revision")
        if "source_manifest" in value:
            verify_source_manifest(value["source_manifest"], f"{path.name}.source_manifest")
        binaries = value.get("binaries")
        if binaries is not None:
            require(isinstance(binaries, dict) and binaries, f"{path.name}.binaries: expected object")
            for name, binary in binaries.items():
                absolute_artifact(binary, f"{path.name}.binaries.{name}", allow_missing=allow_missing)
                checked += 1
        elif "path" in value:
            absolute_artifact(value, path.name, allow_missing=allow_missing)
            checked += 1
    return checked


def check_source_custody() -> int:
    files_path = ROOT / "source-files.json"
    parent_path = ROOT / "parent-sources.json"
    if not files_path.exists() or not parent_path.exists():
        return 0
    files = load(files_path)
    require(isinstance(files, list) and all(isinstance(name, str) for name in files), "source-files.json: invalid source list")
    require(len(files) == len(set(files)), "source-files.json: duplicate source")
    for name in files:
        safe_member(name, f"source-files.json.{name}")
    parent = load(parent_path)
    require(isinstance(parent, dict), "parent-sources.json: expected object")
    revision = text(parent.get("revision"), "parent-sources.revision")
    require(HEX40.fullmatch(revision) is not None, "parent-sources.revision: invalid revision")
    sources = parent.get("sources")
    require(isinstance(sources, dict) and set(sources) == set(files), "parent-sources.sources: source set mismatch")
    for name, value in sources.items():
        if value is not None:
            digest(value, f"parent-sources.sources.{name}")
            snapshot = ROOT / "candidate" / f"before-{Path(name).name}.txt"
            if snapshot.exists():
                require(sha(snapshot.read_bytes()) == value, f"candidate snapshot mismatch for {name}")
    return len(files)


def check_protocol(*, require_bound_files: bool = False) -> dict[str, Any]:
    path = ROOT / "protocol.json"
    require(path.exists(), "protocol.json is required")
    require(sha(path.read_bytes()) == ORIGINAL_PROTOCOL_SHA256, "protocol.json: frozen original protocol hash changed")
    protocol = load(path)
    require(isinstance(protocol, dict), "protocol.json: expected object")
    require(uint(protocol.get("change"), "protocol.change") == 454, "protocol.change: wrong change")
    require(protocol.get("schema") == "pptx_change0454_performance_protocol_v1", "protocol.schema: unexpected schema")
    require(protocol.get("status") == "frozen", "protocol.status: protocol is not frozen")
    for field in ("cpu", "workers", "samples", "warmups"):
        require(uint(protocol.get(field), f"protocol.{field}") > 0, f"protocol.{field}: expected positive integer")
    providers = protocol.get("provider_lanes")
    externals = protocol.get("external_lanes")
    require(isinstance(providers, list) and providers, "protocol.provider_lanes: expected non-empty array")
    require(isinstance(externals, list) and externals, "protocol.external_lanes: expected non-empty array")
    for index, lane in enumerate(providers):
        prefix = f"protocol.provider_lanes[{index}]"
        require(isinstance(lane, dict), f"{prefix}: expected object")
        require(lane.get("suite") == "provider", f"{prefix}.suite: unexpected suite")
        require(lane.get("provider") in {"bytes", "range"}, f"{prefix}.provider: unexpected provider")
        require(lane.get("corpus") in {"plain", "media-rich"}, f"{prefix}.corpus: unexpected corpus")
        require(lane.get("build") in {"baseline", "candidate"}, f"{prefix}.build: unexpected build")
        text(lane.get("repeat"), f"{prefix}.repeat")
        safe_member(lane.get("build_manifest"), f"{prefix}.build_manifest")
    for index, lane in enumerate(externals):
        prefix = f"protocol.external_lanes[{index}]"
        require(isinstance(lane, dict), f"{prefix}: expected object")
        require(lane.get("suite") == "external", f"{prefix}.suite: unexpected suite")
        require(lane.get("fixture") == "libreoffice-smoketest", f"{prefix}.fixture: unexpected fixture")
        require(lane.get("provider") in {"bytes", "range"}, f"{prefix}.provider: unexpected provider")
        require(lane.get("build") == "candidate", f"{prefix}.build: external lane must use candidate")
        text(lane.get("repeat"), f"{prefix}.repeat")
        safe_member(lane.get("build_manifest"), f"{prefix}.build_manifest")
    bound_files = protocol.get("bound_files")
    require(isinstance(bound_files, dict) and bound_files, "protocol.bound_files: expected object")
    require(bound_files.get("derive.py") == ORIGINAL_DERIVE_SHA256, "protocol.bound_files.derive.py: original driver is not bound")
    for name, expected in bound_files.items():
        member = safe_member(name, f"protocol.bound_files.{name}")
        if require_bound_files:
            require(member.is_file(), f"protocol.bound_files.{name}: file is missing")
            require(sha(member.read_bytes()) == digest(expected, f"protocol.bound_files.{name}"), f"protocol.bound_files.{name}: hash mismatch")
    return protocol


def load_build_manifest(name: str, label: str) -> dict[str, Any]:
    path = safe_member(name, f"{label}.path")
    value = load(path)
    require(isinstance(value, dict), f"{label}: expected build object")
    return value


def lane_binary(build: dict[str, Any], lane: dict[str, Any], label: str) -> dict[str, Any]:
    binaries = build.get("binaries")
    if isinstance(binaries, dict):
        key = "probe" if lane.get("suite") == "probe" else ("normal" if lane.get("suite") == "provider" else "external")
        require(key in binaries, f"{label}: build lacks {key} binary")
        value = binaries[key]
    else:
        require(lane.get("suite") == "provider", f"{label}: external lane needs named binary")
        value = build
    require(isinstance(value, dict), f"{label}: invalid binary identity")
    text(value.get("path"), f"{label}.path")
    uint(value.get("bytes"), f"{label}.bytes")
    digest(value.get("sha256"), f"{label}.sha256")
    return value


def passing_receipt(
    path: Path,
    label: str,
    *,
    allow_binary_missing: bool = False,
    allow_output_missing: bool = False,
) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"{label}: receipt is missing or symlinked")
    value = load(path)
    require(isinstance(value, dict), f"{label}: expected receipt object")
    result = check_command_receipt(
        path,
        value,
        strict_running=True,
        allow_binary_missing=allow_binary_missing,
        allow_output_missing=allow_output_missing,
    )
    require(result is not None and result["status"] == "pass", f"{label}: receipt is not a passing completed receipt")
    return value


def check_report_output(report_path: Path, report: dict[str, Any], *, allow_missing: bool) -> dict[str, Any]:
    """Bind an external report to its sibling output, including after cleanup."""
    label = str(report_path.relative_to(ROOT))
    output = report.get("output_artifact")
    require(isinstance(output, dict), f"{label}: output artifact is missing")
    expected_path = report_path.with_suffix(".pptx").resolve()
    reported_path = Path(text(output.get("path"), f"{label}.output_artifact.path")).resolve()
    if reported_path != expected_path:
        # Captures retain the absolute path from their original checkout.  A
        # cloned bundle may relocate the report, so preserve the recorded
        # name and external-runs lane while checking the clone's sibling.
        if historical_attempt(report_path) is not None:
            require(reported_path.name == expected_path.name and reported_path.parent.name == expected_path.parent.name, f"{label}: historical output artifact is not the report sibling")
        else:
            require(reported_path.name == expected_path.name and reported_path.parts[-2:] == expected_path.parts[-2:], f"{label}: output artifact is not the report sibling")
    size = uint(output.get("bytes"), f"{label}.output_artifact.bytes")
    expected_hash = digest(output.get("sha256"), f"{label}.output_artifact.sha256")
    if not expected_path.exists():
        require(allow_missing, f"{label}: output artifact is missing")
        return output
    require(expected_path.is_file() and not expected_path.is_symlink(), f"{label}: output artifact is not a regular file")
    raw = expected_path.read_bytes()
    require(len(raw) == size and sha(raw) == expected_hash, f"{label}: output artifact identity mismatch")
    return output


def check_final_probe_driver(corpus_name: str) -> None:
    path = ROOT / "final-probe-driver.json"
    require(path.is_file(), "final-probe-driver.json is required for final native inventory")
    value = load(path)
    require(isinstance(value, dict), "final-probe-driver.json: expected object")
    expected = sha((ROOT / corpus_name).read_bytes())
    bound = value.get(corpus_name)
    if bound is None:
        bound = value.get("driver_sha256")
    require(digest(bound, f"final-probe-driver.json.{corpus_name}") == expected, "final-probe-driver.json: probe corpus driver hash mismatch")


def check_formal_lane(
    suite: str,
    index: int,
    lane: dict[str, Any],
    protocol: dict[str, Any],
    candidate_build: dict[str, Any],
    *,
    pilot: bool = False,
    allow_binary_missing: bool = False,
    allow_output_missing: bool = False,
) -> None:
    if pilot:
        require(suite == "external", f"pilot lane {suite}/{index}: only external pilot lanes are supported")
    directory_name = "provider-pilots" if pilot and suite == "provider" else (
        "external-pilots" if pilot else ("provider-runs" if suite == "provider" else "external-runs")
    )
    directory = ROOT / directory_name / str(index)
    label = str(directory.relative_to(ROOT))
    receipt_path = directory / "receipt.json"
    receipt = passing_receipt(
        receipt_path,
        f"{label}/receipt.json",
        allow_binary_missing=allow_binary_missing,
        allow_output_missing=allow_output_missing,
    )
    require(receipt.get("change") == 454, f"{label}: wrong change")
    require(receipt.get("suite") == suite, f"{label}: wrong suite")
    require(receipt.get("lane") == index, f"{label}: wrong lane index")
    require(receipt.get("pilot") is pilot, f"{label}: pilot flag differs from lane kind")
    require(receipt.get("lane_definition") == lane, f"{label}: lane definition differs from frozen protocol")
    require(receipt.get("build_manifest", {}).get("path") == lane.get("build_manifest"), f"{label}: build manifest differs from protocol")
    require(receipt.get("source_before") == candidate_build.get("source_manifest"), f"{label}: capture source differs from candidate build")
    require(receipt.get("capture_source_manifest") == candidate_build.get("source_manifest"), f"{label}: capture source binding differs from candidate build")

    build = load_build_manifest(lane["build_manifest"], f"{label}.build_manifest")
    expected_binary = lane_binary(build, lane, f"{label}.binary")
    require(receipt.get("binary") == expected_binary, f"{label}: binary identity differs from build manifest")
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict), f"{label}: artifacts are missing")
    for name in ("report", "resource", "workload", "oracle"):
        require(name in artifacts, f"{label}: missing {name} artifact")
    report_raw = artifact(artifacts["report"], f"{label}.artifacts.report")
    report = json_bytes(report_raw or b"", f"{label}.artifacts.report")
    if suite == "external":
        require("output_artifact" in artifacts, f"{label}: missing output_artifact")
        check_report_output(directory / "report.json", report, allow_missing=allow_output_missing)
        artifact(
            artifacts["output_artifact"],
            f"{label}.artifacts.output_artifact",
            allow_missing=allow_output_missing,
        )
        output = report.get("output_artifact")
        receipt_output = artifacts["output_artifact"]
        require(isinstance(output, dict), f"{label}: external report output identity is missing")
        receipt_output_path = Path(text(receipt_output.get("path"), f"{label}.artifacts.output_artifact.path")).resolve()
        report_output_path = Path(text(output.get("path"), f"{label}.report.output_artifact.path")).resolve()
        require(receipt_output_path == report_output_path or (receipt_output_path.name == report_output_path.name and receipt_output_path.parts[-2:] == report_output_path.parts[-2:]), f"{label}: receipt output path differs from report")
        require(uint(receipt_output.get("bytes"), f"{label}.artifacts.output_artifact.bytes") == uint(output.get("bytes"), f"{label}.report.output_artifact.bytes"), f"{label}: receipt output size differs from report")
        require(digest(receipt_output.get("sha256"), f"{label}.artifacts.output_artifact.sha256") == digest(output.get("sha256"), f"{label}.report.output_artifact.sha256"), f"{label}: receipt output hash differs from report")
    if suite == "provider":
        require(report.get("schema") == "pptx_provider_lifecycle_v1", f"{label}: unexpected provider report schema")
        require(report.get("provider") == lane.get("provider"), f"{label}: provider report lane mismatch")
        require(report.get("corpus") == lane.get("corpus"), f"{label}: provider report corpus mismatch")
        require(report.get("samples") == protocol["samples"], f"{label}: provider sample count mismatch")
        require(report.get("warmup") == protocol["warmups"], f"{label}: provider warmup count mismatch")
    else:
        require(report.get("schema") == "pptx-external-cross-copy-v1", f"{label}: unexpected external report schema")
        require(report.get("provider") == lane.get("provider"), f"{label}: external report lane mismatch")
        expected_samples = 1 if pilot else protocol["samples"]
        expected_warmups = 0 if pilot else protocol["warmups"]
        require(report.get("samples") == expected_samples, f"{label}: external sample count mismatch")
        require(report.get("warmups") == expected_warmups, f"{label}: external warmup count mismatch")
        fixture = receipt.get("fixture")
        require(isinstance(fixture, dict), f"{label}: external fixture identity is missing")
        fixture_metadata = load(ROOT / "external-fixture.json")
        require(isinstance(fixture_metadata, dict), f"{label}: external fixture metadata is invalid")
        for field in ("path", "bytes", "sha256"):
            require(fixture.get(field) == fixture_metadata.get(field), f"{label}: fixture {field} differs from pinned fixture")
        require(receipt.get("revision") == report.get("source_revision"), f"{label}: receipt revision differs from external report")


def check_measurements(protocol: dict[str, Any]) -> dict[str, int]:
    machine_path = ROOT / "machine.json"
    measurements_path = ROOT / "measurements.json"
    markdown_path = ROOT / "measurements.md"
    require(machine_path.is_file(), "machine.json is required for final measurements")
    require(measurements_path.is_file(), "measurements.json is required for final measurements")
    require(markdown_path.is_file() and markdown_path.read_text(encoding="utf-8").startswith("# Change 0454 performance measurements"), "measurements.md is missing or has the wrong heading")
    machine = load(machine_path)
    require(isinstance(machine, dict) and machine.get("schema") == "change0454_machine_v1", "machine.json: unexpected schema")
    measurements = load(measurements_path)
    require(isinstance(measurements, dict), "measurements.json: expected object")
    require(measurements.get("change") == 454 and measurements.get("schema") == "pptx_change0454_measurements_v1", "measurements.json: unexpected schema")
    check_derivation_amendment(protocol)
    protocol_ref = measurements.get("protocol")
    require(isinstance(protocol_ref, dict), "measurements.protocol: expected object")
    require(protocol_ref.get("path") == "protocol.json", "measurements.protocol.path: unexpected path")
    require(protocol_ref.get("sha256") == sha((ROOT / "protocol.json").read_bytes()), "measurements.protocol: protocol hash mismatch")
    for field in ("cpu", "workers", "samples", "warmups"):
        require(protocol_ref.get(field) == protocol[field], f"measurements.protocol.{field}: differs from frozen protocol")
    machine_ref = measurements.get("machine")
    require(isinstance(machine_ref, dict) and machine_ref.get("path") == "machine.json", "measurements.machine: missing machine binding")
    require(machine_ref.get("status") == "captured", "measurements.machine: machine capture is incomplete")
    require(machine_ref.get("sha256") == sha(machine_path.read_bytes()), "measurements.machine: hash mismatch")
    provider_rows = measurements.get("provider_rows")
    external_rows = measurements.get("external_rows")
    provider_lanes = protocol["provider_lanes"]
    external_lanes = protocol["external_lanes"]
    require(isinstance(provider_rows, list) and len(provider_rows) == len(provider_lanes), "measurements.provider_rows: incomplete lane set")
    require(isinstance(external_rows, list) and len(external_rows) == len(external_lanes), "measurements.external_rows: incomplete lane set")
    require({row.get("lane") for row in provider_rows if isinstance(row, dict)} == set(range(len(provider_lanes))), "measurements.provider_rows: lane indexes are incomplete")
    require({row.get("lane") for row in external_rows if isinstance(row, dict)} == set(range(len(external_lanes))), "measurements.external_rows: lane indexes are incomplete")
    pairs = measurements.get("provider_pairs")
    require(isinstance(pairs, list), "measurements.provider_pairs: expected array")
    expected_pairs = {(lane["provider"], lane["corpus"], lane["repeat"]) for lane in provider_lanes}
    actual_pairs = {(row.get("provider"), row.get("corpus"), row.get("repeat")) for row in pairs if isinstance(row, dict)}
    require(actual_pairs == expected_pairs, "measurements.provider_pairs: incomplete control pairing")
    try:
        derive_module = module("derive_final_for_verify_0454", ROOT / "derive-final.py")
        expected = derive_module.derive()
        expected_markdown = derive_module.render(expected)
    except Exception as error:
        fail(f"measurements: cannot reproduce derivation: {error}")
    require(measurements == expected, "measurements.json: does not reproduce from retained reports")
    require(markdown_path.read_text(encoding="utf-8") == expected_markdown, "measurements.md: does not reproduce from measurements")
    return {"provider_rows": len(provider_rows), "external_rows": len(external_rows)}


def check_final_readiness(
    protocol: dict[str, Any],
    *,
    allow_binary_missing: bool = False,
    allow_output_missing: bool = False,
) -> dict[str, Any]:
    candidate_path = ROOT / "candidate-build.json"
    require(candidate_path.is_file(), "candidate-build.json is required before precleanup")
    candidate_build = load(candidate_path)
    require(isinstance(candidate_build, dict), "candidate-build.json: expected object")
    require(HEX40.fullmatch(text(candidate_build.get("revision"), "candidate-build.revision")) is not None, "candidate-build.revision: invalid revision")
    require(isinstance(candidate_build.get("source_manifest"), dict), "candidate-build.source_manifest: missing")
    gates = check_final_release_gates(candidate_build)
    lane_binary(candidate_build, {"suite": "provider"}, "candidate-build.binaries.normal")
    probe_binary = lane_binary(candidate_build, {"suite": "probe"}, "candidate-build.binaries.probe") if isinstance(candidate_build.get("binaries"), dict) and "probe" in candidate_build["binaries"] else None

    build_receipt_path, _build_attempt = latest_gate_receipt("final-candidate-build")
    build_receipt = passing_receipt(
        build_receipt_path,
        str(build_receipt_path.relative_to(ROOT)),
        allow_binary_missing=allow_binary_missing,
        allow_output_missing=allow_output_missing,
    )
    build_label = str(build_receipt_path.relative_to(ROOT))
    require(build_receipt.get("source_after") == candidate_build.get("source_manifest"), f"{build_label}: source differs from candidate build")
    require(build_receipt.get("revision") == candidate_build.get("revision"), f"{build_label}: revision differs from candidate build")
    native_receipt_path, _native_attempt = latest_gate_receipt("final-native-inventory")
    native_receipt = passing_receipt(
        native_receipt_path,
        str(native_receipt_path.relative_to(ROOT)),
        allow_binary_missing=allow_binary_missing,
        allow_output_missing=allow_output_missing,
    )
    native_label = str(native_receipt_path.relative_to(ROOT))
    require(native_receipt.get("source_before") == candidate_build.get("source_manifest"), f"{native_label}: source differs from candidate build")
    require(native_receipt.get("source_after") == candidate_build.get("source_manifest"), f"{native_label}: source differs from candidate build")
    native_argv = native_receipt.get("argv", [])
    probe_names = {Path(argument).name for argument in native_argv}
    probe_corpus = next((name for name in ("final-probe-corpus.py", "candidate-probe-corpus.py") if name in probe_names), None)
    require(probe_corpus is not None, f"{native_label}: receipt is not the final probe inventory")
    check_final_probe_driver(probe_corpus)
    require(probe_binary is not None, "candidate-build.json: probe binary identity is required")

    for suite, lanes, directory_name in (
        ("provider", protocol["provider_lanes"], "provider-runs"),
        ("external", protocol["external_lanes"], "external-runs"),
    ):
        directory = ROOT / directory_name
        require(directory.is_dir() and not directory.is_symlink(), f"{directory_name}: formal capture directory is missing or symlinked")
        children = list(directory.iterdir())
        directories = [path for path in children if path.is_dir()]
        require(len(directories) == len(children), f"{directory_name}: unexpected non-directory entry")
        require(all(not path.is_symlink() for path in directories), f"{directory_name}: symlink lane directory is not accepted")
        actual = {path.name for path in directories}
        require(actual == {str(index) for index in range(len(lanes))}, f"{directory_name}: formal lane set is incomplete")
        for index, lane in enumerate(lanes):
            check_formal_lane(
                suite,
                index,
                lane,
                protocol,
                candidate_build,
                allow_binary_missing=allow_binary_missing,
                allow_output_missing=allow_output_missing,
            )

    pilot_directory = ROOT / "external-pilots"
    require(pilot_directory.is_dir() and not pilot_directory.is_symlink(), "external-pilots: pilot capture directory is missing")
    pilot_children = list(pilot_directory.iterdir())
    pilot_directories = [path for path in pilot_children if path.is_dir()]
    require(len(pilot_directories) == len(pilot_children), "external-pilots: unexpected non-directory entry")
    require(all(not path.is_symlink() for path in pilot_directories), "external-pilots: symlink lane directory is not accepted")
    actual_pilots = {path.name for path in pilot_directories}
    require(actual_pilots == {str(index) for index in range(len(protocol["external_lanes"]))}, "external-pilots: pilot lane set is incomplete")
    for index, lane in enumerate(protocol["external_lanes"]):
        check_formal_lane(
            "external",
            index,
            lane,
            protocol,
            candidate_build,
            pilot=True,
            allow_binary_missing=allow_binary_missing,
            allow_output_missing=allow_output_missing,
        )

    final_inventory_path = ROOT / "final-native-inventory.json"
    final_inventory = load(final_inventory_path)
    require(isinstance(final_inventory, dict), "final-native-inventory.json: expected object")
    rows = final_inventory.get("rows")
    require(isinstance(rows, list) and rows, "final-native-inventory.json: expected rows")
    probe_path = probe_binary["path"]
    for index, row in enumerate(rows):
        require(isinstance(row, dict), f"final-native-inventory.json.rows[{index}]: expected object")
        argv = row.get("argv")
        require(isinstance(argv, list) and len(argv) == 5, f"final-native-inventory.json.rows[{index}].argv: unexpected probe command")
        require(argv[0] == probe_path, f"final-native-inventory.json.rows[{index}]: probe binary differs from candidate build")
        require(argv[1] == row.get("path") and argv[3] == row.get("path"), f"final-native-inventory.json.rows[{index}]: source/destination fixture mismatch")
        slide = str(row.get("slide"))
        require(argv[2] == slide and argv[4] == slide, f"final-native-inventory.json.rows[{index}]: slide selector mismatch")

    measurements = check_measurements(protocol)
    return {
        "provider_lanes": len(protocol["provider_lanes"]),
        "external_lanes": len(protocol["external_lanes"]),
        "release_gates": gates,
        **measurements,
    }


def check_raw_output_cleanup() -> int:
    manifest_path = ROOT / "checks" / "raw-output-artifacts.json"
    cleanup_receipt_path = ROOT / "checks" / "raw-output-cleanup.json"
    require(manifest_path.is_file(), "--cleanup requires checks/raw-output-artifacts.json")
    require(cleanup_receipt_path.is_file(), "--cleanup requires checks/raw-output-cleanup.json")
    manifest = load(manifest_path)
    require(isinstance(manifest, dict) and manifest.get("status") == "pass", "raw-output-artifacts.json: inventory did not pass")
    require(manifest.get("scope") == "external-runs and external-pilots report.pptx siblings; archived failed external-pilot-attempts report.pptx siblings", "raw-output-artifacts.json: unexpected scope")
    require(digest(manifest.get("driver_sha256"), "raw-output-artifacts.json.driver_sha256") == sha((ROOT / "cleanup.py").read_bytes()), "raw-output-artifacts.json: cleanup driver mismatch")
    rows = manifest.get("artifacts")
    require(isinstance(rows, list) and rows, "raw-output-artifacts.json: expected output inventory")
    require(uint(manifest.get("files"), "raw-output-artifacts.json.files") == len(rows), "raw-output-artifacts.json: file count mismatch")
    total = 0
    seen: set[str] = set()
    for index, row in enumerate(rows):
        prefix = f"raw-output-artifacts.json.artifacts[{index}]"
        require(isinstance(row, dict), f"{prefix}: expected object")
        output_path = safe_member(row.get("path"), f"{prefix}.path")
        report_path = safe_member(row.get("report_path"), f"{prefix}.report_path")
        receipt_path = safe_member(row.get("receipt_path"), f"{prefix}.receipt_path")
        require(output_path == report_path.with_suffix(".pptx"), f"{prefix}: output is not the retained report sibling")
        require(receipt_path == report_path.parent / "receipt.json", f"{prefix}: receipt is not the capture sibling")
        relative_output = str(output_path.relative_to(ROOT))
        require(relative_output not in seen, f"{prefix}: duplicate output")
        seen.add(relative_output)
        size = uint(row.get("bytes"), f"{prefix}.bytes")
        expected = digest(row.get("sha256"), f"{prefix}.sha256")
        require(not output_path.exists(), f"{prefix}: raw output remains after cleanup")
        require(report_path.is_file() and not report_path.is_symlink(), f"{prefix}: retained report is missing")
        require(uint(row.get("report_bytes"), f"{prefix}.report_bytes") == report_path.stat().st_size, f"{prefix}: report size changed")
        require(digest(row.get("report_sha256"), f"{prefix}.report_sha256") == sha(report_path.read_bytes()), f"{prefix}: report hash changed")
        require(receipt_path.is_file() and not receipt_path.is_symlink(), f"{prefix}: retained capture receipt is missing")
        require(uint(row.get("receipt_bytes"), f"{prefix}.receipt_bytes") == receipt_path.stat().st_size, f"{prefix}: receipt size changed")
        require(digest(row.get("receipt_sha256"), f"{prefix}.receipt_sha256") == sha(receipt_path.read_bytes()), f"{prefix}: receipt hash changed")
        receipt = load(receipt_path)
        historical = row.get("historical") is True
        if historical:
            output_parts = output_path.relative_to(ROOT).parts
            require(len(output_parts) >= 3 and output_parts[0] == "external-pilot-attempts", f"{prefix}: historical output is outside archived pilot attempts")
            require(row.get("attempt") == output_parts[1], f"{prefix}: historical attempt identity differs")
            require(isinstance(receipt, dict) and receipt.get("status") == "failed" and receipt.get("suite") == "external" and receipt.get("pilot") is True, f"{prefix}: retained historical receipt is not a failed external pilot")
            require(receipt.get("lane") == row.get("lane"), f"{prefix}: historical receipt lane differs")
        else:
            require(isinstance(receipt, dict) and receipt.get("status") == "pass" and receipt.get("suite") == "external", f"{prefix}: retained receipt is not a passing external capture")
        require(receipt.get("lane") == row.get("lane"), f"{prefix}: receipt lane differs")
        require(receipt.get("pilot") is row.get("pilot"), f"{prefix}: receipt pilot identity differs")
        artifacts = receipt.get("artifacts")
        require(isinstance(artifacts, dict), f"{prefix}: receipt artifacts are missing")
        receipt_report = artifacts.get("report")
        receipt_output = artifacts.get("output_artifact")
        require(isinstance(receipt_report, dict) and isinstance(receipt_output, dict), f"{prefix}: receipt raw output identity is incomplete")
        recorded_report = Path(text(receipt_report.get("path"), f"{prefix}.receipt.report.path"))
        if historical:
            require(recorded_report.name == report_path.name and recorded_report.parent.name == report_path.parent.name, f"{prefix}: historical receipt report identity differs")
        else:
            require(safe_member(receipt_report.get("path"), f"{prefix}.receipt.report.path") == report_path, f"{prefix}: receipt report identity differs")
        require(uint(receipt_report.get("bytes"), f"{prefix}.receipt.report.bytes") == report_path.stat().st_size, f"{prefix}: receipt report size differs")
        require(digest(receipt_report.get("sha256"), f"{prefix}.receipt.report.sha256") == sha(report_path.read_bytes()), f"{prefix}: receipt report hash differs")
        recorded_output = Path(text(receipt_output.get("path"), f"{prefix}.receipt.output.path"))
        if historical:
            require(recorded_output.name == output_path.name and recorded_output.parent.name == output_path.parent.name, f"{prefix}: historical receipt output path differs")
        else:
            require(safe_member(receipt_output.get("path"), f"{prefix}.receipt.output.path") == output_path, f"{prefix}: receipt output path differs")
        require(uint(receipt_output.get("bytes"), f"{prefix}.receipt.output.bytes") == size, f"{prefix}: receipt output size differs")
        require(digest(receipt_output.get("sha256"), f"{prefix}.receipt.output.sha256") == expected, f"{prefix}: receipt output hash differs")
        report = load(report_path)
        require(isinstance(report, dict) and report.get("schema") == "pptx-external-cross-copy-v1", f"{prefix}: retained report is not external evidence")
        output = check_report_output(report_path, report, allow_missing=True)
        require(uint(output.get("bytes"), f"{prefix}.report.output_artifact.bytes") == size, f"{prefix}: output size differs from report")
        require(digest(output.get("sha256"), f"{prefix}.report.output_artifact.sha256") == expected, f"{prefix}: output hash differs from report")
        total += size
    require(uint(manifest.get("bytes"), "raw-output-artifacts.json.bytes") == total, "raw-output-artifacts.json: byte total mismatch")

    receipt = load(cleanup_receipt_path)
    require(isinstance(receipt, dict) and receipt.get("status") == "pass", "raw-output-cleanup.json: cleanup did not pass")
    require(receipt.get("scope") == manifest.get("scope"), "raw-output-cleanup.json: scope differs from inventory")
    require(receipt.get("raw_outputs_absent") is True, "raw-output-cleanup.json: raw outputs remain")
    require(uint(receipt.get("files_removed"), "raw-output-cleanup.json.files_removed") == len(rows), "raw-output-cleanup.json: file count mismatch")
    require(uint(receipt.get("bytes_removed"), "raw-output-cleanup.json.bytes_removed") == total, "raw-output-cleanup.json: byte total mismatch")
    require(digest(receipt.get("artifact_manifest_sha256"), "raw-output-cleanup.json.artifact_manifest_sha256") == sha(manifest_path.read_bytes()), "raw-output-cleanup.json: inventory hash mismatch")
    require(digest(receipt.get("driver_sha256"), "raw-output-cleanup.json.driver_sha256") == sha((ROOT / "cleanup.py").read_bytes()), "raw-output-cleanup.json: cleanup driver mismatch")
    return 1


def check_verifier_adapter_amendment(previous_hash: str, current_hash: str) -> dict[str, Any]:
    path = ROOT / "verifier-adapter-amendment.json"
    require(path.is_file() and not path.is_symlink(), "verifier-adapter-amendment.json is required after the verifier changed")
    amendment = load(path)
    require(isinstance(amendment, dict), "verifier-adapter-amendment.json: expected object")
    require(
        set(amendment)
        == {
            "schema",
            "status",
            "scope",
            "changed_component",
            "previous_verifier_sha256",
            "current_verifier_sha256",
            "precleanup_receipt",
            "external_oracle_unchanged",
            "capture_receipts_unchanged",
            "reparsed",
            "no_reparse_claim",
        },
        "verifier-adapter-amendment.json: unexpected fields",
    )
    require(amendment.get("schema") == "change-0454-verifier-adapter-amendment-v1", "verifier-adapter-amendment.json: wrong schema")
    require(amendment.get("status") == "pass", "verifier-adapter-amendment.json: amendment did not pass")
    require(amendment.get("scope") == "post-cleanup replay and cleanup receipt validation", "verifier-adapter-amendment.json: scope differs")
    require(amendment.get("changed_component") == "relocated_external_evidence, check_cleanup_receipts, check_raw_output_cleanup, amendment validation", "verifier-adapter-amendment.json: changed component differs")
    require(digest(amendment.get("previous_verifier_sha256"), "verifier-adapter-amendment.json.previous_verifier_sha256") == previous_hash, "verifier-adapter-amendment.json: previous verifier hash differs")
    require(digest(amendment.get("current_verifier_sha256"), "verifier-adapter-amendment.json.current_verifier_sha256") == current_hash, "verifier-adapter-amendment.json: current verifier hash differs")
    require(amendment.get("precleanup_receipt") == "checks/precleanup.json", "verifier-adapter-amendment.json: precleanup receipt differs")
    require(amendment.get("external_oracle_unchanged") is True, "verifier-adapter-amendment.json: external oracle was changed")
    require(amendment.get("capture_receipts_unchanged") is True, "verifier-adapter-amendment.json: capture receipts were changed")
    require(amendment.get("reparsed") is False, "verifier-adapter-amendment.json: replay was reparsed")
    require(isinstance(amendment.get("no_reparse_claim"), str) and amendment["no_reparse_claim"], "verifier-adapter-amendment.json: no-reparse claim is missing")
    return {
        "previous_verifier_sha256": previous_hash,
        "current_verifier_sha256": current_hash,
        "reparsed": False,
    }


def check_cleanup_receipts(*, cleanup: bool) -> dict[str, Any]:
    precleanup = ROOT / "checks" / "precleanup.json"
    if not cleanup:
        return {"precleanup": precleanup.exists(), "cleanup_receipts": 0}
    require(precleanup.exists(), "--cleanup requires checks/precleanup.json")
    pre = load(precleanup)
    require(isinstance(pre, dict) and pre.get("status") == "pass", "precleanup receipt is not passing")
    previous_verifier_hash = digest(pre.get("verifier_sha256"), "precleanup.verifier_sha256")
    current_verifier_hash = sha(Path(__file__).read_bytes())
    verifier_amendment = None
    if previous_verifier_hash != current_verifier_hash:
        verifier_amendment = check_verifier_adapter_amendment(previous_verifier_hash, current_verifier_hash)
    count = 0
    for path in sorted((ROOT / "checks").glob("*cleanup.json")):
        if path.name == "precleanup.json":
            continue
        value = load(path)
        require(isinstance(value, dict) and value.get("status") == "pass", f"{path.name}: cleanup did not pass")
        if path.name == "raw-output-cleanup.json":
            require(value.get("raw_outputs_absent") is True, f"{path.name}: raw outputs remain")
        else:
            require(value.get("temporary_directory_absent") is True, f"{path.name}: temporary directory remains")
        if path.name == "binary-cleanup.json":
            require(digest(value.get("driver_sha256"), f"{path.name}.driver_sha256") == sha((ROOT / "cleanup.py").read_bytes()), f"{path.name}: cleanup driver mismatch")
            manifest = ROOT / "checks" / "binary-artifacts.json"
            require(manifest.exists(), f"{path.name}: artifact manifest missing")
            require(digest(value.get("artifact_manifest_sha256"), f"{path.name}.artifact_manifest_sha256") == sha(manifest.read_bytes()), f"{path.name}: artifact manifest mismatch")
        if path.name == "fuzz-cleanup.json":
            manifest = ROOT / "checks" / "fuzz-artifacts.json"
            require(manifest.exists(), f"{path.name}: artifact manifest missing")
            require(digest(value.get("artifact_manifest_sha256"), f"{path.name}.artifact_manifest_sha256") == sha(manifest.read_bytes()), f"{path.name}: artifact manifest mismatch")
        count += 1
    for manifest_name, receipt_name in (("binary-artifacts.json", "binary-cleanup.json"), ("fuzz-artifacts.json", "fuzz-cleanup.json")):
        if (ROOT / "checks" / manifest_name).exists():
            require((ROOT / "checks" / receipt_name).exists(), f"--cleanup requires {receipt_name}")
    check_raw_output_cleanup()
    result = {"precleanup": True, "cleanup_receipts": count}
    if verifier_amendment is not None:
        result["verifier_adapter_amendment"] = verifier_amendment
    return result


def write_precleanup(result: dict[str, Any]) -> None:
    target = ROOT / "checks" / "precleanup.json"
    target.parent.mkdir(exist_ok=True)
    target.write_text(
        json.dumps(
            {
                "status": "pass",
                "change": 454,
                "verified_utc": datetime.datetime.now(datetime.timezone.utc).isoformat(),
                "verifier_sha256": sha(Path(__file__).read_bytes()),
                "receipts": result["receipts"],
                "passed_receipts": result["passed_receipts"],
                "failed_attempts": result["failed_attempts"],
                "reports": result["reports"],
            },
            indent=2,
        )
        + "\n"
    )


def check(*, precleanup: bool, cleanup: bool) -> dict[str, Any]:
    protocol = check_protocol(require_bound_files=precleanup or cleanup)
    custody = check_custody_restoration(required=precleanup or cleanup)
    receipt_result = check_receipts(
        strict_running=precleanup or cleanup,
        allow_binary_missing=cleanup,
        allow_output_missing=cleanup,
    )
    inventories = check_inventories(require_final=precleanup or cleanup)
    fixture = check_external_fixture(precleanup=precleanup)
    reports = check_reports(allow_external_fixture_missing=cleanup)
    builds = check_build_identities(allow_missing=cleanup)
    sources = check_source_custody()
    cleanup_result = check_cleanup_receipts(cleanup=cleanup)
    require(not receipt_result["running"] or not (precleanup or cleanup), "incomplete running receipts remain")
    readiness = (
        check_final_readiness(
            protocol,
            allow_binary_missing=cleanup,
            allow_output_missing=cleanup,
        )
        if (precleanup or cleanup)
        else {}
    )
    result = {
        # A plain invocation is a useful custody audit while the batch is in
        # flight.  Only the explicit final gate may report a passing bundle.
        "status": "pass" if (precleanup or cleanup) else "partial",
        "change": 454,
        "precleanup": precleanup,
        "cleanup": cleanup,
        "receipts": receipt_result["receipts"],
        "passed_receipts": receipt_result["passed"],
        "failed_attempts": len(receipt_result["failed"]),
        "reports": reports,
        "inventories": inventories,
        "fixture": fixture,
        "build_identities": builds,
        "source_files": sources,
        "cleanup_receipts": cleanup_result["cleanup_receipts"],
        "custody_restoration": custody,
    }
    if readiness:
        result["final_readiness"] = readiness
    if precleanup:
        write_precleanup(result)
    return result


def seal() -> None:
    path = ROOT / "SHA256SUMS"
    require(path.exists(), "--sealed requires SHA256SUMS")
    seen: set[str] = set()
    for line in path.read_text().splitlines():
        fields = line.split("  ", 1)
        require(len(fields) == 2 and HEX64.fullmatch(fields[0]) is not None, "invalid seal line")
        name = fields[1]
        require(name not in seen, f"duplicate seal member: {name}")
        member = safe_member(name, "seal member")
        require(sha(member.read_bytes()) == fields[0], f"seal hash mismatch: {name}")
        seen.add(name)
    actual = {str(path.relative_to(ROOT)) for path in ROOT.rglob("*") if path.is_file() and path.name != "SHA256SUMS"}
    require(seen == actual, "seal does not cover exactly the evidence bundle")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--precleanup", action="store_true", help="verify final evidence and write checks/precleanup.json")
    parser.add_argument("--cleanup", action="store_true", help="verify evidence after owned cleanup receipts")
    parser.add_argument("--sealed", action="store_true", help="also verify SHA256SUMS coverage")
    args = parser.parse_args()
    try:
        result = check(precleanup=args.precleanup, cleanup=args.cleanup)
        if args.sealed:
            seal()
        print(json.dumps(result, sort_keys=True))
        return 0
    except (VerificationError, OSError, zipfile.BadZipFile) as error:
        print(f"INVALID: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
