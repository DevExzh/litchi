#!/usr/bin/env python3
"""Verify and portably replay the retained native ODP append capture.

This verifier authenticates the capture drivers and receipts before it trusts
any result.  It can validate the original fixture/binary while they still
exist (``--precleanup``), or replay only the retained gzip archives after
cleanup.  Replay invokes the independent ``verify-output.py`` from a private
temporary directory; it never imports the producer and never needs the
original fixture, executable, or repository checkout.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import subprocess
import sys
import tempfile
import zlib
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parent
# The checked-in bundle has a repository four levels above native/. A copied
# portable bundle may live directly under a temporary directory, so importing
# this verifier must not require that depth to exist.
REPO = ROOT.parents[4] if len(ROOT.parents) > 4 else ROOT
TASK = Path("/tmp/litchi-goal-0457/native-final")
SCHEMA = "litchi-0457-native-verification-v1"
MAX_FIXTURES = 10
MAX_RETAINED_GZIP_BYTES = 256 * 1024 * 1024
MAX_DECODED_ARCHIVE_BYTES = 128 * 1024 * 1024
MAX_LOG_BYTES = 32 * 1024 * 1024
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")


class VerificationError(Exception):
    """A bound capture or replay failed independent verification."""


def fail(message: str) -> None:
    raise VerificationError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def load_json(path: Path) -> Any:
    require_regular(path, f"missing or unsafe JSON artifact: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"cannot parse JSON artifact {path}: {error}")
    raise AssertionError


def require_regular(path: Path, label: str | None = None) -> None:
    require(path.is_file() and not path.is_symlink(), label or f"unsafe file: {path}")


def require_directory(path: Path, label: str | None = None) -> None:
    require(path.is_dir() and not path.is_symlink(), label or f"unsafe directory: {path}")


def sha_file(path: Path) -> str:
    require_regular(path)
    digest = hashlib.sha256()
    try:
        with path.open("rb") as source:
            while True:
                chunk = source.read(1024 * 1024)
                if not chunk:
                    break
                digest.update(chunk)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def read_bounded(path: Path, maximum: int, label: str) -> bytes:
    require_regular(path, f"missing or unsafe {label}: {path}")
    try:
        size = path.stat().st_size
        require(size <= maximum, f"{label} exceeds bounded read limit: {path}")
        return path.read_bytes()
    except OSError as error:
        fail(f"cannot read {label} {path}: {error}")
    raise AssertionError


def sha_bytes(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def digest(value: Any, label: str) -> str:
    require(isinstance(value, str) and SHA256_RE.fullmatch(value) is not None,
            f"{label} is not a lowercase SHA-256 digest")
    return value


def integer(value: Any, label: str, minimum: int = 0) -> int:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= minimum,
            f"{label} is not an integer >= {minimum}")
    return value


def text(value: Any, label: str) -> str:
    require(isinstance(value, str) and value != "", f"{label} is not non-empty text")
    return value


def safe_relative(base: Path, value: Any, label: str) -> Path:
    relative = text(value, label)
    candidate = Path(relative)
    require(not candidate.is_absolute(), f"{label} must be relative")
    try:
        resolved = (base / candidate).resolve(strict=True)
    except OSError as error:
        fail(f"{label} cannot be resolved: {error}")
    require(resolved.is_relative_to(base.resolve()), f"{label} escapes its bundle: {relative}")
    path = base / candidate
    require_regular(path, f"{label} is missing, a symlink, or not regular: {path}")
    return path


def safe_artifact(base: Path, value: Any, label: str) -> Path:
    return safe_relative(base, value, f"{label}.path")


def parse_json_log(path: Path, label: str) -> dict[str, Any]:
    raw = read_bounded(path, MAX_LOG_BYTES, label)
    try:
        value = json.loads(raw.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as error:
        fail(f"{label} is not exactly one JSON object: {error}")
    require(isinstance(value, dict), f"{label} JSON root is not an object")
    return value


def gzip_decode(raw: bytes, maximum: int, label: str) -> bytes:
    """Decode exactly one gzip member with a hard decoded-size bound."""

    decompressor = zlib.decompressobj(16 + zlib.MAX_WBITS)
    try:
        decoded = decompressor.decompress(raw, maximum + 1)
        if len(decoded) <= maximum:
            decoded += decompressor.flush()
    except zlib.error as error:
        fail(f"{label} is not a valid gzip stream: {error}")
    require(len(decoded) <= maximum, f"{label} decoded beyond bounded archive limit")
    require(decompressor.eof, f"{label} gzip stream ended before its member EOF")
    require(not decompressor.unused_data and not decompressor.unconsumed_tail,
            f"{label} contains trailing or unconsumed gzip data")
    return decoded


def retained_archive(
    runs: Path,
    metadata: Any,
    expected_name: str,
    expected_decoded_sha256: str,
    expected_decoded_bytes: int,
    label: str,
) -> tuple[Path, bytes]:
    require(isinstance(metadata, dict), f"{label} metadata is not an object")
    path = safe_artifact(runs, metadata.get("path"), label)
    require(path.name == expected_name, f"{label} retained path differs: {path.name}")
    compressed_bytes = integer(metadata.get("bytes"), f"{label}.bytes")
    compressed_sha256 = digest(metadata.get("sha256"), f"{label}.sha256")
    decoded_bytes = integer(metadata.get("decoded_bytes"), f"{label}.decoded_bytes")
    decoded_sha256 = digest(metadata.get("decoded_sha256"), f"{label}.decoded_sha256")
    require(compressed_bytes <= MAX_RETAINED_GZIP_BYTES,
            f"{label} compressed archive exceeds bounded limit")
    raw = read_bounded(path, MAX_RETAINED_GZIP_BYTES, label)
    require(len(raw) == compressed_bytes and sha_bytes(raw) == compressed_sha256,
            f"{label} compressed identity differs")
    decoded = gzip_decode(raw, MAX_DECODED_ARCHIVE_BYTES, label)
    require(len(decoded) == decoded_bytes and sha_bytes(decoded) == decoded_sha256,
            f"{label} decoded gzip identity differs")
    require(decoded_bytes == expected_decoded_bytes and decoded_sha256 == expected_decoded_sha256,
            f"{label} does not match its immutable source/output hash")
    return path, decoded


def binding_ref(
    binding: dict[str, Any], name: str, default_path: str | None = None
) -> tuple[str, str]:
    value = binding.get(name)
    if isinstance(value, dict):
        path = value.get("path", default_path)
        checksum = value.get("sha256")
    else:
        path = binding.get(f"{name}_path", default_path)
        checksum = binding.get(f"{name}_sha256")
    require(path is not None, f"binary-binding.json has no binding.{name}.path")
    require(checksum is not None, f"binary-binding.json has no binding.{name}.sha256")
    return text(path, f"binding.{name}.path"), digest(checksum, f"binding.{name}.sha256")


def binding_artifact_ref(
    binding: dict[str, Any], name: str
) -> tuple[str, str, int | None]:
    """Read a binding reference, including an optional source-file count."""

    value = binding.get(name)
    if isinstance(value, dict):
        path = value.get("path")
        checksum = value.get("sha256")
        files = value.get("files")
    else:
        path = binding.get(f"{name}_path")
        checksum = binding.get(f"{name}_sha256")
        files = binding.get(f"{name}_files")
    require(path is not None, f"binary-binding.json has no binding.{name}.path")
    require(checksum is not None, f"binary-binding.json has no binding.{name}.sha256")
    if files is not None:
        files = integer(files, f"binding.{name}.files")
    return text(path, f"binding.{name}.path"), digest(checksum, f"binding.{name}.sha256"), files


def safe_bundle_relative(value: Any, label: str) -> Path:
    """Resolve a bundle artifact under native/ or its change-0457 parent."""

    relative = text(value, label)
    candidate = Path(relative)
    require(not candidate.is_absolute(), f"{label} must be relative")
    # Inventory/oracle/runner are native-local.  The copied build/source
    # artifacts may instead retain the parent check.py-relative path (for
    # example ``checks/baseline-build.json`` or ``../sources/<sha>.json``),
    # so accept either bounded root while never leaving change-0457/.
    for base in (ROOT, ROOT.parent):
        try:
            resolved = (base / candidate).resolve(strict=True)
        except OSError:
            continue
        if resolved.is_relative_to(ROOT.parent.resolve()):
            path = base / candidate
            if path.is_file() and not path.is_symlink():
                return path
    fail(f"{label} is missing, a symlink, or escapes the evidence bundle: {relative}")
    raise AssertionError


def source_identity(value: Any, label: str) -> tuple[str, int]:
    require(isinstance(value, dict), f"{label} is not a source identity object")
    return digest(value.get("sha256"), f"{label}.sha256"), integer(value.get("files"), f"{label}.files")


def require_source_identity(value: Any, expected: tuple[str, int], label: str) -> None:
    actual = source_identity(value, label)
    require(actual == expected, f"{label} differs from the bound source manifest")


def verify_receipt_log(receipt: dict[str, Any], base: Path, label: str) -> Path:
    log = receipt.get("log")
    require(isinstance(log, dict), f"{label}.log is not an object")
    path_value = text(log.get("path"), f"{label}.log.path")
    candidate = Path(path_value)
    require(not candidate.is_absolute(), f"{label}.log.path must be relative")
    try:
        path = (base / candidate).resolve(strict=True)
    except OSError as error:
        fail(f"{label}.log.path cannot be resolved: {error}")
    require(path.is_relative_to(base.resolve()), f"{label}.log.path escapes the bundle")
    path = base / candidate
    require_regular(path, f"{label}.log is missing, a symlink, or not regular")
    expected_bytes = integer(log.get("bytes"), f"{label}.log.bytes")
    expected_sha256 = digest(log.get("sha256"), f"{label}.log.sha256")
    require(path.stat().st_size == expected_bytes and sha_file(path) == expected_sha256,
            f"{label}.log identity differs")
    return path


def standard_receipt(
    path: Path,
    expected_source: tuple[str, int],
    driver_path: Path,
    label: str,
    verify_log: bool = True,
) -> dict[str, Any]:
    """Authenticate a check.py receipt without producer-specific fields."""

    receipt = load_json(path)
    require(isinstance(receipt, dict), f"{label} root is not an object")
    require(receipt.get("change") == 457, f"{label} is for another change")
    require(receipt.get("status") == "pass" and receipt.get("exit_code") == 0,
            f"{label} is not a successful check")
    require(receipt.get("source_unchanged") is True,
            f"{label} does not prove unchanged source custody")
    require_source_identity(receipt.get("source_before"), expected_source,
                             f"{label}.source_before")
    require_source_identity(receipt.get("source_after"), expected_source,
                             f"{label}.source_after")
    require(text(receipt.get("cwd"), f"{label}.cwd") == str(Path(receipt["cwd"]).resolve()),
            f"{label}.cwd is not an absolute canonical path")
    require(digest(receipt.get("driver_sha256"), f"{label}.driver_sha256") == sha_file(driver_path),
            f"{label}.driver_sha256 differs from check.py")
    if verify_log:
        verify_receipt_log(receipt, ROOT.parent, label)
    else:
        # The final run gate must carry a locally retained log.  A completed
        # build receipt is already hash-bound by binary-binding.json; retain
        # and validate its standard log metadata when the parent log is
        # available, while allowing portable bundles that copied only the
        # receipt and source manifest.
        log = receipt.get("log")
        require(isinstance(log, dict), f"{label}.log is not an object")
        text(log.get("path"), f"{label}.log.path")
        integer(log.get("bytes"), f"{label}.log.bytes")
        digest(log.get("sha256"), f"{label}.log.sha256")
        log_candidate = ROOT.parent / log["path"]
        if log_candidate.is_file() and not log_candidate.is_symlink():
            verify_receipt_log(receipt, ROOT.parent, label)
    return receipt


def runner_from_gate(
    gate: dict[str, Any],
    runner_path: Path,
    expected_cwd: str,
) -> None:
    argv = gate.get("argv")
    require(isinstance(argv, list) and argv and
            all(isinstance(value, str) and value for value in argv),
            "native check receipt argv is not a non-empty string list")
    require(gate.get("cwd") == expected_cwd,
            "native check receipt cwd differs from the authenticated build cwd")
    candidates = [Path(value) for value in argv if Path(value).name == runner_path.name]
    require(len(candidates) == 1, "native check receipt does not identify exactly one run.py")
    token = candidates[0]
    if token.is_absolute():
        resolved = token
    else:
        resolved = Path(expected_cwd) / token
    if not resolved.is_file():
        resolved = ROOT.parent / token
    if not resolved.is_file():
        resolved = ROOT / token
    # A portable copy deliberately does not recreate the historical absolute
    # checkout path recorded by check.py.  The basename plus the prebound
    # runner digest still identifies the command without requiring that path.
    if not resolved.is_file() and token.name == runner_path.name:
        resolved = runner_path
    require_regular(resolved, "native check receipt runner path is missing or unsafe")
    require(sha_file(resolved) == sha_file(runner_path),
            "native check receipt runner hash differs from the bound runner")


def authenticate_bindings(precleanup: bool) -> dict[str, Any]:
    binding_path = ROOT / "binary-binding.json"
    binding = load_json(binding_path)
    require(isinstance(binding, dict), "binary-binding.json root is not an object")
    require(binding.get("change") in (None, 457), "binary binding is for another change")
    binding_sha256 = sha_file(binding_path)

    inventory_name, inventory_sha256 = binding_ref(binding, "inventory", "staticinventory.json")
    oracle_name, oracle_sha256 = binding_ref(binding, "oracle", "verify-output.py")
    runner_name, runner_sha256 = binding_ref(binding, "runner", "run.py")
    inventory_path = safe_relative(ROOT, inventory_name, "binding.inventory")
    oracle_path = safe_relative(ROOT, oracle_name, "binding.oracle")
    runner_path = safe_relative(ROOT, runner_name, "binding.runner")
    require(sha_file(inventory_path) == inventory_sha256, "inventory hash differs from binary binding")
    require(sha_file(oracle_path) == oracle_sha256, "oracle hash differs from binary binding")
    require(sha_file(runner_path) == runner_sha256, "runner hash differs from binary binding")

    # The binding is frozen before capture.  It therefore binds only the
    # completed build receipt and source custody that already exist.  The
    # native run receipt is generated later by check.py at the fixed parent
    # path below and is authenticated by its own driver/source/log fields;
    # putting its final digest in this binding would be circular because
    # run.py records the binding digest in runs/index.json.
    build_name, build_sha256, _ = binding_artifact_ref(binding, "build_receipt")
    source_name, source_sha256, source_files = binding_artifact_ref(binding, "source_manifest")
    build_path = safe_bundle_relative(build_name, "binding.build_receipt")
    source_path = safe_bundle_relative(source_name, "binding.source_manifest")
    require(sha_file(build_path) == build_sha256,
            "completed build receipt hash differs from binary binding")
    require(sha_file(source_path) == source_sha256,
            "source manifest hash differs from binary binding")
    source_value = load_json(source_path)
    require(isinstance(source_value, dict) and source_value,
            "bound source manifest is not a non-empty file map")
    actual_source_files = len(source_value)
    if source_files is None:
        source_files = actual_source_files
    require(source_files == actual_source_files,
            "bound source manifest file count differs")
    expected_source = (source_sha256, source_files)

    check_driver = ROOT.parent / "check.py"
    require_regular(check_driver, "parent check.py is missing or unsafe")
    build = standard_receipt(build_path, expected_source, check_driver, "build receipt", verify_log=False)
    build_cwd = text(build.get("cwd"), "build receipt.cwd")
    build_revision = text(build.get("revision"), "build receipt.revision")

    # This path is deliberately fixed and has no binding hash.  The final
    # bundle SHA256SUMS covers it after capture; before that, check.py's own
    # driver hash, command, source custody, and log hash authenticate it.
    gate_path = ROOT.parent / "checks" / "native-final.json"
    require_regular(gate_path, "fixed native run check receipt is missing or unsafe")
    gate = standard_receipt(gate_path, expected_source, check_driver, "native run check receipt")
    require(gate.get("revision") == build_revision,
            "native run check receipt revision differs from completed build")
    runner_from_gate(gate, runner_path, build_cwd)
    require(gate.get("cwd") == build_cwd,
            "native run check receipt cwd differs from completed build cwd")
    if precleanup:
        require(build_cwd == str(REPO.resolve()),
                "native run check receipt cwd differs from the current repository")

    binary = binding.get("binary")
    require(isinstance(binary, dict), "binary-binding.json has no binary object")
    binary_name = text(binary.get("path"), "binding.binary.path")
    binary_bytes = integer(binary.get("bytes"), "binding.binary.bytes", 1)
    binary_sha256 = digest(binary.get("sha256"), "binding.binary.sha256")
    if precleanup:
        binary_path = Path(binary_name)
        if not binary_path.is_absolute():
            binary_path = REPO / binary_path
        require_regular(binary_path, "bound native binary is missing or unsafe")
        require(binary_path.stat().st_size == binary_bytes and sha_file(binary_path) == binary_sha256,
                "bound native binary identity differs")
    return {
        "binding_path": binding_path,
        "binding_sha256": binding_sha256,
        "binding": binding,
        "binary": {
            "path": binary_name,
            "bytes": binary_bytes,
            "sha256": binary_sha256,
        },
        "inventory_path": inventory_path,
        "inventory_sha256": inventory_sha256,
        "oracle_path": oracle_path,
        "oracle_sha256": oracle_sha256,
        "runner_path": runner_path,
        "runner_sha256": runner_sha256,
        "build_path": build_path,
        "build_sha256": build_sha256,
        "source_path": source_path,
        "source_sha256": source_sha256,
        "source_files": source_files,
        "source_identity": expected_source,
        "build": build,
        "gate_path": gate_path,
        "gate": gate,
    }


def authenticate_inventory(path: Path, precleanup: bool) -> list[dict[str, Any]]:
    inventory = load_json(path)
    require(isinstance(inventory, dict), "static inventory root is not an object")
    fixtures = inventory.get("fixtures")
    require(isinstance(fixtures, list) and len(fixtures) == MAX_FIXTURES,
            "static inventory must contain exactly ten bounded fixtures")
    seen: set[str] = set()
    for index, fixture in enumerate(fixtures):
        require(isinstance(fixture, dict), f"inventory fixture {index} is not an object")
        fixture_path = text(fixture.get("path"), f"inventory fixture {index}.path")
        require(fixture_path not in seen, f"duplicate inventory fixture: {fixture_path}")
        seen.add(fixture_path)
        archive_bytes = integer(fixture.get("archive_bytes"), f"inventory {fixture_path}.archive_bytes")
        archive_sha256 = digest(fixture.get("archive_sha256"), f"inventory {fixture_path}.archive_sha256")
        if precleanup:
            source = REPO / fixture_path
            require(source.resolve().is_relative_to(REPO.resolve()),
                    f"inventory source escapes repository: {fixture_path}")
            require_regular(source, f"inventory source is missing or unsafe: {source}")
            require(source.stat().st_size == archive_bytes and sha_file(source) == archive_sha256,
                    f"immutable inventory source changed: {fixture_path}")
    return fixtures


def record_log(runs: Path, record: dict[str, Any], index: int) -> tuple[Path, dict[str, Any] | None]:
    expected_name = f"{index:02d}-probe.log"
    log_name = record.get("log")
    require(isinstance(log_name, str) and log_name == expected_name,
            f"record {index} probe log name differs")
    log_path = safe_relative(runs, log_name, f"record {index}.log")
    log_sha256 = digest(record.get("log_sha256"), f"record {index}.log_sha256")
    require(sha_file(log_path) == log_sha256, f"record {index} probe log hash differs")
    if record.get("exit_code") == 0:
        return log_path, parse_json_log(log_path, f"record {index} probe log")
    return log_path, None


def oracle_log(runs: Path, record: dict[str, Any], index: int) -> tuple[Path, dict[str, Any]]:
    oracle = record.get("oracle")
    require(isinstance(oracle, dict), f"record {index} has no oracle receipt")
    require(oracle.get("log") == f"{index:02d}-oracle.log",
            f"record {index} oracle log name differs")
    path = safe_relative(runs, oracle["log"], f"record {index}.oracle.log")
    require(sha_file(path) == digest(oracle.get("sha256"), f"record {index}.oracle.sha256"),
            f"record {index} oracle log hash differs")
    return path, parse_json_log(path, f"record {index} oracle log")


def check_probe_oracle(
    index: int,
    record: dict[str, Any],
    probe: dict[str, Any],
    oracle: dict[str, Any],
    fixture: dict[str, Any],
) -> None:
    require(probe.get("status") == "published", f"record {index} probe JSON is not published")
    require(probe.get("semantic_reopen") is True,
            f"record {index} probe JSON lacks semantic reopen success")
    require(oracle.get("status") == "validated", f"record {index} oracle JSON is not validated")
    title = text(record.get("title"), f"record {index}.title")
    body = text(record.get("body"), f"record {index}.body")
    require(oracle.get("request", {}).get("title") == title and
            oracle.get("request", {}).get("body") == body,
            f"record {index} oracle request text differs")
    source = oracle.get("source")
    output = oracle.get("output")
    content = oracle.get("content")
    require(isinstance(source, dict) and isinstance(output, dict) and isinstance(content, dict),
            f"record {index} oracle JSON lacks source/output/content objects")
    require(source.get("sha256") == fixture["archive_sha256"] and
            source.get("bytes") == fixture["archive_bytes"],
            f"record {index} oracle source identity differs")
    retained_output = record.get("retained_output")
    require(isinstance(retained_output, dict), f"record {index} has no retained output")
    output_record = record.get("output")
    require(isinstance(output_record, dict), f"record {index} has no output receipt")
    require(output_record.get("sha256") == retained_output.get("decoded_sha256") and
            output_record.get("bytes") == retained_output.get("decoded_bytes"),
            f"record {index} output receipt differs from retained output")
    require(output.get("sha256") == retained_output.get("decoded_sha256") and
            output.get("bytes") == retained_output.get("decoded_bytes"),
            f"record {index} oracle output identity differs")
    binding = source.get("inventory_binding")
    require(isinstance(binding, dict) and binding.get("sha256") == fixture["archive_sha256"] and
            binding.get("bytes") == fixture["archive_bytes"],
            f"record {index} oracle source inventory binding differs")

    generated = content.get("generated_page")
    require(isinstance(generated, dict), f"record {index} oracle generated page is missing")
    pairs = {
        "source_slides": content.get("source_page_count"),
        "target_slides": content.get("output_page_count"),
        "name": generated.get("name"),
        "insert_at": content.get("source_last_page_end_byte"),
        "source_content_bytes": content.get("source_content_xml_bytes"),
        "target_content_bytes": content.get("output_content_xml_bytes"),
    }
    for key, expected in pairs.items():
        require(probe.get(key) == expected,
                f"record {index} probe/oracle mismatch for {key}: {probe.get(key)!r} != {expected!r}")
    require(content.get("output_generated_page_start_byte") == probe.get("insert_at"),
            f"record {index} probe/oracle insertion offset differs")
    require(probe.get("bytes") == output.get("bytes"),
            f"record {index} probe/oracle output archive bytes differ")
    require(oracle.get("request", {}).get("name") == probe.get("name"),
            f"record {index} probe/oracle generated name differs")


def stable_oracle(value: dict[str, Any]) -> dict[str, Any]:
    """Remove only path and source inventory-binding fields that replay changes."""

    result = json.loads(json.dumps(value, ensure_ascii=False, sort_keys=True))
    for section in ("source", "output"):
        if isinstance(result.get(section), dict):
            result[section].pop("path", None)
    if isinstance(result.get("source"), dict):
        result["source"].pop("inventory_binding", None)
    return result


def replay_oracles(
    records: list[dict[str, Any]],
    fixtures: list[dict[str, Any]],
    runs: Path,
    oracle_path: Path,
) -> int:
    replayed = 0
    with tempfile.TemporaryDirectory(prefix="litchi-0457-native-replay-") as directory:
        temporary = Path(directory)
        script = temporary / "verify-output.py"
        shutil.copyfile(oracle_path, script)
        require_regular(script, "copied raw oracle is unsafe")
        for index, (record, fixture) in enumerate(zip(records, fixtures)):
            if record.get("exit_code") != 0:
                continue
            source_meta = record.get("retained_source")
            output_meta = record.get("retained_output")
            require(isinstance(source_meta, dict) and isinstance(output_meta, dict),
                    f"record {index} successful probe lacks retained archives")
            _, source_bytes = retained_archive(
                runs,
                source_meta,
                f"{index:02d}-source.odp.gz",
                fixture["archive_sha256"],
                fixture["archive_bytes"],
                f"record {index} retained source",
            )
            _, output_bytes = retained_archive(
                runs,
                output_meta,
                f"{index:02d}-output.odp.gz",
                output_meta["decoded_sha256"],
                output_meta["decoded_bytes"],
                f"record {index} retained output",
            )
            source_path = temporary / f"{index:02d}-source.odp"
            output_path = temporary / f"{index:02d}-output.odp"
            source_path.write_bytes(source_bytes)
            output_path.write_bytes(output_bytes)
            command = [
                sys.executable,
                "-B",
                str(script),
                "--source",
                str(source_path),
                "--output",
                str(output_path),
                "--title",
                text(record.get("title"), f"record {index}.title"),
                "--body",
                text(record.get("body"), f"record {index}.body"),
            ]
            result = subprocess.run(command, capture_output=True, text=True)
            recorded_oracle = record.get("oracle")
            require(isinstance(recorded_oracle, dict), f"record {index} lacks oracle receipt")
            expected_exit = integer(recorded_oracle.get("exit_code"), f"record {index}.oracle.exit_code")
            require(result.returncode == expected_exit,
                    f"record {index} replay oracle exit differs: {result.returncode} != {expected_exit}")
            replay_json = parse_json_bytes(result.stdout.encode("utf-8"), f"record {index} replay oracle")
            oracle_log_path, recorded_json = oracle_log(runs, record, index)
            del oracle_log_path
            if expected_exit == 0:
                require(stable_oracle(replay_json) == stable_oracle(recorded_json),
                        f"record {index} replay oracle facts differ from retained oracle JSON")
            else:
                require(replay_json.get("status") == recorded_json.get("status") and
                        replay_json.get("error", {}).get("kind") == recorded_json.get("error", {}).get("kind"),
                        f"record {index} replay refusal/error differs from retained oracle JSON")
            replayed += 1
    return replayed


def parse_json_bytes(raw: bytes, label: str) -> dict[str, Any]:
    try:
        value = json.loads(raw.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as error:
        fail(f"{label} is not exactly one JSON object: {error}")
    require(isinstance(value, dict), f"{label} root is not an object")
    return value


def verify_records(
    inventory: list[dict[str, Any]],
    bindings: dict[str, Any],
    precleanup: bool,
) -> tuple[list[dict[str, Any]], int, int, int]:
    runs = ROOT / "runs"
    require_directory(runs, "native runs directory is missing")
    index_path = runs / "index.json"
    index = load_json(index_path)
    require(isinstance(index, dict), "native runs index root is not an object")
    require(index.get("binding_sha256") == bindings["binding_sha256"],
            "native runs index is not bound to binary-binding.json")
    records = index.get("records")
    require(isinstance(records, list) and len(records) == len(inventory),
            "native runs record count differs from static inventory")
    expected_receipts = {f"{i:02d}-receipt.json" for i in range(len(inventory))}
    actual_receipts = {path.name for path in runs.glob("*-receipt.json") if path.is_file()}
    require(actual_receipts == expected_receipts, "native runs receipt set differs")

    validated = 0
    probe_failures = 0
    oracle_failures = 0
    for index_number, (record, fixture) in enumerate(zip(records, inventory)):
        require(isinstance(record, dict), f"record {index_number} is not an object")
        require(record.get("fixture") == fixture["path"], f"record {index_number} fixture order differs")
        receipt_path = runs / f"{index_number:02d}-receipt.json"
        receipt = load_json(receipt_path)
        require(receipt == record, f"record {index_number} differs from its receipt JSON")
        source_sha256 = digest(record.get("source_sha256"), f"record {index_number}.source_sha256")
        require(source_sha256 == fixture["archive_sha256"], f"record {index_number} source hash differs")
        probe_argv = record.get("argv")
        require(isinstance(probe_argv, list) and len(probe_argv) == 5 and
                all(isinstance(value, str) and value for value in probe_argv),
                f"record {index_number} probe argv shape differs")
        require(probe_argv[3] == record.get("title") and probe_argv[4] == record.get("body"),
                f"record {index_number} probe argv text differs")
        source_meta = record.get("retained_source")
        require(isinstance(source_meta, dict), f"record {index_number} has no retained source")
        retained_archive(
            runs,
            source_meta,
            f"{index_number:02d}-source.odp.gz",
            fixture["archive_sha256"],
            fixture["archive_bytes"],
            f"record {index_number} retained source",
        )
        _, probe = record_log(runs, record, index_number)
        exit_code = record.get("exit_code")
        require(isinstance(exit_code, int) and not isinstance(exit_code, bool),
                f"record {index_number}.exit_code is not an integer")
        output_meta = record.get("retained_output")
        if output_meta is not None:
            require(isinstance(output_meta, dict), f"record {index_number}.retained_output is not an object")
            output_record = record.get("output")
            require(isinstance(output_record, dict), f"record {index_number}.output is not an object")
            output_record_bytes = integer(output_record.get("bytes"), f"record {index_number}.output.bytes")
            output_record_sha256 = digest(output_record.get("sha256"), f"record {index_number}.output.sha256")
            require(output_record_bytes == integer(output_meta.get("decoded_bytes"), f"record {index_number}.retained_output.decoded_bytes") and
                    output_record_sha256 == digest(output_meta.get("decoded_sha256"), f"record {index_number}.retained_output.decoded_sha256"),
                    f"record {index_number} output receipt differs from retained output")
            retained_archive(
                runs,
                output_meta,
                f"{index_number:02d}-output.odp.gz",
                digest(output_meta.get("decoded_sha256"), f"record {index_number}.retained_output.decoded_sha256"),
                integer(output_meta.get("decoded_bytes"), f"record {index_number}.retained_output.decoded_bytes"),
                f"record {index_number} retained output",
            )
        if precleanup:
            source_path = REPO / fixture["path"]
            require(sha_file(source_path) == fixture["archive_sha256"],
                    f"record {index_number} source changed before cleanup")
            argv = record.get("argv")
            require(isinstance(argv, list) and len(argv) == 5,
                    f"record {index_number} probe argv shape differs")
            require(argv[0] == bindings["binary"]["path"] and argv[1] == str(source_path),
                    f"record {index_number} probe argv source/binary differs")
            require(argv[2] == str(TASK / f"{index_number:02d}.odp") and
                    argv[3] == record.get("title") and argv[4] == record.get("body") and
                    record.get("cwd") == str(REPO),
                    f"record {index_number} probe argv/output/cwd differs")
            if output_meta is not None:
                output_record = record.get("output")
                require(isinstance(output_record, dict),
                        f"record {index_number}.output is not an object")
                expected_output = TASK / f"{index_number:02d}.odp"
                output_path = Path(output_record.get("path", ""))
                # run.py writes producer outputs under its fixed task directory
                # and retains a separate gzip under runs/.  Validate the
                # producer path from output.path; output_meta.path is the
                # retained gzip name and must not be mistaken for the live
                # output archive.
                require(output_record.get("path") == str(expected_output) and
                        output_path.is_absolute() and
                        output_path.resolve() == expected_output.resolve(),
                        f"record {index_number} output escapes native task directory")
                require_regular(output_path, f"record {index_number} output is missing before cleanup")
                require(output_path.stat().st_size == output_record["bytes"] and
                        sha_file(output_path) == output_record["sha256"],
                        f"record {index_number} output identity differs before cleanup")

        if exit_code != 0:
            require(record.get("status") == "probe-failed",
                    f"record {index_number} failed probe has invalid status")
            require(record.get("oracle") is None, f"record {index_number} failed probe has oracle receipt")
            probe_failures += 1
            continue
        require(probe is not None, f"record {index_number} successful probe has no JSON")
        require(output_meta is not None, f"record {index_number} successful probe has no retained output")
        oracle_path, oracle_json = oracle_log(runs, record, index_number)
        del oracle_path
        oracle_receipt = record.get("oracle")
        require(isinstance(oracle_receipt, dict), f"record {index_number} successful probe has no oracle receipt")
        oracle_argv = oracle_receipt.get("argv")
        require(isinstance(oracle_argv, list) and len(oracle_argv) == 11 and
                all(isinstance(value, str) and value for value in oracle_argv),
                f"record {index_number} oracle argv shape differs")
        require(oracle_argv[1] == "-B" and Path(oracle_argv[2]).name == "verify-output.py" and
                oracle_argv[3] == "--source" and oracle_argv[5] == "--output" and
                oracle_argv[7] == "--title" and oracle_argv[9] == "--body" and
                oracle_argv[4] == probe_argv[1] and oracle_argv[6] == probe_argv[2] and
                oracle_argv[8] == record.get("title") and oracle_argv[10] == record.get("body"),
                f"record {index_number} oracle argv binding differs")
        oracle_exit = oracle_receipt.get("exit_code")
        require(isinstance(oracle_exit, int) and not isinstance(oracle_exit, bool),
                f"record {index_number} oracle exit is not an integer")
        check_probe_oracle(index_number, record, probe, oracle_json, fixture) if oracle_exit == 0 else None
        if oracle_exit == 0:
            require(record.get("status") == "validated", f"record {index_number} successful oracle has invalid status")
            validated += 1
        else:
            require(record.get("status") == "oracle-failed", f"record {index_number} failed oracle has invalid status")
            oracle_failures += 1
    return records, validated, probe_failures, oracle_failures


def verify_bundle_seal() -> tuple[Path, int]:
    """Verify the final SHA256SUMS over the complete retained bundle."""

    seal_path = ROOT.parent / "SHA256SUMS"
    base = ROOT.parent
    require(seal_path.is_file(),
            "final change-0457 SHA256SUMS is required to seal the native evidence bundle")
    rows: dict[str, str] = {}
    try:
        lines = seal_path.read_text(encoding="utf-8").splitlines()
    except (OSError, UnicodeError) as error:
        fail(f"cannot read final SHA256SUMS: {error}")
    for line in lines:
        fields = line.split("  ", 1)
        require(len(fields) == 2 and SHA256_RE.fullmatch(fields[0]) is not None,
                "malformed SHA256SUMS line")
        name = fields[1]
        candidate = Path(name)
        require(name not in rows and name and not candidate.is_absolute() and
                ".." not in candidate.parts and name != "SHA256SUMS",
                f"unsafe or duplicate SHA256SUMS member: {name}")
        try:
            resolved = (base / candidate).resolve(strict=True)
        except OSError as error:
            fail(f"SHA256SUMS member cannot be resolved: {error}")
        require(resolved.is_relative_to(base.resolve()),
                f"SHA256SUMS member escapes bundle: {name}")
        member = base / candidate
        require_regular(member, f"SHA256SUMS member is missing or unsafe: {name}")
        require(sha_file(member) == fields[0], f"SHA256SUMS hash differs: {name}")
        rows[name] = fields[0]
    actual: set[str] = set()
    for member in base.rglob("*"):
        if member.is_symlink():
            fail(f"bundle contains an unsealed symlink: {member.relative_to(base)}")
        if member.is_file() and member != seal_path:
            actual.add(str(member.relative_to(base)))
    require(set(rows) == actual, "SHA256SUMS does not cover exactly the evidence bundle")
    return seal_path, len(rows)


def portable_copy_check() -> None:
    with tempfile.TemporaryDirectory(prefix="litchi-0457-native-bundle-") as directory:
        # Keep the fixed ../checks/native-final.json gate and its standard build/source
        # artifacts beside native/.  This copied change bundle contains no
        # original fixture or executable; precleanup checks are disabled below.
        exported_bundle = Path(directory) / ROOT.parent.name
        shutil.copytree(ROOT.parent, exported_bundle, symlinks=True)
        exported = exported_bundle / ROOT.name
        result = subprocess.run(
            [sys.executable, "-B", str(exported / "verify.py")],
            cwd=exported,
            capture_output=True,
            text=True,
        )
        require(result.returncode == 0,
                f"portable copied verifier failed: {result.stderr.strip() or result.stdout.strip()}")
        replay = parse_json_bytes(result.stdout.encode("utf-8"), "portable copied verifier output")
        require(replay.get("status") == "pass" and replay.get("portable") is False,
                "portable copied verifier did not complete an ordinary bundle pass")


def verify(precleanup: bool, portable: bool) -> dict[str, Any]:
    require(not (precleanup and portable), "--precleanup and --portable are mutually exclusive")
    bindings = authenticate_bindings(precleanup)
    seal_path, sealed_files = verify_bundle_seal()
    fixtures = authenticate_inventory(bindings["inventory_path"], precleanup)
    records, validated, probe_failures, oracle_failures = verify_records(fixtures, bindings, precleanup)
    replayed = replay_oracles(records, fixtures, ROOT / "runs", bindings["oracle_path"])
    require(oracle_failures == 0,
            f"capture contains {oracle_failures} independent oracle failure(s)")
    result: dict[str, Any] = {
        "schema": SCHEMA,
        "status": "pass",
        "change": 457,
        "precleanup": precleanup,
        "portable": portable,
        "records": len(records),
        "validated_records": validated,
        "probe_failures": probe_failures,
        "oracle_replays": replayed,
        "binary_checked": precleanup,
        "source_fixtures_checked": precleanup,
        "sealed_bundle": str(seal_path.relative_to(ROOT.parent if seal_path.parent == ROOT.parent else ROOT)),
        "sealed_files": sealed_files,
        "static_only_classification": "retained raw ZIP/XML oracle replay; native semantic reopen is authenticated only from probe JSON",
    }
    if portable:
        portable_copy_check()
        result["portable_copy"] = "pass; copied verifier replayed without original fixture or binary"
    return result


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--precleanup", action="store_true",
                        help="also require the original source fixtures and bound binary")
    parser.add_argument("--portable", action="store_true",
                        help="copy this native bundle and replay its verifier")
    args = parser.parse_args(sys.argv[1:] if argv is None else argv)
    try:
        result = verify(args.precleanup, args.portable)
    except (VerificationError, OSError, ValueError, KeyError) as error:
        print(json.dumps({
            "schema": SCHEMA,
            "status": "fail",
            "error": str(error),
            "precleanup": args.precleanup,
            "portable": args.portable,
        }, ensure_ascii=False, sort_keys=True, indent=2))
        return 1
    print(json.dumps(result, ensure_ascii=False, sort_keys=True, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
