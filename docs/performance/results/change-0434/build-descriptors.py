#!/usr/bin/env python3
"""Create the immutable 0434 before/after build descriptors.

The build and binary-copy receipts are produced by earlier serialized steps.
This script only validates those receipts and records their identities next to
the frozen outer protocol and its copied oracle.  It deliberately refuses to
replace an existing descriptor: changing a descriptor after capture would
make the retained executable/source binding ambiguous.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import re
from typing import Any


ROOT = Path(__file__).resolve().parent
HEX40 = re.compile(r"^[0-9a-fA-F]{40}$")
HEX64 = re.compile(r"^[0-9a-fA-F]{64}$")
ROLES = ("before", "after")
EXPECTED_BUILD = (
    "cargo",
    "build",
    "--locked",
    "--release",
    "--manifest-path",
    "tools/perf-baseline/Cargo.toml",
    "--features",
    "allocator-metrics",
    "--bin",
    "litchi-perf-baseline",
    "--bin",
    "litchi-perf-baseline-alloc",
)


class DescriptorError(ValueError):
    """An input receipt cannot support an immutable build descriptor."""


def fail(label: str, message: str) -> None:
    raise DescriptorError(f"{label}: {message}")


def load(path: Path, label: str) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(label, f"invalid JSON: {error}")
    raise AssertionError("unreachable")


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def require_object(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(label, "expected an object")
    return value


def require_text(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        fail(label, "expected a non-empty string")
    return value


def require_digest(value: Any, label: str) -> str:
    if not isinstance(value, str) or HEX64.fullmatch(value) is None:
        fail(label, "expected a SHA-256 digest")
    return value.lower()


def require_relative_file(value: Any, label: str) -> Path:
    relative = require_text(value, label)
    path = Path(relative)
    if path.is_absolute():
        fail(label, "must be bundle-relative")
    resolved = (ROOT / path).resolve()
    if not resolved.is_relative_to(ROOT.resolve()) or not resolved.is_file():
        fail(label, "retained file is missing or escapes the bundle")
    return resolved


def checked_source_manifest(receipt: dict[str, Any], label: str) -> dict[str, Any]:
    before = require_object(receipt.get("source_before"), f"{label}.source_before")
    after = require_object(receipt.get("source_after"), f"{label}.source_after")
    if before != after:
        fail(label, "source_before and source_after differ")
    path = require_relative_file(before.get("path"), f"{label}.source_manifest.path")
    expected_sha = require_digest(before.get("sha256"), f"{label}.source_manifest.sha256")
    files = before.get("files")
    if isinstance(files, bool) or not isinstance(files, int) or files < 1:
        fail(f"{label}.source_manifest.files", "expected a positive integer")
    if sha(path) != expected_sha:
        fail(label, "retained source manifest hash differs from the receipt")
    manifest = load(path, str(path))
    if not isinstance(manifest, dict) or len(manifest) != files:
        fail(label, "retained source manifest file count differs from the receipt")
    return {"path": before["path"], "sha256": expected_sha, "files": files}


def assignment(argv: list[str], name: str, label: str) -> str:
    values = [item[len(name) + 1:] for item in argv if item.startswith(name + "=")]
    if len(values) != 1 or not values[0]:
        fail(label, f"expected one {name}= assignment")
    return values[0]


def checked_build_receipt(role: str) -> tuple[Path, dict[str, Any], dict[str, Any]]:
    path = ROOT / "checks" / f"{role}-build.json"
    receipt = require_object(load(path, str(path)), str(path))
    label = f"checks/{role}-build.json"
    if receipt.get("change") != 434:
        fail(label, "wrong change")
    if receipt.get("status") != "pass" or receipt.get("exit_code") != 0:
        fail(label, "build receipt is not passing")
    if receipt.get("source_unchanged") is not True:
        fail(label, "source custody is not unchanged")
    revision = receipt.get("revision")
    if not isinstance(revision, str) or HEX40.fullmatch(revision) is None:
        fail(label, "malformed source revision")
    check_driver = require_digest(receipt.get("driver_sha256"), f"{label}.driver_sha256")
    check_driver_path = ROOT / "check.py"
    if check_driver != sha(check_driver_path):
        fail(label, "source-custody driver hash differs from check.py")
    source = checked_source_manifest(receipt, label)

    argv = receipt.get("argv")
    if not isinstance(argv, list) or any(not isinstance(item, str) for item in argv):
        fail(label, "build argv must be a string array")
    if argv[:1] != ["env"] or list(argv[3:]) != list(EXPECTED_BUILD):
        fail(label, "build argv is not the frozen release harness build")
    rustflags = assignment(argv, "RUSTFLAGS", f"{label}.argv")
    release_debug_text = assignment(argv, "CARGO_PROFILE_RELEASE_DEBUG", f"{label}.argv")
    try:
        release_debug = int(release_debug_text, 10)
    except ValueError:
        fail(label, "release debug value is not an integer")
    if release_debug != 1:
        fail(label, "release debug must be 1")

    environment = require_object(receipt.get("environment"), f"{label}.environment")
    toolchain = require_text(environment.get("RUSTUP_TOOLCHAIN"), f"{label}.environment.RUSTUP_TOOLCHAIN")
    for key, expected in (("CARGO_INCREMENTAL", "0"), ("PYTHONDONTWRITEBYTECODE", "1")):
        if environment.get(key) != expected:
            fail(label, f"environment {key} must be {expected!r}")
    try:
        jobs = int(require_text(environment.get("CARGO_BUILD_JOBS"), f"{label}.environment.CARGO_BUILD_JOBS"), 10)
    except ValueError:
        fail(label, "CARGO_BUILD_JOBS is not an integer")
    if jobs < 1:
        fail(label, "CARGO_BUILD_JOBS must be positive")

    source_scope = require_text(receipt.get("source_scope"), f"{label}.source_scope")
    if source_scope != "workspace and standalone tools":
        fail(label, "unexpected source scope")
    finished = require_text(receipt.get("finished_utc"), f"{label}.finished_utc")
    receipt_details = {
        "change": 434,
        "revision": revision,
        "source_manifest": source,
        "build_argv": list(argv),
        "build_environment": dict(environment),
        "rustflags": rustflags,
        "release_debug": release_debug,
        "rust_toolchain": toolchain,
        "profile": "release",
        "source_scope": source_scope,
        "created_utc": finished,
        "check_driver_sha256": check_driver,
    }
    return path, receipt_details, receipt


def checked_binary_copies(role: str) -> tuple[Path, dict[str, dict[str, Any]]]:
    path = ROOT / role / "binary-copies.json"
    copies = require_object(load(path, str(path)), str(path))
    if set(copies) != {"normal", "allocator"}:
        fail(str(path), "normal and allocator identities are required")
    checked: dict[str, dict[str, Any]] = {}
    for mode in ("normal", "allocator"):
        identity = require_object(copies[mode], f"{path}.{mode}")
        binary = Path(require_text(identity.get("path"), f"{path}.{mode}.path"))
        if not binary.is_absolute() or not binary.is_file():
            fail(f"{path}.{mode}.path", "copied binary is missing or not absolute")
        expected_sha = require_digest(identity.get("sha256"), f"{path}.{mode}.sha256")
        size = identity.get("bytes")
        if isinstance(size, bool) or not isinstance(size, int) or size <= 0:
            fail(f"{path}.{mode}.bytes", "expected a positive integer")
        if binary.stat().st_size != size or sha(binary) != expected_sha:
            fail(f"{path}.{mode}", "copied binary hash or size differs from receipt")
        checked[mode] = {"path": str(binary), "sha256": expected_sha, "bytes": size}
    return path, checked


def build_descriptor(role: str, protocol: dict[str, Any], protocol_sha: str,
                     hashes: dict[str, str]) -> dict[str, Any]:
    receipt_path, build, receipt = checked_build_receipt(role)
    copies_path, binaries = checked_binary_copies(role)
    oracle = require_object(protocol.get("oracle"), "protocol.oracle")
    oracle_protocol_sha = require_digest(oracle.get("sha256"), "protocol.oracle.sha256")
    oracle_verifier_sha = require_digest(oracle.get("verifier_sha256"), "protocol.oracle.verifier_sha256")
    if oracle_protocol_sha != hashes["oracle_protocol"] or oracle_verifier_sha != hashes["oracle_verifier"]:
        fail("protocol.oracle", "declared oracle hashes differ from retained files")
    descriptor = {
        "change": 434,
        "role": role,
        "revision": build["revision"],
        "created_utc": build["created_utc"],
        "source_manifest": build["source_manifest"],
        "protocol_sha256": protocol_sha,
        # 0434 capture/profile receipts use verifier.py as the outer verifier;
        # retain the explicit alias while keeping the historical key readable.
        "verifier_sha256": hashes["outer_verifier"],
        "outer_verifier_sha256": hashes["outer_verifier"],
        "capture_driver_sha256": hashes["capture"],
        "profile_driver_sha256": hashes["profile"],
        "oracle_protocol_sha256": oracle_protocol_sha,
        "oracle_verifier_sha256": oracle_verifier_sha,
        "binaries": binaries,
        "build_receipt": str(receipt_path.relative_to(ROOT)),
        "build_receipt_sha256": sha(receipt_path),
        "binary_copies_receipt": str(copies_path.relative_to(ROOT)),
        "binary_copies_receipt_sha256": sha(copies_path),
        "check_driver_sha256": build["check_driver_sha256"],
        "build_argv": build["build_argv"],
        "build_environment": build["build_environment"],
        "rustflags": build["rustflags"],
        "release_debug": build["release_debug"],
        "rust_toolchain": build["rust_toolchain"],
        "profile": build["profile"],
        "source_scope": build["source_scope"],
        "capture_script_frozen": True,
        "driver_hashes": {
            "verify.py": hashes["outer_verifier"],
            "capture.py": hashes["capture"],
            "profile.py": hashes["profile"],
            "oracle/protocol.json": hashes["oracle_protocol"],
            "oracle/verify-report.py": hashes["oracle_verifier"],
        },
    }
    # The receipt is kept as an input and is intentionally not copied into
    # the descriptor.  This check makes an accidental role mix-up fail before
    # either output is written.
    if receipt["source_before"] != receipt["source_after"]:
        fail(str(receipt_path), "source receipt changed during descriptor creation")
    return descriptor


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--role", choices=ROLES, action="append", help="descriptor role (default: both)")
    args = parser.parse_args()
    roles = tuple(dict.fromkeys(args.role or ROLES))
    targets = [ROOT / role / "build.json" for role in roles]
    existing = [str(path) for path in targets if path.exists()]
    if existing:
        raise SystemExit("refusing to overwrite existing descriptor(s): " + ", ".join(existing))

    protocol_path = ROOT / "protocol.json"
    protocol = require_object(load(protocol_path, str(protocol_path)), str(protocol_path))
    if protocol.get("change") != 434:
        fail("protocol.change", "must be 434")
    protocol_sha = sha(protocol_path)
    frozen_path = ROOT / "frozen-inputs.json"
    frozen = require_object(load(frozen_path, str(frozen_path)), str(frozen_path))
    if frozen.get("status") != "frozen":
        fail("frozen-inputs.status", "outer inputs are not frozen")
    frozen_files = require_object(frozen.get("files"), "frozen-inputs.files")
    frozen_expected = {
        "protocol.json": protocol_path,
        "oracle/protocol.json": ROOT / "oracle" / "protocol.json",
        "oracle/verify-report.py": ROOT / "oracle" / "verify-report.py",
        "check.py": ROOT / "check.py",
    }
    for name, path in frozen_expected.items():
        expected = require_digest(frozen_files.get(name), f"frozen-inputs.files.{name}")
        if expected != sha(path):
            fail(f"frozen-inputs.files.{name}", "frozen input hash is stale")
    oracle = require_object(protocol.get("oracle"), "protocol.oracle")
    oracle_protocol_path = require_relative_file(oracle.get("path"), "protocol.oracle.path")
    oracle_verifier_path = require_relative_file(oracle.get("verifier_path"), "protocol.oracle.verifier_path")
    files = {
        "outer_verifier": ROOT / "verify.py",
        "capture": ROOT / "capture.py",
        "profile": ROOT / "profile.py",
        "oracle_protocol": oracle_protocol_path,
        "oracle_verifier": oracle_verifier_path,
    }
    for name, path in files.items():
        if not path.is_file():
            fail(name, "frozen driver/oracle is missing")
    hashes = {name: sha(path) for name, path in files.items()}
    if require_digest(oracle.get("sha256"), "protocol.oracle.sha256") != hashes["oracle_protocol"]:
        fail("protocol.oracle.sha256", "does not match retained oracle protocol")
    if require_digest(oracle.get("verifier_sha256"), "protocol.oracle.verifier_sha256") != hashes["oracle_verifier"]:
        fail("protocol.oracle.verifier_sha256", "does not match retained oracle verifier")
    descriptors = {
        role: build_descriptor(role, protocol, protocol_sha, hashes)
        for role in roles
    }
    for role, descriptor in descriptors.items():
        target = ROOT / role / "build.json"
        target.parent.mkdir(parents=True, exist_ok=True)
        with target.open("x", encoding="utf-8") as stream:
            json.dump(descriptor, stream, indent=2)
            stream.write("\n")
        print(target)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
