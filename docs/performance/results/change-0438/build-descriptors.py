#!/usr/bin/env python3
"""Freeze 0438 build and executable identities.

This file only consumes already completed build and binary-copy receipts.  It
does not invoke Cargo, Git, or a workload.  The evidence root is the directory
containing this script; all retained source manifests and copied receipts must
therefore be present below that root before a descriptor can be written.

The ODP report schema belongs to the copied oracle and is deliberately absent
from this descriptor step.  The descriptor binds the source, binaries,
protocol, and driver bytes that the later capture/profile steps consume.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import re
from typing import Any


ROOT = Path(__file__).resolve().parent
ROLES = ("before", "after")
HEX40 = re.compile(r"^[0-9a-fA-F]{40}$")
HEX64 = re.compile(r"^[0-9a-fA-F]{64}$")


class DescriptorError(ValueError):
    pass


def fail(label: str, message: str) -> None:
    raise DescriptorError(f"{label}: {message}")


def load(path: Path, label: str) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(label, f"invalid JSON: {error}")
    raise AssertionError("unreachable")


def obj(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(label, "expected an object")
    return value


def text(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        fail(label, "expected non-empty text")
    return value


def digest(value: Any, label: str) -> str:
    value = text(value, label)
    if HEX64.fullmatch(value) is None:
        fail(label, "expected a SHA-256 digest")
    return value.lower()


def sha(path: Path) -> str:
    digest_value = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest_value.update(block)
    return digest_value.hexdigest()


def relative_file(value: Any, label: str) -> Path:
    relative = Path(text(value, label))
    if relative.is_absolute():
        fail(label, "must be bundle-relative")
    path = (ROOT / relative).resolve()
    if not path.is_relative_to(ROOT.resolve()) or not path.is_file():
        fail(label, "retained file is missing or escapes the evidence root")
    return path


def source_manifest(receipt: dict[str, Any], label: str) -> dict[str, Any]:
    before = obj(receipt.get("source_before"), f"{label}.source_before")
    after = obj(receipt.get("source_after"), f"{label}.source_after")
    if before != after:
        fail(label, "source custody changed during the build")
    path = relative_file(before.get("path"), f"{label}.source_manifest.path")
    expected = digest(before.get("sha256"), f"{label}.source_manifest.sha256")
    files = before.get("files")
    if isinstance(files, bool) or not isinstance(files, int) or files < 1:
        fail(label, "source manifest file count is invalid")
    if sha(path) != expected:
        fail(label, "source manifest hash differs from receipt")
    manifest = obj(load(path, str(path)), str(path))
    if len(manifest) != files:
        fail(label, "source manifest file count differs from receipt")
    return {"path": before["path"], "sha256": expected, "files": files}


def checked_build_receipt(role: str) -> tuple[Path, dict[str, Any]]:
    path = ROOT / "checks" / f"{role}-build.json"
    receipt = obj(load(path, str(path)), str(path))
    label = f"checks/{role}-build.json"
    if receipt.get("change") != 438:
        fail(label, "change must be 438")
    if receipt.get("status") != "pass" or receipt.get("exit_code") != 0:
        fail(label, "build receipt is not passing")
    if receipt.get("source_unchanged") is not True:
        fail(label, "build receipt does not prove source custody")
    revision = text(receipt.get("revision"), f"{label}.revision")
    if HEX40.fullmatch(revision) is None:
        fail(label, "revision is not a Git SHA-1")
    source = source_manifest(receipt, label)
    argv = receipt.get("argv")
    if not isinstance(argv, list) or any(not isinstance(item, str) for item in argv):
        fail(label, "build argv must be a string array")
    environment = obj(receipt.get("environment"), f"{label}.environment")
    source_scope = text(receipt.get("source_scope"), f"{label}.source_scope")
    details = {
        "change": 438,
        "revision": revision,
        "source_manifest": source,
        "build_argv": list(argv),
        "build_environment": dict(environment),
        "source_scope": source_scope,
        "created_utc": text(receipt.get("finished_utc"), f"{label}.finished_utc"),
        "check_driver_sha256": digest(receipt.get("driver_sha256"), f"{label}.driver_sha256"),
    }
    return path, details


def checked_binary_copies(role: str) -> tuple[Path, dict[str, dict[str, Any]]]:
    path = ROOT / role / "binary-copies.json"
    copies = obj(load(path, str(path)), str(path))
    if set(copies) != {"normal", "allocator"}:
        fail(str(path), "normal and allocator copies are required")
    checked: dict[str, dict[str, Any]] = {}
    for mode in ("normal", "allocator"):
        identity = obj(copies[mode], f"{path}.{mode}")
        binary = Path(text(identity.get("path"), f"{path}.{mode}.path"))
        if not binary.is_absolute() or not binary.is_file():
            fail(f"{path}.{mode}.path", "copied executable is missing or not absolute")
        size = identity.get("bytes")
        if isinstance(size, bool) or not isinstance(size, int) or size <= 0:
            fail(f"{path}.{mode}.bytes", "expected positive byte count")
        expected = digest(identity.get("sha256"), f"{path}.{mode}.sha256")
        if binary.stat().st_size != size or sha(binary) != expected:
            fail(f"{path}.{mode}", "copied executable hash or size differs")
        checked[mode] = {"path": str(binary), "bytes": size, "sha256": expected}
    return path, checked


def optional_input(path: Path) -> dict[str, Any] | None:
    if not path.is_file():
        return None
    return {"path": str(path.relative_to(ROOT)), "bytes": path.stat().st_size, "sha256": sha(path)}


def oracle_bindings(protocol: dict[str, Any]) -> dict[str, Any]:
    oracle = protocol.get("oracle")
    if oracle is None:
        return {}
    oracle = obj(oracle, "protocol.oracle")
    bindings: dict[str, Any] = {}
    for key in ("path", "verifier_path", "protocol_path"):
        if key not in oracle:
            continue
        path = relative_file(oracle[key], f"protocol.oracle.{key}")
        bindings[key] = str(path.relative_to(ROOT))
        expected_key = {"path": "sha256", "verifier_path": "verifier_sha256", "protocol_path": "protocol_sha256"}[key]
        if expected_key in oracle:
            expected = digest(oracle[expected_key], f"protocol.oracle.{expected_key}")
            if sha(path) != expected:
                fail(f"protocol.oracle.{key}", "retained oracle hash differs")
            bindings[expected_key] = expected
    return bindings


def build_descriptor(role: str, protocol: dict[str, Any], protocol_path: Path) -> dict[str, Any]:
    build_path, build = checked_build_receipt(role)
    copies_path, binaries = checked_binary_copies(role)
    descriptor: dict[str, Any] = {
        "schema": "litchi-0438-build-descriptor-v1",
        "change": 438,
        "role": role,
        "revision": build["revision"],
        "created_utc": build["created_utc"],
        "source_manifest": build["source_manifest"],
        "protocol_sha256": sha(protocol_path),
        "binaries": binaries,
        "build_receipt": str(build_path.relative_to(ROOT)),
        "build_receipt_sha256": sha(build_path),
        "binary_copies_receipt": str(copies_path.relative_to(ROOT)),
        "binary_copies_receipt_sha256": sha(copies_path),
        "check_driver_sha256": build["check_driver_sha256"],
        "build_argv": build["build_argv"],
        "build_environment": build["build_environment"],
        "source_scope": build["source_scope"],
    }
    driver_paths = {
        "capture.py": ROOT / "capture.py",
        "profile.py": ROOT / "profile.py",
        "check.py": ROOT / "check.py",
    }
    descriptor["driver_hashes"] = {
        name: sha(path) for name, path in driver_paths.items() if path.is_file()
    }
    descriptor["oracle"] = oracle_bindings(protocol)
    return descriptor


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--protocol", type=Path, default=ROOT / "protocol.json")
    parser.add_argument("--role", choices=ROLES, action="append")
    args = parser.parse_args()
    try:
        protocol = obj(load(args.protocol, str(args.protocol)), str(args.protocol))
        if protocol.get("change") != 438:
            fail("protocol.change", "must be 438")
        roles = tuple(dict.fromkeys(args.role or ROLES))
        descriptors = {role: build_descriptor(role, protocol, args.protocol) for role in roles}
        for role, descriptor in descriptors.items():
            target = ROOT / role / "build.json"
            if target.exists():
                fail(str(target), "refusing to overwrite existing descriptor")
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_text(json.dumps(descriptor, indent=2, sort_keys=True) + "\n", encoding="utf-8")
            print(target)
    except (OSError, TypeError, ValueError, DescriptorError) as error:
        print(f"DESCRIPTOR INVALID: {error}")
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
