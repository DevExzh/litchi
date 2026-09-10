#!/usr/bin/env python3
"""Create and verify the bounded 0491 evidence inventory seal.

The coordinator chooses the final receipts explicitly when creating a seal.
Verification reads those bindings back from ``seal.json`` and rechecks the
files and the small set of terminal invariants.  This helper only reads
evidence, apart from its exclusive creation of ``seal.json``; it never builds
or captures a run.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
import os
from pathlib import Path
import stat
import sys
from typing import Any, Mapping

from support import ENV, ENV_KEYS, REPO


sys.dont_write_bytecode = True

ROOT = Path(__file__).resolve().parent
SEAL_NAME = "seal.json"
SCHEMA = "docx-office-0491-evidence-seal-v1"
VERIFICATION_SCHEMA = "docx-office-0491-seal-verification-v1"
VERSION = 1
BYTECODE_SUFFIXES = {".pyc", ".pyo"}
GATE_SCHEMA = "docx-stream-append-gate-v1"
PROVIDER_VERIFICATION_SCHEMA = "docx-provider-matrix-verification-v1"
COLD_VERIFICATION_SCHEMA = "docx-filesystem-cold-matrix-v1-verification"
CLEANUP_SCHEMA = "docx-office-0491-cleanup-v1"
CLEANUP_VERIFICATION_SCHEMA = "docx-office-0491-cleanup-verification-v1"
PROFILE_RESULT_PATH = "profile-providers/profiles2/result.json"
REQUIRED_FINAL_GATE_LABELS = frozenset(
    {
        "build-normal-final5",
        "build-allocator-final5",
        "test-lib-final5",
        "test-allocation-final5",
        "clippy-final5",
        "rustdoc-final5",
        "python-final7",
        "provider-formal1",
        "cold-formal1",
        "profiles2",
        "profile-verify2",
    }
)
PROVIDER_VERIFICATION_KEYS = frozenset(
    {
        "schema", "version", "status", "attempt", "pilot",
        "protocol_sha256", "summary_sha256", "raw_receipts", "verified_utc",
    }
)
COLD_VERIFICATION_KEYS = frozenset(
    {
        "schema", "version", "status", "attempt", "pilot",
        "summary_sha256", "raw_receipts", "verified_utc",
    }
)
EXCLUDED = {
    "self": SEAL_NAME,
    "transient": ["__pycache__/**"],
    "bytecode": "reject",
}


class SealError(RuntimeError):
    """A malformed, changed, or incomplete evidence binding."""


def _fail(message: str) -> None:
    raise SealError(message)


def _require(condition: bool, message: str) -> None:
    if not condition:
        _fail(message)


def _now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def _sha(path: Path) -> str:
    try:
        with path.open("rb") as stream:
            return hashlib.file_digest(stream, "sha256").hexdigest()
    except OSError as error:
        _fail(f"cannot hash {path}: {error}")
    raise AssertionError("unreachable")


def _meta(path: Path, label: str) -> dict[str, int | str]:
    _require(not path.is_symlink(), f"{label}: symlink is not allowed: {path}")
    _require(path.is_file(), f"{label}: regular file is missing: {path}")
    try:
        size = path.stat().st_size
    except OSError as error:
        _fail(f"{label}: cannot stat {path}: {error}")
    return {"bytes": size, "sha256": _sha(path)}


def _read_json(path: Path, label: str) -> Any:
    _meta(path, label)
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as error:
        _fail(f"{label}: invalid JSON: {error}")
    _finite(value, label)
    return value
    raise AssertionError("unreachable")


def _rooted(value: Path) -> Path:
    _require(value.is_dir() and not value.is_symlink(), f"root is not a directory: {value}")
    try:
        root = value.resolve(strict=True)
    except OSError as error:
        _fail(f"cannot resolve root {value}: {error}")
    _require(root.is_dir() and not root.is_symlink(), f"resolved root is not a directory: {root}")
    return root


def _inside(root: Path, value: str | Path, label: str) -> tuple[Path, str]:
    candidate = Path(value)
    if not candidate.is_absolute():
        candidate = root / candidate
    _require(not candidate.is_symlink(), f"{label}: symlink is not allowed: {candidate}")
    try:
        resolved = candidate.resolve(strict=True)
        relative = resolved.relative_to(root).as_posix()
    except (OSError, ValueError) as error:
        _fail(f"{label}: path is missing or outside root: {candidate} ({error})")
    current = root
    for component in Path(relative).parts:
        current /= component
        _require(not current.is_symlink(), f"{label}: symlink path component: {current}")
    _meta(resolved, label)
    return resolved, relative


def _external_file(value: str | Path, label: str, *, base: Path | None = None) -> Path:
    candidate = Path(value)
    if not candidate.is_absolute() and base is not None:
        candidate = base / candidate
    _require(not candidate.is_symlink(), f"{label}: symlink is not allowed: {candidate}")
    try:
        resolved = candidate.resolve(strict=True)
    except OSError as error:
        _fail(f"{label}: retained file is missing: {candidate} ({error})")
    _meta(resolved, label)
    return resolved


def _descriptor(root: Path, path: Path, label: str, *, allow_external: bool = False) -> dict[str, int | str]:
    _require(not path.is_symlink(), f"{label}: symlink is not allowed: {path}")
    try:
        resolved = path.resolve(strict=True)
    except OSError as error:
        _fail(f"{label}: missing path: {path} ({error})")
    try:
        name = resolved.relative_to(root).as_posix()
    except ValueError:
        _require(allow_external, f"{label}: path is outside root: {path}")
        name = str(resolved)
    return {"path": name, **_meta(resolved, label)}


def _descriptor_path(root: Path, value: Any, label: str, *, allow_external: bool = False) -> Path:
    _require(isinstance(value, Mapping), f"{label}: descriptor must be an object")
    path_value = value.get("path")
    _require(isinstance(path_value, str) and path_value, f"{label}.path: missing path")
    candidate = Path(path_value)
    if candidate.is_absolute():
        path = _external_file(candidate, label)
        try:
            path.relative_to(root)
        except ValueError:
            _require(allow_external, f"{label}: external path is not allowed: {path}")
    else:
        path, _ = _inside(root, candidate, label)
    actual = _meta(path, label)
    _require(value.get("bytes") == actual["bytes"], f"{label}: byte length changed")
    _require(value.get("sha256") == actual["sha256"], f"{label}: SHA-256 changed")
    return path


def _finite(value: Any, label: str = "json") -> None:
    if isinstance(value, float):
        _require(value == value and abs(value) != float("inf"), f"{label}: non-finite number")
    elif isinstance(value, Mapping):
        for key, child in value.items():
            _require(isinstance(key, str), f"{label}: JSON key is not a string")
            _finite(child, f"{label}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            _finite(child, f"{label}[{index}]")


def _timestamp(value: Any, label: str) -> _datetime.datetime:
    _require(isinstance(value, str) and value, f"{label}: timestamp is missing")
    try:
        parsed = _datetime.datetime.fromisoformat(value)
    except ValueError as error:
        _fail(f"{label}: invalid timestamp: {error}")
    _require(parsed.tzinfo is not None, f"{label}: timestamp lacks timezone")
    return parsed


def _source_binding(root: Path, value: Any, label: str) -> dict[str, Any]:
    _require(isinstance(value, Mapping), f"{label}: source binding is missing")
    _require(set(value) == {"files", "path", "sha256"}, f"{label}: source fields differ")
    _require(type(value["files"]) is int and value["files"] > 0,
             f"{label}.files: expected a positive integer")
    digest = value["sha256"]
    _require(isinstance(digest, str) and len(digest) == 64 and
             all(character in "0123456789abcdef" for character in digest),
             f"{label}.sha256: malformed")
    manifest_path, _ = _inside(root, value["path"], f"{label}.path")
    _require(_sha(manifest_path) == digest,
             f"{label}: source manifest bytes differ from their binding")
    manifest = _read_json(manifest_path, label)
    _finite(manifest, label)
    _require(isinstance(manifest, Mapping) and len(manifest) == value["files"],
             f"{label}: source manifest file count differs")
    for name, item in manifest.items():
        candidate = Path(name)
        _require(isinstance(name, str) and name and not candidate.is_absolute()
                 and ".." not in candidate.parts, f"{label}: unsafe source path")
        _require(isinstance(item, str) and len(item) == 64 and
                 all(character in "0123456789abcdef" for character in item),
                 f"{label}.{name}: malformed source digest")
    encoded = (json.dumps(manifest, indent=2, sort_keys=True) + "\n").encode("utf-8")
    _require(hashlib.sha256(encoded).hexdigest() == digest,
             f"{label}: source manifest digest differs")
    return {"files": value["files"], "path": str(manifest_path), "sha256": digest}


def _write_exclusive(path: Path, value: Mapping[str, Any]) -> None:
    _require(not path.exists(), f"refusing to replace existing seal: {path}")
    try:
        with path.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(value, stream, indent=2, sort_keys=True, allow_nan=False)
            stream.write("\n")
    except OSError as error:
        _fail(f"cannot write {path}: {error}")


def _inventory(root: Path) -> dict[str, dict[str, int | str]]:
    """Return the regular-file inventory, rejecting links and bytecode.

    A transient ``__pycache__`` tree is ignored as a whole.  Bytecode outside
    that explicitly excluded tree remains a hard error.
    """

    result: dict[str, dict[str, int | str]] = {}
    try:
        paths = sorted(root.rglob("*"), key=lambda item: item.relative_to(root).as_posix())
    except OSError as error:
        _fail(f"cannot enumerate evidence: {error}")
    for path in paths:
        relative = path.relative_to(root).as_posix()
        if path.is_symlink():
            _fail(f"evidence contains a symlink: {relative}")
        if "__pycache__" in path.parts:
            if path.is_dir() or path.is_file():
                continue
            _fail(f"evidence contains a non-regular transient entry: {relative}")
        if path.is_dir():
            continue
        if not path.is_file():
            _fail(f"evidence contains a non-regular entry: {relative}")
        if path.suffix.lower() in BYTECODE_SUFFIXES:
            _fail(f"bytecode is not allowed in evidence: {relative}")
        if relative == SEAL_NAME:
            continue
        result[relative] = _meta(path, f"evidence file {relative}")
    return result


def _gate(root: Path, value: str | Path) -> dict[str, Any]:
    path, relative = _inside(root, value, "gate receipt")
    _require(
        path.suffix == ".json" and Path(relative).parent == Path("validation")
        and relative == f"validation/{path.stem}.json",
        f"gate receipt {relative}: path is not canonical",
    )
    receipt = _read_json(path, f"gate receipt {relative}")
    _require(isinstance(receipt, Mapping), f"gate receipt {relative}: expected an object")
    required = {
        "argv", "artifacts", "attempt", "common_sha256", "cwd", "driver_sha256",
        "environment", "exit_code", "finished_utc", "label", "schema",
        "source_after", "source_before", "source_unchanged", "started_utc",
    }
    _require(set(receipt) == required, f"gate receipt {relative}: fields differ")
    _require(receipt["schema"] == GATE_SCHEMA and receipt["label"] == path.stem,
             f"gate receipt {relative}: schema or label differs")
    _require(type(receipt["attempt"]) is str or receipt["attempt"] is None,
             f"gate receipt {relative}: attempt is malformed")
    if receipt["attempt"] is not None:
        _require(bool(receipt["attempt"]) and "/" not in receipt["attempt"] and
                 "\\" not in receipt["attempt"] and
                 all(not character.isspace() for character in receipt["attempt"]),
                 f"gate receipt {relative}: attempt is not path-safe")
    _require(type(receipt["exit_code"]) is int and receipt["exit_code"] == 0,
             f"gate receipt {relative}: exit_code is not zero")
    _require(receipt["source_unchanged"] is True,
             f"gate receipt {relative}: source_unchanged is not true")
    _require(isinstance(receipt["argv"], list) and receipt["argv"] and
             all(isinstance(item, str) and item for item in receipt["argv"]),
             f"gate receipt {relative}: argv is malformed")
    _require(receipt["cwd"] == str(REPO), f"gate receipt {relative}: cwd differs")
    _require(receipt["environment"] == {key: ENV[key] for key in ENV_KEYS},
             f"gate receipt {relative}: environment differs")
    _require(receipt["driver_sha256"] == _sha(root / "gate.py"),
             f"gate receipt {relative}: gate driver changed")
    _require(receipt["common_sha256"] == _sha(root / "support.py"),
             f"gate receipt {relative}: support helper changed")
    started_at = _timestamp(receipt["started_utc"], f"gate receipt {relative}.started_utc")
    finished_at = _timestamp(receipt["finished_utc"], f"gate receipt {relative}.finished_utc")
    _require(finished_at > started_at, f"gate receipt {relative}: chronology is invalid")
    source_before = _source_binding(root, receipt["source_before"], f"{relative}.source_before")
    source_after = _source_binding(root, receipt["source_after"], f"{relative}.source_after")
    _require(source_before == source_after, f"gate receipt {relative}: source snapshots differ")
    artifact_names = {f"{path.stem}.stdout", f"{path.stem}.stderr"}
    _require(set(receipt["artifacts"]) == artifact_names,
             f"gate receipt {relative}: artifact inventory differs")
    artifacts: dict[str, dict[str, Any]] = {}
    for name in sorted(artifact_names):
        value = receipt["artifacts"][name]
        _require(isinstance(value, Mapping) and set(value) == {"bytes", "sha256"},
                 f"gate receipt {relative}: malformed artifact {name}")
        artifact_path = path.with_suffix("." + name.rsplit(".", 1)[1])
        actual = _meta(artifact_path, f"gate artifact {relative}/{name}")
        _require(dict(value) == actual, f"gate receipt {relative}: artifact {name} changed")
        artifacts[name] = dict(value)
    return {
        "path": relative,
        **_meta(path, f"gate receipt {relative}"),
        "schema": GATE_SCHEMA,
        "label": receipt["label"],
        "attempt": receipt["attempt"],
        "exit_code": 0,
        "source_unchanged": True,
        "source": source_after,
        "argv": list(receipt["argv"]),
        "artifacts": artifacts,
    }


def _gates(root: Path, values: list[str]) -> list[dict[str, Any]]:
    _require(values, "seal requires at least one --gate receipt")
    result = [_gate(root, value) for value in values]
    paths = [item["path"] for item in result]
    _require(len(paths) == len(set(paths)), "selected gates contain duplicate paths")
    labels = [item["label"] for item in result]
    _require(len(labels) == len(set(labels)), "selected gates contain duplicate labels")
    _require(REQUIRED_FINAL_GATE_LABELS <= set(labels),
             f"selected gates omit required final labels: {sorted(REQUIRED_FINAL_GATE_LABELS - set(labels))}")
    return result


def _verify_gates(root: Path, values: Any) -> list[dict[str, Any]]:
    _require(isinstance(values, list) and values, "seal has no selected final gates")
    actual: list[dict[str, Any]] = []
    for index, item in enumerate(values):
        _require(isinstance(item, Mapping), f"seal.gates[{index}]: expected an object")
        path = item.get("path")
        _require(isinstance(path, str) and path, f"seal.gates[{index}].path: missing path")
        current = _gate(root, path)
        _require(dict(item) == current, f"seal.gates[{index}]: binding changed")
        actual.append(current)
    paths = [item["path"] for item in actual]
    labels = [item["label"] for item in actual]
    _require(len(paths) == len(set(paths)) and len(labels) == len(set(labels)),
             "seal selects duplicate gates")
    _require(REQUIRED_FINAL_GATE_LABELS <= set(labels),
             f"sealed gates omit required final labels: {sorted(REQUIRED_FINAL_GATE_LABELS - set(labels))}")
    return actual


def _source(value: Any, label: str) -> dict[str, Any]:
    _require(isinstance(value, Mapping), f"{label}: source identity is missing")
    _require(isinstance(value.get("path"), str) and value["path"], f"{label}.path: missing")
    _require(isinstance(value.get("sha256"), str) and len(value["sha256"]) == 64,
             f"{label}.sha256: missing or malformed")
    _require(all(character in "0123456789abcdef" for character in value["sha256"]),
             f"{label}.sha256: expected lowercase hexadecimal")
    if "files" in value:
        _require(type(value["files"]) is int and value["files"] >= 0,
                 f"{label}.files: expected a non-negative integer")
    if "bytes" in value:
        _require(type(value["bytes"]) is int and value["bytes"] >= 0,
                 f"{label}.bytes: expected a non-negative integer")
    return dict(value)


def _binary(root: Path, build_path: Path, value: Any, label: str) -> dict[str, Any]:
    _require(isinstance(value, Mapping), f"{label}: retained binary binding is missing")
    path_value = value.get("path")
    _require(isinstance(path_value, str) and path_value, f"{label}.path: missing")
    candidate = Path(path_value)
    if candidate.is_absolute():
        path = _external_file(candidate, label)
    else:
        # Build records normally store absolute retained paths.  Accept a
        # relative path rooted either at the evidence tree or beside the record.
        root_candidate = root / candidate
        record_candidate = build_path.parent / candidate
        path = _external_file(root_candidate if root_candidate.exists() else record_candidate, label)
    actual = _meta(path, label)
    _require(type(value.get("bytes")) is int and value["bytes"] == actual["bytes"],
             f"{label}: byte length changed")
    _require(value.get("sha256") == actual["sha256"], f"{label}: SHA-256 changed")
    executable = value.get("executable", True)
    _require(executable is True, f"{label}: receipt does not mark the binary executable")
    _require(bool(path.stat().st_mode & (stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)),
             f"{label}: retained file is not executable")
    return {"path": str(path), **actual, "executable": True}


def _strict_builds(root: Path) -> dict[str, dict[str, Any]]:
    """Load both canonical receipts through the frozen provider validator."""

    try:
        import provider_matrix as driver
    except ImportError as error:
        _fail(f"strict retained build validator is unavailable: {error}")
    try:
        builds = driver.load_builds(root)
    except (OSError, driver.ProviderMatrixError, KeyError, TypeError, ValueError) as error:
        _fail(f"strict retained build validation failed: {error}")
    _require(isinstance(builds, Mapping) and set(builds) == {"normal", "allocator"},
             "strict retained build inventory differs")
    for role in ("normal", "allocator"):
        expected = (root / f"build-{role}.json").resolve()
        actual = Path(str(builds[role].get("path", ""))).resolve()
        _require(actual == expected, f"{role} build receipt is not canonical")
    return {role: dict(builds[role]) for role in ("normal", "allocator")}


def _build(root: Path, value: str | Path, role: str,
           builds: Mapping[str, Mapping[str, Any]] | None = None) -> dict[str, Any]:
    expected = f"build-{role}.json"
    path, relative = _inside(root, value, f"{role} build receipt")
    _require(relative == expected, f"{role} build receipt must be canonical: {expected}")
    selected = _strict_builds(root) if builds is None else builds
    _require(role in selected, f"strict {role} build receipt is missing")
    record = dict(selected[role])
    _require(record.get("path") == str(path.resolve()),
             f"{role} strict build path differs")
    return record


def _bind_build_gates(root: Path, builds: Mapping[str, Mapping[str, Any]],
                      gates: list[Mapping[str, Any]]) -> None:
    by_label = {str(item["label"]): item for item in gates}
    for role in ("normal", "allocator"):
        build = builds[role]
        expected_label = f"build-{role}-final5"
        gate = build.get("gate")
        _require(isinstance(gate, Mapping), f"{role} build gate binding is missing")
        gate_path = Path(str(gate.get("path", ""))).resolve()
        expected_path = (root / "validation" / f"{expected_label}.json").resolve()
        _require(gate_path == expected_path, f"{role} build gate path differs")
        selected = by_label.get(expected_label)
        _require(selected is not None and selected["path"] == str(expected_path.relative_to(root)),
                 f"{role} build gate is not selected as a final gate")
        _require(gate.get("sha256") == selected["sha256"],
                 f"{role} build gate hash differs from selected gate")
        _require(gate.get("source") == build.get("source"),
                 f"{role} build gate source differs from build source")


def _verification(root: Path, value: str | Path, kind: str,
                   provider_builds: Mapping[str, Mapping[str, Any]]) -> dict[str, Any]:
    path, relative = _inside(root, value, f"{kind} verification")
    record = _read_json(path, f"{kind} verification {relative}")
    _require(isinstance(record, Mapping), f"{kind} verification {relative}: expected an object")
    expected = PROVIDER_VERIFICATION_SCHEMA if kind == "provider" else COLD_VERIFICATION_SCHEMA
    expected_keys = PROVIDER_VERIFICATION_KEYS if kind == "provider" else COLD_VERIFICATION_KEYS
    _require(set(record) == expected_keys,
             f"{kind} verification {relative}: fields differ")
    _require(record.get("schema") == expected,
             f"{kind} verification {relative}: schema differs")
    _require(record.get("version") == 1,
             f"{kind} verification {relative}: version differs")
    _require(record.get("status") == "pass",
             f"{kind} verification {relative}: terminal status is not pass")
    _require(record.get("pilot") is False, f"{kind} verification must describe formal captures")
    attempt = record.get("attempt")
    _require(isinstance(attempt, str) and attempt and
             all(character.isalnum() or character in "_.-" for character in attempt),
             f"{kind} verification {relative}: attempt is not path-safe")
    _require(relative == f"verification/{attempt}.json",
             f"{kind} verification {relative}: canonical path differs")
    _require(isinstance(record.get("summary_sha256"), str) and
             len(record["summary_sha256"]) == 64 and
             all(character in "0123456789abcdef" for character in record["summary_sha256"]),
             f"{kind} verification {relative}: summary hash is malformed")
    _timestamp(record.get("verified_utc"), f"{kind} verification {relative}.verified_utc")
    # A green receipt alone is insufficient: recompute through the frozen
    # validators, including executable/source custody and every raw report.
    if kind == "provider":
        import provider_matrix as driver
        builds = {role: dict(provider_builds[role]) for role in ("normal", "allocator")}
        try:
            protocol, protocol_hash = driver._load_protocol()
            entries = driver._collect(attempt, builds, protocol, protocol_hash)
        except (driver.ProviderMatrixError, OSError, KeyError, TypeError, ValueError) as error:
            _fail(f"provider verification recomputation failed: {error}")
        _require(record.get("protocol_sha256") == protocol_hash, "provider protocol binding differs")
    else:
        import cold_matrix as driver
        try:
            builds = driver.load_builds(root)
            _require(set(builds) == {"normal", "allocator"},
                     "cold strict build inventory differs")
            for role in ("normal", "allocator"):
                _require(
                    Path(str(builds[role]["path"])).resolve()
                    == Path(str(provider_builds[role]["path"])).resolve()
                    and builds[role]["receipt_sha256"] == provider_builds[role]["receipt_sha256"]
                    and builds[role]["binary"] == provider_builds[role]["binary"]
                    and builds[role]["source"] == provider_builds[role]["source"]
                    and builds[role]["git_revision"] == provider_builds[role]["git_revision"],
                    f"cold {role} build differs from strict provider build",
                )
            entries = driver._collect(attempt, builds)
        except (driver.ColdMatrixError, OSError, KeyError, TypeError, ValueError) as error:
            _fail(f"cold verification recomputation failed: {error}")
    summary = root / "analysis" / f"{attempt}.json"
    _meta(summary, f"{kind} analysis summary")
    _require(_sha(summary) == record.get("summary_sha256"), f"{kind} summary binding differs")
    try:
        expected_summary = driver.analyze_data(entries, builds)
    except (OSError, KeyError, TypeError, ValueError, RuntimeError) as error:
        _fail(f"{kind} summary recomputation failed: {error}")
    _require(_read_json(summary, kind) == expected_summary, f"{kind} summary recomputation differs")
    expected_receipts = [
        {"label": entry["spec"]["label"], "terminal_sha256": entry["terminal_sha256"], "report_sha256": entry["report_sha256"]}
        for entry in entries
    ]
    _require(record.get("raw_receipts") == expected_receipts, f"{kind} raw receipt inventory differs")
    result: dict[str, Any] = {
        "path": relative,
        **_meta(path, f"{kind} verification {relative}"),
    }
    result.update(dict(record))
    return result


def _cleanup(root: Path, value: str | Path) -> dict[str, Any]:
    path, relative = _inside(root, value, "cleanup receipt")
    _require(relative == "cleanup.json", "cleanup receipt must be canonical: cleanup.json")
    try:
        import cleanup as driver
        checked = driver.verify(root=root)
    except (ImportError, OSError, KeyError, TypeError, ValueError) as error:
        _fail(f"cleanup verification helper failed: {error}")
    except RuntimeError as error:
        _fail(f"cleanup verification helper failed: {error}")
    record = _read_json(path, f"cleanup receipt {relative}")
    _require(isinstance(record, Mapping), f"cleanup receipt {relative}: expected an object")
    _require(record.get("schema") == CLEANUP_SCHEMA and record.get("version") == 1 and
             record.get("status") == "pass", f"cleanup receipt {relative}: status/schema differs")
    _require(isinstance(checked, Mapping) and
             checked.get("schema") == CLEANUP_VERIFICATION_SCHEMA and
             checked.get("version") == 1 and checked.get("status") == "pass" and
             checked.get("remaining") == [] and checked.get("build_target_removed") is True,
             f"cleanup verification helper returned an incomplete result")
    expected_descriptor = {"path": str(path), "relative_path": relative,
                           **_meta(path, f"cleanup receipt {relative}")}
    _require(checked.get("cleanup_receipt") == expected_descriptor,
             f"cleanup verification helper receipt binding differs")
    result: dict[str, Any] = {
        "path": relative,
        **_meta(path, f"cleanup receipt {relative}"),
        "schema": CLEANUP_SCHEMA,
        "version": 1,
        "status": "pass",
        "remaining": [],
        "verification_schema": CLEANUP_VERIFICATION_SCHEMA,
    }
    return result


def _profile(root: Path, provider_builds: Mapping[str, Mapping[str, Any]],
             gates: list[Mapping[str, Any]]) -> dict[str, Any]:
    """Verify the canonical post-capture profile and bind its outer gate."""

    path, relative = _inside(root, PROFILE_RESULT_PATH, "profile result")
    _require(relative == PROFILE_RESULT_PATH, "profile result path is not canonical")
    try:
        import profile_providers as driver
        driver.verify("profiles2")
    except (ImportError, OSError, KeyError, TypeError, ValueError) as error:
        _fail(f"profile verification failed: {error}")
    except RuntimeError as error:
        _fail(f"profile verification failed: {error}")
    result = _read_json(path, "profile result")
    _require(isinstance(result, Mapping) and result.get("status") == "pass" and
             result.get("attempt") == "profiles2", "profile result is not final and passing")
    _require(result.get("build", {}).get("source") == provider_builds["normal"].get("source"),
             "profile source differs from retained build")
    profile_gate = _gate(root, "validation/profiles2.json")
    selected = {item["label"]: item for item in gates}
    _require(selected.get("profiles2") == profile_gate,
             "profile gate is not selected as a final gate")
    return {
        "path": relative,
        **_meta(path, "profile result"),
        "schema": result.get("schema"),
        "version": result.get("version"),
        "status": result.get("status"),
        "attempt": result.get("attempt"),
        "gate": profile_gate,
    }


def _manifest(root: Path, gates: list[str], build_normal: str, build_allocator: str,
              provider: str, cold: str, cleanup: str) -> dict[str, Any]:
    gate_records = _gates(root, gates)
    strict_builds = _strict_builds(root)
    normal = _build(root, build_normal, "normal", strict_builds)
    allocator = _build(root, build_allocator, "allocator", strict_builds)
    _require(normal["source"] == allocator["source"],
             "normal and allocator build source identities differ")
    _bind_build_gates(root, strict_builds, gate_records)
    provider_record = _verification(root, provider, "provider", strict_builds)
    cold_record = _verification(root, cold, "cold", strict_builds)
    profile_record = _profile(root, strict_builds, gate_records)
    cleanup_record = _cleanup(root, cleanup)
    return {
        "schema": SCHEMA,
        "version": VERSION,
        "root": ".",
        "sealed_utc": _now(),
        "excluded": EXCLUDED,
        "gates": gate_records,
        "builds": {"normal": normal, "allocator": allocator},
        "provider_verification": provider_record,
        "cold_verification": cold_record,
        "profile": profile_record,
        "cleanup": cleanup_record,
        "files": _inventory(root),
    }


def _verify_manifest(root: Path, manifest: Any) -> dict[str, Any]:
    _require(isinstance(manifest, Mapping), "seal manifest: expected an object")
    _require(
        set(manifest) == {
            "schema", "version", "root", "sealed_utc", "excluded", "gates", "builds",
            "provider_verification", "cold_verification", "profile", "cleanup", "files",
        },
        "seal manifest fields differ",
    )
    _require(manifest.get("schema") == SCHEMA and manifest.get("version") == VERSION,
             "seal schema or version differs")
    _require(manifest.get("root") == ".", "seal root binding differs")
    _require(isinstance(manifest.get("sealed_utc"), str) and manifest["sealed_utc"],
             "seal timestamp is missing")
    _require(manifest.get("excluded") == EXCLUDED, "seal exclusion policy differs")
    _require(manifest.get("files") == _inventory(root),
             "sealed regular-file inventory or hash differs")
    gates = _verify_gates(root, manifest.get("gates"))
    builds = manifest.get("builds")
    _require(isinstance(builds, Mapping) and set(builds) == {"normal", "allocator"},
             "seal.builds: missing or incomplete")
    normal_path = builds.get("normal", {}).get("path") if isinstance(builds.get("normal"), Mapping) else None
    allocator_path = builds.get("allocator", {}).get("path") if isinstance(builds.get("allocator"), Mapping) else None
    _require(isinstance(normal_path, str) and isinstance(allocator_path, str),
             "seal.builds: receipt paths are missing")
    strict_builds = _strict_builds(root)
    normal = _build(root, normal_path, "normal", strict_builds)
    allocator = _build(root, allocator_path, "allocator", strict_builds)
    _require(dict(builds["normal"]) == normal and dict(builds["allocator"]) == allocator,
             "sealed build binding changed")
    _require(normal["source"] == allocator["source"],
             "normal and allocator build source identities differ")
    _bind_build_gates(root, strict_builds, gates)
    provider_path = manifest.get("provider_verification", {}).get("path") if isinstance(manifest.get("provider_verification"), Mapping) else None
    cold_path = manifest.get("cold_verification", {}).get("path") if isinstance(manifest.get("cold_verification"), Mapping) else None
    cleanup_path = manifest.get("cleanup", {}).get("path") if isinstance(manifest.get("cleanup"), Mapping) else None
    _require(isinstance(provider_path, str) and isinstance(cold_path, str) and isinstance(cleanup_path, str),
             "seal verification or cleanup paths are missing")
    provider = _verification(root, provider_path, "provider", strict_builds)
    cold = _verification(root, cold_path, "cold", strict_builds)
    profile_path = manifest.get("profile", {}).get("path") if isinstance(manifest.get("profile"), Mapping) else None
    _require(profile_path == PROFILE_RESULT_PATH,
             "seal profile result path is missing or noncanonical")
    profile = _profile(root, strict_builds, gates)
    cleanup = _cleanup(root, cleanup_path)
    _require(dict(manifest["provider_verification"]) == provider,
             "provider verification binding changed")
    _require(dict(manifest["cold_verification"]) == cold,
             "cold verification binding changed")
    _require(dict(manifest["profile"]) == profile,
             "profile binding changed")
    _require(dict(manifest["cleanup"]) == cleanup, "cleanup binding changed")
    return {"files": manifest["files"], "gates": gates, "builds": builds,
            "provider": provider, "cold": cold, "profile": profile, "cleanup": cleanup}


def _parse_args(argv: list[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("command", nargs="?", choices=("seal", "verify"), default="verify")
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--gate", action="append", default=[],
                        help="final gate receipt path; repeat for each selected gate")
    parser.add_argument("--build-normal", help="final normal build receipt path (seal only)")
    parser.add_argument("--build-allocator", help="final allocator build receipt path (seal only)")
    parser.add_argument("--provider", "--provider-verification", dest="provider",
                        help="final provider verification JSON path (seal only)")
    parser.add_argument("--cold", "--cold-verification", dest="cold",
                        help="final cold verification JSON path (seal only)")
    parser.add_argument("--cleanup", help="cleanup.json path (seal only)")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)
    try:
        root = _rooted(args.root)
        seal_path = root / SEAL_NAME
        selection = (args.gate, args.build_normal, args.build_allocator,
                     args.provider, args.cold, args.cleanup)
        if args.command == "seal":
            _require(not seal_path.exists(), f"refusing to replace existing seal: {seal_path}")
            _require(all(isinstance(item, str) and item for item in selection[1:]),
                     "seal requires --build-normal, --build-allocator, --provider, --cold, and --cleanup")
            value = _manifest(root, args.gate, args.build_normal, args.build_allocator,
                              args.provider, args.cold, args.cleanup)
            _write_exclusive(seal_path, value)
            print(json.dumps({"schema": SCHEMA, "status": "sealed",
                              "files": len(value["files"]),
                              "final_gates": len(value["gates"]),
                              "manifest": _descriptor(root, seal_path, "seal manifest")},
                             sort_keys=True))
            return 0
        _require(not any(selection), "selection options are only valid with seal")
        manifest = _read_json(seal_path, "seal manifest")
        checked = _verify_manifest(root, manifest)
        result = {
            "schema": VERIFICATION_SCHEMA,
            "status": "pass",
            "verified_utc": _now(),
            "manifest": _descriptor(root, seal_path, "seal manifest"),
            "files": len(checked["files"]),
            "final_gates": len(checked["gates"]),
            "builds": {"normal": checked["builds"]["normal"]["path"],
                       "allocator": checked["builds"]["allocator"]["path"]},
            "provider_verification": checked["provider"]["path"],
            "cold_verification": checked["cold"]["path"],
            "profile": checked["profile"]["path"],
            "cleanup": checked["cleanup"]["path"],
        }
        print(json.dumps(result, sort_keys=True))
        return 0
    except (SealError, OSError, TypeError, ValueError, KeyError) as error:
        print(f"seal.py: error: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
