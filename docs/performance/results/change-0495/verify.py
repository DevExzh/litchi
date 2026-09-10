#!/usr/bin/env python3
"""Authenticate the 0495 managed DOCX evidence bundle.

The measurement driver authenticates one report at a time.  This module is the
bundle boundary: it binds the four retained builds, the frozen before/after
protocol, source and fixture custody, raw formal and pilot captures, failed
development receipts, final validation commands, helper archives, and the
post-capture cleanup receipt.  It is deliberately read-only in ``verify``
mode; ``create`` only writes the immutable bundle seal after verification has
passed.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import re
import stat
import subprocess
import sys
from typing import Any, Iterable, Mapping

import cleanup as cleanup_driver
import measure
from support import ROOT, TEMP, TARGET_DIR, meta, read, sha, write


BUNDLE_SCHEMA = "docx-edit-provider-bundle-v1"
BUNDLE_VERSION = 1
ACCEPTED_SCHEMA = "docx-edit-provider-accepted-v1"
FINAL_GATES_SCHEMA = "docx-edit-provider-final-gates-v1"
FINAL_GATES_VERSION = 1
SEAL_NAME = "seal.json"
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")

# ``support.REPO`` deliberately follows LITCHI_MANAGED_EDIT_SOURCE_ROOT so
# build and gate receipts can bind the disposable sparse checkout.  Bundle
# provenance files live beside this verifier in the real workspace, however:
# the sparse checkout contains only the Rust build inputs and does not carry
# test-data, ADRs, or the evidence directory itself.
WORKSPACE = ROOT.parents[3].resolve()

# These are the source files owned by the 0495 harness.  The frozen common
# harness manifest is checked by measure._shared_harness_binding(); spelling
# the tuple here makes the rustfmt gate part of the final command contract.
FORMAT_FILES = (
    "tools/perf-baseline/src/lib.rs",
    "tools/perf-baseline/src/main.rs",
    "tools/perf-baseline/src/bin/litchi-perf-baseline-alloc.rs",
    "tools/perf-baseline/src/docx_managed_edit.rs",
)

# The production patch is spread across these owned DOCX modules and focused
# tests.  Keep formatting scoped to these manifests because unrelated
# workspace crates may carry pre-existing rustfmt differences during the frozen
# build.
DOCX_FORMAT_FILES = (
    "crates/litchi-docx/src/document/mod.rs",
    "crates/litchi-docx/src/document/transaction.rs",
    "crates/litchi-docx/src/document/transaction/durable.rs",
    "crates/litchi-docx/src/paragraph/codec/content.rs",
    "crates/litchi-docx/src/paragraph/codec/editing.rs",
    "crates/litchi-docx/src/paragraph/codec/inlines.rs",
    "crates/litchi-docx/src/paragraph/codec/paragraph_properties.rs",
    "crates/litchi-docx/src/paragraph/codec/run.rs",
    "crates/litchi-docx/src/paragraph/codec/run_contents.rs",
    "crates/litchi-docx/src/paragraph/codec/run_properties.rs",
    "crates/litchi-docx/src/paragraph/codec/runs.rs",
    "crates/litchi-docx/src/paragraph/codec/text.rs",
    "crates/litchi-docx/src/paragraph/collapsed/codec.rs",
    "crates/litchi-docx/src/paragraph/collapsed/package.rs",
    "crates/litchi-docx/src/paragraph/extensions/codec.rs",
    "crates/litchi-docx/src/paragraph/extensions/package.rs",
    "crates/litchi-docx/src/paragraph/model.rs",
    "crates/litchi-docx/src/paragraph/package.rs",
    "crates/litchi-docx/src/run_symbols/codec.rs",
    "crates/litchi-docx/src/run_symbols/package.rs",
    "crates/litchi-docx/src/source_backed.rs",
    "crates/litchi-docx/src/source_backed/document_policy.rs",
    "crates/litchi-docx/src/source_backed/story_text.rs",
    "crates/litchi-docx/tests/source_backed_managed.rs",
    "crates/litchi-docx/tests/source_backed_managed_document_edit.rs",
    "crates/litchi-docx/tests/source_backed_semantic.rs",
)
OPC_FORMAT_FILES = (
    "crates/litchi-opc/src/xml_splice.rs",
)

# The OPC all-targets production check compiles these literal include_bytes! /
# include_str! inputs.  Keep this set explicit: a receipt which names only the
# ordinary test-data tree could omit the signed local-only corpus input while
# still looking superficially complete.
OPC_COMPILE_FIXTURE_NAMES = frozenset({
    "test-data/poi/test-data/xmldsign/hello-world-signed.xlsx",
    "test-data/poi/test-data/xmldsign/hello-world-signed.pptx",
    "test-data/poi/test-data/xmldsign/ms-office-2010-signed.xlsx",
    "test-data/poi/test-data/xmldsign/ms-office-2010-signed.pptx",
    "docs/performance/results/change-0416/corpus/opc-local-only-signed.zip",
    "test-data/poi/test-data/slideshow/EmbeddedVideo.pptx",
    "test-data/poi/test-data/slideshow/bug62513.pptx",
})
OPC_RUNTIME_FIXTURE_NAMES = frozenset({
    "test-data/poi/test-data/openxml4j/PackageRelsHasEntities.ooxml",
    "test-data/poi/test-data/openxml4j/50154.xlsx",
})
HARNESS_RUNTIME_FIXTURE_NAMES = frozenset({
    "test-data/ooxml/pptx/shapes.pptx",
    "test-data/office-interop/libreoffice-resaved/shapes-litchi.pptx",
})

_ALLOCATION_FIELDS = frozenset({
    "status", "scope", "allocation_calls", "deallocation_calls",
    "reallocation_calls", "failed_allocation_calls", "allocated_bytes",
    "deallocated_bytes", "live_bytes_before", "live_bytes_after",
    "peak_live_bytes_before", "peak_live_bytes_after",
    "region_peak_live_bytes",
})
_ALLOCATION_UINT_FIELDS = _ALLOCATION_FIELDS - {"status", "scope"}
def _require(condition: bool, message: str) -> None:
    if not condition:
        measure.fail(message)


def _safe_token(value: Any, path: str) -> str:
    _require(
        isinstance(value, str) and bool(value)
        and len(value) <= 128 and value not in {".", ".."}
        and all(char.isalnum() or char in "._-" for char in value),
        f"{path}: unsafe token",
    )
    return value


def _safe_label(value: Any, path: str) -> str:
    value = _safe_token(value, path)
    _require("/" not in value and "\\" not in value and not any(c.isspace() for c in value),
             f"{path}: unsafe label")
    return value


def _path(value: Any, path: str, *, base: Path = ROOT) -> Path:
    _require(isinstance(value, str) and value, f"{path}: path is missing")
    result = Path(value)
    if not result.is_absolute():
        _require(".." not in result.parts, f"{path}: parent traversal is forbidden")
        result = base / result
    return result


def _regular(path: Path, label: str, *, allow_missing: bool = False) -> Path:
    _require(not path.is_symlink(), f"{label}: symlink is forbidden")
    if allow_missing and not path.exists():
        return path
    _require(path.is_file(), f"{label}: regular file is required: {path}")
    mode = path.stat().st_mode
    _require(stat.S_ISREG(mode), f"{label}: special file is forbidden: {path}")
    return path


def _directory(path: Path, label: str, *, allow_missing: bool = False) -> Path:
    _require(not path.is_symlink(), f"{label}: symlink is forbidden")
    if allow_missing and not path.exists():
        return path
    _require(path.is_dir(), f"{label}: directory is required: {path}")
    _require(stat.S_ISDIR(path.stat().st_mode), f"{label}: special path is forbidden: {path}")
    return path


def _digest_descriptor(path: Path, label: str, *, root: Path | None = None) -> dict[str, Any]:
    _regular(path, label)
    resolved = path.resolve(strict=True)
    value: dict[str, Any] = {
        "path": str(resolved),
        "bytes": resolved.stat().st_size,
        "sha256": sha(resolved),
    }
    if root is not None:
        try:
            value["relative_path"] = resolved.relative_to(root.resolve()).as_posix()
        except ValueError as error:
            measure.fail(f"{label}: path escapes root: {error}")
    return value


def _json(path: Path, label: str) -> Any:
    _regular(path, label)
    try:
        value = read(path)
    except (OSError, UnicodeError, ValueError, json.JSONDecodeError) as error:
        measure.fail(f"{label}: invalid JSON: {error}")
    measure._finite(value, label)
    return value


def _hash(value: Any, path: str) -> str:
    _require(isinstance(value, str) and SHA256_RE.fullmatch(value) is not None,
             f"{path}: lowercase SHA-256 required")
    return value


def _exact(value: Any, fields: Iterable[str], path: str) -> None:
    measure._exact(value, fields, path)


def _descriptor(value: Any, path: str, *, root: Path | None = None,
                allow_missing: bool = False) -> dict[str, Any]:
    _exact(value, ("bytes", "path", "sha256"), path)
    artifact = _path(value["path"], f"{path}.path")
    _hash(value["sha256"], f"{path}.sha256")
    _require(type(value["bytes"]) is int and value["bytes"] >= 0,
             f"{path}.bytes: non-negative integer required")
    if allow_missing and not artifact.exists():
        return {"path": str(artifact), "bytes": value["bytes"],
                "sha256": value["sha256"]}
    actual = _digest_descriptor(artifact, path, root=root)
    _require(actual["bytes"] == value["bytes"] and actual["sha256"] == value["sha256"],
             f"{path}: artifact changed")
    return {"path": str(artifact.resolve()), "bytes": value["bytes"],
            "sha256": value["sha256"]}


def _expected_final_commands() -> tuple[tuple[str, ...], ...]:
    """Exact final command vectors.

    The two build vectors are intentionally present once each.  Before and
    after build receipts use the same Cargo argv but bind different frozen
    source manifests; their per-receipt gate bindings are checked separately.
    The harness and production DOCX checks are separate command vectors so a
    green harness cannot stand in for validation of the owned crate.
    """

    cargo = ("cargo",)
    harness = (
        cargo + (
            "clippy", "--release", "--locked", "--manifest-path",
            "tools/perf-baseline/Cargo.toml", "--all-targets", "--features",
            "allocator-metrics", "--", "-D", "warnings",
        ),
        ("env", "RUSTDOCFLAGS=-Dwarnings", *cargo, "doc", "--release", "--locked",
         "--manifest-path", "tools/perf-baseline/Cargo.toml", "--lib",
         "--features", "allocator-metrics", "--no-deps"),
        cargo + (
            "test", "--release", "--locked", "--manifest-path",
            "tools/perf-baseline/Cargo.toml", "--lib", "--features",
            "allocator-metrics", "--", "--test-threads=1",
        ),
        ("rustfmt", "+1.98.1", "--edition", "2024", "--check", "--config",
         "skip_children=true", *FORMAT_FILES),
        ("python3", "-B", "tools/check_crate_boundaries.py"),
        ("python3", "-B", "-m", "unittest", "discover", "-s", str(ROOT),
         "-p", "test_*.py"),
        tuple(measure._expected_build_command("normal")),
        tuple(measure._expected_build_command("allocator")),
    )
    production = (
        cargo + ("test", "--release", "--locked", "-p", "litchi-docx", "--lib"),
        cargo + ("test", "--release", "--locked", "-p", "litchi-docx",
                 "--test", "source_backed_managed_document_edit", "--test",
                 "source_backed_managed", "--test", "source_backed", "--test",
                 "source_backed_semantic", "--test", "source_backed_story_text",
                 "--test", "source_backed_secondary_story_text", "--test",
                 "source_backed_paragraph_copy", "--test",
                 "source_backed_paragraph_removal"),
        cargo + ("check", "--release", "--locked", "-p", "litchi-docx",
                 "--all-features", "--all-targets"),
        cargo + ("clippy", "--release", "--locked", "-p", "litchi-docx",
                 "--all-features", "--all-targets", "--", "-D", "warnings"),
        ("env", "RUSTDOCFLAGS=-Dwarnings", *cargo, "doc", "--release", "--locked",
         "-p", "litchi-docx", "--all-features", "--no-deps"),
        ("rustfmt", "+1.98.1", "--edition", "2024", "--check", "--config",
         "skip_children=true", *DOCX_FORMAT_FILES, *OPC_FORMAT_FILES),
        cargo + ("test", "--release", "--locked", "-p", "litchi-opc",
                 "--lib"),
        ("env", "RUSTDOCFLAGS=-Dwarnings", *cargo, "doc", "--release", "--locked",
         "-p", "litchi-opc", "--no-deps"),
        cargo + ("clippy", "--release", "--locked", "-p", "litchi-opc",
                 "--all-targets", "--", "-D", "warnings"),
    )
    return harness + production


def _validate_final_gates(value: Any) -> list[tuple[str, list[str]]]:
    _exact(value, ("schema", "version", "commands"), "final-gates.json")
    _require(value["schema"] == FINAL_GATES_SCHEMA and value["version"] == FINAL_GATES_VERSION,
             "final-gates.json schema differs")
    commands = value["commands"]
    _require(isinstance(commands, dict) and commands, "final-gates.json commands are empty")
    actual: list[tuple[str, ...]] = []
    result: list[tuple[str, list[str]]] = []
    for label, command in commands.items():
        _safe_label(label, "final-gates label")
        _require(isinstance(command, list)
                 and all(isinstance(item, str) and item for item in command),
                 f"{label}: malformed command")
        actual.append(tuple(command))
        result.append((label, list(command)))
    expected = set(_expected_final_commands())
    _require(len(actual) == len(expected) and len(set(actual)) == len(actual)
             and set(actual) == expected,
             "final-gate command set differs")
    return result


def _validate_gate_receipt(label: str, command: list[str], path: Path,
                           source: Mapping[str, Any] | None = None) -> dict[str, Any]:
    """Validate one final gate, optionally binding it to a frozen source."""

    _regular(path, f"{label} gate receipt")
    receipt = _json(path, f"{label} gate receipt")
    _require(isinstance(receipt, dict) and receipt.get("label") == label
             and path.stem == label, f"{label}: receipt label differs")
    binding = measure._gate_binding({"path": str(path), "sha256": sha(path)}, path)
    _require(binding["argv"] == command, f"{label}: command differs")
    if source is not None:
        _require(binding["source"] == dict(source), f"{label}: source differs")
    return binding


def _source_file_manifest(path: Path, label: str) -> dict[str, str]:
    value = _json(path, label)
    _require(isinstance(value, dict) and value, f"{label}: source manifest is empty")
    for name, digest in value.items():
        _require(isinstance(name, str) and name and not Path(name).is_absolute()
                 and ".." not in Path(name).parts, f"{label}: unsafe source name")
        _hash(digest, f"{label}.{name}")
    canonical = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode()
    _require(hashlib.sha256(canonical).hexdigest() == sha(path),
             f"{label}: manifest digest differs")
    return value


def _tracked_at_head(names: Iterable[str], label: str) -> set[str]:
    """Return the workspace HEAD paths and require every named input there."""

    try:
        output = subprocess.check_output(
            ["git", "ls-tree", "-r", "-z", "--name-only", "HEAD", "--"],
            cwd=WORKSPACE,
        )
    except (OSError, subprocess.CalledProcessError) as error:
        measure.fail(f"{label}: cannot inspect workspace HEAD: {error}")
    tracked = {item for item in output.decode("utf-8").split("\0") if item}
    expected = set(names)
    _require(expected <= tracked, f"{label}: input is not tracked at workspace HEAD")
    return tracked


def _validate_baseline_inputs() -> dict[str, Any]:
    fixture_path = ROOT / "baseline-fixture-inputs.json"
    fixtures = _json(fixture_path, str(fixture_path))
    _require(isinstance(fixtures, dict) and fixtures, "baseline fixture inventory is empty")
    checked_fixtures: dict[str, str] = {}
    for name, digest in fixtures.items():
        _require(isinstance(name, str) and name and not Path(name).is_absolute()
                 and ".." not in Path(name).parts, f"fixture path is unsafe: {name}")
        _hash(digest, f"fixture {name}")
        path = WORKSPACE / name
        _regular(path, f"fixture {name}")
        _require(sha(path) == digest, f"fixture {name}: content changed")
        checked_fixtures[name] = digest

    harness_runtime_path = ROOT / "harness-runtime-fixture-inputs.json"
    harness_runtime_fixtures = _json(harness_runtime_path,
                                     str(harness_runtime_path))
    _require(isinstance(harness_runtime_fixtures, dict)
             and set(harness_runtime_fixtures) == HARNESS_RUNTIME_FIXTURE_NAMES,
             "harness runtime fixture inventory differs")
    checked_harness_runtime_fixtures: dict[str, dict[str, Any]] = {}
    harness_runtime_names: list[str] = []
    for name, descriptor in harness_runtime_fixtures.items():
        _require(isinstance(name, str) and name and not Path(name).is_absolute()
                 and ".." not in Path(name).parts
                 and Path(name).suffix.lower() == ".pptx",
                 f"harness runtime fixture path is unsafe: {name}")
        _exact(descriptor, ("bytes", "sha256"),
               f"harness runtime fixture {name}")
        _require(type(descriptor["bytes"]) is int and descriptor["bytes"] >= 0,
                 f"harness runtime fixture {name}.bytes is malformed")
        _hash(descriptor["sha256"], f"harness runtime fixture {name}.sha256")
        harness_runtime_names.append(name)
    _tracked_at_head(harness_runtime_names, "harness runtime fixture inventory")
    for name in harness_runtime_names:
        descriptor = harness_runtime_fixtures[name]
        path = WORKSPACE / name
        actual = _digest_descriptor(path, f"harness runtime fixture {name}")
        _require(actual["bytes"] == descriptor["bytes"]
                 and actual["sha256"] == descriptor["sha256"],
                 f"harness runtime fixture {name}: content changed")
        checked_harness_runtime_fixtures[name] = dict(descriptor)

    compile_fixture_path = ROOT / "docx-compile-fixture-inputs.json"
    compile_fixtures = _json(compile_fixture_path, str(compile_fixture_path))
    _require(isinstance(compile_fixtures, dict) and compile_fixtures,
             "DOCX compile fixture inventory is empty")
    checked_compile_fixtures: dict[str, dict[str, Any]] = {}
    for name, descriptor in compile_fixtures.items():
        _require(isinstance(name, str) and name and not Path(name).is_absolute()
                 and ".." not in Path(name).parts,
                 f"DOCX compile fixture path is unsafe: {name}")
        _exact(descriptor, ("bytes", "sha256"),
               f"DOCX compile fixture {name}")
        _require(type(descriptor["bytes"]) is int and descriptor["bytes"] >= 0,
                 f"DOCX compile fixture {name}.bytes is malformed")
        _hash(descriptor["sha256"], f"DOCX compile fixture {name}.sha256")
        path = WORKSPACE / name
        _regular(path, f"DOCX compile fixture {name}")
        actual = _digest_descriptor(path, f"DOCX compile fixture {name}")
        _require(actual["bytes"] == descriptor["bytes"]
                 and actual["sha256"] == descriptor["sha256"],
                 f"DOCX compile fixture {name}: content changed")
        checked_compile_fixtures[name] = dict(descriptor)

    runtime_fixture_path = ROOT / "docx-runtime-fixture-inputs.json"
    runtime_fixtures = _json(runtime_fixture_path, str(runtime_fixture_path))
    _require(isinstance(runtime_fixtures, dict) and runtime_fixtures,
             "DOCX runtime fixture inventory is empty")
    checked_runtime_fixtures: dict[str, dict[str, Any]] = {}
    runtime_names: list[str] = []
    for name, descriptor in runtime_fixtures.items():
        _require(isinstance(name, str) and name and not Path(name).is_absolute()
                 and ".." not in Path(name).parts
                 and Path(name).suffix.lower() in {".docx", ".dotx"},
                 f"DOCX runtime fixture path is unsafe: {name}")
        _exact(descriptor, ("bytes", "sha256"),
               f"DOCX runtime fixture {name}")
        _require(type(descriptor["bytes"]) is int and descriptor["bytes"] >= 0,
                 f"DOCX runtime fixture {name}.bytes is malformed")
        _hash(descriptor["sha256"], f"DOCX runtime fixture {name}.sha256")
        runtime_names.append(name)
    _tracked_at_head(runtime_names, "DOCX runtime fixture inventory")
    for name in runtime_names:
        descriptor = runtime_fixtures[name]
        path = WORKSPACE / name
        actual = _digest_descriptor(path, f"DOCX runtime fixture {name}")
        _require(actual["bytes"] == descriptor["bytes"]
                 and actual["sha256"] == descriptor["sha256"],
                 f"DOCX runtime fixture {name}: content changed")
        checked_runtime_fixtures[name] = dict(descriptor)

    opc_fixture_path = ROOT / "opc-compile-fixture-inputs.json"
    opc_fixtures = _json(opc_fixture_path, str(opc_fixture_path))
    _require(isinstance(opc_fixtures, dict)
             and set(opc_fixtures) == OPC_COMPILE_FIXTURE_NAMES,
             "OPC compile fixture inventory differs")
    checked_opc_fixtures: dict[str, dict[str, Any]] = {}
    opc_names: list[str] = []
    for name, descriptor in opc_fixtures.items():
        _require(isinstance(name, str) and name and not Path(name).is_absolute()
                 and ".." not in Path(name).parts
                 and Path(name).suffix.lower() in {".docx", ".pptx", ".xlsx", ".zip"},
                 f"OPC compile fixture path is unsafe: {name}")
        _exact(descriptor, ("bytes", "sha256"),
               f"OPC compile fixture {name}")
        _require(type(descriptor["bytes"]) is int and descriptor["bytes"] >= 0,
                 f"OPC compile fixture {name}.bytes is malformed")
        _hash(descriptor["sha256"], f"OPC compile fixture {name}.sha256")
        opc_names.append(name)
    # All seven inputs are tracked at the frozen workspace HEAD.  In
    # particular, retain the historical change-0416 ZIP in this check rather
    # than silently treating it as an untracked convenience fixture.
    _tracked_at_head(opc_names, "OPC compile fixture inventory")
    for name in opc_names:
        descriptor = opc_fixtures[name]
        path = WORKSPACE / name
        actual = _digest_descriptor(path, f"OPC compile fixture {name}")
        _require(actual["bytes"] == descriptor["bytes"]
                 and actual["sha256"] == descriptor["sha256"],
                 f"OPC compile fixture {name}: content changed")
        checked_opc_fixtures[name] = dict(descriptor)

    opc_runtime_fixture_path = ROOT / "opc-runtime-fixture-inputs.json"
    opc_runtime_fixtures = _json(opc_runtime_fixture_path,
                                 str(opc_runtime_fixture_path))
    _require(isinstance(opc_runtime_fixtures, dict)
             and set(opc_runtime_fixtures) == OPC_RUNTIME_FIXTURE_NAMES,
             "OPC runtime fixture inventory differs")
    checked_opc_runtime_fixtures: dict[str, dict[str, Any]] = {}
    opc_runtime_names: list[str] = []
    for name, descriptor in opc_runtime_fixtures.items():
        _require(isinstance(name, str) and name and not Path(name).is_absolute()
                 and ".." not in Path(name).parts
                 and Path(name).suffix.lower() in {".ooxml", ".xlsx"},
                 f"OPC runtime fixture path is unsafe: {name}")
        _exact(descriptor, ("bytes", "sha256"),
               f"OPC runtime fixture {name}")
        _require(type(descriptor["bytes"]) is int and descriptor["bytes"] >= 0,
                 f"OPC runtime fixture {name}.bytes is malformed")
        _hash(descriptor["sha256"], f"OPC runtime fixture {name}.sha256")
        opc_runtime_names.append(name)
    _tracked_at_head(opc_runtime_names, "OPC runtime fixture inventory")
    for name in opc_runtime_names:
        descriptor = opc_runtime_fixtures[name]
        path = WORKSPACE / name
        actual = _digest_descriptor(path, f"OPC runtime fixture {name}")
        _require(actual["bytes"] == descriptor["bytes"]
                 and actual["sha256"] == descriptor["sha256"],
                 f"OPC runtime fixture {name}: content changed")
        checked_opc_runtime_fixtures[name] = dict(descriptor)

    harness_path = ROOT / "baseline-harness-manifest.json"
    harness = measure._shared_harness_binding()
    _require(harness["path"] == "baseline-harness-manifest.json"
             and Path(harness_path).is_file(), "common harness binding is missing")

    lock_path = ROOT / "workspace-Cargo.lock"
    lock = _digest_descriptor(lock_path, "workspace Cargo.lock", root=ROOT)
    workspace_lock = WORKSPACE / "Cargo.lock"
    _regular(workspace_lock, "workspace Cargo.lock source")
    _require(sha(workspace_lock) == lock["sha256"],
             "workspace Cargo.lock differs from retained dependency input")

    patch_path = ROOT / "baseline-tracked.patch"
    patch = _digest_descriptor(patch_path, "baseline tracked patch", root=ROOT)
    adr_path = ROOT / "adr-refresh.json"
    adr = _json(adr_path, str(adr_path))
    _exact(adr, ("files", "reference"), "adr-refresh.json")
    _require(isinstance(adr["files"], dict) and adr["files"],
             "adr-refresh.json files are empty")
    for name, digest in adr["files"].items():
        _require(isinstance(name, str) and name.startswith("docs/adr/")
                 and ".." not in Path(name).parts, f"ADR path is unsafe: {name}")
        _hash(digest, f"adr-refresh.{name}")
        source = WORKSPACE / name
        _regular(source, f"ADR {name}")
        _require(sha(source) == digest, f"ADR {name}: content changed")
    reference = adr["reference"]
    _exact(reference, ("path", "sha256"), "adr-refresh.reference")
    _hash(reference["sha256"], "adr-refresh.reference.sha256")
    reference_path = WORKSPACE / reference["path"]
    _regular(reference_path, "ADR refresh reference")
    _require(sha(reference_path) == reference["sha256"],
             "ADR refresh reference changed")

    boundary_check_path = ROOT / "boundary-check-inputs.json"
    boundary_checks = _json(boundary_check_path, str(boundary_check_path))
    boundary_names = {
        "tools/check_crate_boundaries.py",
        "tools/crate_boundaries.json",
    }
    _require(isinstance(boundary_checks, dict)
             and set(boundary_checks) == boundary_names,
             "boundary checker input inventory differs")
    checked_boundary_inputs: dict[str, dict[str, Any]] = {}
    for name, descriptor in boundary_checks.items():
        _exact(descriptor, ("bytes", "sha256"),
               f"boundary checker input {name}")
        _require(type(descriptor["bytes"]) is int and descriptor["bytes"] >= 0,
                 f"boundary checker input {name}.bytes is malformed")
        _hash(descriptor["sha256"], f"boundary checker input {name}.sha256")
        source = WORKSPACE / name
        actual = _digest_descriptor(source, f"boundary checker input {name}")
        _require(actual["bytes"] == descriptor["bytes"]
                 and actual["sha256"] == descriptor["sha256"],
                 f"boundary checker input {name}: content changed")
        checked_boundary_inputs[name] = dict(descriptor)

    boundary_adr_path = ROOT / "boundary-adr-inputs.json"
    boundary_adrs = _json(boundary_adr_path, str(boundary_adr_path))
    _require(isinstance(boundary_adrs, dict) and boundary_adrs,
             "boundary ADR input inventory is empty")
    _require(boundary_adrs == adr["files"],
             "boundary ADR inputs differ from ADR refresh custody")
    checked_boundary_adrs: dict[str, str] = {}
    for name, digest in boundary_adrs.items():
        _require(isinstance(name, str) and name.startswith("docs/adr/")
                 and ".." not in Path(name).parts,
                 f"boundary ADR path is unsafe: {name}")
        _hash(digest, f"boundary ADR {name}")
        checked_boundary_adrs[name] = digest

    isolated_path = ROOT / "isolated-source.json"
    isolated = _json(isolated_path, str(isolated_path))
    _exact(isolated, ("base", "reason", "sparse_paths", "worktree", "workspace_lock"),
           "isolated-source.json")
    _require(isinstance(isolated["base"], str)
             and re.fullmatch(r"[0-9a-f]{7,40}", isolated["base"]),
             "isolated source base is malformed")
    _require(isinstance(isolated["reason"], str) and isolated["reason"],
             "isolated source reason is missing")
    _require(isinstance(isolated["sparse_paths"], list) and isolated["sparse_paths"]
             and all(isinstance(item, str) and item and ".." not in Path(item).parts
                     for item in isolated["sparse_paths"]),
             "isolated source sparse paths are malformed")
    _exact(isolated["workspace_lock"], ("path", "sha256"),
           "isolated-source.workspace_lock")
    _require(isolated["workspace_lock"]["path"] == "workspace-Cargo.lock"
             and isolated["workspace_lock"]["sha256"] == lock["sha256"],
             "isolated source dependency binding differs")
    worktree = Path(isolated["worktree"])
    _require(worktree.is_absolute() and worktree.resolve(strict=False) == (TEMP / "source").resolve(strict=False),
             "isolated source worktree is outside the owned scratch root")

    transfers = []
    for path in sorted(ROOT.glob("development-candidate*-transfer.json")):
        value = _json(path, str(path))
        _require(isinstance(value, dict) and value, f"{path}: transfer inventory is empty")
        if {"note", "path", "sha256"} <= set(value):
            # A failed isolated attempt may transfer one explicit candidate
            # file instead of a complete relative-path inventory.  Keep that
            # provenance honest by requiring the path to stay in the owned
            # sparse worktree and by checking its current content hash.
            _exact(value, ("note", "path", "sha256"), f"{path} transfer")
            _require(isinstance(value["note"], str) and value["note"],
                     f"{path}: transfer note is missing")
            _hash(value["sha256"], f"{path}.sha256")
            candidate = Path(value["path"])
            _require(candidate.is_absolute(), f"{path}: transferred candidate path must be absolute")
            try:
                candidate.resolve(strict=False).relative_to(worktree.resolve(strict=False))
            except ValueError:
                measure.fail(f"{path}: transferred candidate escapes isolated source")
            _require(candidate.resolve(strict=False) != worktree.resolve(strict=False),
                     f"{path}: transferred candidate is the source root")
            # The isolated checkout is disposable and is intentionally
            # removed before post-cleanup bundle verification.  This is a
            # historical transfer descriptor: if the source file is still
            # present, only its path/type are checked because later candidate
            # attempts may have changed its contents; the recorded digest is
            # the digest of that historical attempt.
            if candidate.exists():
                _regular(candidate, f"{path} transferred candidate")
            else:
                _require(not candidate.is_symlink(),
                         f"{path}: transferred candidate became a symlink")
        else:
            for name, digest in value.items():
                _require(isinstance(name, str) and name.startswith("crates/")
                         and not Path(name).is_absolute() and ".." not in Path(name).parts,
                         f"{path}: unsafe transferred source path")
                _hash(digest, f"{path}.{name}")
        transfers.append(_digest_descriptor(path, path.name, root=ROOT))
    _require(transfers, "candidate transfer history is missing")
    reproduction = None
    reproduction_path = ROOT / "source-reproduction.json"
    if reproduction_path.exists():
        reproduction_value = _json(reproduction_path, "source-reproduction.json")
        _require(isinstance(reproduction_value, dict)
                 and {"patch", "source"} <= set(reproduction_value)
                 and set(reproduction_value) <= {"base_revision", "patch",
                                                "preexisting_source_delta", "source"},
                 "source-reproduction.json fields differ")
        _require(isinstance(reproduction_value["patch"], dict)
                 and set(reproduction_value["patch"]) <= {"bytes", "path", "sha256"}
                 and {"bytes", "sha256"} <= set(reproduction_value["patch"]),
                 "source-reproduction.patch fields differ")
        _require(reproduction_value["patch"]["bytes"] == patch["bytes"]
                 and reproduction_value["patch"]["sha256"] == patch["sha256"],
                 "source reproduction patch differs")
        _exact(reproduction_value["source"], ("files", "path", "sha256"),
               "source-reproduction.source")
        measure._source_binding(reproduction_value["source"], "source-reproduction.source")
        reproduction = _digest_descriptor(reproduction_path, "source-reproduction", root=ROOT)
    return {
        "fixtures": checked_fixtures,
        "harness_runtime_fixtures": checked_harness_runtime_fixtures,
        "docx_compile_fixtures": checked_compile_fixtures,
        "docx_runtime_fixtures": checked_runtime_fixtures,
        "opc_compile_fixtures": checked_opc_fixtures,
        "opc_runtime_fixtures": checked_opc_runtime_fixtures,
        "boundary_check_inputs": checked_boundary_inputs,
        "boundary_adr_inputs": checked_boundary_adrs,
        "harness": harness,
        "lock": lock,
        "patch": patch,
        "adr": _digest_descriptor(adr_path, "adr-refresh", root=ROOT),
        "isolated_source": _digest_descriptor(isolated_path, "isolated-source", root=ROOT),
        "transfers": transfers,
        "source_reproduction": reproduction,
    }


def _helper_archive(path: Path, label: str) -> dict[str, Any]:
    _directory(path, label)
    children = {item.name for item in path.iterdir()}
    allowed = {"gate.py", "support.py", "helper-custody.json"}
    _require(children <= allowed and {"gate.py", "support.py"} <= children,
             f"{label}: helper archive inventory differs")
    gate = _digest_descriptor(path / "gate.py", f"{label}.gate")
    support = _digest_descriptor(path / "support.py", f"{label}.support")
    _require(gate["sha256"] == sha(ROOT / "gate.py"),
             f"{label}: archived gate helper differs")
    custody = None
    if (path / "helper-custody.json").exists():
        custody = _json(path / "helper-custody.json", f"{label}.helper-custody")
        _exact(custody, ("schema", "version", "helpers"), f"{label}.helper-custody")
        _require(custody["schema"] == "docx-edit-provider-helper-archive-v1"
                 and custody["version"] == 1, f"{label}: helper custody schema differs")
        _require(isinstance(custody["helpers"], dict)
                 and set(custody["helpers"]) == {"gate.py", "support.py"}
                 and _matches_meta(custody["helpers"]["gate.py"], gate,
                                   f"{label}.helper-custody.gate.py")
                 and _matches_meta(custody["helpers"]["support.py"], support,
                                   f"{label}.helper-custody.support.py"),
                 f"{label}: helper custody differs")
    return {"gate": gate, "support": support, "custody": custody}


def _known_support_hashes() -> set[str]:
    values = {sha(ROOT / "support.py")}
    for path in sorted(ROOT.glob("development-*")):
        if path.is_dir() and (path / "support.py").is_file():
            values.add(sha(path / "support.py"))
    return values


def _validate_historical_receipt(path: Path) -> dict[str, Any]:
    """Validate a gate receipt without rewriting history as current success."""

    label = path.stem
    _safe_label(label, f"validation receipt {path}")
    value = _json(path, str(path))
    required = {
        "argv", "artifacts", "attempt", "common_sha256", "cwd", "driver_sha256",
        "environment", "exit_code", "finished_utc", "label", "schema",
        "source_after", "source_before", "source_unchanged", "started_utc",
        "termination", "timed_out", "timeout_seconds",
    }
    _exact(value, required, str(path))
    _require(value["schema"] == "docx-edit-provider-gate-v1"
             and value["label"] == label, f"{path}: receipt identity differs")
    _require(isinstance(value["argv"], list) and value["argv"]
             and all(isinstance(item, str) and item for item in value["argv"]),
             f"{path}: argv is malformed")
    _require(type(value["exit_code"]) is int and type(value["timed_out"]) is bool,
             f"{path}: process result is malformed")
    _require(value["termination"] is None or value["termination"] in {"SIGTERM", "SIGKILL"},
             f"{path}: termination is malformed")
    measure._timestamp(value["started_utc"], f"{path}.started_utc")
    _require(measure._timestamp(value["finished_utc"], f"{path}.finished_utc")
             > measure._timestamp(value["started_utc"], f"{path}.started_utc"),
             f"{path}: receipt chronology is invalid")
    _require(type(value["source_unchanged"]) is bool,
             f"{path}: source_unchanged is malformed")
    measure._source_binding(value["source_before"], f"{path}.source_before")
    measure._source_binding(value["source_after"], f"{path}.source_after")
    if value["source_unchanged"]:
        _require(value["source_before"] == value["source_after"],
                 f"{path}: unchanged receipt has differing source bindings")
    else:
        _require(value["source_before"] != value["source_after"],
                 f"{path}: changed-source receipt hides equal source bindings")
    _hash(value["driver_sha256"], f"{path}.driver_sha256")
    _require(value["driver_sha256"] == sha(ROOT / "gate.py"),
             f"{path}: gate helper changed")
    _hash(value["common_sha256"], f"{path}.common_sha256")
    _require(value["common_sha256"] in _known_support_hashes(),
             f"{path}: support helper has no retained archive")
    _require(isinstance(value["environment"], dict) and value["environment"],
             f"{path}: environment is malformed")
    _require(value["cwd"] and isinstance(value["cwd"], str),
             f"{path}: cwd is malformed")
    artifacts = value["artifacts"]
    _require(isinstance(artifacts, dict)
             and set(artifacts) == {f"{label}.stdout", f"{label}.stderr"},
             f"{path}: gate artifact inventory differs")
    for name, descriptor in artifacts.items():
        _exact(descriptor, ("bytes", "sha256"), f"{path}.artifacts.{name}")
        _require(type(descriptor["bytes"]) is int and descriptor["bytes"] >= 0,
                 f"{path}.artifacts.{name}.bytes is malformed")
        _hash(descriptor["sha256"], f"{path}.artifacts.{name}.sha256")
        artifact_path = path.with_name(name)
        actual = _digest_descriptor(artifact_path, f"{path}.{name}")
        _require(actual["bytes"] == descriptor["bytes"]
                 and actual["sha256"] == descriptor["sha256"],
                 f"{path}: gate artifact changed")
    started_path = path.with_name(f"{label}.started.json")
    _regular(started_path, f"{label} start receipt")
    started = _json(started_path, f"{label} start receipt")
    _require(isinstance(started, dict) and started.get("label") == label,
             f"{path}: start receipt identity differs")
    for key in ("argv", "attempt", "common_sha256", "cwd", "driver_sha256",
                "environment", "source_before", "started_utc", "schema"):
        _require(started.get(key) == value.get(key), f"{path}: start/terminal {key} differs")
    return value


def _validate_validation_history() -> dict[str, Any]:
    validation = ROOT / "validation"
    _directory(validation, "validation directory")
    receipts: list[dict[str, Any]] = []
    for started in sorted(validation.glob("*.started.json")):
        _regular(started, f"validation start {started.name}")
        terminal = started.with_name(started.name[:-len(".started.json")] + ".json")
        _regular(terminal, f"validation terminal {terminal.name}")
        receipts.append(_validate_historical_receipt(terminal))
    _require(receipts, "no validation receipts found")
    # A validation directory must not contain an unpaired terminal receipt or
    # an unrelated special path.  Final gates are also gate receipts and are
    # covered by the same schema.
    for path in validation.iterdir():
        _require(path.is_file() and not path.is_symlink(),
                 f"validation directory contains unsafe path: {path}")
        if path.name.endswith(".started.json"):
            continue
        _require(path.suffix in {".json", ".stdout", ".stderr"},
                 f"validation directory contains unexpected artifact: {path}")
    archive_bindings = {}
    for archive in sorted(ROOT.glob("development-*")):
        if not archive.is_dir():
            continue
        if (archive / "gate.py").is_file() or (archive / "support.py").is_file():
            archive_bindings[archive.name] = _helper_archive(
                archive, f"helper archive {archive.name}"
            )
    _require(archive_bindings, "historical helper archives are missing")
    # Development receipts deliberately include both failed and successful
    # attempts.  Preserve their status; only the final accepted captures may
    # be used as workload evidence.
    _require(any(item["exit_code"] != 0 for item in receipts),
             "failed development receipt history is missing")
    _require(any(item["exit_code"] == 0 for item in receipts),
             "successful validation history is missing")
    started_labels = {
        path.name[:-len(".started.json")]
        for path in validation.glob("*.started.json")
    }
    for path in validation.iterdir():
        if path.name.endswith(".started.json"):
            continue
        if path.suffix == ".json":
            _require(path.stem in started_labels,
                     f"validation terminal has no start receipt: {path.name}")
        elif path.suffix in {".stdout", ".stderr"}:
            _require(path.stem in started_labels,
                     f"validation artifact has no start receipt: {path.name}")
    negative_path = validation / "before-managed-refusal1.json"
    _regular(negative_path, "managed refusal receipt")
    negative = _json(negative_path, "managed refusal receipt")
    _require(isinstance(negative, dict) and negative.get("exit_code", 0) != 0,
             "managed refusal is not an explicit nonzero development result")
    _require("docx-managed-edit" in negative.get("argv", [])
             and "managed-api" in negative.get("argv", []),
             "managed refusal does not select the managed harness/API")
    negative_argv = negative["argv"]
    if "--output" in negative_argv:
        output_index = negative_argv.index("--output")
        _require(output_index + 1 < len(negative_argv),
                 "managed refusal output option is incomplete")
        refusal_output = Path(negative_argv[output_index + 1])
        _require(refusal_output.is_absolute()
                 and refusal_output.resolve(strict=False).parent == ROOT.resolve()
                 and not refusal_output.exists()
                 and not refusal_output.is_symlink(),
                 "managed refusal unexpectedly retained a report output")
    return {"count": len(receipts), "receipts": receipts,
            "negative_receipt": _digest_descriptor(negative_path, "managed refusal receipt", root=ROOT),
            "helper_archives": archive_bindings}


def _strict_uint(value: Any, path: str) -> int:
    _require(type(value) is int and value >= 0, f"{path}: strict unsigned integer required")
    return value


def _validate_allocation(allocation: Any, path: str) -> None:
    _require(isinstance(allocation, dict) and set(allocation) == _ALLOCATION_FIELDS,
             f"{path}: allocator fields differ")
    _require(allocation["status"] == "measured"
             and allocation["scope"] == measure.SAMPLE_ALLOCATION_SCOPE,
             f"{path}: allocator scope differs")
    for field in _ALLOCATION_UINT_FIELDS:
        _strict_uint(allocation[field], f"{path}.{field}")
    before = allocation["live_bytes_before"]
    allocated = allocation["allocated_bytes"]
    deallocated = allocation["deallocated_bytes"]
    after = allocation["live_bytes_after"]
    _require(after == before + allocated - deallocated,
             f"{path}: allocator live-byte conservation differs")
    peak_before = allocation["peak_live_bytes_before"]
    peak_after = allocation["peak_live_bytes_after"]
    region_peak = allocation["region_peak_live_bytes"]
    _require(peak_before >= before, f"{path}: pre-operation peak is below live bytes")
    _require(peak_after >= peak_before and peak_after >= after,
             f"{path}: post-operation peak bound differs")
    _require(region_peak >= before and region_peak >= after and region_peak <= peak_after,
             f"{path}: region peak bound differs")
    _require(allocation["allocation_calls"] >= allocation["reallocation_calls"],
             f"{path}: reallocation calls exceed allocation calls")


def _validate_raw_entries(entries: list[dict[str, Any]], *, pilot: bool) -> dict[str, int]:
    expected_processes = len(measure.formal_inventory(pilot=pilot))
    _require(len(entries) == expected_processes,
             f"{'pilot' if pilot else 'formal'} process inventory differs")
    allocator_rows = 0
    managed_rows = 0
    managed_budget_rows = 0
    for index, entry in enumerate(entries):
        spec = entry["spec"]
        rows = entry["report"]["rows"]
        _require(len(rows) == spec["samples"], f"raw entry {index}: sample count differs")
        if spec["api"] == measure.MANAGED_API:
            managed_rows += len(rows)
        for row_index, row in enumerate(rows):
            # Re-run the canonical phase-aware budget validator at the bundle
            # boundary with the row's exact source and sink evidence.  This
            # keeps the bundle check fail-closed if report validation is ever
            # bypassed, while preserving the distinct cache/resource sampling
            # points in the harness lifecycle.
            measure._check_budget(
                row["budget"], spec["api"],
                f"raw[{index}].rows[{row_index}].budget",
                reads=row["reads"],
                sink_bytes=row["sink"]["accepted_bytes"],
            )
            if spec["api"] == measure.MANAGED_API:
                managed_budget_rows += 1
            allocation = row.get("allocation")
            if spec["role"] == "allocator":
                _validate_allocation(allocation, f"raw[{index}].rows[{row_index}].allocation")
                allocator_rows += 1
            else:
                _require(allocation is None, f"raw[{index}].rows[{row_index}]: normal row has allocation")
    expected_allocator_processes = sum(
        item["role"] == "allocator" for item in measure.formal_inventory(pilot=pilot)
    )
    expected_allocator_rows = expected_allocator_processes * (
        measure.PILOT_SAMPLES if pilot else measure.FORMAL_SAMPLES
    )
    _require(allocator_rows == expected_allocator_rows,
             f"{'pilot' if pilot else 'formal'} allocator row inventory differs")
    return {"processes": len(entries), "allocator_rows": allocator_rows,
            "managed_rows": managed_rows, "managed_budget_rows": managed_budget_rows}


def _verify_accepted(value: Any, builds: dict[str, dict[str, Any]],
                     protocol_hash: str) -> dict[str, Any]:
    _require(isinstance(value, dict), "accepted evidence is not an object")
    required = {"formal_attempt", "formal_samples", "pilot_attempt",
                "pilot_samples", "schema"}
    optional = {"claim_authorized", "version", "negative_attempt", "negative_reason"}
    _require(set(value) >= required and set(value) <= required | optional,
             "accepted-evidence.json fields differ")
    _require(value["schema"] == ACCEPTED_SCHEMA
             and value.get("version", 1) == 1
             and value.get("claim_authorized", False) is False,
             "accepted evidence identity differs")
    formal_attempt = _safe_token(value["formal_attempt"], "accepted.formal_attempt")
    pilot_attempt = _safe_token(value["pilot_attempt"], "accepted.pilot_attempt")
    _require(formal_attempt != pilot_attempt, "formal and pilot attempts must differ")
    formal_expected = sum(int(item["samples"]) for item in measure.formal_inventory())
    pilot_expected = sum(int(item["samples"]) for item in measure.formal_inventory(pilot=True))
    _require(value["formal_samples"] == formal_expected
             and value["pilot_samples"] == pilot_expected,
             "accepted sample inventory differs from frozen matrix")
    if "negative_attempt" in value:
        _require(value["negative_attempt"] == "before-managed-refusal1",
                 "accepted negative attempt differs")
    if "negative_reason" in value:
        _require(isinstance(value["negative_reason"], str) and value["negative_reason"],
                 "accepted negative reason is missing")
    formal_path = measure.verify(formal_attempt, pilot=False)
    pilot_path = measure.verify(pilot_attempt, pilot=True)
    for path in (formal_path, pilot_path):
        proof = _json(Path(path), str(path))
        _require(isinstance(proof, dict) and proof.get("status") == "pass"
                 and proof.get("protocol_sha256") == protocol_hash,
                 f"{path}: verification proof is not bound to protocol")
    formal_protocol, formal_hash = measure._load_protocol(builds)
    _require(formal_hash == protocol_hash and formal_protocol == measure.protocol_value(builds),
             "protocol changed while accepting reports")
    formal_entries = measure._collect(formal_attempt, builds, formal_protocol,
                                      protocol_hash, pilot=False)
    pilot_entries = measure._collect(pilot_attempt, builds, formal_protocol,
                                     protocol_hash, pilot=True)
    formal_inventory = _validate_raw_entries(formal_entries, pilot=False)
    pilot_inventory = _validate_raw_entries(pilot_entries, pilot=True)
    return {
        "formal_attempt": formal_attempt, "pilot_attempt": pilot_attempt,
        "formal_samples": formal_expected, "pilot_samples": pilot_expected,
        "formal_inventory": formal_inventory, "pilot_inventory": pilot_inventory,
        "formal_verification": _digest_descriptor(Path(formal_path), "formal verification", root=ROOT),
        "pilot_verification": _digest_descriptor(Path(pilot_path), "pilot verification", root=ROOT),
    }


def _expected_helpers() -> set[str]:
    return set(measure.DRIVER_FILES) | {
        "cleanup.py", "verify.py", "profile.py", "test_cleanup.py", "test_measure.py",
        "test_profile.py", "test_verify.py",
    }


def _matches_meta(value: Any, actual: Mapping[str, Any], path: str) -> bool:
    """Accept the compact historical ``meta`` form or a rooted descriptor."""

    _require(isinstance(value, dict), f"{path}: descriptor is missing")
    if set(value) == {"bytes", "sha256"}:
        return value == {"bytes": actual["bytes"], "sha256": actual["sha256"]}
    if set(value) == {"bytes", "path", "sha256"}:
        candidate = Path(value["path"])
        if not candidate.is_absolute():
            candidate = ROOT / candidate
        return (value["bytes"] == actual["bytes"] and value["sha256"] == actual["sha256"]
                and candidate.resolve() == Path(actual["path"]).resolve())
    if set(value) == {"bytes", "path", "relative_path", "sha256"}:
        return value == dict(actual)
    measure.fail(f"{path}: descriptor fields differ")
    return False


def _validate_helper_custody(gates: Mapping[str, Any], gate_bindings: Mapping[str, Any]) -> dict[str, Any]:
    path = ROOT / "helper-test-custody.json"
    value = _json(path, "helper-test-custody.json")
    _require(isinstance(value, dict)
             and {"gate", "helpers", "label", "source"} <= set(value)
             and set(value) <= {"gate", "helpers", "label", "schema", "source", "version"},
             "helper-test-custody.json fields differ")
    if "schema" in value or "version" in value:
        _require(value.get("schema") == "docx-edit-provider-helper-custody-v1"
                 and value.get("version") == 1, "helper custody schema differs")
    label = _safe_label(value["label"], "helper custody label")
    _require(label in gates["commands"] and label in gate_bindings,
             "helper custody gate is not a final gate")
    _require(tuple(gates["commands"][label]) == tuple(_expected_final_commands()[5]),
             "helper custody does not select Python helper gate")
    helpers = value["helpers"]
    _require(isinstance(helpers, dict) and set(helpers) == _expected_helpers(),
             "helper custody inventory differs")
    expected_helpers = {name: _digest_descriptor(ROOT / name, f"helper {name}", root=ROOT)
                        for name in sorted(_expected_helpers())}
    _require(set(helpers) == set(expected_helpers)
             and all(_matches_meta(helpers[name], expected_helpers[name], f"helper {name}")
                     for name in expected_helpers),
             "tested helper changed")
    gate_path = ROOT / "validation" / f"{label}.json"
    expected_gate = _digest_descriptor(gate_path, "helper gate", root=ROOT)
    _require(_matches_meta(value["gate"], expected_gate, "helper gate"),
             "helper gate custody changed")
    _require(value["source"] == gate_bindings[label]["source"],
             "helper gate source custody differs")
    return {"label": label, "helpers": expected_helpers,
            "gate": value["gate"], "source": value["source"]}


def _verify_cleanup() -> dict[str, Any]:
    proof = cleanup_driver.verify(root=ROOT, temp=TEMP, target=TARGET_DIR)
    _require(isinstance(proof, dict) and proof.get("status") == "pass",
             "cleanup verification did not pass")
    return proof


def inventory(root: Path = ROOT) -> dict[str, dict[str, Any]]:
    root = Path(root)
    _directory(root, "evidence root")
    _require(root.resolve(strict=True) == root.absolute(),
             "evidence root is not canonical")
    files: dict[str, dict[str, Any]] = {}
    for path in sorted(root.rglob("*")):
        _require(not path.is_symlink(), f"symlink in evidence: {path}")
        if path == root / SEAL_NAME:
            _regular(path, "seal")
            continue
        if path.is_file():
            _require("__pycache__" not in path.parts and path.suffix != ".pyc",
                     f"generated bytecode in evidence: {path}")
            files[path.relative_to(root).as_posix()] = meta(path)
        elif not path.is_dir():
            measure.fail(f"special file in evidence: {path}")
    return files


def verify_evidence() -> dict[str, Any]:
    builds = measure.load_builds()
    protocol, protocol_hash = measure._load_protocol(builds)
    _require(protocol == measure.protocol_value(builds), "protocol/build mismatch")
    baseline = _validate_baseline_inputs()
    history = _validate_validation_history()
    accepted = _json(ROOT / "accepted-evidence.json", "accepted-evidence.json")
    accepted_result = _verify_accepted(accepted, builds, protocol_hash)

    gates = _json(ROOT / "final-gates.json", "final-gates.json")
    final_commands = _validate_final_gates(gates)
    gate_bindings: dict[str, Any] = {}
    final_source = builds["after-normal"]["source"]
    helper_command = tuple(_expected_final_commands()[5])
    for label, command in final_commands:
        path = ROOT / "validation" / f"{label}.json"
        binding_source = None if tuple(command) == helper_command else final_source
        gate_bindings[label] = _validate_gate_receipt(label, command, path, binding_source)
    helper = _validate_helper_custody(gates, gate_bindings)

    cleanup_proof = _verify_cleanup()
    cleanup_path = ROOT / cleanup_driver.SEAL_NAME
    cleanup_receipt = _digest_descriptor(cleanup_path, "cleanup receipt", root=ROOT)
    expected_binaries = {
        Path(build["binary"]["path"]).resolve()
        for build in builds.values()
    }
    actual_binaries = {
        path.resolve() for path in TEMP.rglob("*")
        if path.is_file() and not path.is_symlink()
    }
    _require(actual_binaries == expected_binaries,
             "temporary root does not retain exactly the four build executables")
    _require(not TARGET_DIR.exists() and not TARGET_DIR.is_symlink(),
             "Cargo target remains after cleanup")
    return {
        "status": "pass", "formal_samples": accepted_result["formal_samples"],
        "pilot_samples": accepted_result["pilot_samples"],
        "protocol_sha256": protocol_hash,
        "baseline": baseline, "validation": history,
        "accepted": accepted_result, "gates": sorted(gates["commands"]),
        "helper_custody": helper, "cleanup_receipt": cleanup_receipt,
        "cleanup_verification_schema": cleanup_proof["schema"],
    }


def _write_seal(value: dict[str, Any]) -> Path:
    path = ROOT / SEAL_NAME
    _require(not path.exists() and not path.is_symlink(),
             f"refusing to replace existing seal: {path}")
    write(path, value)
    return path


def main(argv: list[str] | None = None) -> int:
    values = list(sys.argv[1:] if argv is None else argv)
    if values not in (["create"], ["verify"]):
        raise SystemExit("usage: verify.py create|verify")
    evidence = verify_evidence()
    result = {"schema": BUNDLE_SCHEMA, "version": BUNDLE_VERSION,
              "evidence": evidence, "files": inventory()}
    path = ROOT / SEAL_NAME
    if values == ["create"]:
        _write_seal(result)
    else:
        _require(path.is_file() and not path.is_symlink(), "seal is missing")
        _require(_json(path, "seal.json") == result,
                 "bundle inventory or evidence changed")
    print(json.dumps({"status": "pass", "files": len(result["files"]),
                      **meta(path)}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
