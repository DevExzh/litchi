#!/usr/bin/env python3
"""Source-bound consumer for the signed-native policy v2.

The policy module owns the immutable candidate fence.  This module owns the
external consumer project and the receipts which consume that fence.  It does
not copy a production tree or patch a format crate: the generated Rust bridge
is deliberately a consumer-only adapter for the two API spellings which the
historical probe must compile against.

Preparation is source-only.  ``build`` is the only function which may invoke
Cargo, and it does so only after revalidating the preparation receipt.  Native
route receipts are consumed by ``execute``; this module never treats a route
status, a binary path, or a source digest supplied by a caller as trusted
without checking the corresponding receipt artifact.
"""

from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import tarfile
from typing import Any, Mapping, Sequence


PREPARATION_SCHEMA = "litchi.signed-native.preparation.v2"
BUILD_SCHEMA = "litchi.signed-native.build.v2"
CAPABILITIES_SCHEMA = "litchi.signed-native.capabilities.v2"
EXECUTION_SCHEMA = "litchi.signed-native.execution.v2"
ADAPTER_SCHEMA = "litchi.signed-native.adapter.v2"
FIXTURE_SCHEMA = "litchi.signed-native.fixture-binding.v2"

TARGET_CAP_BYTES = 20_000_000_000
TMP_CAP_BYTES = 20_000_000_000
CHILD_TIMEOUT_SECONDS = 300.0
CHILD_OUTPUT_LIMIT = 4 * 1024 * 1024

_ALLOWED_ROUTE_STATUS = frozenset(
    {"available", "capability_missing", "expected_refusal", "not_applicable"}
)
_REQUIRED_SOURCE_KEYS = (
    "commit",
    "git_tree",
    "tree_sha256",
    "manifests",
    "locks",
    "vendor_tree",
    "toolchain",
    "probe_inputs",
    "fixture",
    "archive",
)


class ConsumerError(ValueError):
    """A preparation, build, or execution receipt violates the v2 contract."""


def _policy() -> Any:
    # Import lazily.  The policy module delegates to us after it has checked
    # its own execution gate, while direct callers still get the same gate.
    from native_qualification import prepare_signed_native_inputs_v2 as module

    return module


def _expect(condition: bool, message: str) -> None:
    if not condition:
        raise ConsumerError(message)


def _canonical(value: Any) -> bytes:
    return (
        json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        + "\n"
    ).encode("utf-8")


def _sha256(payload: bytes) -> str:
    return hashlib.sha256(payload).hexdigest()


def _reject_symlink_components(path: Path, label: str) -> None:
    lexical = path.absolute()
    cursor = Path(lexical.anchor)
    for component in lexical.parts[1:]:
        cursor /= component
        _expect(not cursor.is_symlink(), f"{label} path contains a symlink: {cursor}")


def _safe_external(path: Path, label: str) -> Path:
    _reject_symlink_components(path, label)
    lexical = path.absolute()
    _expect("3rd" + "party" not in lexical.parts, f"{label} may not use an upstream checkout")
    repository = _policy().REPOSITORY.resolve()
    resolved = lexical.resolve(strict=False)
    for forbidden in (
        repository / "docs" / "performance" / "results",
        repository / "docs" / "report",
        repository / "binaries",
    ):
        _expect(
            not (resolved == forbidden.resolve() or forbidden.resolve() in resolved.parents),
            f"{label} must remain outside {forbidden}",
        )
    return resolved


def _under(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
    except ValueError:
        return False
    return True


def _assert_disjoint(paths: Mapping[str, Path]) -> None:
    resolved: list[tuple[str, Path]] = []
    for label, path in paths.items():
        _reject_symlink_components(path, label)
        resolved.append((label, path.absolute().resolve(strict=False)))
    for index, (left_label, left) in enumerate(resolved):
        for right_label, right in resolved[index + 1 :]:
            _expect(
                left != right and not _under(left, right) and not _under(right, left),
                f"paths must be pairwise disjoint: {left_label}={left}, {right_label}={right}",
            )


def _artifact(path: Path, label: str = "artifact", *, executable: bool = False) -> dict[str, Any]:
    _reject_symlink_components(path, label)
    _expect(path.is_file() and not path.is_symlink(), f"{label} is not a regular file: {path}")
    if executable:
        _expect(os.access(path, os.X_OK), f"{label} is not executable: {path}")
    resolved = path.resolve()
    payload = resolved.read_bytes()
    return {"path": str(resolved), "bytes": len(payload), "sha256": _sha256(payload)}


def _verify_artifact(value: Any, label: str, *, executable: bool = False) -> dict[str, Any]:
    _expect(isinstance(value, Mapping), f"{label} artifact is missing")
    path = value.get("path")
    expected_bytes = value.get("bytes")
    expected_sha = value.get("sha256")
    _expect(isinstance(path, str) and path, f"{label} artifact path is missing")
    _expect(isinstance(expected_bytes, int) and expected_bytes >= 0, f"{label} artifact size is malformed")
    _expect(isinstance(expected_sha, str) and len(expected_sha) == 64, f"{label} artifact digest is malformed")
    actual = _artifact(Path(path), label, executable=executable)
    _expect(actual == dict(value), f"{label} artifact changed since its receipt")
    return actual


def _load_canonical(path: Path, label: str) -> tuple[dict[str, Any], dict[str, Any]]:
    _reject_symlink_components(path, label)
    _expect(path.is_file() and not path.is_symlink(), f"{label} is missing: {path}")
    try:
        raw = path.read_bytes()
        value = json.loads(raw.decode("utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise ConsumerError(f"{label} is invalid JSON: {error}") from error
    _expect(isinstance(value, dict), f"{label} is not an object")
    _expect(raw == _canonical(value), f"{label} is not canonical JSON")
    return value, {"path": str(path.resolve()), "bytes": len(raw), "sha256": _sha256(raw)}


def _write_new(path: Path, value: Any, label: str) -> dict[str, Any]:
    _reject_symlink_components(path, label)
    _expect(not path.exists() and not path.is_symlink(), f"refusing to overwrite {label}: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(_canonical(value))
    return _artifact(path, label)


def _write_bytes_new(path: Path, payload: bytes, label: str) -> dict[str, Any]:
    _reject_symlink_components(path, label)
    _expect(not path.exists() and not path.is_symlink(), f"refusing to overwrite {label}: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(payload)
    return _artifact(path, label)


def _copy_regular(source: Path, destination: Path, label: str) -> dict[str, Any]:
    _reject_symlink_components(source, label)
    _expect(source.is_file() and not source.is_symlink(), f"{label} source is not regular: {source}")
    _write_bytes_new(destination, source.read_bytes(), label)
    return _artifact(destination, label)


def _adapter_source() -> str:
    """Return the consumer-only API bridge used by both source endpoints."""

    return r'''//! Consumer-only bridge for the signed-native v2 probe.
#![allow(dead_code)]

/// The owner-confirmed batch policy is intentionally explicit in the
/// generated consumer.  A refusal happens before the iterator is touched.
pub const SIGNED_BATCH_SELECTION: &str = "immediate_refusal";

pub fn signed_batch_refusal() -> (&'static str, u64, u64) {
    // (status, iterator_reads, publication_attempts)
    ("expected_refusal", 0, 0)
}

#[cfg(feature = "legacy_api")]
pub fn exercise_api() {
    fn commit_fields<T>(commit: &litchi_ops::transaction::Commit<T>) {
        let _ = &commit.file;
        let _ = &commit.patch;
        let _ = &commit.diagnostics;
    }
    fn edit_entry(file: &litchi_ops::File) {
        let _ = file.edit_with_limits(
            &litchi_ops::transaction::SemanticLimits::default(),
            |_edit| Ok::<(), litchi_ops::Error>(()),
        );
    }
    fn xlsb_entry(package: &litchi_xlsb::Package) {
        let _ = package.edit_cell_values(0usize);
    }
    let _ = (commit_fields::<()>, edit_entry, xlsb_entry);
}

#[cfg(not(feature = "legacy_api"))]
pub fn exercise_api() {
    fn commit_methods<T>(commit: &litchi_ops::transaction::Commit<T>) {
        let _ = commit.file();
        let _ = commit.patch();
        let _ = commit.diagnostics();
    }
    fn edit_entry(file: &litchi_ops::File) {
        let options = litchi_ops::transaction::EditOptions::new()
            .semantic_limits(litchi_ops::transaction::SemanticLimits::default());
        let _ = file.edit_with_options(options, |_edit| Ok::<(), litchi_ops::Error>(()));
    }
    fn xlsb_entry(package: &litchi_xlsb::Package) {
        let _ = package.edit_cell_values(0usize).resolve();
    }
    let _ = (commit_methods::<()>, edit_entry, xlsb_entry);
}
'''


def _project_toml(source: Path) -> str:
    dependencies = (
        ("litchi-ops", "crates/litchi-ops"),
        ("litchi-core", "crates/litchi-core"),
        ("litchi-opc", "crates/litchi-opc"),
        ("litchi-cfb", "crates/litchi-cfb"),
        ("litchi-docx", "crates/litchi-docx"),
        ("litchi-pptx", "crates/litchi-pptx"),
        ("litchi-xlsx", "crates/litchi-xlsx"),
        ("litchi-xlsb", "crates/litchi-xlsb"),
        ("litchi-doc", "crates/litchi-doc"),
        ("litchi-ppt", "crates/litchi-ppt"),
        ("litchi-xls", "crates/litchi-xls"),
        ("serde_json", None),
        ("sha2", None),
    )
    lines = [
        "[package]",
        'name = "signed-native-consumer"',
        'version = "0.0.0"',
        'edition = "2024"',
        "",
        "[features]",
        'default = ["legacy_api"]',
        'legacy_api = []',
        "",
        "[dependencies]",
    ]
    for name, relative in dependencies:
        if relative is None:
            version = "1" if name == "serde_json" else "0.10"
            lines.append(f'{name} = "{version}"')
        else:
            lines.append(f'{name} = {{ path = {json.dumps((source / relative).as_posix())} }}')
    return "\n".join(lines) + "\n"


def _fixture_value(source: Path) -> dict[str, Any]:
    module = _policy()
    manifest_path = source / module.FIXTURE_INPUT["path"]
    module._fixture_policy(source)
    try:
        value = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise ConsumerError(f"cannot read signed fixture manifest: {error}") from error
    _expect(isinstance(value, dict), "signed fixture manifest is not an object")
    _expect(_canonical(value) == manifest_path.read_bytes(), "signed fixture manifest is not canonical")
    return value


def _stage_fixtures(source: Path, output: Path, value: Mapping[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
    root = output / "fixture-root"
    _expect(not root.exists() and not root.is_symlink(), f"fixture root is not fresh: {root}")
    rows = value.get("fixtures")
    _expect(isinstance(rows, list) and rows, "signed fixture manifest is empty")
    root.mkdir(parents=True)
    for index, row in enumerate(rows):
        _expect(isinstance(row, Mapping), f"fixture row {index} is malformed")
        relative = _policy()._safe_relative(row.get("relative_path"), f"fixture row {index} path")
        source_file = source / relative
        destination = root / relative
        _expect(not destination.is_symlink(), f"fixture destination is a symlink: {destination}")
        _expect(_policy()._file_identity(source, relative)["sha256"] == row.get("input_sha256"), f"fixture row {index} changed")
        _copy_regular(source_file, destination, f"fixture row {index}")
    manifest_relative = Path(_policy().FIXTURE_INPUT["path"])
    manifest_path = root / manifest_relative
    manifest_artifact = _write_bytes_new(manifest_path, _canonical(value), "staged fixture manifest")
    binding = {
        "schema": FIXTURE_SCHEMA,
        "root": str(root.resolve()),
        "manifest": manifest_artifact,
        "source_manifest": _policy()._file_identity(source, _policy().FIXTURE_INPUT["path"]),
        "rows": len(rows),
        "cryptographic_verification": "not_performed",
    }
    return binding, {"path": str(root.resolve()), "manifest": manifest_artifact, "rows": len(rows)}


def _stage_probes(source: Path, output: Path) -> dict[str, Any]:
    module = _policy()
    destination = output / "probe-inputs"
    destination.mkdir(parents=True)
    rows: dict[str, Any] = {}
    for name, expected in module.PROBE_INPUTS.items():
        source_path = source / expected["path"]
        target = destination / f"{name}-{Path(expected['path']).name}"
        source_artifact = module._file_identity(source, expected["path"])
        _expect(source_artifact == expected, f"probe input changed: {expected['path']}")
        staged = _copy_regular(source_path, target, f"staged {name} probe")
        rows[name] = {"source": source_artifact, "staged": staged}
    return {"schema": "litchi.signed-native.probes.v2", "inputs": rows}


def _source_closure(source: Path, identity: Mapping[str, Any]) -> dict[str, Any]:
    module = _policy()
    binding = module.CANDIDATE_BINDING
    _expect(isinstance(binding, Mapping), "candidate source identity is not frozen")
    source_binding = binding.get("source")
    _expect(isinstance(source_binding, Mapping), "candidate source binding is missing")
    closure = {
        "source_root": str(source.resolve()),
        "commit": identity.get("commit"),
        "git_tree": identity.get("git_tree"),
        "tree_sha256": identity.get("tree_sha256"),
        "archive": identity.get("archive", source_binding.get("archive")),
        "manifests": identity.get("manifests", binding.get("manifests")),
        "locks": identity.get("locks"),
        "lock_graph": identity.get("lock_graph", binding.get("lock_graph")),
        "vendor_tree": identity.get("vendor_tree", binding.get("vendor_tree")),
        "toolchain": identity.get("toolchain", binding.get("toolchain")),
        "toolchain_paths": list(module.TOOLCHAIN_PATHS),
        "fixture": identity.get("fixture", module.FIXTURE_INPUT),
        "probe_inputs": identity.get("probe_inputs", module.PROBE_INPUTS),
        "source_manifests": list(module.SOURCE_MANIFEST_PATHS),
        "vendor_roots": list(module.VENDOR_ROOTS),
        "role_locks": module.ROLE_LOCKS,
    }
    for key in _REQUIRED_SOURCE_KEYS:
        _expect(closure.get(key) is not None, f"source closure is missing {key}")
    return closure


def _policy_identity() -> dict[str, Any]:
    module = _policy()
    policy_path = module.POLICY_PATH
    return _artifact(policy_path, "signed-native policy")


def _adapter_receipt(adapter: Path, main: Path, manifest: Path) -> dict[str, Any]:
    module = _policy()
    return {
        "schema": ADAPTER_SCHEMA,
        "source": "consumer-only generated bridge; production crates are read-only",
        "legacy_feature": "legacy_api",
        "api_branches": ["legacy_commit_fields", "current_commit_accessors", "legacy_xlsb_option", "current_xlsb_resolve"],
        "signed_batch": dict(module.SIGNED_BATCH_POLICY),
        "module": _artifact(adapter, "consumer adapter"),
        "probe": _artifact(main, "generated signed probe"),
        "manifest": _artifact(manifest, "consumer manifest"),
    }


def prepare(source_root: Path, output_root: Path, *, identity: Mapping[str, Any] | None = None) -> dict[str, Any]:
    """Generate a v2 consumer project after the policy source gate passes."""

    module = _policy()
    if identity is None:
        identity = module.candidate_identity(source_root)
    source = module._root(source_root)
    output = _safe_external(Path(output_root), "preparation output")
    archive_path = Path(str(module.CANDIDATE_BINDING["source"]["archive"]["path"]))
    _assert_disjoint({"candidate source": source, "candidate archive": archive_path, "preparation output": output})
    _expect(not output.exists() and not output.is_symlink(), f"refusing to overwrite preparation output: {output}")
    closure = _source_closure(source, identity)
    fixture_value = _fixture_value(source)
    output.mkdir(parents=True)
    project = output / "consumer"
    (project / "src").mkdir(parents=True)
    probe_path = source / module.PROBE_INPUTS["probe"]["path"]
    probe_text = probe_path.read_text(encoding="utf-8")
    _expect("fn main()" in probe_text, "signed probe has no main function")
    generated_main = "mod api_compat;\n" + probe_text.replace(
        "fn main() {", "fn main() {\n    api_compat::exercise_api();", 1
    )
    adapter_path = project / "src/api_compat.rs"
    main_path = project / "src/main.rs"
    manifest_path = project / "Cargo.toml"
    _write_bytes_new(adapter_path, _adapter_source().encode("utf-8"), "consumer adapter")
    _write_bytes_new(main_path, generated_main.encode("utf-8"), "generated signed probe")
    _write_bytes_new(manifest_path, _project_toml(source).encode("utf-8"), "consumer manifest")
    fixture_binding, fixture_summary = _stage_fixtures(source, output, fixture_value)
    probe_receipt = _stage_probes(source, output)
    policy_identity = _policy_identity()
    adapter = _adapter_receipt(adapter_path, main_path, manifest_path)
    source_identity = dict(identity)
    source_identity["source_root"] = str(source.resolve())
    source_identity["adapter"] = adapter
    source_identity["source_capture"] = str((output / "source-identity.json").resolve())
    source_identity_artifact = _write_new(output / "source-identity.json", source_identity, "source identity")
    fixture_binding_artifact = _write_new(output / "fixture-manifest-receipt.json", fixture_binding, "fixture binding")
    probes_artifact = _write_new(output / "probe-inputs.json", probe_receipt, "probe inputs")
    plan = {
        "schema": PREPARATION_SCHEMA,
        "status": "prepared",
        "policy": policy_identity,
        "candidate_binding_sha256": _sha256(_canonical(module.CANDIDATE_BINDING)),
        "source": closure,
        "adapter": adapter,
        "fixture": fixture_summary,
        "probes": probe_receipt,
        "formats": list(module.NATIVE_FORMATS),
        "routes": list(module.ROUTES),
        "correctness_only": sorted(module.CORRECTNESS_ONLY),
        "baseline_not_applicable": sorted(module.BASELINE_NOT_APPLICABLE),
        "signed_batch": dict(module.SIGNED_BATCH_POLICY),
        "build_policy": {
            "offline": True,
            "locked": True,
            "jobs": 1,
            "incremental": 0,
            "memory_cap": "20G",
            "target_cap_bytes": TARGET_CAP_BYTES,
            "tmp_cap_bytes": TMP_CAP_BYTES,
            "timing": "not run by this source-only preparation",
        },
        "commands": [
            "<pinned-cargo-shim> generate-lockfile --offline",
            "<pinned-cargo-shim> metadata --offline --locked --format-version=1",
            "<pinned-cargo-shim> build --offline --locked --release --bin signed-native-consumer",
        ],
        "receipts": {
            "source_identity": source_identity_artifact,
            "fixture_binding": fixture_binding_artifact,
            "probes": probes_artifact,
        },
    }
    plan_artifact = _write_new(output / "plan.json", plan, "preparation plan")
    receipt = {
        "schema": PREPARATION_SCHEMA,
        "status": "prepared",
        "execution_admitted": True,
        "policy": policy_identity,
        "candidate_binding_sha256": plan["candidate_binding_sha256"],
        "source_identity": source_identity_artifact,
        "source": closure,
        "adapter": adapter,
        "fixture": fixture_summary,
        "probes": probe_receipt,
        "formats": list(module.NATIVE_FORMATS),
        "routes": list(module.ROUTES),
        "correctness_only": sorted(module.CORRECTNESS_ONLY),
        "baseline_not_applicable": sorted(module.BASELINE_NOT_APPLICABLE),
        "signed_batch": dict(module.SIGNED_BATCH_POLICY),
        "plan": plan_artifact,
        "fixture_binding": fixture_binding_artifact,
        "probe_receipt": probes_artifact,
    }
    _write_new(output / "preparation-receipt.json", receipt, "preparation receipt")
    return {"status": "prepared", "output": str(output), "receipt": receipt}


def _source_identity_from_preparation(prepared: Path, receipt: Mapping[str, Any]) -> dict[str, Any]:
    module = _policy()
    source_artifact = receipt.get("source_identity")
    _verify_artifact(source_artifact, "source identity")
    source_path = Path(str(source_artifact["path"]))
    source_value, actual = _load_canonical(source_path, "source identity")
    _expect(actual == dict(source_artifact), "source identity artifact changed")
    source_root = source_value.get("source_root")
    _expect(isinstance(source_root, str), "source identity root is missing")
    current = module.candidate_identity(Path(source_root))
    current["source_root"] = str(Path(source_root).resolve())
    current["adapter"] = source_value.get("adapter")
    current["source_capture"] = str(source_path.resolve())
    _expect(current == source_value, "candidate source identity changed after preparation")
    _expect(receipt.get("source") == _source_closure(Path(source_root), source_value), "preparation source closure changed")
    return source_value


def _concrete_source_identity(state: Mapping[str, Any]) -> dict[str, Any]:
    """Return the endpoint-facing source identity, not its receipt artifact."""

    source = dict(state["source"])
    source["source_ref"] = source["commit"]
    source["source_capture"] = state["receipt"]["source_identity"]["path"]
    source["adapter"] = state["adapter"]
    return source


def _verify_fixture_binding(prepared: Path, value: Mapping[str, Any]) -> dict[str, Any]:
    module = _policy()
    binding_artifact = value.get("fixture_binding")
    _verify_artifact(binding_artifact, "fixture binding")
    path = Path(str(binding_artifact["path"]))
    binding, actual = _load_canonical(path, "fixture binding")
    _expect(actual == dict(binding_artifact), "fixture binding artifact changed")
    _expect(binding.get("schema") == FIXTURE_SCHEMA, "fixture binding schema differs")
    root = Path(str(binding.get("root", "")))
    manifest = binding.get("manifest")
    _verify_artifact(manifest, "staged fixture manifest")
    manifest_path = Path(str(manifest["path"]))
    _expect(manifest_path == root / module.FIXTURE_INPUT["path"], "staged fixture manifest path differs")
    fixture_value, _ = _load_canonical(manifest_path, "staged fixture manifest")
    rows = fixture_value.get("fixtures")
    _expect(isinstance(rows, list) and len(rows) == binding.get("rows"), "staged fixture rows changed")
    for index, row in enumerate(rows):
        _expect(isinstance(row, Mapping), f"staged fixture row {index} is malformed")
        relative = module._safe_relative(row.get("relative_path"), f"staged fixture row {index} path")
        fixture = root / relative
        _reject_symlink_components(fixture, f"staged fixture row {index}")
        _expect(fixture.is_file() and not fixture.is_symlink(), f"staged fixture row {index} is missing")
        _expect(_sha256(fixture.read_bytes()) == row.get("input_sha256"), f"staged fixture row {index} changed")
    return binding


def _revalidate_preparation(prepared_root: Path) -> dict[str, Any]:
    prepared = _safe_external(prepared_root, "prepared output")
    _expect(prepared.is_dir() and not prepared.is_symlink(), f"prepared output is not a directory: {prepared}")
    receipt, receipt_artifact = _load_canonical(prepared / "preparation-receipt.json", "preparation receipt")
    _expect(receipt.get("schema") == PREPARATION_SCHEMA, "preparation receipt schema differs")
    _expect(receipt.get("status") == "prepared" and receipt.get("execution_admitted") is True, "preparation receipt is not executable")
    module = _policy()
    _expect(receipt.get("formats") == list(module.NATIVE_FORMATS), "preparation format order differs")
    _expect(receipt.get("routes") == list(module.ROUTES), "preparation route matrix differs")
    _expect(receipt.get("signed_batch") == module.SIGNED_BATCH_POLICY, "preparation signed-batch policy differs")
    _expect(receipt.get("policy") == _policy_identity(), "policy artifact changed after preparation")
    _expect(receipt.get("candidate_binding_sha256") == _sha256(_canonical(module.CANDIDATE_BINDING)), "candidate binding changed after preparation")
    source = _source_identity_from_preparation(prepared, receipt)
    fixture = _verify_fixture_binding(prepared, receipt)
    adapter = receipt.get("adapter")
    _expect(isinstance(adapter, Mapping), "preparation adapter is missing")
    _verify_artifact(adapter.get("module"), "consumer adapter")
    _verify_artifact(adapter.get("probe"), "generated signed probe")
    _verify_artifact(adapter.get("manifest"), "consumer manifest")
    plan_artifact = receipt.get("plan")
    _verify_artifact(plan_artifact, "preparation plan")
    plan, actual_plan = _load_canonical(Path(str(plan_artifact["path"])), "preparation plan")
    _expect(actual_plan == dict(plan_artifact), "preparation plan artifact changed")
    _expect(plan.get("source") == receipt.get("source"), "plan source closure differs")
    probes = receipt.get("probes")
    _expect(isinstance(probes, Mapping) and probes.get("schema") == "litchi.signed-native.probes.v2", "probe receipt schema differs")
    return {
        "prepared": prepared,
        "receipt": receipt,
        "receipt_artifact": receipt_artifact,
        "source": source,
        "fixture": fixture,
        "adapter": adapter,
        "plan": plan,
    }


def _directory_bytes(path: Path) -> int:
    total = 0
    for current, directories, files in os.walk(path, topdown=True, followlinks=False):
        for name in directories:
            _expect(not (Path(current) / name).is_symlink(), f"directory contains a symlink: {Path(current) / name}")
        for name in files:
            child = Path(current) / name
            _expect(not child.is_symlink() and child.is_file(), f"directory contains a non-regular file: {child}")
            total += child.stat().st_size
    return total


def build(prepared_root: Path, cargo: Path | None) -> dict[str, Any]:
    """Build the generated consumer with the caller's pinned Cargo shim."""

    state = _revalidate_preparation(prepared_root)
    prepared = state["prepared"]
    _expect(cargo is not None, "build requires the approved pinned cargo shim")
    cargo_path = _safe_external(Path(cargo), "cargo shim")
    cargo_artifact = _artifact(cargo_path, "cargo shim", executable=True)
    project = prepared / "consumer"
    target = prepared / "target"
    temporary = prepared / "tmp"
    _expect(not target.exists() and not target.is_symlink(), f"refusing to reuse target directory: {target}")
    _expect(not temporary.exists() and not temporary.is_symlink(), f"refusing to reuse TMPDIR: {temporary}")
    target.mkdir()
    temporary.mkdir()
    fixture_binding = state["fixture"]
    fixture_root = str(fixture_binding["root"])
    fixture_manifest = fixture_binding["manifest"]
    environment = dict(os.environ)
    environment.update(
        {
            "CARGO_TARGET_DIR": str(target),
            "TMPDIR": str(temporary),
            "CARGO_BUILD_JOBS": "1",
            "CARGO_INCREMENTAL": "0",
            "CARGO_NET_OFFLINE": "true",
            "CARGO_TERM_COLOR": "never",
            "LITCHI_MEM_MAX": "20G",
            "LITCHI_SIGNED_FIXTURE_ROOT": fixture_root,
            "LITCHI_SIGNED_FIXTURE_MANIFEST": str(fixture_manifest["path"]),
            "LITCHI_SIGNED_FIXTURE_MANIFEST_SHA256": str(fixture_manifest["sha256"]),
        }
    )
    commands = [
        [str(cargo_path), "generate-lockfile", "--offline"],
        [str(cargo_path), "metadata", "--offline", "--locked", "--format-version=1"],
        [str(cargo_path), "build", "--offline", "--locked", "--release", "--bin", "signed-native-consumer"],
    ]
    try:
        subprocess.run(commands[0], cwd=project, env=environment, check=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
        metadata = subprocess.run(commands[1], cwd=project, env=environment, check=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True).stdout
        metadata_value = json.loads(metadata)
        _expect(isinstance(metadata_value, dict), "cargo metadata did not return an object")
        (prepared / "metadata.json").write_bytes(_canonical(metadata_value))
        subprocess.run(commands[2], cwd=project, env=environment, check=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
    except (OSError, subprocess.CalledProcessError, json.JSONDecodeError) as error:
        raise ConsumerError(f"bounded signed-native consumer build failed: {error}") from error
    binary = target / "release" / "signed-native-consumer"
    binary_artifact = _artifact(binary, "signed-native consumer binary", executable=True)
    target_bytes = _directory_bytes(target)
    tmp_bytes = _directory_bytes(temporary)
    _expect(target_bytes < TARGET_CAP_BYTES, f"target exceeds 20 GB cap: {target_bytes}")
    _expect(tmp_bytes < TMP_CAP_BYTES, f"TMPDIR exceeds 20 GB cap: {tmp_bytes}")
    receipt = {
        "schema": BUILD_SCHEMA,
        "status": "built",
        "preparation_receipt": state["receipt_artifact"],
        "source_identity": _concrete_source_identity(state),
        "source_identity_receipt": state["receipt"]["source_identity"],
        "source": state["receipt"]["source"],
        "source_archive": state["receipt"]["source"]["archive"],
        "lock": state["receipt"]["source"]["locks"],
        "adapter": state["adapter"],
        "fixture": state["receipt"]["fixture"],
        "routes": state["receipt"]["routes"],
        "signed_batch": state["receipt"]["signed_batch"],
        "cargo": cargo_artifact,
        "binary": binary_artifact,
        "metadata": _artifact(prepared / "metadata.json", "cargo metadata"),
        "command": commands[2],
        "environment": {key: environment[key] for key in (
            "CARGO_TARGET_DIR", "TMPDIR", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL",
            "CARGO_NET_OFFLINE", "CARGO_TERM_COLOR", "LITCHI_MEM_MAX",
            "LITCHI_SIGNED_FIXTURE_ROOT", "LITCHI_SIGNED_FIXTURE_MANIFEST",
            "LITCHI_SIGNED_FIXTURE_MANIFEST_SHA256",
        )},
        "target_bytes": target_bytes,
        "tmp_bytes": tmp_bytes,
        "timing": None,
    }
    receipt_path = prepared / "build-receipt.json"
    _write_new(receipt_path, receipt, "build receipt")
    return {"status": "built", "receipt": str(receipt_path), "binary": binary_artifact}


def _route_rows(value: Mapping[str, Any]) -> dict[str, Mapping[str, Any]]:
    module = _policy()
    routes = value.get("routes")
    if isinstance(routes, Mapping):
        # Canonical JSON sorts object keys.  Route maps are therefore
        # validated as an exact set; the ordered policy matrix remains bound
        # by the preparation receipt and by the emitted route_receipts list.
        _expect(set(routes) == set(module.ROUTES), "capability route order or set differs")
        result = {str(route): {"route": str(route), "status": routes[route]} for route in module.ROUTES}
    elif isinstance(routes, list):
        _expect(len(routes) == len(module.ROUTES), "capability route count differs")
        result = {}
        for row in routes:
            _expect(isinstance(row, Mapping) and isinstance(row.get("route"), str), "capability route row is malformed")
            _expect(row["route"] not in result, "capability route is duplicated")
            result[str(row["route"])] = row
        _expect(tuple(result) == module.ROUTES, "capability route order or set differs")
    else:
        raise ConsumerError("capability receipt routes are missing")
    for route, row in result.items():
        status = row.get("status")
        _expect(status in _ALLOWED_ROUTE_STATUS, f"capability route status is invalid: {route}")
    return result


def _validate_batch(value: Mapping[str, Any]) -> dict[str, Any]:
    module = _policy()
    batch = value.get("signed_batch")
    _expect(isinstance(batch, Mapping), "capability receipt lacks signed-batch evidence")
    _expect(dict(batch).get("selection") == module.SIGNED_BATCH_POLICY["selection"], "signed-batch selection differs")
    _expect(dict(batch).get("status") == "expected_refusal", "signed-batch did not refuse immediately")
    _expect(dict(batch).get("iterator_reads") == 0, "signed-batch refusal touched the iterator")
    _expect(dict(batch).get("publication_attempts") == 0, "signed-batch refusal attempted publication")
    return dict(batch)


def execute(
    prepared_root: Path | None = None,
    capabilities: Path | None = None,
    *,
    output: Path | None = None,
) -> dict[str, Any]:
    """Consume a typed v2 capabilities receipt into an execution receipt.

    Native process ownership, timeout, RSS, and child-receipt semantics live in
    the reviewed protocol runner.  The consumer checks that runner's concrete
    artifacts are tied to this preparation/build before publishing an
    execution receipt.
    """

    if prepared_root is None:
        raise ConsumerError("execute requires a prepared v2 consumer output")
    state = _revalidate_preparation(prepared_root)
    prepared = state["prepared"]
    build_path = prepared / "build-receipt.json"
    build_value, build_artifact = _load_canonical(build_path, "build receipt")
    _expect(build_value.get("schema") == BUILD_SCHEMA and build_value.get("status") == "built", "build receipt is not a v2 build")
    _expect(build_value.get("preparation_receipt") == state["receipt_artifact"], "build receipt is not bound to preparation")
    _verify_artifact(build_value.get("binary"), "signed-native consumer binary", executable=True)
    _expect(build_value.get("source_identity") == _concrete_source_identity(state), "build concrete source identity differs")
    _expect(build_value.get("source_identity_receipt") == state["receipt"]["source_identity"], "build source identity receipt differs")
    _expect(build_value.get("source_archive") == state["receipt"]["source"]["archive"], "build source archive differs")
    _expect(build_value.get("lock") == state["receipt"]["source"]["locks"], "build role lock binding differs")
    _expect(build_value.get("adapter") == state["adapter"], "build adapter differs")
    _expect(build_value.get("routes") == state["receipt"]["routes"], "build route matrix differs")
    _expect(capabilities is not None, "execute requires the protocol capabilities receipt")
    capability_value, capability_artifact = _load_canonical(Path(capabilities), "capabilities receipt")
    _expect(capability_value.get("schema") == CAPABILITIES_SCHEMA, "capabilities receipt schema differs")
    _expect(capability_value.get("preparation_receipt") == state["receipt_artifact"], "capabilities are not bound to preparation")
    _expect(capability_value.get("build_receipt") == build_artifact, "capabilities are not bound to build receipt")
    _expect(capability_value.get("source_identity") == _concrete_source_identity(state), "capabilities concrete source identity differs")
    _expect(capability_value.get("source_identity_receipt") == state["receipt"]["source_identity"], "capabilities source identity receipt differs")
    _expect(capability_value.get("source") == state["receipt"]["source"], "capabilities source closure differs")
    _expect(capability_value.get("source_archive") == state["receipt"]["source"]["archive"], "capabilities source archive differs")
    _expect(capability_value.get("lock") == state["receipt"]["source"]["locks"], "capabilities role lock binding differs")
    _expect(capability_value.get("binary") == build_value["binary"], "capabilities binary differs")
    _expect(
        capability_value.get("build_receipt_binding") == {
            "preparation_receipt": state["receipt_artifact"],
            "build_receipt": build_artifact,
        },
        "capabilities build receipt binding differs",
    )
    _expect(capability_value.get("adapter") == state["adapter"], "capabilities adapter differs")
    _expect(capability_value.get("fixture") == state["receipt"]["fixture"], "capabilities fixture binding differs")
    rows = _route_rows(capability_value)
    batch = _validate_batch(capability_value)
    for route, row in rows.items():
        evidence = row.get("evidence")
        if evidence is not None:
            _verify_artifact(evidence, f"capability evidence for {route}")
            child, _ = _load_canonical(Path(str(evidence["path"])), f"capability evidence for {route}")
            _expect(child.get("route") == route, f"capability evidence route differs: {route}")
            _expect(child.get("timeout_seconds") == CHILD_TIMEOUT_SECONDS, f"capability timeout differs: {route}")
            _expect(child.get("exit_code") == 0 and child.get("timed_out") is False, f"capability route did not exit cleanly: {route}")
            _expect(child.get("guard_error") is None and child.get("after_guard_error") is None, f"capability guard failed: {route}")
    execution = {
        "schema": EXECUTION_SCHEMA,
        "status": "completed",
        "preparation_receipt": state["receipt_artifact"],
        "build_receipt": build_artifact,
        "capabilities_receipt": capability_artifact,
        "source_identity": _concrete_source_identity(state),
        "source_identity_receipt": state["receipt"]["source_identity"],
        "source": state["receipt"]["source"],
        "source_archive": state["receipt"]["source"]["archive"],
        "lock": state["receipt"]["source"]["locks"],
        "adapter": state["adapter"],
        "binary": build_value["binary"],
        "build_receipt_binding": {
            "preparation_receipt": state["receipt_artifact"],
            "build_receipt": build_artifact,
        },
        "fixture": state["receipt"]["fixture"],
        "formats": state["receipt"]["formats"],
        "routes": state["receipt"]["routes"],
        "route_receipts": list(rows.values()),
        "signed_batch": batch,
        "protocol": {
            "schema": capability_value.get("protocol_schema", CAPABILITIES_SCHEMA),
            "timeout_seconds": CHILD_TIMEOUT_SECONDS,
            "output_limit_bytes": CHILD_OUTPUT_LIMIT,
        },
    }
    destination = _safe_external(output or (prepared / "execution-receipt.json"), "execution receipt")
    _expect(not destination.exists() and not destination.is_symlink(), f"refusing to overwrite execution receipt: {destination}")
    _write_new(destination, execution, "execution receipt")
    return {"status": "completed", "receipt": str(destination), "routes": dict(rows)}


def main(argv: Sequence[str]) -> int:
    import argparse

    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    prepare_parser = sub.add_parser("prepare")
    prepare_parser.add_argument("--source-root", type=Path, required=True)
    prepare_parser.add_argument("--output-root", type=Path, required=True)
    build_parser = sub.add_parser("build")
    build_parser.add_argument("--prepared", type=Path, required=True)
    build_parser.add_argument("--cargo", type=Path, required=True)
    execute_parser = sub.add_parser("execute")
    execute_parser.add_argument("--prepared", type=Path, required=True)
    execute_parser.add_argument("--capabilities", type=Path, required=True)
    execute_parser.add_argument("--output", type=Path)
    args = parser.parse_args(list(argv))
    if args.command == "prepare":
        value = prepare(args.source_root, args.output_root)
    elif args.command == "build":
        value = build(args.prepared, args.cargo)
    else:
        value = execute(args.prepared, args.capabilities, output=args.output)
    print(json.dumps(value, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(__import__("sys").argv[1:]))
