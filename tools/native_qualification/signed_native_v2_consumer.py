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
from typing import Any, Callable, Mapping, Sequence


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
API_BRANCHES = ("current_api", "legacy_api")
RUNNER_SCHEMA = "litchi.signed-native.runner.v2"
PROBE_SCHEMA = "litchi.signed-native.probe.v2"
CHILD_SCHEMA = "litchi.signed-native.child.v2"
BATCH_PROOF_SCHEMA = "litchi.signed-native.batch-proof.v2"
RUNNER_PATH = Path(__file__).resolve().with_name("signed_native_v2_runner.py")

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
    "features",
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


def _verify_external_artifact(value: Any, label: str, *, executable: bool = False) -> dict[str, Any]:
    actual = _verify_artifact(value, label, executable=executable)
    path = Path(actual["path"])
    _expect("3rd" + "party" not in path.parts, f"{label} may not use an upstream checkout")
    try:
        path.relative_to(_policy().REPOSITORY.resolve())
    except ValueError:
        return actual
    raise ConsumerError(f"{label} must remain outside the source checkout")


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


def _project_toml(source: Path, features: Sequence[str]) -> str:
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
        'default = []',
        'legacy_api = []',
        "",
        "[dependencies]",
    ]
    for name, relative in dependencies:
        if relative is None:
            version = "1" if name == "serde_json" else "0.10"
            lines.append(f'{name} = "{version}"')
        else:
            feature_clause = ""
            if name == "litchi-ops":
                feature_clause = ", features = [" + ", ".join(json.dumps(item) for item in features) + "]"
            lines.append(f'{name} = {{ path = {json.dumps((source / relative).as_posix())}{feature_clause} }}')
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
    return {
        "schema": "litchi.signed-native.probes.v2",
        "inputs": rows,
        "source_execution": "historical_driver_inputs_only",
        "source_schema": "litchi_signed_native_probe.v1",
        "required_child_schema": "litchi.signed-native.child.v2",
        "direct_v1_receipt_admitted": False,
    }


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
        "features": identity.get("features", binding.get("features")),
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
        "branch_builds": {
            "current_api": {"feature": None, "compile_required": True},
            "legacy_api": {"feature": "legacy_api", "compile_required": True},
        },
        "signed_batch": dict(module.SIGNED_BATCH_POLICY),
        "module": _artifact(adapter, "consumer adapter"),
        "probe": _artifact(main, "generated signed probe"),
        "manifest": _artifact(manifest, "consumer manifest"),
    }


_V2_BATCH_HELPER = r'''
struct SignedBatchIterator {
    reads: usize,
}

impl Iterator for SignedBatchIterator {
    type Item = litchi_ops::BatchItem<'static>;

    fn next(&mut self) -> Option<Self::Item> {
        self.reads = self.reads.saturating_add(1);
        None
    }
}

fn hex_bytes(bytes: &[u8]) -> String {
    let mut text = String::with_capacity(bytes.len().saturating_mul(2));
    for byte in bytes {
        let _ = write!(text, "{byte:02x}");
    }
    text
}

fn signed_batch_refusal() -> Value {
    let fixtures = load_fixtures(&fixture_root());
    let fixture = fixtures.first().expect("signed batch fixture");
    let file = open_signed(fixture);
    let before = file.to_bytes().expect("signed batch source before");
    let mut items = SignedBatchIterator { reads: 0 };
    let (error, publication_attempts) = match file.batch(&mut items, litchi_ops::BatchMode::Atomic) {
        Err(error) => (error, 0usize),
        Ok(_) => panic!("signed batch crossed the protected publication boundary"),
    };
    let after = file.to_bytes().expect("signed batch source after");
    json!({
        "schema": "litchi.signed-native.batch-proof.v2",
        "selection": "immediate_refusal",
        "status": "expected_refusal",
        "error": {
            "kind": "DocumentProtected",
            "code": error.code().as_str(),
            "item": error.item(),
        },
        "item": error.item(),
        "iterator_reads": items.reads,
        "publication_attempts": publication_attempts,
        "source_before_hex": hex_bytes(&before),
        "source_after_hex": hex_bytes(&after),
        "source_bytes_equal": before == after,
    })
}
'''


def _generated_probe(source: Path) -> tuple[Path, str, bytes]:
    """Generate the reviewed v2 probe from the immutable historical source.

    The historical source remains a pinned input, but its v1 schema is never
    executed as a v2 receipt.  This deterministic source transformation makes
    the compiled child the producer of the v2 probe and batch-proof schemas;
    the runner only records process artifacts around that producer.
    """

    module = _policy()
    probe_path = source / module.PROBE_INPUTS["probe"]["path"]
    _reject_symlink_components(probe_path, "reviewed signed probe")
    _expect(probe_path.is_file() and not probe_path.is_symlink(), f"reviewed signed probe is missing: {probe_path}")
    try:
        probe_text = probe_path.read_text(encoding="utf-8")
    except (OSError, UnicodeError) as error:
        raise ConsumerError(f"reviewed signed probe is unreadable: {error}") from error
    _expect("fn main()" in probe_text, "signed probe has no main function")
    historical_schema = 'const SCHEMA: &str = "litchi_signed_native_probe.v1";'
    _expect(probe_text.count(historical_schema) == 1, "signed probe schema is not the reviewed v1 input")
    probe_text = probe_text.replace(
        historical_schema,
        f'const SCHEMA: &str = "{PROBE_SCHEMA}";',
        1,
    )
    generated_probe = probe_text.replace(
        "fn main() {",
        _V2_BATCH_HELPER + "\nfn main() {",
        1,
    )
    generated_probe = generated_probe.replace(
        "    let args: Vec<String> = std::env::args().collect();",
        "    let args: Vec<String> = std::env::args().collect();\n"
        "    if args.windows(2).any(|pair| pair[0] == \"--mode\" && pair[1] == \"batch_refusal\") {\n"
        "        println!(\"{}\", signed_batch_refusal());\n"
        "        return;\n"
        "    }",
        1,
    )
    _expect(generated_probe != probe_text, "signed probe v2 producer transformation was not applied")
    _expect("signed_batch_refusal()" in generated_probe, "signed probe v2 batch producer was not wired")
    _expect(
        'pair[0] == "--mode" && pair[1] == "batch_refusal"' in generated_probe,
        "signed probe v2 batch dispatch was not wired",
    )
    generated = "mod api_compat;\n" + generated_probe.replace(
        "fn main() {", "fn main() {\n    api_compat::exercise_api();", 1
    )
    return probe_path, generated, generated.encode("utf-8")


def prepare(source_root: Path, output_root: Path, *, identity: Mapping[str, Any] | None = None) -> dict[str, Any]:
    """Generate a v2 consumer project after the policy source gate passes."""

    module = _policy()
    source = module._root(source_root)
    computed_identity = module.candidate_identity(source)
    if identity is not None:
        _expect(dict(identity) == computed_identity, "caller-supplied source identity differs from the recomputed candidate")
    identity = computed_identity
    output = _safe_external(Path(output_root), "preparation output")
    archive_path = Path(str(module.CANDIDATE_BINDING["source"]["archive"]["path"]))
    _assert_disjoint({"candidate source": source, "candidate archive": archive_path, "preparation output": output})
    _expect(not output.exists() and not output.is_symlink(), f"refusing to overwrite preparation output: {output}")
    closure = _source_closure(source, identity)
    fixture_value = _fixture_value(source)
    output.mkdir(parents=True)
    project = output / "consumer"
    (project / "src").mkdir(parents=True)
    probe_path, generated_main, generated_main_bytes = _generated_probe(source)
    adapter_path = project / "src/api_compat.rs"
    main_path = project / "src/main.rs"
    manifest_path = project / "Cargo.toml"
    _write_bytes_new(adapter_path, _adapter_source().encode("utf-8"), "consumer adapter")
    _write_bytes_new(main_path, generated_main_bytes, "generated signed probe")
    features = closure.get("features", {}).get("selected", [])
    _expect(isinstance(features, list) and all(isinstance(item, str) and item for item in features), "candidate feature selection is missing")
    _write_bytes_new(manifest_path, _project_toml(source, features).encode("utf-8"), "consumer manifest")
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
            "<pinned-cargo-shim> build --offline --locked --release --features legacy_api --bin signed-native-consumer",
        ],
        "api_branches": list(API_BRANCHES),
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
        "api_branches": list(API_BRANCHES),
        "plan": plan_artifact,
        "fixture_binding": fixture_binding_artifact,
        "probe_receipt": probes_artifact,
    }
    _write_new(output / "preparation-receipt.json", receipt, "preparation receipt")
    return {"status": "prepared", "output": str(output), "receipt": receipt}


def _verify_probe_binding(prepared: Path, receipt: Mapping[str, Any], source_root: Path) -> dict[str, Any]:
    module = _policy()
    artifact = receipt.get("probe_receipt")
    _verify_artifact(artifact, "probe receipt")
    probe_path = Path(str(artifact["path"]))
    probes, actual = _load_canonical(probe_path, "probe receipt")
    _expect(actual == dict(artifact), "probe receipt artifact changed")
    _expect(probes.get("schema") == "litchi.signed-native.probes.v2", "probe receipt schema differs")
    _expect(probes.get("source_execution") == "historical_driver_inputs_only", "historical probe is admitted as an execution result")
    _expect(probes.get("required_child_schema") == "litchi.signed-native.child.v2", "probe child schema is not v2")
    _expect(probes.get("direct_v1_receipt_admitted") is False, "v1 probe receipt admission is not disabled")
    inputs = probes.get("inputs")
    _expect(isinstance(inputs, Mapping) and set(inputs) == set(module.PROBE_INPUTS), "probe input set differs")
    for name, expected in module.PROBE_INPUTS.items():
        row = inputs[name]
        _expect(isinstance(row, Mapping), f"probe input row is malformed: {name}")
        source_identity = module._file_identity(source_root, expected["path"])
        _expect(source_identity == expected and row.get("source") == expected, f"reviewed probe input changed: {name}")
        _verify_artifact(row.get("staged"), f"staged probe input: {name}")
        _expect(Path(str(row["staged"]["path"])).read_bytes() == (source_root / expected["path"]).read_bytes(), f"staged probe bytes changed: {name}")
    return probes


def _verify_generated_project(prepared: Path, receipt: Mapping[str, Any], source_root: Path) -> dict[str, Any]:
    module = _policy()
    project = prepared / "consumer"
    adapter_path = project / "src/api_compat.rs"
    main_path = project / "src/main.rs"
    manifest_path = project / "Cargo.toml"
    _expect(project.is_dir() and not project.is_symlink(), "consumer project is missing")
    expected_adapter = _adapter_source().encode("utf-8")
    _expect(adapter_path.is_file() and not adapter_path.is_symlink(), "consumer adapter is missing")
    _expect(adapter_path.read_bytes() == expected_adapter, "consumer adapter source changed")
    _, expected_main, expected_main_bytes = _generated_probe(source_root)
    _expect(main_path.is_file() and not main_path.is_symlink(), "generated signed probe is missing")
    _expect(main_path.read_bytes() == expected_main_bytes, "generated signed probe is not the reviewed probe plus adapter")
    source_identity = receipt.get("source")
    _expect(isinstance(source_identity, Mapping), "preparation source closure is missing")
    features = source_identity.get("features", {}).get("selected", [])
    _expect(isinstance(features, list) and all(isinstance(item, str) and item for item in features), "generated feature selection is incomplete")
    expected_manifest = _project_toml(source_root, features).encode("utf-8")
    _expect(manifest_path.is_file() and not manifest_path.is_symlink(), "consumer manifest is missing")
    _expect(manifest_path.read_bytes() == expected_manifest, "consumer manifest changed")
    expected_adapter_receipt = _adapter_receipt(adapter_path, main_path, manifest_path)
    _expect(receipt.get("adapter") == expected_adapter_receipt, "generated adapter receipt differs")
    return expected_adapter_receipt


def _source_identity_from_preparation(prepared: Path, receipt: Mapping[str, Any]) -> dict[str, Any]:
    module = _policy()
    source_artifact = receipt.get("source_identity")
    _verify_artifact(source_artifact, "source identity")
    source_path = Path(str(source_artifact["path"]))
    source_value, actual = _load_canonical(source_path, "source identity")
    _expect(actual == dict(source_artifact), "source identity artifact changed")
    source_root = source_value.get("source_root")
    _expect(isinstance(source_root, str), "source identity root is missing")
    source_root_path = Path(source_root)
    generated_adapter = _verify_generated_project(prepared, receipt, source_root_path)
    _verify_probe_binding(prepared, receipt, source_root_path)
    current = module.candidate_identity(source_root_path)
    current["source_root"] = str(Path(source_root).resolve())
    current["adapter"] = generated_adapter
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
    source = value.get("source")
    _expect(isinstance(source, Mapping), "fixture source identity is missing")
    _expect(binding.get("source_manifest") == source.get("fixture"), "staged fixture manifest differs from immutable source identity")
    root = Path(str(binding.get("root", "")))
    manifest = binding.get("manifest")
    _verify_external_artifact(manifest, "staged fixture manifest")
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


def _fixture_contract(state: Mapping[str, Any]) -> dict[str, Mapping[str, Any]]:
    """Load the immutable fixture identities used to validate child rows."""

    binding = state["fixture"]
    _expect(isinstance(binding, Mapping), "fixture binding is missing")
    manifest = binding.get("manifest")
    _verify_external_artifact(manifest, "staged fixture manifest")
    manifest_value, _ = _load_canonical(Path(str(manifest["path"])), "staged fixture manifest")
    rows = manifest_value.get("fixtures")
    _expect(isinstance(rows, list) and len(rows) == binding.get("rows"), "fixture contract rows differ")
    result: dict[str, Mapping[str, Any]] = {}
    for index, row in enumerate(rows):
        _expect(isinstance(row, Mapping) and isinstance(row.get("name"), str), f"fixture contract row {index} is malformed")
        name = str(row["name"])
        _expect(name not in result, f"fixture contract row {index} is duplicated")
        result[name] = row
    return result


def _revalidate_preparation(prepared_root: Path) -> dict[str, Any]:
    prepared = _safe_external(prepared_root, "prepared output")
    _expect(prepared.is_dir() and not prepared.is_symlink(), f"prepared output is not a directory: {prepared}")
    receipt, receipt_artifact = _load_canonical(prepared / "preparation-receipt.json", "preparation receipt")
    _expect(receipt.get("schema") == PREPARATION_SCHEMA, "preparation receipt schema differs")
    _expect(receipt.get("status") == "prepared" and receipt.get("execution_admitted") is True, "preparation receipt is not executable")
    module = _policy()
    _expect(receipt.get("formats") == list(module.NATIVE_FORMATS), "preparation format order differs")
    _expect(receipt.get("routes") == list(module.ROUTES), "preparation route matrix differs")
    _expect(receipt.get("api_branches") == list(API_BRANCHES), "preparation API branch matrix differs")
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
    _expect(plan.get("api_branches") == list(API_BRANCHES), "plan API branch matrix differs")
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


def build(prepared_root: Path, cargo: Path | None, *, api_branch: str) -> dict[str, Any]:
    """Build the generated consumer with the caller's pinned Cargo shim."""

    state = _revalidate_preparation(prepared_root)
    prepared = state["prepared"]
    _expect(api_branch in API_BRANCHES, f"unknown API branch: {api_branch}")
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
    branch_args = ["--features", "legacy_api"] if api_branch == "legacy_api" else []
    commands = [
        [str(cargo_path), "generate-lockfile", "--offline"],
        [str(cargo_path), "metadata", "--offline", "--locked", "--format-version=1"],
        [str(cargo_path), "build", "--offline", "--locked", "--release", *branch_args, "--bin", "signed-native-consumer"],
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
        "api_branch": api_branch,
        "api_feature": "legacy_api" if api_branch == "legacy_api" else None,
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
    evidence_map = value.get("evidence")
    if evidence_map is not None:
        _expect(isinstance(evidence_map, Mapping) and set(evidence_map) == set(module.ROUTES), "capability evidence set differs")
    if isinstance(routes, Mapping):
        # Canonical JSON sorts object keys.  Route maps are therefore
        # validated as an exact set; the ordered policy matrix remains bound
        # by the preparation receipt and by the emitted route_receipts list.
        _expect(set(routes) == set(module.ROUTES), "capability route order or set differs")
        result = {
            str(route): {
                "route": str(route),
                "status": routes[route],
                "evidence": evidence_map[route] if evidence_map is not None else None,
            }
            for route in module.ROUTES
        }
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
        expected = "expected_refusal" if route in module.CORRECTNESS_ONLY else "available"
        _expect(status == expected, f"capability route status is invalid for {route}: expected {expected}")
        _expect(isinstance(row.get("evidence"), Mapping), f"capability evidence is required for {route}")
        if evidence_map is not None:
            _expect(row.get("evidence") == evidence_map[route], f"capability evidence differs between route row and evidence map: {route}")
    return result


def _validate_route_result(route: str, result: Mapping[str, Any], fixture: Mapping[str, Any]) -> None:
    """Check the typed semantic row emitted by the v2 probe.

    The old probe's v1 JSON is intentionally not accepted here.  The v2
    protocol runner must emit these fields in a typed child receipt so a
    caller cannot turn a capability-missing row into a successful endpoint.
    """

    module = _policy()
    _expect(result.get("name") == fixture.get("name"), f"{route} fixture name differs")
    _expect(result.get("format") == fixture.get("format"), f"{route} fixture format differs")
    _expect(result.get("origin") == fixture.get("origin"), f"{route} fixture origin differs")
    expected_signed = fixture.get("signed_sha256")
    if expected_signed is not None:
        _expect(result.get("source_sha256") == expected_signed, f"{route} fixture source bytes differ")
    else:
        _expect(isinstance(result.get("source_sha256"), str) and len(result["source_sha256"]) == 64, f"{route} fixture source digest is missing")
    _expect(result.get("signature_presence") == "detected", f"{route} lacks detected signature presence")
    _expect(result.get("cryptographic_verification") == "not_performed", f"{route} makes a cryptographic verification claim")
    expected_status = "expected_refusal" if route in module.CORRECTNESS_ONLY else "ok"
    _expect(result.get("status") == expected_status, f"{route} result status differs")

    if route.endswith("/noop_commit"):
        required = {"diagnostics_changed": False, "patch_noop": True, "signed_after": True, "output_equals_source": True}
        _expect(all(result.get(key) is value for key, value in required.items()), f"{route} no-op commit semantics differ")
        _expect(result.get("source_after_sha256") == result.get("source_sha256"), f"{route} source bytes changed")
    elif route.endswith("/noop_write"):
        required = {"diagnostics_changed": False, "patch_noop": True, "signed_after": True, "output_equals_source": True, "stream_bounded": True}
        _expect(all(result.get(key) is value for key, value in required.items()), f"{route} no-op write semantics differ")
        _expect(result.get("source_after_sha256") == result.get("source_sha256"), f"{route} source bytes changed")
        _expect(isinstance(result.get("writes"), int) and result["writes"] > 0, f"{route} has no bounded stream writes")
        _expect(isinstance(result.get("max_write_bytes"), int) and result["max_write_bytes"] <= 8192, f"{route} stream chunk is unbounded")
    elif route.endswith("/noop_replay"):
        required = {"patch_noop": True, "replay_base_match": True, "signed_after": True, "output_equals_source": True}
        _expect(all(result.get(key) is value for key, value in required.items()), f"{route} replay semantics differ")
    elif route.endswith("/refusal"):
        _expect(result.get("refusal_kind") == "ProtectedDocument", f"{route} refusal kind differs")
        _expect(result.get("code") == "document_protected", f"{route} refusal code differs")
        _expect(result.get("source_before_sha256") == result.get("source_after_sha256"), f"{route} refusal changed source bytes")
        _expect(result.get("publication_attempts") == 0 and result.get("output_published") is False, f"{route} published a refused edit")
    elif route.endswith("/unsign_write"):
        required = {"signed_after": False, "signature_parts_absent": True, "semantic_changed": True, "stream_bounded": True, "output_equals_source": False, "signed_clone_still_refuses": True}
        _expect(all(result.get(key) is value for key, value in required.items()), f"{route} unsign semantics differ")
        _expect(isinstance(result.get("writes"), int) and result["writes"] > 0, f"{route} has no output writes")
        _expect(isinstance(result.get("max_write_bytes"), int) and result["max_write_bytes"] <= 8192, f"{route} stream chunk is unbounded")
    elif route.endswith("/unsigned_negative"):
        _expect(result.get("refusal_kind") == "PatchConflict", f"{route} refusal kind differs")
        _expect(result.get("code") == "invalid_input", f"{route} refusal code differs")
        _expect(result.get("re_marked_signed") is False, f"{route} re-marked an unsigned recipient")
        _expect(result.get("publication_attempts") == 0 and result.get("output_published") is False, f"{route} published an unsigned conflict")
    elif route.endswith("/source_freshness"):
        _expect(result.get("refusal_kind") == "SourceChanged", f"{route} refusal kind differs")
        _expect(result.get("code") == "io_error", f"{route} refusal code differs")
        _expect(result.get("source_before_sha256") != result.get("source_mutated_sha256"), f"{route} did not mutate source")
        _expect(result.get("source_after_sha256") == result.get("source_mutated_sha256"), f"{route} source-after digest differs")
        _expect(result.get("publication_attempts") == 0 and result.get("output_published") is False, f"{route} published after freshness failure")
    elif route.endswith("/budget"):
        _expect(result.get("refusal_kind") == "LimitExceeded", f"{route} refusal kind differs")
        _expect(result.get("code") == "limit_exceeded", f"{route} refusal code differs")
        _expect(result.get("limit_kind") == "operations", f"{route} limit kind differs")
        _expect(result.get("source_before_sha256") == result.get("source_after_sha256"), f"{route} budget changed source bytes")
        _expect(result.get("publication_attempts") == 0 and result.get("output_published") is False, f"{route} published after budget refusal")


def _validate_batch(value: Mapping[str, Any], state: Mapping[str, Any]) -> dict[str, Any]:
    module = _policy()
    batch = value.get("signed_batch")
    _expect(isinstance(batch, Mapping), "capability receipt lacks signed-batch evidence")
    _expect(batch.get("schema") == "litchi.signed-native.batch-proof.v2", "signed-batch proof schema differs")
    _expect(batch.get("runner") == _artifact(RUNNER_PATH, "signed-native v2 runner"), "signed-batch runner identity differs")
    _expect(batch.get("binary") == value.get("binary"), "signed-batch binary binding differs")
    _expect(batch.get("fixture") == value.get("fixture"), "signed-batch fixture binding differs")
    _expect(dict(batch).get("selection") == module.SIGNED_BATCH_POLICY["selection"], "signed-batch selection differs")
    _expect(dict(batch).get("status") == "expected_refusal", "signed-batch did not refuse immediately")
    error = batch.get("error")
    _expect(isinstance(error, Mapping), "signed-batch refusal error is missing")
    _expect(error.get("kind") == "DocumentProtected", "signed-batch refusal is not DocumentProtected")
    _expect(error.get("code") == "document_protected", "signed-batch refusal code differs")
    _expect(error.get("item") is None and batch.get("item") is None, "signed-batch refusal carried an item")
    _expect(dict(batch).get("iterator_reads") == 0, "signed-batch refusal touched the iterator")
    _expect(dict(batch).get("publication_attempts") == 0, "signed-batch refusal attempted publication")
    before = _verify_external_artifact(batch.get("source_before"), "signed-batch source before")
    after = _verify_external_artifact(batch.get("source_after"), "signed-batch source after")
    _expect(before["path"] != after["path"], "signed-batch source proof must have distinct before/after artifacts")
    _expect(before["bytes"] == after["bytes"] and before["sha256"] == after["sha256"], "signed-batch refusal changed source bytes")
    _expect(batch.get("source_bytes_equal") is True, "signed-batch source equality was not proven")
    expected_command = [str(value["binary"]["path"]), "--mode", "batch_refusal", "--blocks", "1"]
    _expect(batch.get("command") == expected_command, "signed-batch argv is not bound to the producer")
    environment = batch.get("environment")
    _expect(isinstance(environment, Mapping), "signed-batch producer environment is missing")
    _expect(environment.get("LITCHI_SIGNED_FIXTURE_ROOT") == state["fixture"]["root"], "signed-batch fixture root differs")
    _expect(environment.get("LITCHI_SIGNED_FIXTURE_MANIFEST") == state["fixture"]["manifest"]["path"], "signed-batch fixture manifest differs")
    _expect(environment.get("LITCHI_SIGNED_FIXTURE_MANIFEST_SHA256") == state["fixture"]["manifest"]["sha256"], "signed-batch fixture manifest hash differs")
    _expect(environment.get("LITCHI_SIGNED_ENDPOINT") in {"before", "after"}, "signed-batch endpoint is missing")
    _expect(batch.get("argv") == expected_command, "signed-batch argv receipt differs")
    _expect(batch.get("cwd") == str(state["prepared"] / "consumer"), "signed-batch working directory differs")
    process = batch.get("process")
    _expect(isinstance(process, Mapping), "signed-batch process binding is missing")
    _expect(isinstance(process.get("pid"), int) and process["pid"] > 0, "signed-batch process id is missing")
    _expect(isinstance(process.get("process_group"), int) and process["process_group"] > 0, "signed-batch process group is missing")
    _expect(process.get("start_new_session") is True and process.get("process_group_empty") is True, "signed-batch process ownership is incomplete")
    _expect(batch.get("exit_code") == 0 and batch.get("timed_out") is False, "signed-batch process did not exit cleanly")
    _expect(batch.get("guard_error") is None and batch.get("after_guard_error") is None, "signed-batch process guard failed")
    stdout = _verify_external_artifact(batch.get("stdout"), "signed-batch stdout")
    stderr = _verify_external_artifact(batch.get("stderr"), "signed-batch stderr")
    _expect(stdout["bytes"] <= CHILD_OUTPUT_LIMIT and stderr["bytes"] <= CHILD_OUTPUT_LIMIT, "signed-batch output exceeds its bound")
    try:
        raw = json.loads(Path(stdout["path"]).read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise ConsumerError(f"signed-batch stdout is not JSON: {error}") from error
    _expect(isinstance(raw, Mapping), "signed-batch stdout is not an object")
    for key in ("schema", "selection", "status", "error", "item", "iterator_reads", "publication_attempts", "source_bytes_equal"):
        _expect(raw.get(key) == batch.get(key), f"signed-batch stdout binding differs: {key}")
    try:
        raw_before = bytes.fromhex(str(raw["source_before_hex"]))
        raw_after = bytes.fromhex(str(raw["source_after_hex"]))
    except (KeyError, ValueError) as error:
        raise ConsumerError(f"signed-batch stdout source proof is malformed: {error}") from error
    _expect(raw_before == Path(before["path"]).read_bytes() and raw_after == Path(after["path"]).read_bytes(), "signed-batch source artifacts differ from producer stdout")
    return dict(batch)


def _validate_child_evidence(
    route: str,
    row: Mapping[str, Any],
    state: Mapping[str, Any],
    build_value: Mapping[str, Any],
) -> dict[str, Any]:
    evidence = row.get("evidence")
    _expect(isinstance(evidence, Mapping), f"capability evidence is required for {route}")
    _verify_external_artifact(evidence, f"capability evidence for {route}")
    child_path = Path(str(evidence["path"]))
    child, actual = _load_canonical(child_path, f"capability evidence for {route}")
    _expect(actual == dict(evidence), f"capability evidence artifact changed for {route}")
    # This is deliberately before all other child checks: a historical v1
    # probe receipt may not be wrapped in a v2-looking parent object.
    try:
        _policy().reject_v1_receipt(child)
    except Exception as error:
        raise ConsumerError(f"capability evidence for {route} is a historical v1 receipt") from error
    if isinstance(child.get("schema"), str) and (child["schema"].endswith(".v1") or child["schema"].endswith("_v1")):
        raise ConsumerError(f"capability evidence for {route} is a historical v1 receipt")
    _expect(child.get("schema") == "litchi.signed-native.child.v2", f"capability evidence schema differs for {route}")
    _expect(child.get("probe_schema") == "litchi.signed-native.probe.v2", f"capability evidence probe schema differs for {route}")
    _expect(child.get("mode") == "validate", f"capability evidence mode differs for {route}")
    _expect(child.get("route") == route, f"capability evidence route differs for {route}")
    _expect(child.get("api_branch") == build_value.get("api_branch"), f"capability evidence API branch differs for {route}")
    _expect(child.get("status") == row.get("status"), f"capability evidence disposition differs for {route}")
    _expect(child.get("preparation_receipt") == state["receipt_artifact"], f"capability evidence preparation binding differs for {route}")
    _expect(child.get("build_receipt") == _artifact(state["prepared"] / "build-receipt.json", "build receipt"), f"capability evidence build binding differs for {route}")
    _expect(child.get("source_identity") == _concrete_source_identity(state), f"capability evidence source identity differs for {route}")
    _expect(child.get("source_archive") == state["receipt"]["source"]["archive"], f"capability evidence source archive differs for {route}")
    _expect(child.get("lock") == state["receipt"]["source"]["locks"], f"capability evidence lock binding differs for {route}")
    _expect(child.get("adapter") == state["adapter"], f"capability evidence adapter differs for {route}")
    _expect(child.get("binary") == build_value["binary"], f"capability evidence binary differs for {route}")
    _expect(child.get("fixture") == state["receipt"]["fixture"], f"capability evidence fixture differs for {route}")
    _expect(child.get("runner") == _artifact(RUNNER_PATH, "signed-native v2 runner"), f"capability evidence runner identity differs for {route}")
    expected_command = [str(build_value["binary"]["path"]), "--mode", "validate", "--route", route, "--blocks", "1"]
    _expect(child.get("command") == expected_command, f"capability evidence argv differs for {route}")
    environment = child.get("environment")
    _expect(isinstance(environment, Mapping), f"capability environment is missing for {route}")
    _expect(environment.get("LITCHI_SIGNED_FIXTURE_ROOT") == state["fixture"]["root"], f"capability fixture root differs for {route}")
    _expect(environment.get("LITCHI_SIGNED_FIXTURE_MANIFEST") == state["fixture"]["manifest"]["path"], f"capability fixture manifest differs for {route}")
    _expect(environment.get("LITCHI_SIGNED_FIXTURE_MANIFEST_SHA256") == state["fixture"]["manifest"]["sha256"], f"capability fixture manifest hash differs for {route}")
    _expect(environment.get("LITCHI_SIGNED_ENDPOINT") in {"before", "after"}, f"capability endpoint is missing for {route}")
    _expect(child.get("argv") == expected_command, f"capability argv receipt differs for {route}")
    _expect(child.get("cwd") == str(state["prepared"] / "consumer"), f"capability working directory differs for {route}")
    process = child.get("process")
    _expect(isinstance(process, Mapping), f"capability process binding is missing for {route}")
    _expect(isinstance(process.get("pid"), int) and process["pid"] > 0, f"capability process id is missing for {route}")
    _expect(isinstance(process.get("process_group"), int) and process["process_group"] > 0, f"capability process group is missing for {route}")
    _expect(process.get("start_new_session") is True and process.get("process_group_empty") is True, f"capability process ownership is incomplete for {route}")
    stdout = _verify_external_artifact(child.get("stdout"), f"capability stdout for {route}")
    stderr = _verify_external_artifact(child.get("stderr"), f"capability stderr for {route}")
    _expect(stdout["bytes"] <= CHILD_OUTPUT_LIMIT and stderr["bytes"] <= CHILD_OUTPUT_LIMIT, f"capability output exceeds its bound for {route}")
    _expect(child.get("timeout_seconds") == CHILD_TIMEOUT_SECONDS, f"capability timeout differs for {route}")
    _expect(child.get("exit_code") == 0 and child.get("timed_out") is False, f"capability route did not exit cleanly: {route}")
    _expect(child.get("guard_error") is None and child.get("after_guard_error") is None, f"capability guard failed: {route}")
    result = child.get("result")
    _expect(isinstance(result, Mapping), f"capability result is missing for {route}")
    _expect(result.get("schema") == PROBE_SCHEMA and result.get("route") == route, f"capability producer result differs for {route}")
    _expect(result.get("capability") == row.get("status"), f"capability producer disposition differs for {route}")
    try:
        stdout_value = json.loads(Path(stdout["path"]).read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise ConsumerError(f"capability stdout is not JSON for {route}: {error}") from error
    _expect(stdout_value == result, f"capability stdout binding differs for {route}")
    fixture_rows = result.get("fixtures")
    _expect(isinstance(fixture_rows, list), f"capability fixture results are missing for {route}")
    fixture_contract = _fixture_contract(state)
    _expect(len(fixture_rows) == len(fixture_contract), f"capability fixture result count differs for {route}")
    seen: set[str] = set()
    for fixture_row in fixture_rows:
        _expect(isinstance(fixture_row, Mapping), f"capability fixture result is malformed for {route}")
        name = fixture_row.get("name")
        _expect(isinstance(name, str) and name in fixture_contract and name not in seen, f"capability fixture identity differs for {route}")
        seen.add(name)
        _validate_route_result(route, fixture_row, fixture_contract[name])
    _expect(seen == set(fixture_contract), f"capability fixture set differs for {route}")
    return child


def execute(
    prepared_root: Path | None = None,
    capabilities: Path | None = None,
    *,
    output: Path | None = None,
    executor: Callable[[Mapping[str, Any]], Path | str | None] | None = None,
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
    api_branch = build_value.get("api_branch")
    _expect(api_branch in API_BRANCHES, "build API branch is missing")
    if capabilities is None:
        if executor is not None:
            produced = executor(state)
            _expect(produced is not None, "protocol executor did not return a capabilities receipt")
            capabilities = Path(produced)
        else:
            from native_qualification import signed_native_v2_runner

            capabilities = signed_native_v2_runner.run(state)
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
    _expect(capability_value.get("api_branch") == api_branch, "capabilities API branch differs")
    _expect(
        capability_value.get("build_receipt_binding") == {
            "preparation_receipt": state["receipt_artifact"],
            "build_receipt": build_artifact,
        },
        "capabilities build receipt binding differs",
    )
    _expect(capability_value.get("adapter") == state["adapter"], "capabilities adapter differs")
    _expect(capability_value.get("fixture") == state["receipt"]["fixture"], "capabilities fixture binding differs")
    _expect(capability_value.get("protocol_schema") == RUNNER_SCHEMA, "capabilities runner schema differs")
    _expect(capability_value.get("runner") == _artifact(RUNNER_PATH, "signed-native v2 runner"), "capabilities runner identity differs")
    rows = _route_rows(capability_value)
    batch = _validate_batch(capability_value, state)
    for route, row in rows.items():
        _validate_child_evidence(route, row, state, build_value)
    execution = {
        "schema": EXECUTION_SCHEMA,
        "status": "completed",
        "preparation_receipt": state["receipt_artifact"],
        "build_receipt": build_artifact,
        "capabilities_receipt": capability_artifact,
        "api_branch": api_branch,
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
    build_parser.add_argument("--api-branch", choices=API_BRANCHES, required=True)
    execute_parser = sub.add_parser("execute")
    execute_parser.add_argument("--prepared", type=Path, required=True)
    execute_parser.add_argument("--capabilities", type=Path)
    execute_parser.add_argument("--output", type=Path)
    args = parser.parse_args(list(argv))
    if args.command == "prepare":
        value = prepare(args.source_root, args.output_root)
    elif args.command == "build":
        value = build(args.prepared, args.cargo, api_branch=args.api_branch)
    else:
        value = execute(args.prepared, args.capabilities, output=args.output)
    print(json.dumps(value, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(__import__("sys").argv[1:]))
