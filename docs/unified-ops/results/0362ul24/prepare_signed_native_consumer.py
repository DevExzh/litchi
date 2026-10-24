#!/usr/bin/env python3
"""Prepare and build the signed-native consumer outside a source checkout.

The signed qualification binary is an external consumer.  This driver keeps
the two endpoint source trees immutable, makes the f2 API differences explicit
in a tiny compile-time adapter, and writes endpoint manifests only after a
real release binary exists.  ``plan`` and ``prepare`` are source/provenance
operations; ``build`` is deliberately a separate, owner-granted operation.

No generated project, target directory, binary, archive, lockfile, or raw log
may be placed in the repository.  The caller supplies an external output
directory for all of them.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
from typing import Any


HERE = Path(__file__).resolve().parent
PROBE = HERE / "signed_native_probe.rs"
FIXTURES = HERE / "fixtures.json"
F2_COMMIT = "f2bbbc8d07ab5da26c180f1ff74813d39f632d54"
FINAL_COMMIT = "472b6705475b073e7fe08f6e11ee17fb1e29f838"
SOURCE_KINDS = {"f2": F2_COMMIT, "472b": FINAL_COMMIT}
FULL_SHA = re.compile(r"^[0-9a-f]{40}$")
SHA256 = re.compile(r"^[0-9a-f]{64}$")
FEATURES = (
    "doc",
    "docx",
    "ppt",
    "pptx",
    "xls",
    "xlsx",
    "xlsb",
    "opc",
)
ROUTES = (
    "signed_leaf/noop_commit",
    "signed_facade/noop_commit",
    "signed_leaf/noop_write",
    "signed_facade/noop_write",
    "signed_leaf/noop_replay",
    "signed_facade/noop_replay",
    "signed_leaf/refusal",
    "signed_facade/refusal",
    "signed_leaf/unsign_write",
    "signed_facade/unsign_write",
    "signed_leaf/unsigned_negative",
    "signed_facade/unsigned_negative",
    "signed_leaf/source_freshness",
    "signed_facade/source_freshness",
    "signed_leaf/budget",
    "signed_facade/budget",
)
BASELINE_NOT_APPLICABLE = frozenset(
    {
        "signed_leaf/unsigned_negative",
        "signed_facade/unsigned_negative",
        "signed_leaf/source_freshness",
        "signed_facade/source_freshness",
        "signed_leaf/budget",
        "signed_facade/budget",
    }
)
TARGET_CAP_BYTES = 20_000_000_000


class PreparationError(RuntimeError):
    pass


def fail(message: str) -> PreparationError:
    return PreparationError(message)


def canonical(value: object) -> bytes:
    return (
        json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        + "\n"
    ).encode("utf-8")


def digest(path: Path) -> str:
    h = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                h.update(block)
    except OSError as error:
        raise fail(f"cannot read {path}: {error}") from error
    return h.hexdigest()


def artifact(path: Path) -> dict[str, Any]:
    path = path.resolve()
    if not path.is_file() or path.is_symlink():
        raise fail(f"artifact is not a regular non-symlink file: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": digest(path)}


def safe_external(path: Path, label: str) -> Path:
    lexical = path.absolute()
    cursor = lexical
    while cursor != cursor.parent:
        if cursor.is_symlink():
            raise fail(f"{label} may not traverse a symlink: {cursor}")
        cursor = cursor.parent
    path = lexical.resolve()
    if "3rdparty" in path.parts:
        raise fail(f"{label} may not use 3rdparty: {path}")
    # ``HERE`` is ``<checkout>/docs/unified-ops/results/0362ul24``.
    # Worktrees use a `.git` file, so this script must not infer checkout
    # validity from that path; the repository root is the fourth parent.
    repository = HERE.parents[3]
    forbidden = (
        repository / "docs" / "performance" / "results",
        repository / "docs" / "report",
        repository / "binaries",
    )
    for root in forbidden:
        try:
            path.relative_to(root.resolve())
        except ValueError:
            continue
        raise fail(f"{label} must be outside {root}")
    return path


def require_commit(value: str, label: str) -> str:
    if FULL_SHA.fullmatch(value) is None:
        raise fail(f"{label} is not a full commit id")
    return value


def run(
    command: list[str],
    cwd: Path,
    *,
    capture: bool = True,
    environment: dict[str, str] | None = None,
) -> str:
    try:
        result = subprocess.run(
            command,
            cwd=cwd,
            check=True,
            text=True,
            env=environment,
            stdout=subprocess.PIPE if capture else None,
            stderr=subprocess.STDOUT if capture else None,
        )
    except (OSError, subprocess.CalledProcessError) as error:
        raise fail(f"command failed in {cwd}: {' '.join(command)}: {error}") from error
    return result.stdout or ""


def _verified_receipt_artifact(value: object, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise fail(f"{label} artifact is missing")
    path = value.get("path")
    expected_bytes = value.get("bytes")
    expected_sha = value.get("sha256")
    if not isinstance(path, str) or not isinstance(expected_bytes, int) or not isinstance(expected_sha, str):
        raise fail(f"{label} artifact is malformed")
    if SHA256.fullmatch(expected_sha) is None:
        raise fail(f"{label} artifact hash is malformed")
    actual = artifact(Path(path))
    if actual != {"path": str(Path(path).resolve()), "bytes": expected_bytes, "sha256": expected_sha}:
        raise fail(f"{label} artifact changed since its receipt")
    return actual


def source_identity(
    source: Path,
    kind: str,
    *,
    source_archive: Path | None = None,
    source_assembly: Path | None = None,
) -> dict[str, Any]:
    source_lexical = source.absolute()
    if source_lexical.is_symlink():
        raise fail(f"source checkout may not be a symlink: {source}")
    source = source.resolve()
    if source_archive is not None:
        source_archive = safe_external(source_archive, "source archive")
    if source_assembly is not None:
        source_assembly = safe_external(source_assembly, "source assembly")
    expected = SOURCE_KINDS[kind]
    if not source.is_dir() or not (source / "Cargo.toml").is_file():
        raise fail(f"source is not a usable checkout: {source}")
    git_head = subprocess.run(
        ["git", "rev-parse", "HEAD"], cwd=source, text=True,
        stdout=subprocess.PIPE, stderr=subprocess.DEVNULL,
    )
    git_checkout = git_head.returncode == 0
    tree_hash: str | None = None
    assembly_artifact: dict[str, Any] | None = None
    source_archive_artifact: dict[str, Any] | None = None
    source_tree: dict[str, Any] | None = None
    if git_head.returncode == 0:
        actual = git_head.stdout.strip()
        if actual != expected:
            raise fail(f"{kind} source HEAD {actual} differs from required {expected}")
        status = run(["git", "status", "--porcelain=v1"], source)
        if status:
            raise fail(f"{kind} source checkout is dirty:\n{status}")
        tree = subprocess.run(
            ["git", "ls-tree", "-r", "-z", expected],
            cwd=source,
            check=True,
            stdout=subprocess.PIPE,
        ).stdout
        tree_hash = hashlib.sha256(tree).hexdigest()
        source_tree = {
            "files": tree.count(b"\0"),
            "sha256": tree_hash,
            "source_role": "pinned_git_checkout",
        }
        if source_archive is not None:
            source_archive_artifact = artifact(source_archive)
    else:
        if kind != "f2" or source_archive is None or source_assembly is None:
            raise fail("non-git source requires the pinned f2 archive and source-assembly receipt")
        source_archive_artifact = artifact(source_archive)
        assembly_artifact = artifact(source_assembly)
        try:
            assembly = json.loads(source_assembly.read_text(encoding="utf-8"))
        except (OSError, UnicodeError, json.JSONDecodeError) as error:
            raise fail(f"invalid source-assembly receipt: {error}") from error
        if assembly.get("schema") != "litchi_w8_source_assembly.v1":
            raise fail("f2 source-assembly receipt schema mismatch")
        if assembly.get("source_commit") != expected or assembly.get("resolved_commit") != expected:
            raise fail("f2 source-assembly receipt commit mismatch")
        recorded_archive = _verified_receipt_artifact(assembly.get("source_archive"), "f2 source archive")
        if recorded_archive != source_archive_artifact:
            raise fail("f2 source archive differs from source-assembly receipt")
        overlay = assembly.get("overlay")
        if not isinstance(overlay, dict) or overlay.get("kind") != "compatibility":
            raise fail("f2 source-assembly must declare compatibility-only overlay")
        if overlay.get("base_commit") != expected:
            raise fail("f2 compatibility overlay base commit differs")
        patch_value = overlay.get("patch")
        if isinstance(patch_value, dict) and isinstance(patch_value.get("path"), str):
            safe_external(Path(patch_value["path"]), "f2 adapter patch")
        patch = _verified_receipt_artifact(overlay.get("patch"), "f2 adapter patch")
        source_tree = {
            "files": 459676,
            "sha256": "23e332254abb024cbf9b038b8219b861e363fe474423bccc6122c3183b4d1437",
            "source_role": "pristine_f2_before_leg",
            "adapter_patch": patch,
        }
        tree_hash = source_tree["sha256"]
    lock = source / "Cargo.lock"
    toolchain = source / "rust-toolchain.toml"
    cargo_config = source / ".cargo" / "config.toml"
    for path, label in ((lock, "Cargo.lock"), (toolchain, "rust-toolchain.toml"), (cargo_config, ".cargo/config.toml")):
        if not path.is_file():
            raise fail(f"{kind} source lacks {label}")
    return {
        "kind": kind,
        "commit": expected,
        "source_ref": (
            f"git:{expected}"
            if git_checkout
            else f"archive:{source_archive_artifact['sha256']}"
        ),
        "source_root": str(source),
        "tree_sha256": tree_hash,
        "lock": artifact(lock),
        "toolchain": artifact(toolchain),
        "cargo_config": artifact(cargo_config),
        "source_archive": source_archive_artifact,
        "source_assembly": assembly_artifact,
        "source_tree": source_tree,
    }


def fixture_manifest() -> dict[str, Any]:
    if not FIXTURES.is_file():
        raise fail(f"missing signed fixture manifest: {FIXTURES}")
    try:
        value = json.loads(FIXTURES.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise fail(f"invalid signed fixture manifest: {error}") from error
    if not isinstance(value, dict) or value.get("schema") != "litchi_signed_fixture_manifest.v1":
        raise fail("signed fixture manifest schema mismatch")
    rows = value.get("fixtures")
    if not isinstance(rows, list) or not rows:
        raise fail("signed fixture manifest has no fixtures")
    formats: set[str] = set()
    names: set[str] = set()
    for index, row in enumerate(rows):
        if not isinstance(row, dict):
            raise fail(f"fixture row {index} is not an object")
        name = row.get("name")
        relative = row.get("relative_path")
        if not isinstance(name, str) or not name or name in names:
            raise fail(f"fixture row {index} has a duplicate/empty name")
        if not isinstance(relative, str) or not relative:
            raise fail(f"fixture row {index} has no relative path")
        path = Path(relative)
        if path.is_absolute() or ".." in path.parts or "\\" in relative or "3rdparty" in path.parts:
            raise fail(f"fixture row {index} has an unsafe relative path")
        if row.get("origin") not in {"Opc", "Cfb"}:
            raise fail(f"fixture row {index} has an invalid origin")
        if row.get("fixture_mode") not in {"producer_signed", "synthetic_presence"}:
            raise fail(f"fixture row {index} has an invalid fixture mode")
        if row.get("signature_presence") != "detected" or row.get("cryptographic_verification") != "not_performed":
            raise fail(f"fixture row {index} has an invalid signature claim")
        if not isinstance(row.get("input_sha256"), str) or SHA256.fullmatch(row["input_sha256"]) is None:
            raise fail(f"fixture row {index} input hash is invalid")
        formats.add(str(row.get("format")))
        names.add(name)
    if formats != {"DOCX", "PPTX", "XLSX", "XLSB", "DOC", "PPT", "XLS"}:
        raise fail(f"signed fixture formats are incomplete: {sorted(formats)}")
    return value


def adapter_source(kind: str) -> str:
    """Compile-time API bridge checks, kept outside both production trees."""
    if kind == "f2":
        return """//! Consumer-only compatibility assertions for original f2 APIs.
#![allow(dead_code)]
pub fn assert_api() {
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
"""
    return """//! Consumer-only compatibility assertions for the final API.
#![allow(dead_code)]
pub fn assert_api() {
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
        let _ = package.edit_cell_values(0usize);
    }
    let _ = (commit_methods::<()>, edit_entry, xlsb_entry);
}
"""


def project_toml(source: Path, kind: str) -> str:
    features = ", ".join(json.dumps(feature) for feature in FEATURES)
    deps = [
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
    ]
    lines = [
        "[package]",
        'name = "signed-native-consumer"',
        'version = "0.0.0"',
        'edition = "2024"',
        "",
        "[features]",
        'default = ["legacy_api"]' if kind == "f2" else 'default = []',
        'legacy_api = []',
        "",
        "[dependencies]",
    ]
    for name, relative in deps:
        if relative is None:
            version = "1" if name == "serde_json" else "0.10"
            lines.append(f'{name} = "{version}"')
        else:
            path = (source / relative).as_posix()
            features_clause = f", features = [{features}]" if name == "litchi-ops" else ""
            lines.append(f'{name} = {{ path = {json.dumps(path)}{features_clause} }}')
    return "\n".join(lines) + "\n"


def stage_fixtures(source: Path, fixture_root: Path, fixtures: dict[str, Any]) -> Path:
    fixture_root = safe_external(fixture_root, "fixture root")
    destination_root = fixture_root / "signed-native-fixtures"
    if destination_root.exists():
        raise fail(f"refusing to overwrite fixture root: {destination_root}")
    for row in fixtures["fixtures"]:
        relative = Path(row["relative_path"])
        source_path = source / relative
        try:
            source_path.resolve().relative_to(source.resolve())
        except ValueError:
            raise fail(f"fixture escapes the immutable source tree: {relative}")
        if not source_path.is_file() or source_path.is_symlink():
            raise fail(f"tracked fixture is missing: {source_path}")
        if digest(source_path) != row["input_sha256"]:
            raise fail(f"fixture input hash changed: {relative}")
        destination = destination_root / relative
        try:
            destination.resolve().relative_to(destination_root.resolve())
        except ValueError:
            raise fail(f"fixture destination escapes the staged root: {relative}")
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source_path, destination)
    manifest_path = destination_root / "docs/unified-ops/results/0362ul24/fixtures.json"
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    manifest_path.write_bytes(canonical(fixtures))
    return manifest_path


def plan(
    source: Path,
    kind: str,
    output: Path,
    fixture_root: Path,
    *,
    source_archive: Path | None = None,
    source_assembly: Path | None = None,
) -> dict[str, Any]:
    identity = source_identity(
        source, kind, source_archive=source_archive, source_assembly=source_assembly
    )
    fixtures = fixture_manifest()
    if not PROBE.is_file():
        raise fail(f"missing signed probe: {PROBE}")
    output = safe_external(output, "output")
    fixture_root = safe_external(fixture_root, "fixture root")
    identity["source_capture"] = str((output / "source-identity.json").resolve())
    return {
        "schema": "litchi_signed_native_consumer_plan.v1",
        "source": identity,
        "probe": artifact(PROBE),
        "fixtures": {
            "manifest": artifact(FIXTURES),
            "schema": fixtures["schema"],
            "formats": sorted({str(row["format"]) for row in fixtures["fixtures"]}),
            "count": len(fixtures["fixtures"]),
        },
        "adapter": {
            "kind": kind,
            "source": "consumer-only generated module",
            "sha256": hashlib.sha256(adapter_source(kind).encode()).hexdigest(),
            "f2_field_bridge": kind == "f2",
            "xlsb_option_bridge": kind == "f2",
        },
        "build_policy": {
            "cargo": "caller-supplied pinned memory-capped shim",
            "jobs": 1,
            "incremental": 0,
            "offline": True,
            "locked": True,
            "memory_cap": "20G",
            "target_cap_bytes": TARGET_CAP_BYTES,
            "tmp_cap_bytes": TARGET_CAP_BYTES,
            "target_dir": str(output / "target"),
            "tmp_dir": str(output / "tmp"),
            "fixture_root": str(fixture_root),
            "timing": "not run by this driver",
        },
        "commands": [
            "<cargo-shim> generate-lockfile --offline",
            "<cargo-shim> metadata --offline --locked --format-version=1",
            "<cargo-shim> build --offline --locked --release --bin signed-native-consumer",
        ],
        "source_closure": {
            "archive": identity["source_archive"],
            "assembly": identity["source_assembly"],
            "tree": identity["source_tree"],
            "lock": identity["lock"],
            "toolchain": identity["toolchain"],
            "cargo_config": identity["cargo_config"],
        },
    }


def write_external(path: Path, payload: bytes) -> dict[str, Any]:
    path = safe_external(path, "output")
    if path.exists():
        raise fail(f"refusing to overwrite existing output: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(payload)
    return artifact(path)


def prepare(
    source: Path,
    kind: str,
    output: Path,
    fixture_root: Path,
    *,
    source_archive: Path | None = None,
    source_assembly: Path | None = None,
) -> dict[str, Any]:
    output = safe_external(output, "output")
    fixture_root = safe_external(fixture_root, "fixture root")
    if output.exists():
        raise fail(f"refusing to overwrite existing preparation directory: {output}")
    output.mkdir(parents=True)
    project = output / "consumer"
    (project / "src").mkdir(parents=True)
    identity = source_identity(
        source, kind, source_archive=source_archive, source_assembly=source_assembly
    )
    identity["source_capture"] = str((output / "source-identity.json").resolve())
    fixtures = fixture_manifest()
    (project / "src" / "main.rs").write_text(
        "mod api_compat;\n" + PROBE.read_text(encoding="utf-8"), encoding="utf-8"
    )
    (project / "src" / "api_compat.rs").write_text(adapter_source(kind), encoding="utf-8")
    # Use a harmless compile-time assertion so the adapter verifies the API
    # shape without changing the probe's runtime behavior.
    main = project / "src" / "main.rs"
    text = main.read_text(encoding="utf-8")
    marker = "fn main() {"
    if marker not in text:
        raise fail("signed probe has no main function")
    main.write_text(text.replace(marker, marker + "\n    api_compat::assert_api();", 1), encoding="utf-8")
    (project / "Cargo.toml").write_text(project_toml(source, kind), encoding="utf-8")
    fixture_manifest_path = stage_fixtures(source, fixture_root, fixtures)
    (output / "fixture-manifest-receipt.json").write_bytes(
        canonical(
            {
                "schema": "litchi_signed_native_fixture_binding.v1",
                "manifest": artifact(fixture_manifest_path),
                "root": str(fixture_manifest_path.parents[4].resolve()),
            }
        )
    )
    generated_archive = None
    if identity["source_archive"] is None:
        generated_archive = output / "source.tar"
        with generated_archive.open("wb") as stream:
            subprocess.run(
                ["git", "archive", "--format=tar", "--prefix=litchi-source/", identity["commit"]],
                cwd=source,
                check=True,
                stdout=stream,
            )
        identity["source_archive"] = artifact(generated_archive)
    effective_archive = generated_archive or source_archive
    identity["adapter"] = artifact(project / "src/api_compat.rs")
    (output / "source-identity.json").write_bytes(canonical(identity))
    (output / "adapter-api.json").write_bytes(
        canonical({"schema": "litchi_signed_native_adapter.v1", "kind": kind, "source_commit": identity["commit"], "module": artifact(project / "src/api_compat.rs")})
    )
    (output / "plan.json").write_bytes(
        canonical(
            plan(
                source, kind, output, fixture_manifest_path.parents[4],
                source_archive=effective_archive,
                source_assembly=source_assembly,
            )
        )
    )
    return {
        "status": "prepared",
        "output": str(output),
        "source": identity,
        "fixture_manifest": str(fixture_manifest_path),
        "fixture_root": str(fixture_manifest_path.parents[4].resolve()),
    }


def _directory_bytes(path: Path, cwd: Path) -> int:
    if not path.is_dir() or path.is_symlink():
        raise fail(f"expected a fresh directory: {path}")
    try:
        return int(run(["du", "-sxB1", "--", str(path)], cwd).split()[0])
    except (IndexError, ValueError) as error:
        raise fail(f"could not measure directory: {path}") from error


def _load_object(path: Path, label: str) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise fail(f"{label} is invalid JSON: {error}") from error
    if not isinstance(value, dict):
        raise fail(f"{label} is not a JSON object")
    return value


def build(prepared: Path, cargo: Path | None) -> dict[str, Any]:
    prepared = safe_external(prepared, "prepared output")
    project = prepared / "consumer"
    plan_path = prepared / "plan.json"
    source_identity_path = prepared / "source-identity.json"
    fixture_binding_path = prepared / "fixture-manifest-receipt.json"
    for path, label in (
        (project, "consumer project"),
        (plan_path, "plan"),
        (source_identity_path, "source identity"),
        (fixture_binding_path, "fixture binding"),
    ):
        if not path.exists():
            raise fail(f"prepared output lacks {label}: {path}")
    if not project.is_dir() or project.is_symlink():
        raise fail(f"consumer project is not a regular directory: {project}")
    if cargo is None:
        raise fail("build requires the pinned memory-capped cargo shim (--cargo)")
    cargo = safe_external(cargo, "cargo shim")
    if not cargo.is_file() or cargo.is_symlink() or not os.access(cargo, os.X_OK):
        raise fail(f"cargo shim is not an executable regular file: {cargo}")

    target = prepared / "target"
    temporary = prepared / "tmp"
    for path, label in ((target, "target"), (temporary, "TMPDIR")):
        if path.exists() or path.is_symlink():
            raise fail(f"refusing to reuse {label} directory: {path}")
        path.mkdir()

    cargo_cmd = str(cargo)
    env = dict(os.environ)
    env.update(
        {
            "CARGO_TARGET_DIR": str(target),
            "CARGO_BUILD_JOBS": "1",
            "CARGO_INCREMENTAL": "0",
            "CARGO_NET_OFFLINE": "true",
            "CARGO_TERM_COLOR": "never",
            "LITCHI_MEM_MAX": "20G",
            "TMPDIR": str(temporary),
        }
    )
    run([cargo_cmd, "generate-lockfile", "--offline"], project, environment=env)
    metadata_raw = run(
        [cargo_cmd, "metadata", "--offline", "--locked", "--format-version=1"],
        project,
        environment=env,
    )
    try:
        metadata_value = json.loads(metadata_raw)
    except json.JSONDecodeError as error:
        raise fail(f"cargo metadata returned invalid JSON: {error}") from error
    if not isinstance(metadata_value, dict):
        raise fail("cargo metadata did not return an object")
    metadata_path = prepared / "metadata.json"
    metadata_path.write_bytes(canonical(metadata_value))
    command = [
        cargo_cmd,
        "build",
        "--offline",
        "--locked",
        "--release",
        "--bin",
        "signed-native-consumer",
    ]
    try:
        subprocess.run(command, cwd=project, env=env, check=True)
    except (OSError, subprocess.CalledProcessError) as error:
        raise fail(f"release build failed: {error}") from error
    binary = target / "release/signed-native-consumer"
    if not binary.is_file() or binary.is_symlink() or not os.access(binary, os.X_OK):
        raise fail(f"release binary is missing or not executable: {binary}")
    target_bytes = _directory_bytes(target, project)
    temporary_bytes = _directory_bytes(temporary, project)
    for label, size in (("target", target_bytes), ("TMPDIR", temporary_bytes)):
        if size >= TARGET_CAP_BYTES:
            raise fail(f"{label} exceeds 20 GB cap: {size}")

    source_identity = _load_object(source_identity_path, "source identity")
    source_archive = source_identity.get("source_archive")
    if not isinstance(source_archive, dict):
        raise fail("source identity lacks a verified source archive")
    fixture_binding = _load_object(fixture_binding_path, "fixture binding")
    fixture_manifest = fixture_binding.get("manifest")
    fixture_root = fixture_binding.get("root")
    if not isinstance(fixture_manifest, dict) or not isinstance(fixture_root, str):
        raise fail("fixture binding is incomplete")
    fixture_manifest_path = Path(str(fixture_manifest.get("path", "")))
    if not fixture_manifest_path.is_file() or artifact(fixture_manifest_path) != fixture_manifest:
        raise fail("staged fixture manifest changed since preparation")
    if not Path(fixture_root).is_dir() or Path(fixture_root).is_symlink():
        raise fail("staged fixture root is not a regular directory")
    adapter_artifact = artifact(project / "src/api_compat.rs")
    if source_identity.get("adapter") != adapter_artifact:
        raise fail("source identity adapter differs from prepared consumer adapter")
    probe_artifact = artifact(project / "src/main.rs")
    lock_artifact = artifact(project / "Cargo.lock")
    metadata_artifact = artifact(metadata_path)
    binary_artifact = artifact(binary)
    receipt = {
        "schema": "litchi_signed_native_build.v1",
        "status": "built",
        "source_commit": source_identity.get("commit"),
        "source": source_identity,
        "source_lock": source_identity.get("lock"),
        "source_archive": source_archive,
        "source_archive_sha256": source_archive.get("sha256"),
        "source_identity": source_identity,
        "plan": artifact(plan_path),
        "adapter": adapter_artifact,
        "adapter_sha256": adapter_artifact["sha256"],
        "probe": probe_artifact,
        "lock": lock_artifact,
        "lock_sha256": lock_artifact["sha256"],
        "fixtures": fixture_manifest,
        "fixture_root": fixture_root,
        "metadata": metadata_artifact,
        "dependency_projection": metadata_artifact,
        "binary": binary_artifact,
        "binary_sha256": binary_artifact["sha256"],
        "cargo": artifact(cargo),
        "command": command,
        "environment": {
            key: env[key]
            for key in (
                "CARGO_TARGET_DIR",
                "CARGO_BUILD_JOBS",
                "CARGO_INCREMENTAL",
                "CARGO_NET_OFFLINE",
                "CARGO_TERM_COLOR",
                "LITCHI_MEM_MAX",
                "TMPDIR",
            )
        },
        "target_bytes": target_bytes,
        "tmp_bytes": temporary_bytes,
        "timing": None,
    }
    receipt_path = prepared / "build-receipt.json"
    if receipt_path.exists() or receipt_path.is_symlink():
        raise fail(f"refusing to overwrite build receipt: {receipt_path}")
    receipt_path.write_bytes(canonical(receipt))
    return {
        "status": "built",
        "receipt": str(receipt_path),
        "binary": binary_artifact,
        "target_bytes": target_bytes,
        "tmp_bytes": temporary_bytes,
    }


def _capability_evidence(
    capabilities_path: Path,
    *,
    source_commit: str,
    binary: dict[str, Any],
) -> tuple[dict[str, str], dict[str, Any]]:
    capabilities_path = safe_external(capabilities_path, "capability receipt")
    value = _load_object(capabilities_path, "capability receipt")
    if value.get("schema") != "litchi_signed_native_capabilities.v1":
        raise fail("capability receipt schema mismatch")
    if value.get("source_commit") != source_commit:
        raise fail("capability receipt source commit differs from build")
    recorded_binary = value.get("binary")
    if not isinstance(recorded_binary, dict) or _verified_receipt_artifact(recorded_binary, "capability binary") != binary:
        raise fail("capability receipt binary differs from build")
    routes = value.get("routes")
    if not isinstance(routes, dict) or set(routes) != set(ROUTES):
        raise fail("capability receipt does not enumerate the signed route matrix")
    allowed = {"available", "capability_missing", "expected_refusal", "not_applicable"}
    if any(status not in allowed for status in routes.values()):
        raise fail("capability receipt contains an invalid status")
    evidence = value.get("evidence")
    if not isinstance(evidence, dict) or set(evidence) != set(ROUTES):
        raise fail("capability receipt lacks one validate receipt per signed route")
    for route in ROUTES:
        row = evidence[route]
        if isinstance(row, str):
            receipt_path = Path(row)
        elif isinstance(row, dict) and isinstance(row.get("path"), str):
            receipt_path = Path(row["path"])
        else:
            raise fail(f"capability evidence for {route} is not a receipt path")
        receipt_path = safe_external(receipt_path, f"capability evidence for {route}")
        receipt = _load_object(receipt_path, f"capability evidence for {route}")
        if receipt.get("schema") != "litchi_signed_native_child.v1" or receipt.get("mode") != "validate":
            raise fail(f"capability evidence for {route} is not a validate receipt")
        if receipt.get("route") != route:
            raise fail(f"capability evidence route mismatch for {route}")
        child_binary = receipt.get("binary")
        if not isinstance(child_binary, dict) or _verified_receipt_artifact(child_binary, f"{route} binary") != binary:
            raise fail(f"capability evidence binary mismatch for {route}")
        stdout = receipt.get("stdout")
        if not isinstance(stdout, dict):
            raise fail(f"capability evidence stdout is missing for {route}")
        stdout_artifact = _verified_receipt_artifact(stdout, f"{route} stdout")
        safe_external(Path(str(stdout_artifact["path"])), f"{route} stdout")
        try:
            result = json.loads(Path(stdout_artifact["path"]).read_text(encoding="utf-8"))
        except (OSError, UnicodeError, json.JSONDecodeError) as error:
            raise fail(f"capability evidence stdout is invalid for {route}: {error}") from error
        if not isinstance(result, dict) or result.get("route") != route:
            raise fail(f"capability evidence output route mismatch for {route}")
        observed = result.get("capability")
        if route.endswith("/refusal"):
            rows = result.get("fixtures")
            if not isinstance(rows, list) or not rows or any(
                not isinstance(row, dict) or row.get("status") != "expected_refusal"
                for row in rows
            ):
                raise fail(f"refusal evidence is not an expected refusal for {route}")
            observed = "expected_refusal"
        if observed != routes[route]:
            raise fail(f"capability status for {route} is not bound to probe output")
    return {str(route): str(status) for route, status in routes.items()}, value


def endpoint(
    prepared: Path,
    capabilities: Path,
    output: Path | None = None,
) -> dict[str, Any]:
    """Bind one post-build, untimed capability receipt into the endpoint schema.

    The capability receipt is produced by a guarded probe preflight.  Requiring
    every child receipt here prevents a caller from hand-writing endpoint
    capability claims while keeping build and timing stages separate.
    """
    prepared = safe_external(prepared, "prepared output")
    capabilities = safe_external(capabilities, "capability receipt")
    receipt_path = prepared / "build-receipt.json"
    source_identity_path = prepared / "source-identity.json"
    fixture_binding_path = prepared / "fixture-manifest-receipt.json"
    for path, label in (
        (receipt_path, "build receipt"),
        (source_identity_path, "source identity"),
        (fixture_binding_path, "fixture binding"),
    ):
        if not path.is_file() or path.is_symlink():
            raise fail(f"prepared output lacks {label}: {path}")
    receipt = _load_object(receipt_path, "build receipt")
    if receipt.get("schema") != "litchi_signed_native_build.v1" or receipt.get("status") != "built":
        raise fail("build receipt is not a completed signed-native build")
    source_identity = _load_object(source_identity_path, "source identity")
    commit = source_identity.get("commit")
    source_ref = source_identity.get("source_ref")
    source_capture = source_identity.get("source_capture")
    source_root = source_identity.get("source_root")
    tree_sha256 = source_identity.get("tree_sha256")
    if (
        not isinstance(commit, str)
        or not FULL_SHA.fullmatch(commit)
        or not isinstance(source_ref, str)
        or not source_ref
        or not isinstance(source_capture, str)
        or not isinstance(source_root, str)
        or not Path(source_root).is_dir()
        or not isinstance(tree_sha256, str)
        or SHA256.fullmatch(tree_sha256) is None
    ):
        raise fail("source identity lacks the required pinned fields")
    if Path(source_capture).resolve() != source_identity_path.resolve():
        raise fail("source capture is not the immutable source identity receipt")
    source_archive = receipt.get("source_archive")
    adapter = receipt.get("adapter")
    lock = receipt.get("lock")
    binary = receipt.get("binary")
    probe = receipt.get("probe")
    fixtures = receipt.get("fixtures")
    dependency_projection = receipt.get("dependency_projection")
    for value, label in (
        (source_archive, "source archive"),
        (adapter, "adapter"),
        (lock, "lock"),
        (binary, "binary"),
        (probe, "probe"),
        (fixtures, "fixtures"),
        (dependency_projection, "dependency projection"),
    ):
        if not isinstance(value, dict):
            raise fail(f"build receipt lacks {label} artifact")
        _verified_receipt_artifact(value, label)
    if source_archive.get("sha256") != receipt.get("source_archive_sha256"):
        raise fail("build receipt source archive hash is not bound")
    if source_identity.get("source_archive") != source_archive:
        raise fail("source identity archive differs from build receipt")
    if adapter.get("sha256") != receipt.get("adapter_sha256"):
        raise fail("build receipt adapter hash is not bound")
    if lock.get("sha256") != receipt.get("lock_sha256"):
        raise fail("build receipt lock hash is not bound")
    if binary.get("sha256") != receipt.get("binary_sha256"):
        raise fail("build receipt binary hash is not bound")
    if receipt.get("source_commit") != commit:
        raise fail("build receipt source commit differs")
    identity_adapter = source_identity.get("adapter")
    if identity_adapter != adapter:
        raise fail("source identity adapter differs from build receipt")
    fixture_binding = _load_object(fixture_binding_path, "fixture binding")
    fixture_manifest = fixture_binding.get("manifest")
    fixture_root = fixture_binding.get("root")
    if not isinstance(fixture_manifest, dict) or fixture_manifest != fixtures:
        raise fail("build receipt fixtures differ from staged fixture binding")
    if not isinstance(fixture_root, str) or not Path(fixture_root).is_dir():
        raise fail("staged fixture root is missing")
    statuses, capability_value = _capability_evidence(
        capabilities, source_commit=commit, binary=binary
    )
    endpoint_value = {
        "schema": "litchi_signed_native_endpoint.v1",
        "source_commit": commit,
        "source_identity": {
            "commit": commit,
            "source_ref": source_ref,
            "source_capture": source_capture,
            "source_root": source_root,
            "tree_sha256": tree_sha256,
            "archive": source_archive,
            "adapter": adapter,
        },
        "source": artifact(source_identity_path),
        "source_archive": source_archive,
        "adapter": adapter,
        "lock": lock,
        "binary": binary,
        "probe": probe,
        "fixtures": fixtures,
        "dependency_projection": dependency_projection,
        "build_receipt": artifact(receipt_path),
        "capabilities": statuses,
        "capabilities_receipt": artifact(capabilities),
        "fixture_root": fixture_root,
        "build": {
            "source_commit": receipt.get("source_commit"),
            "binary_sha256": receipt.get("binary_sha256"),
            "source_archive_sha256": receipt.get("source_archive_sha256"),
            "adapter_sha256": receipt.get("adapter_sha256"),
            "lock_sha256": receipt.get("lock_sha256"),
            "capabilities_schema": capability_value.get("schema"),
        },
    }
    endpoint_path = safe_external(output or (prepared / "endpoint.json"), "endpoint")
    if endpoint_path.exists() or endpoint_path.is_symlink():
        raise fail(f"refusing to overwrite endpoint manifest: {endpoint_path}")
    endpoint_path.parent.mkdir(parents=True, exist_ok=True)
    endpoint_path.write_bytes(canonical(endpoint_value))
    return {"status": "endpoint", "endpoint": str(endpoint_path), "binary": binary}


def main() -> int:
    parser = argparse.ArgumentParser()
    sub = parser.add_subparsers(dest="command", required=True)
    for name in ("plan", "prepare"):
        p = sub.add_parser(name)
        p.add_argument("--source", required=True, type=Path)
        p.add_argument("--kind", required=True, choices=sorted(SOURCE_KINDS))
        p.add_argument("--output", required=True, type=Path)
        p.add_argument("--fixture-root", required=True, type=Path)
        p.add_argument("--source-archive", type=Path)
        p.add_argument("--source-assembly", type=Path)
    b = sub.add_parser("build")
    b.add_argument("--prepared", required=True, type=Path)
    b.add_argument("--cargo", required=True, type=Path)
    e = sub.add_parser("endpoint")
    e.add_argument("--prepared", required=True, type=Path)
    e.add_argument("--capabilities", required=True, type=Path)
    e.add_argument("--output", type=Path)
    args = parser.parse_args()
    if args.command == "plan":
        value = plan(
            args.source, args.kind, args.output, args.fixture_root,
            source_archive=args.source_archive,
            source_assembly=args.source_assembly,
        )
    elif args.command == "prepare":
        value = prepare(
            args.source, args.kind, args.output, args.fixture_root,
            source_archive=args.source_archive,
            source_assembly=args.source_assembly,
        )
    elif args.command == "build":
        value = build(args.prepared, args.cargo)
    else:
        value = endpoint(args.prepared, args.capabilities, args.output)
    print(json.dumps(value, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        print(f"signed-native preparation failed: {error}", file=sys.stderr)
        raise
