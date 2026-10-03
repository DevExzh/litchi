#!/usr/bin/env python3
"""Capture the bounded 0894 baseline/candidate paired workload.

This driver is intentionally a packet-local protocol adapter.  It launches
one fresh child for each frozen row, validates the semantic report before
advancing, and stops at the first failure without retrying or replacing a
child.  Building binaries and preparing source trees belong to the separate
0894 source/build receipt helpers.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
import os
from pathlib import Path
import re
import stat
import subprocess
from typing import Any


PACKET = Path(__file__).resolve().parent
BASELINE_QUALIFICATION = PACKET.parent / "change-0887" / "baseline-qualification"
GUARD_REFERENCE = PACKET.parent / "change-0893" / "qualification-after-cleanup" / "report.json"
SOURCE_SCHEMA = "litchi.xlsx-source-snapshot-0894.v1"
CANDIDATE_FREEZE_SCHEMA = "litchi.xlsx-selected-metadata-candidate-freeze-0894.v1"
QUALITY_SOURCE_SCHEMA = "litchi.isolated-quality-source-0894.v1"

SCHEMA = "litchi.xlsx-selected-metadata-experiment-0894.v1"
COMMAND_SCHEMA = "litchi.xlsx-selected-metadata-child-command-0894.v1"
RECEIPT_SCHEMA = "litchi.xlsx-selected-metadata-child-receipt-0894.v1"
EXPERIMENT_SCHEMA = "litchi.xlsx-selected-metadata-experiment-receipt-0894.v1"
PROFILE = "release"
CPU_AFFINITY = "12"
TASKSET = "/usr/bin/taskset"
TIME = "/usr/bin/time"
BASE_COMMIT = "506aa2edae6666d9c76ae6827c2f5941a8e47896"
SOURCE_FILE_COUNT = 10_771
SOURCE_ARCHIVE_BYTES = 89_047_170
SOURCE_ARCHIVE_SHA256 = "f6f9d13813866ea24fd9cc108fc949aa42cdd042bacd3ca960897e8c88c63ce9"
SOURCE_SELECTION = [
    ".cargo", "Cargo.toml", "README.md", "clippy.toml", "deny.toml",
    "rust-toolchain.toml", "rustfmt.toml", "crates", "test-data",
    "tools/perf-baseline",
]
CANDIDATE_PATHS = [
    "crates/litchi-xlsx/src/raw/worksheet/selected.rs",
    "crates/litchi-xlsx/src/raw/worksheet/selected/producer_metadata.rs",
    "crates/litchi-xlsx/src/raw/worksheet/selected/producer_metadata_tests.rs",
    "tools/perf-baseline/src/producer_shape/selected_route_tests.rs",
    "crates/litchi-xlsx/src/raw/worksheet/tests.rs",
]
HELPER_PATHS = {
    "experiment.py": PACKET / "experiment.py",
    "freeze.py": PACKET / "freeze.py",
    "verify.py": PACKET / "verify.py",
    "reduce.py": PACKET / "reduce.py",
    "reduce-0887.py": PACKET.parent / "change-0887" / "reduce.py",
    "quality.py": PACKET / "quality.py",
    "quality_source.py": PACKET / "quality_source.py",
}
QUALITY_SUPPORT_PATHS = (
    "tools/check_crate_boundaries.py",
    "tools/crate_boundaries.json",
    "docs/adr/0008-migration-and-verification.md",
    "docs/adr/0028-iwa-monolith-exit.md",
)

CASES = (
    "xlsx_producer_medium_source_full_cell_scan",
    "xlsx_full_cell_scan",
    "docx_producer_source_selected_paragraph",
    "pptx_producer_source_selected_slide",
)
PRIMARY_CASE = CASES[0]
NATIVE_BLOCKS = 8
NATIVE_WARMUP = 3
NATIVE_SAMPLES = 50
ALLOCATOR_BLOCKS = 4
ALLOCATOR_WARMUP = 3
ALLOCATOR_SAMPLES = 15
TOTAL_NATIVE_CHILDREN = NATIVE_BLOCKS * len(CASES) * 2
TOTAL_ALLOCATOR_CHILDREN = ALLOCATOR_BLOCKS * 2
TOTAL_CHILDREN = TOTAL_NATIVE_CHILDREN + TOTAL_ALLOCATOR_CHILDREN
BOOTSTRAP_SEED = 894
BOOTSTRAP_DRAWS = 10_000
BOOTSTRAP_CI_INDICES = (249, 9_749)
PAIR_RATIO_REVIEW_THRESHOLD = 1.05
P50_SPREAD_REVIEW_THRESHOLD = 1.20

REQUIRED_BINDINGS = (
    "source_archive", "source_manifest", "source_identity", "candidate_freeze",
    "baseline_build", "candidate_build", "baseline_native_depfile",
    "baseline_allocator_depfile", "candidate_native_depfile",
    "candidate_allocator_depfile", "baseline_flags", "candidate_flags",
    "baseline_quality", "candidate_quality", "source_review",
    "candidate_review", "methodology_review", "quality_source",
)

# The manifest is intentionally role-aware.  A path plus a hash is not enough:
# each retained receipt must state which build, binary lane, or review scope it
# belongs to, and the verifier checks those relationships against receipt data.
BINDING_SPECS = {
    "source_archive": {"kind": "source_archive", "role": "source"},
    "source_manifest": {"kind": "source_manifest", "role": "source"},
    "source_identity": {"kind": "source_identity", "role": "source"},
    "candidate_freeze": {"kind": "candidate_freeze", "role": "candidate"},
    "baseline_build": {"kind": "build_receipt", "role": "baseline"},
    "candidate_build": {"kind": "build_receipt", "role": "candidate"},
    "baseline_native_depfile": {"kind": "depfile", "role": "baseline", "binary_role": "native", "build_binding": "baseline_build"},
    "baseline_allocator_depfile": {"kind": "depfile", "role": "baseline", "binary_role": "allocator", "build_binding": "baseline_build"},
    "candidate_native_depfile": {"kind": "depfile", "role": "candidate", "binary_role": "native", "build_binding": "candidate_build"},
    "candidate_allocator_depfile": {"kind": "depfile", "role": "candidate", "binary_role": "allocator", "build_binding": "candidate_build"},
    "baseline_flags": {"kind": "build_flags", "role": "baseline", "build_binding": "baseline_build"},
    "candidate_flags": {"kind": "build_flags", "role": "candidate", "build_binding": "candidate_build"},
    "baseline_quality": {"kind": "quality_receipt", "role": "baseline", "review_binding": "source_review"},
    "candidate_quality": {"kind": "quality_receipt", "role": "candidate", "build_binding": "candidate_build", "candidate_binding": "candidate_freeze"},
    "quality_source": {"kind": "quality_source", "role": "candidate", "candidate_binding": "candidate_freeze", "archive_binding": "source_archive"},
    "source_review": {"kind": "review", "review_scope": "source", "source_binding": "source_manifest"},
    "candidate_review": {"kind": "review", "review_scope": "candidate", "candidate_binding": "candidate_freeze"},
    "methodology_review": {"kind": "review", "review_scope": "methodology", "source_binding": "source_review", "candidate_binding": "candidate_review"},
}

BUILD_CONFIGURATION = {
    "toolchain": "1.95.0",
    "profile": "release",
    "opt_level": 3,
    "debug": 1,
    "lto": "fat",
    "codegen_units": 1,
    "jobs": 2,
}


def binary_evidence_bindings(role: str, mode: str) -> dict[str, str]:
    _require(role in ("before", "after") and mode in ("native", "allocator"), "invalid binary evidence role")
    build = "baseline_build" if role == "before" else "candidate_build"
    quality = "baseline_quality" if role == "before" else "candidate_quality"
    depfile = f"{'baseline' if role == 'before' else 'candidate'}_{mode}_depfile"
    return {
        "build": build,
        "depfile": depfile,
        "flags": f"{'baseline' if role == 'before' else 'candidate'}_flags",
        "source_archive": "source_archive",
        "source_manifest": "source_manifest",
        "candidate_freeze": "candidate_freeze",
        "quality": quality,
        "source_review": "source_review",
        "candidate_review": "candidate_review",
        "methodology_review": "methodology_review",
    }
BASELINE_QUALITY_SOURCE_HASHES = {
    "tools/perf-baseline/src/producer_shape.rs": "0a81997839bfe6f23edc09707bd6b695a8d11aecbd330b1d9191ffff4c8cfa77",
    "tools/perf-baseline/src/producer_shape/full_cell_guard_tests.rs": "645d9aeafc54d6a9febdc0cf86cfa8d43a66c378b1d3e466a47abc0d68accccd",
    "tools/perf-baseline/src/producer_shape/selected_route_tests.rs": "f70148c654408689b9a970d309da5266408d4cd51eb28e7f171efad985bd396a",
}
BASELINE_QUALITY_HEAD = "1e5120f0818b754c31f25d6c1b033269600e38f6"
GUARD_TEST_SHA256 = BASELINE_QUALITY_SOURCE_HASHES["tools/perf-baseline/src/producer_shape/full_cell_guard_tests.rs"]

GUARD_CONTRACT = "post-endpoint-per-worksheet-tolerance-v1"
PRIMARY_ARCHIVE_SHA256 = "55e0901c093c27d80bcd4e5fef888f00798654cdef8c0d79d130a2ce7cbe0ae8"
PRIMARY_ARCHIVE_BYTES = 22_739
PRIMARY_DIGEST_SHA256 = "8d29b40718aa9d5b08c746f841195bba17ab8e01608b12e6093daf160c00764b"
PRIMARY_CENSUS_SHA256 = "552c35c8b8f958de28792f00c1e9e831ae884471aa28b5b2a7a1479bd06c5207"

CHILD_ENV = {
    "PATH": "/usr/bin:/bin",
    "LC_ALL": "C",
    "LANG": "C",
    "TZ": "UTC",
}


def _require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def _utc_now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def _sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    path = Path(path)
    _require(path.is_file() and not path.is_symlink(), f"expected regular non-symlink file: {path}")
    path = path.resolve()
    info = path.stat()
    _require(stat.S_ISREG(info.st_mode), f"expected regular file: {path}")
    return {
        "path": str(path),
        "bytes": info.st_size,
        "sha256": _sha(path),
        "mode": stat.S_IMODE(info.st_mode),
        "executable": bool(info.st_mode & stat.S_IXUSR),
    }


def _read(path: Path) -> Any:
    with path.open(encoding="utf-8") as handle:
        return json.load(handle)


def _write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as handle:
        json.dump(value, handle, indent=2, sort_keys=True)
        handle.write("\n")


def _entry(value: Any) -> dict[str, Any]:
    if isinstance(value, str):
        return {"path": value}
    if isinstance(value, dict):
        if isinstance(value.get("identity"), dict):
            merged = dict(value["identity"])
            merged.update({key: item for key, item in value.items() if key != "identity"})
            return merged
        return value
    raise ValueError(f"invalid binary descriptor: {value!r}")


def _descriptor_fields(value: dict[str, Any]) -> dict[str, Any]:
    return {key: value.get(key) for key in ("path", "bytes", "sha256", "mode", "executable")}


def _same_descriptor(left: Any, right: Any, label: str) -> None:
    _require(isinstance(left, dict) and isinstance(right, dict), f"{label} descriptor is missing")
    for key in ("path", "bytes", "sha256", "mode", "executable"):
        _require(left.get(key) == right.get(key), f"{label} {key} mismatch")


def _regular_path(raw: Any, base: Path, label: str) -> Path:
    _require(isinstance(raw, (str, Path)), f"{label} has no path")
    path = Path(raw)
    if not path.is_absolute():
        path = base / path
    _require(path.is_file() and not path.is_symlink(), f"{label} is not a regular non-symlink file: {path}")
    return path.resolve()


def _fixture_paths() -> dict[str, Path]:
    return {
        "xlsx_primary_report": BASELINE_QUALIFICATION / CASES[0] / "report.json",
        "xlsx_primary_producer": BASELINE_QUALIFICATION / CASES[0] / "producer.json",
        "xlsx_marker_free_report": BASELINE_QUALIFICATION / CASES[1] / "report.json",
        "docx_report": BASELINE_QUALIFICATION / CASES[2] / "report.json",
        "docx_producer": BASELINE_QUALIFICATION / CASES[2] / "producer.json",
        "pptx_report": BASELINE_QUALIFICATION / CASES[3] / "report.json",
        "pptx_producer": BASELINE_QUALIFICATION / CASES[3] / "producer.json",
        "guard_reference": GUARD_REFERENCE,
    }


def fixture_descriptors() -> dict[str, dict[str, Any]]:
    return {name: descriptor(path) for name, path in _fixture_paths().items()}


def helper_descriptors() -> dict[str, dict[str, Any]]:
    return {name: descriptor(path) for name, path in HELPER_PATHS.items()}


def _validate_helper_bindings(plan: dict[str, Any]) -> None:
    helpers = plan.get("helper_hashes")
    _require(isinstance(helpers, dict) and set(helpers) == set(HELPER_PATHS), "helper hash binding set is incomplete")
    for name, path in HELPER_PATHS.items():
        _same_descriptor(helpers.get(name), descriptor(path), f"helper {name}")


def _validate_fixture_bindings(plan: dict[str, Any]) -> None:
    fixtures = plan.get("fixtures")
    expected_paths = _fixture_paths()
    _require(isinstance(fixtures, dict) and set(fixtures) == set(expected_paths), "sealed 0887/0893 fixture binding set is incomplete")
    for name, expected_path in expected_paths.items():
        actual = descriptor(expected_path)
        item = fixtures.get(name)
        _same_descriptor(item, actual, f"fixture {name}")
        _require(item.get("path") == actual["path"], f"fixture {name} path changed")


def _resolve_descriptor(packet: Path, raw: Any, name: str) -> dict[str, Any]:
    item = _entry(raw)
    path = _regular_path(item.get("path"), packet, name)
    resolved = path
    _require(resolved.is_file() and os.access(resolved, os.X_OK), f"{name} is not executable: {resolved}")
    actual = descriptor(resolved)
    for key in ("bytes", "sha256", "mode", "executable"):
        if key in item:
            _require(item[key] == actual[key], f"{name} {key} does not match the retained binary")
    _require(item.get("executable", True) is True, f"{name} is not declared executable")
    # Keep the complete role/name/features/build-receipt/depfile/ELF metadata
    # that was frozen with the binary.  Capture receipts compare this complete
    # object, so a descriptor cannot be rebound to another build later.
    result = {"path": str(resolved), "identity": actual, "declared": item}
    for key in ("role", "name", "features", "profile", "elf_build_id", "build_receipt", "depfile", "evidence_bindings"):
        if key in item:
            result[key] = item[key]
    return result


def _binary(plan: dict[str, Any], role: str, mode: str) -> dict[str, Any]:
    binaries = plan.get("binaries")
    _require(isinstance(binaries, dict), "plan binaries must be an object")
    side = binaries.get(role)
    _require(isinstance(side, dict), f"plan has no {role} binary set")
    return _resolve_descriptor(PACKET, side.get(mode), f"{role}/{mode}")


def _binding_path(bindings: dict[str, Any], name: str) -> Path:
    item = _entry(bindings.get(name))
    return _regular_path(item.get("path"), PACKET, f"binding {name}")


def _binding_descriptor(bindings: dict[str, Any], name: str) -> dict[str, Any]:
    item = _entry(bindings.get(name))
    actual = descriptor(_binding_path(bindings, name))
    _same_descriptor(item, actual, f"binding {name}")
    return actual


def _json_binding(bindings: dict[str, Any], name: str) -> tuple[Path, dict[str, Any]]:
    path = _binding_path(bindings, name)
    try:
        value = _read(path)
    except (OSError, ValueError) as error:
        raise ValueError(f"binding {name} must be a JSON receipt: {path}") from error
    _require(isinstance(value, dict), f"binding {name} JSON root must be an object")
    return path, value


def _candidate_freeze_contract(
    path: Path,
    source: dict[str, Any],
    source_path: Path,
    source_archive: dict[str, Any],
) -> dict[str, Any]:
    _require(path.resolve() == (PACKET / "candidate-freeze.json").resolve(), "candidate freeze is not the packet receipt")
    receipt = _read(path)
    _require(receipt.get("schema") == CANDIDATE_FREEZE_SCHEMA, "candidate freeze schema mismatch")
    _require(receipt.get("status") == "frozen", "candidate freeze is not frozen")
    for key in ("base_commit", "baseline_commit"):
        _require(receipt.get(key) == BASE_COMMIT, f"candidate freeze {key} mismatch")
    _require(receipt.get("timed_harness_guard_contract") == GUARD_CONTRACT, "candidate freeze guard contract mismatch")
    _require(receipt.get("candidate_paths") == CANDIDATE_PATHS, "candidate freeze path inventory mismatch")
    _require(receipt.get("baseline_file_count") == SOURCE_FILE_COUNT, "candidate freeze baseline census mismatch")
    _require(receipt.get("candidate_file_count") == SOURCE_FILE_COUNT + 2, "candidate freeze candidate census mismatch")
    _same_descriptor(receipt.get("source_manifest"), descriptor(source_path), "candidate freeze source manifest")
    _same_descriptor(receipt.get("source_archive"), source_archive, "candidate freeze source archive")

    patch_item = receipt.get("candidate_patch")
    _require(isinstance(patch_item, dict), "candidate freeze patch descriptor is missing")
    patch_path = _regular_path(patch_item.get("path"), path.parent, "candidate freeze patch")
    _require(patch_path == (PACKET / "candidate.patch").resolve(), "candidate freeze patch is not the packet patch")
    _same_descriptor(patch_item, descriptor(patch_path), "candidate freeze patch")
    overlay_root_raw = receipt.get("candidate_overlay_root")
    _require(isinstance(overlay_root_raw, str), "candidate freeze overlay root is missing")
    overlay_root = Path(overlay_root_raw)
    _require(overlay_root.is_dir() and not overlay_root.is_symlink(), "candidate freeze overlay root is not a regular directory")
    _require(overlay_root.resolve() == (PACKET / "candidate-overlay").resolve(), "candidate freeze overlay is not the packet overlay")
    files = receipt.get("files")
    _require(isinstance(files, dict) and set(files) == set(CANDIDATE_PATHS), "candidate freeze file inventory mismatch")
    _require(receipt.get("replacement_files") == 3 and receipt.get("added_files") == 2, "candidate freeze replacement/addition census mismatch")
    source_files = source.get("files")
    _require(isinstance(source_files, dict) and len(source_files) == SOURCE_FILE_COUNT, "source file census is incomplete")
    for name in CANDIDATE_PATHS:
        item = files[name]
        _require(isinstance(item, dict), f"candidate freeze file record is invalid: {name}")
        current = item.get("current")
        overlay = item.get("overlay")
        _require(isinstance(current, dict) and isinstance(overlay, dict), f"candidate freeze file identity is missing: {name}")
        _require(current.get("path") == name and overlay.get("path") == name, f"candidate freeze relative path mismatch: {name}")
        overlay_path = _regular_path(overlay_root / name, overlay_root, f"candidate overlay {name}")
        observed = descriptor(overlay_path)
        _same_descriptor(overlay, {**observed, "path": name}, f"candidate overlay {name}")
        _same_descriptor(current, overlay, f"candidate current/overlay {name}")
        baseline = item.get("baseline")
        if name in source_files:
            _require(isinstance(baseline, dict), f"candidate replacement lacks baseline identity: {name}")
            _require(baseline.get("bytes") == source_files[name].get("bytes") and baseline.get("sha256") == source_files[name].get("sha256") and baseline.get("mode") == source_files[name].get("mode"), f"candidate baseline identity mismatch: {name}")
            _require(item.get("replacement") is True, f"candidate replacement flag mismatch: {name}")
        else:
            _require(baseline is None and item.get("replacement") is False, f"candidate addition flag mismatch: {name}")
    return receipt


def _role_build_contract(
    plan: dict[str, Any],
    bindings: dict[str, Any],
    build_role: str,
    build_path: Path,
    build: dict[str, Any],
) -> None:
    schema_word = "baseline" if build_role == "baseline" else "candidate"
    expected_build_path = PACKET / ("build.json" if build_role == "baseline" else "candidate-build.json")
    _require(build_path.resolve() == expected_build_path.resolve(), f"{build_role} build receipt is not the packet terminal receipt")
    _require(build.get("status") == "pass", f"{build_role} build receipt does not pass")
    _require(build.get("schema") == f"litchi.xlsx-{schema_word}-build-0894.v1" or build.get("schema") == f"litchi.xlsx-selected-metadata-{schema_word}-build-0894.v1", f"{build_role} build schema mismatch")
    for key in ("base_commit", "baseline_commit", "source_revision"):
        _require(build.get(key) == BASE_COMMIT, f"{build_role} build {key} mismatch")
    _require(build.get("profile") == PROFILE, f"{build_role} build profile mismatch")
    _require(build.get("guard_contract") == GUARD_CONTRACT and build.get("timed_harness_guard_contract") == GUARD_CONTRACT, f"{build_role} build guard contract mismatch")
    _require(build.get("flags") == BUILD_CONFIGURATION, f"{build_role} build flags mismatch")
    _same_descriptor(build.get("source_manifest"), _binding_descriptor(bindings, "source_manifest"), f"{build_role} source manifest binding")
    _same_descriptor(build.get("source_archive"), _binding_descriptor(bindings, "source_archive"), f"{build_role} source archive binding")
    if build_role == "candidate":
        _same_descriptor(build.get("candidate_freeze"), _binding_descriptor(bindings, "candidate_freeze"), "candidate build freeze binding")
        candidate_freeze = _read(_binding_path(bindings, "candidate_freeze"))
        _same_descriptor(build.get("candidate_overlay"), candidate_freeze.get("candidate_patch"), "candidate build overlay binding")
        _validate_candidate_source_receipt(build, bindings, candidate_freeze, "source_before")
        _validate_candidate_source_receipt(build, bindings, candidate_freeze, "source_after")
    else:
        _validate_baseline_source_receipt(build, bindings, "source_before")
        _validate_baseline_source_receipt(build, bindings, "source_after")

    binaries = build.get("binaries")
    _require(isinstance(binaries, dict) and set(binaries) == {"native", "allocator"}, f"{build_role} build binary map is incomplete")
    plan_side = plan.get("binaries", {}).get("before" if build_role == "baseline" else "after", {})
    for mode, expected_name, expected_features, binding_name in (
        ("native", "litchi-perf-baseline", [], f"{build_role}_native_depfile"),
        ("allocator", "litchi-perf-baseline-alloc", ["allocator-metrics"], f"{build_role}_allocator_depfile"),
    ):
        item = binaries.get(mode)
        _require(isinstance(item, dict), f"{build_role}/{mode} build metadata is missing")
        _require(item.get("role") == mode and item.get("name") == expected_name and item.get("features") == expected_features, f"{build_role}/{mode} build role metadata mismatch")
        _require(isinstance(item.get("elf_build_id"), str) and item["elf_build_id"], f"{build_role}/{mode} ELF build ID is missing")
        binary_path = _regular_path(item.get("path"), PACKET, f"{build_role}/{mode} binary")
        _require(binary_path == (PACKET / "binaries" / f"{build_role}-{mode}").resolve(), f"{build_role}/{mode} binary is not the retained packet artifact")
        _same_descriptor(item, descriptor(binary_path), f"{build_role}/{mode} binary")
        depfile = item.get("depfile")
        _same_descriptor(depfile, _binding_descriptor(bindings, binding_name), f"{build_role}/{mode} depfile")
        build_receipt = item.get("build_receipt")
        _require(isinstance(build_receipt, dict), f"{build_role}/{mode} role build receipt is missing")
        role_receipt_path = _regular_path(build_receipt.get("path"), PACKET, f"{build_role}/{mode} role build receipt")
        expected_role_receipt = PACKET / (f"{mode}-build.json" if build_role == "baseline" else f"candidate-{mode}-build.json")
        _require(role_receipt_path == expected_role_receipt.resolve(), f"{build_role}/{mode} role receipt is not the packet terminal receipt")
        _same_descriptor(build_receipt, descriptor(role_receipt_path), f"{build_role}/{mode} role build receipt")
        role_receipt = _read(role_receipt_path)
        _require(role_receipt.get("status") == "pass" and role_receipt.get("role") == mode, f"{build_role}/{mode} role build receipt status/role mismatch")
        _require(role_receipt.get("binary") == expected_name and role_receipt.get("features") == expected_features, f"{build_role}/{mode} role build receipt name/features mismatch")
        plan_item = _entry(plan_side.get(mode))
        for key in ("path", "bytes", "sha256", "mode", "executable", "role", "name", "features", "elf_build_id", "build_receipt", "depfile"):
            _require(plan_item.get(key) == item.get(key), f"plan {build_role}/{mode} {key} is not linked to build receipt")
        expected_role = "before" if build_role == "baseline" else "after"
        _require(plan_item.get("evidence_bindings") == binary_evidence_bindings(expected_role, mode), f"plan {build_role}/{mode} evidence bindings are incomplete")

    # A flags binding may be the final build receipt itself, or a dedicated
    # JSON sidecar.  In either case its manifest metadata and content must bind
    # to this exact build and exact settings.
    flags_name = f"{build_role}_flags"
    flags_path = _binding_path(bindings, flags_name)
    flags_value: Any
    try:
        flags_value = _read(flags_path)
    except (OSError, ValueError):
        flags_value = None
    if flags_path == build_path:
        _require(flags_value == build, f"{flags_name} does not equal its build receipt")
    else:
        _require(isinstance(flags_value, dict), f"{flags_name} must be a JSON build-flags sidecar")
        _require(flags_value.get("flags") == BUILD_CONFIGURATION, f"{flags_name} settings mismatch")
        linkage = flags_value.get("build_receipt") or flags_value.get("build")
        _same_descriptor(linkage, _binding_descriptor(bindings, "baseline_build" if build_role == "baseline" else "candidate_build"), f"{flags_name} build linkage")


def _validate_quality_binding(bindings: dict[str, Any], name: str, role: str, candidate_freeze: dict[str, Any] | None = None) -> None:
    path, value = _json_binding(bindings, name)
    _require(value.get("status") == "pass", f"{name} is not passing")
    if role == "baseline":
        _require(value.get("schema") == "litchi.xlsx-producer-shape-quality-0893.v1", f"{name} schema mismatch")
        _require(value.get("source_head_before") == BASELINE_QUALITY_HEAD and value.get("source_head_after") == BASELINE_QUALITY_HEAD, f"{name} is not the retained 0893 quality-v3 receipt")
        affected = value.get("affected_source_paths")
        _require(affected == list(BASELINE_QUALITY_SOURCE_HASHES), f"{name} affected source inventory mismatch")
        before = value.get("affected_sources_before")
        after = value.get("affected_sources_after")
        for records in (before, after):
            _require(isinstance(records, list), f"{name} source snapshots are missing")
            observed = {record.get("relative_path"): record for record in records if isinstance(record, dict)}
            _require(set(observed) == set(BASELINE_QUALITY_SOURCE_HASHES), f"{name} source snapshot paths mismatch")
            for relative_path, expected_sha in BASELINE_QUALITY_SOURCE_HASHES.items():
                record = observed[relative_path]
                source = record.get("source") or record.get("snapshot") or record
                _require(source.get("sha256") == expected_sha, f"{name} {relative_path} hash mismatch")
    else:
        _require(value.get("schema") == "litchi.xlsx-selected-metadata-quality-receipt-0894.v1", f"{name} schema mismatch")
        paths = value.get("candidate_source_paths")
        # The freeze keeps the candidate helper's source order.  The quality
        # helper records its own traversal order, so bind the identity as a
        # unique exact set while retaining the receipt's recorded order.
        _require(
            isinstance(paths, list)
            and len(paths) == len(CANDIDATE_PATHS)
            and len(set(paths)) == len(CANDIDATE_PATHS)
            and set(paths) == set(CANDIDATE_PATHS),
            f"{name} candidate source inventory mismatch",
        )
        quality_root = value.get("quality_source_root_binding")
        _require(isinstance(quality_root, dict), f"{name} isolated quality source binding is missing")
        _require(quality_root.get("path") == value.get("quality_source_root"), f"{name} isolated quality source root linkage mismatch")
        bound_paths = quality_root.get("candidate_source_paths")
        _require(
            isinstance(bound_paths, list)
            and len(bound_paths) == len(CANDIDATE_PATHS)
            and len(set(bound_paths)) == len(CANDIDATE_PATHS)
            and set(bound_paths) == set(CANDIDATE_PATHS),
            f"{name} isolated quality source inventory mismatch",
        )
        for descriptor_name in ("manifest", "perf_manifest", "lockfile"):
            item = quality_root.get(descriptor_name)
            _require(
                isinstance(item, dict)
                and isinstance(item.get("path"), str)
                and isinstance(item.get("bytes"), int)
                and isinstance(item.get("sha256"), str)
                and isinstance(item.get("mode"), int)
                and isinstance(item.get("executable"), bool),
                f"{name} isolated quality {descriptor_name} identity is missing",
            )
        commands = value.get("commands")
        expected_commands = [
            "rustfmt_candidate_sources", "focused_producer_metadata_tests",
            "xlsx_lib_tests", "xlsx_all_features_check", "xlsx_clippy_all_features",
            "xlsx_rustdoc", "producer_shape_harness_tests",
            "producer_full_cell_selector_test", "crate_boundaries",
        ]
        _require(isinstance(commands, list) and [item.get("name") for item in commands] == expected_commands, f"{name} nine-gate command inventory mismatch")
        _require(all(item.get("pass") is True and item.get("returncode") == 0 for item in commands), f"{name} quality gate did not pass")
        _require(value.get("failures") == [], f"{name} reports quality failures")
        _require(
            value.get("source_head_before") == BASE_COMMIT
            and value.get("source_head_after") == BASE_COMMIT,
            f"{name} quality source HEAD is not the frozen 506aa2eda baseline",
        )
        _require(isinstance(candidate_freeze, dict), f"{name} candidate freeze identity is missing")
        expected_files = candidate_freeze.get("files")
        _require(isinstance(expected_files, dict), f"{name} candidate freeze files are missing")
        for label in ("candidate_sources_before", "candidate_sources_after"):
            records = value.get(label)
            _require(isinstance(records, list), f"{name} {label} is missing")
            observed = {record.get("relative_path"): record for record in records if isinstance(record, dict)}
            _require(set(observed) == set(CANDIDATE_PATHS), f"{name} {label} path inventory mismatch")
            for relative_path in CANDIDATE_PATHS:
                expected = expected_files[relative_path]["current"]
                source = observed[relative_path].get("source")
                _require(isinstance(source, dict) and source.get("sha256") == expected.get("sha256") and source.get("bytes") == expected.get("bytes"), f"{name} {label} identity mismatch: {relative_path}")
    _require(path.is_file(), f"{name} disappeared")


def _validate_quality_source_binding(bindings: dict[str, Any], candidate_freeze: dict[str, Any]) -> None:
    path, value = _json_binding(bindings, "quality_source")
    _require(path.resolve() == (PACKET / "quality-source.json").resolve(), "quality source receipt is not the packet evidence")
    _require(value.get("schema") == QUALITY_SOURCE_SCHEMA and value.get("status") == "pass", "quality source receipt schema/status mismatch")
    _require(value.get("baseline_revision") == BASE_COMMIT, "quality source baseline revision mismatch")
    _require(value.get("exit_code") == 0 and value.get("shared_root_lock_consumed") is False, "quality source preparation did not pass in isolation")

    def identity_subset(actual: Any, expected: Any, label: str, keys: tuple[str, ...] = ("path", "bytes", "sha256")) -> None:
        _require(isinstance(actual, dict) and isinstance(expected, dict), f"{label} identity is missing")
        for key in keys:
            _require(actual.get(key) == expected.get(key), f"{label} {key} mismatch")

    identity_subset(value.get("archive"), _binding_descriptor(bindings, "source_archive"), "quality source archive")
    frozen_files = candidate_freeze.get("files")
    overlays = value.get("overlays")
    _require(isinstance(frozen_files, dict) and set(frozen_files) == set(CANDIDATE_PATHS), "quality source candidate freeze files are missing")
    _require(isinstance(overlays, dict) and set(overlays) == set(CANDIDATE_PATHS), "quality source overlay inventory mismatch")
    for relative_path in CANDIDATE_PATHS:
        identity_subset(overlays[relative_path], frozen_files[relative_path]["current"], f"quality source overlay {relative_path}", ("bytes", "sha256"))

    support = value.get("support")
    _require(isinstance(support, dict) and set(support) == set(QUALITY_SUPPORT_PATHS), "quality source support inventory mismatch")
    for relative_path, item in support.items():
        _require(
            isinstance(item, dict)
            and isinstance(item.get("path"), str)
            and isinstance(item.get("bytes"), int)
            and isinstance(item.get("sha256"), str),
            f"quality source support identity is incomplete: {relative_path}",
        )

    lock = value.get("lock")
    _require(isinstance(lock, dict), "quality source private lock identity is missing")
    _require(
        isinstance(lock.get("path"), str)
        and isinstance(lock.get("bytes"), int)
        and isinstance(lock.get("sha256"), str),
        "quality source private lock identity is incomplete",
    )
    _quality_path, quality = _json_binding(bindings, "candidate_quality")
    quality_root = quality.get("quality_source_root_binding")
    _require(isinstance(quality_root, dict), "candidate quality isolated source binding is missing")
    _require(quality_root.get("path") == value.get("source_root"), "candidate quality source root differs from quality source receipt")
    identity_subset(quality_root.get("lockfile"), lock, "candidate quality private lock")

    source_manifest = _read(_binding_path(bindings, "source_manifest"))
    source_files = source_manifest.get("files")
    _require(isinstance(source_files, dict), "quality source cannot resolve source manifest files")
    for descriptor_name, relative_path in (
        ("manifest", "crates/litchi-xlsx/Cargo.toml"),
        ("perf_manifest", "tools/perf-baseline/Cargo.toml"),
    ):
        item = quality_root.get(descriptor_name)
        expected = source_files.get(relative_path)
        _require(isinstance(item, dict) and isinstance(expected, dict), f"candidate quality {descriptor_name} identity is missing")
        for key in ("bytes", "sha256", "mode"):
            _require(item.get(key) == expected.get(key), f"candidate quality {descriptor_name} {key} differs from sealed archive")
        _require(item.get("executable") == bool(expected.get("mode", 0) & stat.S_IXUSR), f"candidate quality {descriptor_name} executable bit mismatch")

    runner = quality.get("runner")
    helper = descriptor(HELPER_PATHS["quality.py"])
    for key in ("bytes", "sha256"):
        _require(isinstance(runner, dict) and runner.get(key) == helper[key], f"candidate quality runner copy {key} differs from sealed quality.py")


def _validate_baseline_source_receipt(build: dict[str, Any], bindings: dict[str, Any], label: str) -> None:
    raw = build.get(label)
    _require(isinstance(raw, dict), f"baseline build {label} descriptor is missing")
    path = _regular_path(raw.get("path"), PACKET, f"baseline build {label}")
    expected_name = "source-before.json" if label == "source_before" else "source-after.json"
    _require(path == (PACKET / expected_name).resolve(), f"baseline build {label} is not the packet snapshot")
    _same_descriptor(raw, descriptor(path), f"baseline build {label}")
    receipt = _read(path)
    expected_schema = "litchi.xlsx-source-before-0894.v1" if label == "source_before" else "litchi.xlsx-source-after-0894.v1"
    _require(receipt.get("schema") == expected_schema and receipt.get("status") == "pass", f"baseline build {label} schema/status mismatch")
    _require(receipt.get("file_count") == SOURCE_FILE_COUNT, f"baseline build {label} file census mismatch")
    source_manifest = _read(_binding_path(bindings, "source_manifest"))
    expected_files = source_manifest.get("files")
    files = receipt.get("files")
    _require(isinstance(expected_files, dict) and len(expected_files) == SOURCE_FILE_COUNT, "baseline source manifest file census is incomplete")
    _require(isinstance(files, dict) and set(files) == set(expected_files), f"baseline build {label} source path census differs from source manifest")
    for relative_path, expected in expected_files.items():
        observed = files.get(relative_path)
        _require(isinstance(observed, dict), f"baseline build {label} omits {relative_path}")
        for key in ("bytes", "sha256", "mode"):
            _require(observed.get(key) == expected.get(key), f"baseline build {label} {relative_path} {key} mismatch")


def _validate_candidate_source_receipt(build: dict[str, Any], bindings: dict[str, Any], candidate_freeze: dict[str, Any], label: str) -> None:
    raw = build.get(label)
    _require(isinstance(raw, dict), f"candidate build {label} descriptor is missing")
    path = _regular_path(raw.get("path"), PACKET, f"candidate build {label}")
    expected_name = "candidate-source-before.json" if label == "source_before" else "candidate-source-after.json"
    _require(path == (PACKET / expected_name).resolve(), f"candidate build {label} is not the packet snapshot")
    _same_descriptor(raw, descriptor(path), f"candidate build {label}")
    receipt = _read(path)
    expected_schema = "litchi.xlsx-candidate-source-before-0894.v1" if label == "source_before" else "litchi.xlsx-candidate-source-after-0894.v1"
    _require(receipt.get("schema") == expected_schema and receipt.get("status") == "pass", f"candidate build {label} schema/status mismatch")
    _require(receipt.get("file_count") == SOURCE_FILE_COUNT + 2, f"candidate build {label} file census mismatch")
    files = receipt.get("files")
    frozen_files = candidate_freeze.get("files")
    _require(isinstance(files, dict) and len(files) == SOURCE_FILE_COUNT + 2, f"candidate build {label} source census is incomplete")
    _require(isinstance(frozen_files, dict) and set(frozen_files) == set(CANDIDATE_PATHS), "candidate freeze file identities are missing")
    source_manifest_path = _binding_path(bindings, "source_manifest")
    source_manifest = _read(source_manifest_path)
    baseline_files = source_manifest.get("files")
    _require(isinstance(baseline_files, dict) and len(baseline_files) == SOURCE_FILE_COUNT, "candidate build baseline source census is incomplete")
    expected_files = dict(baseline_files)
    for relative_path in CANDIDATE_PATHS:
        expected_files[relative_path] = frozen_files[relative_path]["current"]
    _require(set(files) == set(expected_files), f"candidate build {label} source path census is not the frozen baseline plus overlay")
    for relative_path, expected in expected_files.items():
        observed = files.get(relative_path)
        _require(isinstance(observed, dict), f"candidate build {label} omits {relative_path}")
        for key in ("bytes", "sha256", "mode"):
            _require(observed.get(key) == expected.get(key), f"candidate build {label} {relative_path} {key} mismatch")


def _validate_review_binding(bindings: dict[str, Any], name: str) -> None:
    _path, value = _json_binding(bindings, name)
    _require(value.get("status") in ("pass", "approved", "reviewed", "approved_for_measurement"), f"{name} has no approved status")
    if name == "source_review":
        source = value.get("source")
        _require(isinstance(source, dict), "source review source identity is missing")
        _require(source.get("expected_head_from_quality") == BASELINE_QUALITY_HEAD and source.get("head_matches_quality") is True, "source review quality HEAD is not the retained 0893 v3 HEAD")
        _require(source.get("affected_rust_sources") == list(BASELINE_QUALITY_SOURCE_HASHES), "source review affected-source inventory mismatch")
        files = source.get("files")
        _require(isinstance(files, dict), "source review source file identities are missing")
        for relative_path, expected_sha in BASELINE_QUALITY_SOURCE_HASHES.items():
            _require(files.get(relative_path, {}).get("sha256") == expected_sha, f"source review {relative_path} hash mismatch")
        evidence = value.get("guard_evidence")
        _require(isinstance(evidence, dict) and evidence.get("contract") == GUARD_CONTRACT, "source review guard contract is not exact")
        qualification = value.get("qualification", {}).get("verifications", {}).get("qualification-after-cleanup", {})
        _require(qualification.get("status") == "pass" and qualification.get("schema") == "litchi.full-cell-guard-qualification-0893.v1", "source review retained qualification is not passing")
    elif name == "candidate_review":
        schema = str(value.get("schema", ""))
        _require("0894" in schema and "review" in schema.lower(), "candidate review schema is not packet-specific")
        _require(value.get("base_commit") == BASE_COMMIT or value.get("candidate", {}).get("base_commit") == BASE_COMMIT, "candidate review base commit is missing")
        paths = value.get("candidate_paths") or value.get("candidate", {}).get("candidate_paths")
        _require(
            isinstance(paths, list)
            and len(paths) == len(CANDIDATE_PATHS)
            and len(set(paths)) == len(CANDIDATE_PATHS)
            and set(paths) == set(CANDIDATE_PATHS),
            "candidate review source inventory mismatch",
        )
        freeze = _read(_binding_path(bindings, "candidate_freeze"))
        frozen_files = freeze.get("files")
        _require(isinstance(frozen_files, dict), "candidate review cannot resolve candidate freeze files")
        hashes = value.get("candidate_source_hashes") or value.get("candidate", {}).get("candidate_source_hashes")
        if isinstance(hashes, dict):
            _require(set(hashes) == set(CANDIDATE_PATHS), "candidate review source hash inventory mismatch")
            for relative_path in CANDIDATE_PATHS:
                _require(hashes.get(relative_path) == frozen_files[relative_path]["current"].get("sha256"), f"candidate review source hash mismatch: {relative_path}")
        else:
            records = value.get("candidate_sources") or value.get("candidate", {}).get("candidate_sources")
            _require(isinstance(records, list), "candidate review source hashes are missing")
            observed = {record.get("relative_path"): record for record in records if isinstance(record, dict)}
            _require(set(observed) == set(CANDIDATE_PATHS), "candidate review source hash records are incomplete")
            for relative_path in CANDIDATE_PATHS:
                source = observed[relative_path].get("source") or observed[relative_path]
                _require(source.get("sha256") == frozen_files[relative_path]["current"].get("sha256"), f"candidate review source hash mismatch: {relative_path}")
    else:
        schema = str(value.get("schema", ""))
        _require("0894" in schema and "methodology" in schema.lower(), "methodology review schema is not packet-specific")
        guard = value.get("guard_contract") or value.get("timed_harness_guard_contract") or value.get("guard", {}).get("contract")
        _require(guard == GUARD_CONTRACT, "methodology review guard contract is not exact")
        guard_sha = value.get("guard_test_sha256") or value.get("guard_tests_sha256") or value.get("guard", {}).get("test_sha256")
        _require(guard_sha == GUARD_TEST_SHA256, "methodology review guard-test source hash is not the 0893 v3 hash")
        protocol = value.get("protocol") or value.get("schedule")
        _require(isinstance(protocol, dict), "methodology review protocol is missing")
        _require(protocol.get("total_children") == TOTAL_CHILDREN or (protocol.get("native", {}).get("children") == TOTAL_NATIVE_CHILDREN and protocol.get("allocator", {}).get("children") == TOTAL_ALLOCATOR_CHILDREN), "methodology review child protocol mismatch")
        review_helpers = value.get("helper_hashes")
        _require(isinstance(review_helpers, dict) and set(review_helpers) == set(HELPER_PATHS), "methodology review helper hashes are incomplete")
        for helper_name, helper_identity in helper_descriptors().items():
            _require(review_helpers.get(helper_name) == helper_identity.get("sha256"), f"methodology review helper hash mismatch: {helper_name}")
        _require(value.get("performance_claim") is None, "methodology review contains a performance claim")


def _validate_bindings(plan: dict[str, Any]) -> None:
    bindings = plan.get("bindings")
    _require(isinstance(bindings, dict), "plan evidence bindings are missing")
    _require(set(bindings) == set(REQUIRED_BINDINGS), "plan evidence binding set is incomplete or unexpected")
    for name in REQUIRED_BINDINGS:
        item = bindings.get(name)
        _require(isinstance(item, dict), f"binding {name} must include role metadata")
        expected = BINDING_SPECS[name]
        for key, value in expected.items():
            _require(item.get(key) == value, f"binding {name} {key} is not {value!r}")
        _binding_descriptor(bindings, name)

    source = plan.get("source")
    _require(isinstance(source, dict), "plan source contract is missing")
    source_path = _regular_path(source.get("receipt", {}).get("path"), PACKET, "plan source receipt")
    _require(source_path == (PACKET / "source.json").resolve(), "plan source receipt is not the packet manifest")
    source_receipt = _read(source_path)
    _require(isinstance(source_receipt, dict), "plan source receipt is not an object")
    source_files = source_receipt.get("files")
    _require(isinstance(source_files, dict), "source receipt file census is missing")
    for relative_path, expected_sha in BASELINE_QUALITY_SOURCE_HASHES.items():
        _require(source_files.get(relative_path, {}).get("sha256") == expected_sha, f"source archive {relative_path} does not match 0893 v3 quality")
    _same_descriptor(bindings["source_manifest"], descriptor(source_path), "source manifest binding")
    _same_descriptor(bindings["source_identity"], descriptor(source_path), "source identity binding")
    archive = source_receipt.get("archive")
    _require(isinstance(archive, dict), "source receipt archive is missing")
    archive_path = _regular_path(archive.get("path"), source_path.parent, "source archive")
    _require(archive_path == (PACKET / "source.tar.gz").resolve(), "source archive path is not the packet archive")
    _same_descriptor(bindings["source_archive"], descriptor(archive_path), "source archive binding")
    _same_descriptor(archive, descriptor(archive_path), "source archive receipt")

    candidate_path = _binding_path(bindings, "candidate_freeze")
    candidate_receipt = _candidate_freeze_contract(candidate_path, source_receipt, source_path, descriptor(archive_path))
    _same_descriptor(plan.get("source", {}).get("candidate_overlay"), candidate_receipt.get("candidate_patch"), "plan candidate overlay")

    for build_role, ref_name, plan_name in (("baseline", "baseline_build", "before"), ("candidate", "candidate_build", "after")):
        build_path = _binding_path(bindings, ref_name)
        _same_descriptor(plan.get("builds", {}).get(plan_name), descriptor(build_path), f"plan {build_role} build binding")
        build = _read(build_path)
        _require(isinstance(build, dict), f"{build_role} build receipt is not an object")
        _role_build_contract(plan, bindings, build_role, build_path, build)
    _validate_quality_binding(bindings, "baseline_quality", "baseline")
    _validate_quality_binding(bindings, "candidate_quality", "candidate", candidate_receipt)
    _validate_quality_source_binding(bindings, candidate_receipt)
    _validate_review_binding(bindings, "source_review")
    _validate_review_binding(bindings, "candidate_review")
    _validate_review_binding(bindings, "methodology_review")


def _load_expected_corpora() -> dict[str, dict[str, Any]]:
    expected: dict[str, dict[str, Any]] = {}
    for case in CASES:
        folder = BASELINE_QUALIFICATION / case
        report_path = folder / "report.json"
        _require(report_path.is_file(), f"sealed baseline qualification is missing {report_path}")
        report = _read(report_path)
        results = report.get("results")
        _require(type(results) is list and len(results) == 1 and results[0].get("case") == case, f"baseline case mismatch: {case}")
        elapsed = results[0].get("elapsed_ns", {})
        _require(type(elapsed.get("samples")) is list and elapsed["samples"] == sorted(elapsed["samples"]), f"baseline elapsed vector invalid: {case}")
        producer_path = folder / "producer.json"
        producer = _read(producer_path) if producer_path.is_file() else None
        if producer is not None:
            _require(producer.get("schema") == "litchi.perf-baseline.producer-shape-evidence.v1", f"producer schema mismatch: {case}")
        if case in (PRIMARY_CASE, CASES[2], CASES[3]):
            _require(producer is not None, f"producer sidecar missing: {case}")
        expected[case] = {
            "corpus": results[0].get("corpus"),
            "report_sha256": _sha(report_path),
            "producer": producer,
            "producer_sha256": _sha(producer_path) if producer is not None else None,
        }
    return expected


def _guard_reference() -> dict[str, Any]:
    _require(GUARD_REFERENCE.is_file(), f"0893 guard reference is missing: {GUARD_REFERENCE}")
    report = _read(GUARD_REFERENCE)
    scan = report.get("source", {}).get("producer_full_cell_scan")
    _require(isinstance(scan, dict), "0893 reference lacks producer scan")
    return scan


def validate_plan(plan: dict[str, Any]) -> None:
    _require(plan.get("schema") == SCHEMA, f"unexpected plan schema: {plan.get('schema')!r}")
    _require(plan.get("status") == "frozen", "0894 plan is not frozen")
    _require(plan.get("profile") == PROFILE, "0894 requires release binaries")
    _require(str(plan.get("cpu_affinity")) == CPU_AFFINITY, "0894 is pinned to CPU 12")
    _require(plan.get("cache") == "warm", "0894 requires warm cache")
    _require(plan.get("cases") == list(CASES), "0894 case order is not frozen")
    _require(plan.get("base_commit") == BASE_COMMIT, "0894 baseline commit is not 506aa2eda")
    _require(plan.get("native_blocks") == NATIVE_BLOCKS, "native block count mismatch")
    _require(plan.get("native_warmup") == NATIVE_WARMUP and plan.get("native_samples") == NATIVE_SAMPLES, "native protocol mismatch")
    _require(plan.get("allocator_blocks") == ALLOCATOR_BLOCKS, "allocator block count mismatch")
    _require(plan.get("allocator_warmup") == ALLOCATOR_WARMUP and plan.get("allocator_samples") == ALLOCATOR_SAMPLES, "allocator protocol mismatch")
    _require(plan.get("bootstrap_seed") == BOOTSTRAP_SEED, "bootstrap seed mismatch")
    _require(plan.get("bootstrap_draws") == BOOTSTRAP_DRAWS and plan.get("bootstrap_ci_indices") == list(BOOTSTRAP_CI_INDICES), "bootstrap protocol mismatch")
    _require(plan.get("fresh_child") is True and plan.get("fail_fast") is True, "0894 requires fresh-child fail-fast execution")
    _require(plan.get("guard_contract") == GUARD_CONTRACT, "0893 guard contract is not bound")
    build_configuration = plan.get("build_configuration")
    _require(build_configuration == BUILD_CONFIGURATION, "0894 build configuration is not exact")
    source = plan.get("source")
    _require(isinstance(source, dict), "source baseline identity is missing")
    _require(source.get("schema") == SOURCE_SCHEMA and source.get("status") == "pass", "source snapshot schema/status is not bound")
    for key in ("baseline_commit", "source_revision", "revision"):
        _require(source.get(key) == BASE_COMMIT, f"source {key} identity is missing")
    _require(source.get("timed_harness_guard_contract") == GUARD_CONTRACT, "timed harness guard identity is missing")
    _require(source.get("selected_file_count") == SOURCE_FILE_COUNT and source.get("source_file_count") == SOURCE_FILE_COUNT, "source file census is not exact")
    _require(source.get("archive_sha256") == SOURCE_ARCHIVE_SHA256 and source.get("archive_bytes") == SOURCE_ARCHIVE_BYTES, "source archive identity is missing")
    _require(isinstance(source.get("candidate_overlay"), dict), "candidate overlay identity is missing")
    _require(source.get("selection") == SOURCE_SELECTION, "source selection is not exact")
    _require(isinstance(source.get("receipt"), dict), "source receipt descriptor is missing")
    _validate_fixture_bindings(plan)
    _validate_helper_bindings(plan)
    binaries = plan.get("binaries")
    _require(isinstance(binaries, dict), "plan binaries are missing")
    for role in ("before", "after"):
        for mode in ("native", "allocator"):
            item = _entry(binaries.get(role, {}).get(mode))
            _require(isinstance(item.get("path"), str), f"missing {role}/{mode} binary")
            _require(isinstance(item.get("sha256"), str) and isinstance(item.get("bytes"), int), f"incomplete {role}/{mode} binary identity")
            _require(item.get("executable") is True, f"{role}/{mode} binary is not executable")
            _require(item.get("role") == mode, f"{role}/{mode} binary role metadata is missing")
            expected_name = "litchi-perf-baseline" if mode == "native" else "litchi-perf-baseline-alloc"
            expected_features = [] if mode == "native" else ["allocator-metrics"]
            _require(item.get("name") == expected_name and item.get("features") == expected_features, f"{role}/{mode} binary name/features metadata mismatch")
            _require(isinstance(item.get("build_receipt"), dict) and isinstance(item.get("depfile"), dict), f"{role}/{mode} build/depfile binding metadata missing")
            _require(isinstance(item.get("elf_build_id"), str) and item.get("elf_build_id"), f"{role}/{mode} ELF build metadata missing")
            _require(item.get("evidence_bindings") == binary_evidence_bindings(role, mode), f"{role}/{mode} evidence binding map is incomplete")
    _require(plan.get("total_children") == TOTAL_CHILDREN, "total child count mismatch")


def load_plan(path: Path) -> dict[str, Any]:
    plan = _read(path)
    _require(isinstance(plan, dict), "plan must be an object")
    validate_plan(plan)
    _validate_bindings(plan)
    _load_expected_corpora()
    _guard_reference()
    for role in ("before", "after"):
        for mode in ("native", "allocator"):
            _binary(plan, role, mode)
    return plan


def _raw_stats(samples: list[int]) -> dict[str, int]:
    _require(type(samples) is list and samples and all(type(value) is int and value > 0 for value in samples), "elapsed samples must be positive integers")
    ordered = sorted(samples)
    midpoint = len(ordered) // 2
    p50 = (ordered[midpoint - 1] + ordered[midpoint]) // 2 if len(ordered) % 2 == 0 else ordered[midpoint]
    def nearest(rank: int) -> int:
        index = max(0, (len(ordered) * rank + 99) // 100 - 1)
        return ordered[index]
    return {"min": ordered[0], "p50": p50, "p95": nearest(95), "p99": nearest(99), "max": ordered[-1]}


def _validate_binary_identity(report: dict[str, Any], binary: dict[str, Any], mode: str) -> None:
    identity = report.get("binary_identity")
    _require(isinstance(identity, dict), "report binary identity is missing")
    actual = binary["identity"]
    for key, report_key in (("path", "path"), ("sha256", "binary_sha256"), ("bytes", "binary_bytes"), ("mode", "mode_bits")):
        _require(identity.get(report_key) == actual[key], f"report binary {report_key} mismatch")
    _require(identity.get("executable") is True and identity.get("profile") == PROFILE, "report binary profile/executable mismatch")
    expected_name = "litchi-perf-baseline" if mode == "native" else "litchi-perf-baseline-alloc"
    _require(report.get("tool", {}).get("binary") == expected_name, "report tool binary mismatch")


def _validate_common_report(report: dict[str, Any], case: str, mode: str, warmup: int, samples: int, expected: dict[str, Any], binary: dict[str, Any]) -> dict[str, Any]:
    _require(report.get("schema_version") == 1, "report schema version mismatch")
    tool = report.get("tool", {})
    _require(tool.get("name") == "litchi-perf-baseline" and tool.get("version") == "0.1.0", "report tool mismatch")
    _require(tool.get("profile") == PROFILE and tool.get("target_os") == "linux" and tool.get("target_arch") == "x86_64", "report target mismatch")
    expected_instrumentation = "none" if mode == "native" else "system_allocator_operation_scoped"
    _require(tool.get("instrumentation") == expected_instrumentation, "report instrumentation mismatch")
    if mode == "allocator":
        _require(tool.get("allocator_counter_revision") == "serialized_region_peak_v3", "allocator counter revision mismatch")
    config = report.get("configuration", {})
    _require(config.get("samples_per_case") == samples and config.get("warmup_iterations_per_case") == warmup, "report sample protocol mismatch")
    _require(config.get("filesystem_cache_states") == ["warm"] and config.get("filesystem_fresh_child_per_sample") is True and config.get("filesystem_process_isolated") is True, "report filesystem protocol mismatch")
    _require(config.get("xlsx_shapes") == ["medium"] and config.get("cases") == [case], "report case selection mismatch")
    _require(report.get("environment", {}).get("cpu_affinity") == CPU_AFFINITY, "report CPU affinity mismatch")
    _validate_binary_identity(report, binary, mode)
    results = report.get("results")
    _require(type(results) is list and len(results) == 1 and results[0].get("case") == case, "report result case mismatch")
    result = results[0]
    elapsed = result.get("elapsed_ns", {})
    raw = elapsed.get("samples")
    _require(type(raw) is list and len(raw) == samples, "elapsed vector length mismatch")
    order = elapsed.get("sample_order")
    _require(type(order) is list and sorted(order) == list(range(samples)), "sample-order permutation mismatch")
    stats = _raw_stats(raw)
    for key, value in stats.items():
        _require(elapsed.get(key) == value, f"elapsed {key} mismatch")
    _require(result.get("corpus") == expected["corpus"], f"{case} corpus differs from sealed fixture")
    return {"result": result, "stats": stats, "sample_order": order}


def _validate_guard(result: dict[str, Any], sample_order: list[int], warmup: int, samples: int, expected: dict[str, Any]) -> None:
    source = result.get("source")
    _require(isinstance(source, dict), "primary source evidence is missing")
    scan = source.get("producer_full_cell_scan")
    _require(isinstance(scan, dict), "primary full-cell scan evidence is missing")
    operation = result.get("operation_metrics")
    _require(isinstance(operation, dict), "primary operation metrics are missing")
    _require(operation.get("sample_count") == samples and operation.get("sample_indices") == sample_order, "operation metric sample alignment mismatch")
    reference = _guard_reference()
    _require(scan.get("guard_contract") == GUARD_CONTRACT, "0893 guard contract missing or wrong")
    _require(scan.get("sample_indices") == sample_order, "guard sample alignment mismatch")
    _require(scan.get("archive_sha256") == PRIMARY_ARCHIVE_SHA256 and scan.get("archive_bytes") == PRIMARY_ARCHIVE_BYTES, "guard archive identity mismatch")
    for field, value in (("expected_cell_count", 4096), ("expected_numeric_count", 3686), ("expected_text_count", 410)):
        _require(scan.get(field) == value, f"{field} mismatch")
    _require(scan.get("marker_refusal_census_sha256") == PRIMARY_CENSUS_SHA256, "marker census mismatch")
    _require(scan.get("oracle_digest_sha256") == PRIMARY_DIGEST_SHA256, "oracle digest mismatch")
    _require(scan.get("observed_cell_counts") == [4096] * samples, "observed cell counts mismatch")
    _require(scan.get("observed_numeric_counts") == [3686] * samples, "observed numeric counts mismatch")
    _require(scan.get("observed_text_counts") == [410] * samples, "observed text counts mismatch")
    _require(scan.get("observed_digest_sha256") == [PRIMARY_DIGEST_SHA256] * samples, "observed digest mismatch")
    _require(scan.get("source_version_stable") == [True] * samples, "post-guard source fence mismatch")
    _require(scan.get("forced_worksheet_counts") == [4] * samples, "measured worksheet guard coverage mismatch")
    _require(scan.get("forced_worksheet_tolerance_clean") == [True] * samples, "measured worksheet guard cleanliness mismatch")
    _require(scan.get("warmup_forced_worksheet_counts") == [4] * warmup, "warmup worksheet guard coverage mismatch")
    _require(scan.get("warmup_forced_worksheet_tolerance_clean") == [True] * warmup, "warmup worksheet guard cleanliness mismatch")
    _require(scan.get("forced_worksheet_guard_iterations") == warmup + samples, "worksheet guard iteration total mismatch")
    _require(scan.get("runtime_route_observed") is False and scan.get("physical_cold_eligible") is False, "route/cold scope claim mismatch")
    timing_scope = scan.get("timing_scope")
    _require(isinstance(timing_scope, str), "timing scope missing")
    for phrase in ("process/allocation endpoint snapshots", "post-timing per-worksheet tolerance guard", "aggregate tolerance report", "post-guard source-freshness fence"):
        _require(phrase in timing_scope, f"timing scope omits {phrase}")
    # Bind the static producer metadata to the reviewed 0893 shape while
    # allowing only measured vectors, sample alignment, and guard fields to
    # differ across children.
    dynamic = {
        "sample_indices", "observed_cell_counts", "observed_numeric_counts",
        "observed_text_counts", "observed_digest_sha256", "source_version_stable",
        "forced_worksheet_counts", "forced_worksheet_tolerance_clean",
        "warmup_forced_worksheet_counts", "warmup_forced_worksheet_tolerance_clean",
        "forced_worksheet_guard_iterations", "timing_scope",
    }
    for key, value in reference.items():
        if key not in dynamic:
            _require(scan.get(key) == value, f"guard static field {key} differs from 0893 reference")
    expected_scan = expected.get("guard_static")
    if expected_scan is not None:
        for key, value in expected_scan.items():
            if key not in dynamic:
                _require(scan.get(key) == value, f"guard static field {key} differs from sealed source")


def _validate_child(report: dict[str, Any], producer: dict[str, Any] | None, case: str, mode: str, warmup: int, samples: int, expected: dict[str, Any], binary: dict[str, Any]) -> dict[str, Any]:
    checked = _validate_common_report(report, case, mode, warmup, samples, expected, binary)
    result = checked["result"]
    if case == PRIMARY_CASE:
        _require(producer is not None, "primary producer sidecar is missing")
        _require(producer == expected["producer"], "primary producer sidecar differs from sealed fixture")
        _validate_guard(result, checked["sample_order"], warmup, samples, expected)
    elif expected["producer"] is None:
        _require(producer is None, f"{case} unexpectedly emitted producer evidence")
    else:
        _require(producer == expected["producer"], f"{case} producer sidecar differs from sealed fixture")
    if case != PRIMARY_CASE:
        _require(result.get("operation_metrics") in (None, {}), f"control {case} unexpectedly has operation metrics")
    allocation = result.get("operation_metrics", {}).get("allocation") if mode == "allocator" else None
    if mode == "allocator":
        _require(isinstance(allocation, dict) and allocation.get("status") == "measured", "allocator metrics are missing")
    return {"stats": checked["stats"], "result": result, "allocation": allocation}


def _read_peak(path: Path) -> int:
    _require(path.is_file(), f"peak RSS output missing: {path}")
    value = path.read_text(encoding="utf-8").strip()
    _require(re.fullmatch(r"[0-9]+", value) is not None, f"peak RSS output is not one integer: {path}")
    return int(value)


def _artifacts(folder: Path) -> dict[str, dict[str, Any]]:
    result = {}
    for name in ("stdout.txt", "stderr.txt", "report.json", "peak-rss-kib.txt", "producer.json"):
        path = folder / name
        if path.is_file():
            result[name] = descriptor(path)
    return result


def _command(binary: str, case: str, warmup: int, samples: int, report: Path, producer: Path, peak: Path) -> list[str]:
    return [
        TIME, "-f", "%M", "-o", str(peak), TASKSET, "-c", CPU_AFFINITY,
        binary, "--case", case, "--xlsx-shape", "medium", "--warmup", str(warmup),
        "--samples", str(samples), "--filesystem-cache", "warm", "--json",
        str(report), "--producer-evidence", str(producer),
    ]


def _schedule() -> list[dict[str, Any]]:
    rows = []
    ordinal = 0
    for block in range(NATIVE_BLOCKS):
        for case_index, case in enumerate(CASES):
            order = ("before", "after") if (block + case_index) % 2 == 0 else ("after", "before")
            pair = f"native-b{block:02d}-c{case_index:02d}"
            for role in order:
                rows.append({"ordinal": ordinal, "lane": "native", "block": block, "case_index": case_index, "case": case, "role": role, "pair": pair, "warmup": NATIVE_WARMUP, "samples": NATIVE_SAMPLES})
                ordinal += 1
    for block in range(ALLOCATOR_BLOCKS):
        order = ("before", "after") if block % 2 == 0 else ("after", "before")
        pair = f"allocator-b{block:02d}-c00"
        for role in order:
            rows.append({"ordinal": ordinal, "lane": "allocator", "block": block, "case_index": 0, "case": PRIMARY_CASE, "role": role, "pair": pair, "warmup": ALLOCATOR_WARMUP, "samples": ALLOCATOR_SAMPLES})
            ordinal += 1
    return rows


def validate_schedule(rows: list[dict[str, Any]]) -> None:
    expected = _schedule()
    _require(rows == expected, "experiment schedule differs from frozen alternating AB/BA schedule")
    _require(len(rows) == TOTAL_CHILDREN, "child population mismatch")
    _require(sum(row["lane"] == "native" for row in rows) == TOTAL_NATIVE_CHILDREN, "native population mismatch")
    _require(sum(row["lane"] == "allocator" for row in rows) == TOTAL_ALLOCATOR_CHILDREN, "allocator population mismatch")


def _parse_time(value: str) -> _datetime.datetime:
    return _datetime.datetime.fromisoformat(value)


def run(plan_path: Path, output_root: Path, *, popen: Any = subprocess.Popen) -> dict[str, Any]:
    plan = load_plan(plan_path)
    expected_corpora = _load_expected_corpora()
    if output_root.exists():
        raise FileExistsError(f"experiment output already exists: {output_root}")
    output_root.mkdir(parents=True)
    rows = _schedule()
    validate_schedule(rows)
    binaries = {(role, mode): _binary(plan, role, mode) for role in ("before", "after") for mode in ("native", "allocator")}
    receipts = []
    previous_finished: _datetime.datetime | None = None
    for row in rows:
        mode = "native" if row["lane"] == "native" else "allocator"
        binary = binaries[(row["role"], mode)]
        folder = output_root / "runs" / f"{row['ordinal']:03d}-{row['lane']}-{row['pair']}-{row['role']}"
        folder.mkdir(parents=True, exist_ok=False)
        report_path = folder / "report.json"
        producer_path = folder / "producer.json"
        peak_path = folder / "peak-rss-kib.txt"
        stdout_path = folder / "stdout.txt"
        stderr_path = folder / "stderr.txt"
        argv = _command(binary["path"], row["case"], row["warmup"], row["samples"], report_path, producer_path, peak_path)
        started = _utc_now()
        command = {"schema": COMMAND_SCHEMA, "row": row, "argv": argv, "cwd": str(PACKET), "environment": CHILD_ENV, "binary": binary, "plan_sha256": _sha(plan_path), "started_utc": started}
        with stdout_path.open("xb") as stdout, stderr_path.open("xb") as stderr:
            process = popen(argv, cwd=PACKET, env=CHILD_ENV.copy(), stdout=stdout, stderr=stderr)
            exit_code = process.wait()
        finished = _utc_now()
        command.update({"finished_utc": finished, "exit_code": exit_code})
        _write_json(folder / "command.json", command)
        receipt = {"schema": RECEIPT_SCHEMA, "row": row, "argv": argv, "cwd": str(PACKET), "environment": CHILD_ENV, "binary": binary, "plan_sha256": command["plan_sha256"], "started_utc": started, "finished_utc": finished, "exit_code": exit_code, "command": descriptor(folder / "command.json"), "artifacts": _artifacts(folder)}
        _write_json(folder / "receipt.json", receipt)
        receipts.append({"path": str((folder / "receipt.json").relative_to(output_root)), "row": row})
        if exit_code != 0:
            raise RuntimeError(f"child {row['ordinal']} exited {exit_code}; capture stops without retry")
        current_started = _parse_time(started)
        if previous_finished is not None:
            _require(current_started >= previous_finished, f"children were not launched sequentially: {row['ordinal']}")
        previous_finished = _parse_time(finished)
        _require(report_path.is_file(), f"child report is missing: {row['ordinal']}")
        report = _read(report_path)
        producer = _read(producer_path) if producer_path.is_file() else None
        _validate_child(report, producer, row["case"], mode, row["warmup"], row["samples"], expected_corpora[row["case"]], binary)
        _read_peak(peak_path)
        semantic = {"status": "pass", "report_sha256": _sha(report_path), "producer_sha256": _sha(producer_path) if producer_path.is_file() else None, "guard_contract": GUARD_CONTRACT if row["case"] == PRIMARY_CASE else None}
        _write_json(folder / "semantic.json", semantic)
    result = {"schema": EXPERIMENT_SCHEMA, "status": "pass", "plan": str(plan_path.resolve()), "plan_sha256": _sha(plan_path), "schedule": "native then allocator; fixed alternating AB/BA by (block + case_index) % 2", "rows": receipts, "children": TOTAL_CHILDREN, "native_children": TOTAL_NATIVE_CHILDREN, "allocator_children": TOTAL_ALLOCATOR_CHILDREN, "no_retries_or_replacements": True, "performance_claim": None}
    _write_json(output_root / "experiment.json", result)
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--output-root", type=Path, required=True)
    args = parser.parse_args()
    print(json.dumps(run(args.plan.resolve(), args.output_root.resolve()), indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
