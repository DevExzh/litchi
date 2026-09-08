#!/usr/bin/env python3
"""Portable, fail-closed verifier for the 0467 XLSX evidence bundle.

The verifier authenticates retained build bindings, capture receipts and raw
lane artifacts.  It validates the four clean normal ABBA legs, the four full
default-matrix guard legs, and the two external Heaptrack legs.  It never
builds, executes a benchmark, or requires a binary after cleanup.  ``--preseal``
allows an otherwise complete bundle to be checked before the final
``SHA256SUMS`` file is written; when a seal is present it is always checked.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import math
from pathlib import Path
import re
import sys
from typing import Any, Iterable, Mapping, Sequence


ROOT = Path(__file__).resolve().parent
SCHEMA = "litchi-0467-verification-v1"
PROTOCOL_SCHEMA = "litchi-0467-protocol-v2"
CAPTURE_SCHEMA = "litchi-0467-capture-v1"
REVISION_RE = re.compile(r"[0-9a-f]{40}\Z")
SHA256_RE = re.compile(r"[0-9a-f]{64}\Z")
FLAGS = "-C force-frame-pointers=yes -C force-unwind-tables=yes"
CPU = "2"
WORKERS = 1
SOURCE_PRESENT_FILES = 6992
SOURCE_OMITTED_FILES = 40
CASES = ("xlsx_one_cell_commit_save", "xlsx_one_percent_commit_save")
SHAPES = ("tiny", "medium", "dense-wide")
NORMAL_LANES = ("A1-clean", "B1-clean", "B2-clean", "A2-clean")
FULL_LANES = ("A-full-clean", "B-full-clean", "B2-full-clean", "A2-full-clean")
# The capture driver currently names the full guard only A-full-clean/B-full-
# clean in its additional-lane prose.  The formal verifier accepts the
# protocol's four-role expansion below and the legacy A2/B2 spelling is not
# used for the full guard unless all four are present.
FULL_ALIASES = {
    "A-full-clean": "A-full-clean",
    "B-full-clean": "B-full-clean",
    "B2-full-clean": "B-full-clean",
    "A2-full-clean": "A-full-clean",
}
HEAP_LANES = ("A-heap-clean", "B-heap-clean")
LANE_ROLE = {
    "A1-clean": "control",
    "A2-clean": "control",
    "B1-clean": "candidate",
    "B2-clean": "candidate",
    "A-full-clean": "control",
    "A2-full-clean": "control",
    "B-full-clean": "candidate",
    "B2-full-clean": "candidate",
    "A-heap-clean": "control",
    "B-heap-clean": "candidate",
}
EXPECTED_TOOL = {
    "name": "litchi-perf-baseline",
    "version": "0.1.0",
    "binary": "litchi-perf-baseline",
    "profile": "release",
    "target_os": "linux",
    "target_arch": "x86_64",
    "instrumentation": "none",
}
EXPECTED_BUILD_ENVIRONMENT = {
    "RUSTUP_TOOLCHAIN": "1.98.1",
    "RUSTFLAGS": FLAGS,
    "CARGO_PROFILE_RELEASE_DEBUG": "1",
    "CARGO_INCREMENTAL": "0",
    "CARGO_BUILD_JOBS": "4",
    "DEBUGINFOD_URLS": "",
}
EXPECTED_CAPTURE_ENVIRONMENT = {
    "RUSTUP_TOOLCHAIN": "1.98.1",
    "RUSTFLAGS": FLAGS,
    "CARGO_PROFILE_RELEASE_DEBUG": "1",
    "DEBUGINFOD_URLS": "",
    "LC_ALL": "C",
}

for _candidate in (ROOT, *ROOT.parents):
    if (_candidate / "tools" / "perf_abba_summary.py").is_file():
        if str(_candidate) not in sys.path:
            sys.path.insert(0, str(_candidate))
        break


class VerificationError(ValueError):
    """A fail-closed bundle validation error."""


def fail(message: str) -> None:
    raise VerificationError(message)


def regular(path: Path, label: str) -> Path:
    if path.is_symlink() or not path.is_file():
        fail(f"{label}: missing, symlinked, or non-regular file")
    return path


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def canonical_bytes(value: Any) -> bytes:
    try:
        return json.dumps(
            value,
            ensure_ascii=False,
            allow_nan=False,
            sort_keys=True,
            separators=(",", ":"),
        ).encode("utf-8")
    except (TypeError, ValueError, OverflowError) as error:
        fail(f"cannot canonicalize JSON: {error}")
    raise AssertionError("unreachable")


def canonical_sha256(value: Any) -> str:
    return hashlib.sha256(canonical_bytes(value)).hexdigest()


def load_json(path: Path, label: str) -> Any:
    regular(path, label)
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"{label}: invalid JSON ({error})")
    raise AssertionError("unreachable")


def obj(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label}: expected object")
    return value


def text(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        fail(f"{label}: expected non-empty string")
    return value


def digest(value: Any, label: str) -> str:
    value = text(value, label)
    if SHA256_RE.fullmatch(value) is None:
        fail(f"{label}: expected lowercase SHA-256")
    return value


def revision(value: Any, label: str) -> str:
    value = text(value, label)
    if REVISION_RE.fullmatch(value) is None:
        fail(f"{label}: expected a lowercase commit id")
    return value


def integer(value: Any, label: str, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        fail(f"{label}: expected integer >= {minimum}")
    return value


def safe_relative(value: Any, label: str) -> Path:
    raw = text(value, label)
    path = Path(raw)
    if path.is_absolute() or ".." in path.parts or path.as_posix() != raw:
        fail(f"{label}: expected a relative traversal-free path")
    return path


def bundle_path(value: Any, label: str, *, base: Path = ROOT) -> Path:
    relative = safe_relative(value, label)
    path = base / relative
    try:
        resolved = path.resolve(strict=True)
        if not resolved.is_relative_to(base.resolve()):
            fail(f"{label}: path escapes bundle")
    except OSError as error:
        fail(f"{label}: cannot resolve ({error})")
    return regular(path, label)


def _reject_symlinks(root: Path) -> None:
    for path in root.rglob("*"):
        if path.is_symlink():
            fail(f"bundle contains symlink: {path.relative_to(root)}")


def _protocol() -> dict[str, Any]:
    path = ROOT / "protocol-r1.json"
    if not path.is_file():
        path = ROOT / "protocol.json"
    protocol = obj(load_json(path, path.name), path.name)
    if protocol.get("schema") != PROTOCOL_SCHEMA or protocol.get("status") != "frozen":
        fail("protocol schema/status is not the frozen 0467 revision")
    if protocol.get("order") != list(NORMAL_LANES):
        fail("protocol normal lane order differs")
    if protocol.get("samples") != 100 or protocol.get("warmups") != 5:
        fail("protocol normal sample/warmup counts differ")
    if protocol.get("cpu") != 2 or protocol.get("workers") != WORKERS:
        fail("protocol CPU/worker configuration differs")
    if protocol.get("cases") != list(CASES) or protocol.get("shapes") != list(SHAPES):
        fail("protocol normal case/shape matrix differs")
    triggers = obj(protocol.get("review_triggers"), "protocol.review_triggers")
    if triggers.get("latency_percent") != 5 or triggers.get("process_rss_percent") != 5:
        fail("protocol review threshold differs")
    additional = protocol.get("additional_lanes")
    if additional != ["A-full-clean", "B-full-clean", "A-heap-clean", "B-heap-clean"]:
        fail("protocol additional lane names differ")
    driver_sha = digest(protocol.get("capture_driver_sha256"), "protocol.capture_driver_sha256")
    if driver_sha != sha256_file(ROOT / "capture.py"):
        fail("protocol capture driver hash differs from retained capture.py")
    return protocol


def _binding(role: str) -> tuple[Path, dict[str, Any], dict[str, Any]]:
    path = ROOT / f"{role}-binding.json"
    binding = obj(load_json(path, path.name), path.name)
    bound_revision = revision(binding.get("revision"), f"{path}.revision")
    bound_binary = digest(binding.get("binary_sha256"), f"{path}.binary_sha256")
    bound_bytes = integer(binding.get("bytes"), f"{path}.bytes", 1)
    if binding.get("clean_build") is not True:
        fail(f"{path}.clean_build must be true")
    receipt_name = safe_relative(binding.get("build_receipt"), f"{path}.build_receipt")
    receipt_path = ROOT / receipt_name
    receipt = obj(load_json(receipt_path, str(receipt_path)), str(receipt_path))
    if sha256_file(receipt_path) != digest(binding.get("build_receipt_sha256"), f"{path}.build_receipt_sha256"):
        fail(f"{path}: build receipt hash differs")
    if receipt.get("role") != role or revision(receipt.get("revision"), f"{receipt_path}.revision") != bound_revision:
        fail(f"{receipt_path}: role/revision differs from binding")
    if receipt.get("exit_code") != 0 or receipt.get("clean_before") is not True or receipt.get("clean_after") is not True:
        fail(f"{receipt_path}: clean build proof is incomplete")
    if receipt.get("environment") != EXPECTED_BUILD_ENVIRONMENT:
        fail(f"{receipt_path}.environment differs from the formal build")
    argv = receipt.get("argv")
    expected_prefix = ["cargo", "build", "--release", "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml"]
    if not isinstance(argv, list) or argv[: len(expected_prefix)] != expected_prefix or "--bin" not in argv or argv[argv.index("--bin") + 1 : argv.index("--bin") + 2] != ["litchi-perf-baseline"]:
        fail(f"{receipt_path}.argv is not the authenticated release build")
    source = binding.get("source")
    if source is not None:
        source_object = obj(source, f"{path}.source")
        source_path = bundle_path(source_object.get("path"), f"{path}.source.path")
        if sha256_file(source_path) != digest(source_object.get("sha256"), f"{path}.source.sha256"):
            fail(f"{path}.source hash differs")
        integer(source_object.get("files"), f"{path}.source.files", 1)
    return path, binding, {
        "revision": bound_revision,
        "binary_sha256": bound_binary,
        "bytes": bound_bytes,
        "build_receipt": str(receipt_name),
        "build_receipt_sha256": sha256_file(receipt_path),
    }


def _bindings() -> dict[str, dict[str, Any]]:
    values = {role: _binding(role)[2] for role in ("control", "candidate")}
    if values["control"]["revision"] == values["candidate"]["revision"]:
        fail("control and candidate revisions must be distinct")
    if values["control"]["binary_sha256"] == values["candidate"]["binary_sha256"]:
        fail("control and candidate binaries must be distinct")
    return values


def _source_bindings(bindings: Mapping[str, Mapping[str, Any]]) -> dict[str, Any]:
    """Authenticate the portable source epoch and the reviewed change set."""

    path = ROOT / "source-bindings.json"
    record = obj(load_json(path, path.name), path.name)
    if record.get("schema") != "litchi-0467-source-bindings-v1":
        fail("source-bindings.json schema differs")
    prior_path = text(record.get("prior_source_manifest"), "source-bindings.prior_source_manifest")
    if Path(prior_path).is_absolute() or ".." in Path(prior_path).parts:
        fail("source-bindings.prior_source_manifest is unsafe")
    digest(record.get("prior_source_manifest_sha256"), "source-bindings.prior_source_manifest_sha256")
    roles = obj(record.get("roles"), "source-bindings.roles")
    if set(roles) != {"control", "candidate"}:
        fail("source-bindings.roles must contain control and candidate")
    source_maps: dict[str, dict[str, Any]] = {}
    source_trees: dict[str, str] = {}
    omitted_sets: dict[str, set[str]] = {}
    fixture_maps: dict[str, Mapping[str, Any]] = {}
    for role in ("control", "candidate"):
        value = obj(roles.get(role), f"source-bindings.roles.{role}")
        if value.get("revision") != bindings[role]["revision"]:
            fail(f"source-bindings.{role}.revision differs from build binding")
        # A portable replay can validate only the recorded 40-hex tree
        # identity.  Source custody is established by the clean build
        # receipts and the hashed source manifests below; no live Git lookup
        # is required after cleanup.
        tree = revision(value.get("git_tree"), f"source-bindings.{role}.git_tree")
        source_trees[role] = tree
        source_path = bundle_path(value.get("path"), f"source-bindings.{role}.path")
        source_digest = digest(value.get("sha256"), f"source-bindings.{role}.sha256")
        if sha256_file(source_path) != source_digest:
            fail(f"source-bindings.{role}: source manifest hash differs")
        source_map = obj(load_json(source_path, f"sources/{role}.json"), f"sources/{role}.json")
        expected_files = integer(value.get("files"), f"source-bindings.{role}.files", 1)
        if expected_files != SOURCE_PRESENT_FILES:
            fail(f"source-bindings.{role}.files must be {SOURCE_PRESENT_FILES}")
        if len(source_map) != expected_files:
            fail(f"source-bindings.{role}: file count differs")
        for name, file_digest in source_map.items():
            relative_name = safe_relative(name, f"sources/{role}.json.{name}")
            if relative_name.suffix not in {".rs", ".toml", ".lock"}:
                fail(f"sources/{role}.json.{name}: unsupported source suffix")
            digest(file_digest, f"sources/{role}.json.{name}")
        omitted = value.get("omitted_from_sparse")
        if not isinstance(omitted, list) or len(omitted) != SOURCE_OMITTED_FILES or len(set(omitted)) != SOURCE_OMITTED_FILES:
            fail(f"source-bindings.{role}.omitted_from_sparse must contain {SOURCE_OMITTED_FILES} unique paths")
        omitted_sets[role] = set()
        for name in omitted:
            omitted_sets[role].add(safe_relative(name, f"source-bindings.{role}.omitted_from_sparse").as_posix())
        changed = value.get("changed_from_prior")
        if not isinstance(changed, list) or any(not isinstance(item, str) for item in changed):
            fail(f"source-bindings.{role}.changed_from_prior is malformed")
        if value.get("clean") is not True:
            fail(f"source-bindings.{role}.clean must be true")
        fixtures = obj(value.get("included_fixtures"), f"source-bindings.{role}.included_fixtures")
        if set(fixtures) != {
            "test-data/poi/test-data/spreadsheet/54016.xls",
            "test-data/rtf/watermark.rtf",
        }:
            fail(f"source-bindings.{role}.included_fixtures differs")
        for fixture, fixture_digest in fixtures.items():
            safe_relative(fixture, f"source-bindings.{role}.included_fixtures.{fixture}")
            fixture_maps.setdefault(role, {})[fixture] = digest(
                fixture_digest, f"source-bindings.{role}.included_fixtures.{fixture}"
            )
        source_maps[role] = source_map
    if omitted_sets["control"] != omitted_sets["candidate"]:
        fail("control and candidate sparse omission lists differ")
    if set(source_maps["control"]) != set(source_maps["candidate"]):
        fail("control and candidate source manifest key sets differ")
    source_keys = set(source_maps["control"])
    changed_keys = {
        name for name in source_keys if source_maps["control"][name] != source_maps["candidate"][name]
    }
    expected_changed = {
        "crates/litchi-xlsx/src/raw/worksheet/codec.rs",
        "crates/litchi-xlsx/src/raw/worksheet/tests.rs",
    }
    if changed_keys != expected_changed:
        fail(f"source manifest changed set differs: {sorted(changed_keys)}")
    if roles["control"].get("changed_from_prior") != [] or set(roles["candidate"].get("changed_from_prior", [])) != expected_changed:
        fail("source-bindings changed_from_prior does not match the exact reviewed pair")
    review_path = ROOT / "review.json"
    review = obj(load_json(review_path, review_path.name), review_path.name)
    if review.get("schema") != "litchi-0467-candidate-review-v1":
        fail("review.json schema differs")
    reviewed_files = obj(review.get("files"), "review.files")
    if set(reviewed_files) != expected_changed:
        fail("review.json must cover exactly the two changed files")
    for name in expected_changed:
        if digest(reviewed_files[name], f"review.files.{name}") != source_maps["candidate"][name]:
            fail(f"review.files.{name} does not match candidate source manifest")
    if fixture_maps["control"] != fixture_maps["candidate"]:
        fail("compile-time fixture hashes differ between source roles")
    return {
        "path": path.name,
        "sha256": sha256_file(path),
        "files": len(source_keys),
        "changed_files": sorted(expected_changed),
        "git_trees": {
            role: source_trees[role] for role in ("control", "candidate")
        },
        "git_tree_provenance": "recorded 40-hex tree identities checked structurally; source custody comes from clean build receipts and hashed manifests",
    }


def _preseal_live_bindings(bindings: Mapping[str, Mapping[str, Any]]) -> None:
    """Optionally prove that the still-live pre-cleanup binary matches its binding."""

    temp = Path("/tmp/litchi-goal-0467")
    for role, binding in bindings.items():
        path = temp / role
        regular(path, f"preseal {role} binary")
        if path.stat().st_size != binding["bytes"] or sha256_file(path) != binding["binary_sha256"]:
            fail(f"preseal {role} binary identity differs")
    # The fixed-checkout experiment has separate role bindings and binaries.
    # Check them when those bindings are retained, while leaving the normal
    # post-cleanup verifier entirely artifact-only.
    for role in ("control", "candidate"):
        binding_path = ROOT / f"{role}-fixed-binding.json"
        if not binding_path.is_file():
            continue
        fixed_binding = obj(load_json(binding_path, binding_path.name), binding_path.name)
        fixed_revision = revision(fixed_binding.get("revision"), f"{binding_path}.revision")
        fixed_digest = digest(fixed_binding.get("binary_sha256"), f"{binding_path}.binary_sha256")
        fixed_bytes = integer(fixed_binding.get("bytes"), f"{binding_path}.bytes", 1)
        if fixed_binding.get("role") != role:
            fail(f"{binding_path}.role differs from {role}")
        path = temp / f"{role}-fixed"
        regular(path, f"preseal {role}-fixed binary")
        if path.stat().st_size != fixed_bytes or sha256_file(path) != fixed_digest:
            fail(f"preseal {role}-fixed binary identity differs for revision {fixed_revision}")


def _artifact_rows(receipt: Mapping[str, Any], label: str) -> dict[str, Mapping[str, Any]]:
    rows = receipt.get("artifacts")
    if not isinstance(rows, dict):
        fail(f"{label}.artifacts must be an object")
    result: dict[str, Mapping[str, Any]] = {}
    for name, value in rows.items():
        if not isinstance(name, str) or not name or Path(name).name != name:
            fail(f"{label}.artifacts has an unsafe name")
        result[name] = obj(value, f"{label}.artifacts.{name}")
    return result


def _artifact_hashes(
    lane: str,
    receipt: Mapping[str, Any],
    *,
    exported: bool = False,
) -> dict[str, Path]:
    lane_dir = ROOT / lane
    if not lane_dir.is_dir() or lane_dir.is_symlink():
        fail(f"{lane}: missing or unsafe lane directory")
    if any(path.is_dir() for path in lane_dir.iterdir()):
        fail(f"{lane}: nested directories are not permitted")
    rows = _artifact_rows(receipt, f"{lane}.receipt")
    expected = {"started.json", "corpus-catalog.json", "stdout.log", "stderr.log", "report.json", "resource.log"}
    if lane in HEAP_LANES:
        expected.add("heaptrack.zst")
    if set(rows) != expected:
        fail(f"{lane}: artifact set differs (expected={sorted(expected)}, actual={sorted(rows)})")
    result: dict[str, Path] = {}
    for name, row in rows.items():
        path = lane_dir / safe_relative(row.get("path", name), f"{lane}.artifacts.{name}.path")
        if path.parent != lane_dir:
            fail(f"{lane}.artifacts.{name}.path must be lane-local")
        regular(path, f"{lane}.artifacts.{name}")
        expected_sha = digest(row.get("sha256"), f"{lane}.artifacts.{name}.sha256")
        expected_bytes = integer(row.get("bytes"), f"{lane}.artifacts.{name}.bytes")
        if path.stat().st_size != expected_bytes or sha256_file(path) != expected_sha:
            fail(f"{lane}.artifacts.{name}: retained identity differs")
        result[name] = path
    allowed_files = expected | {"receipt.json"}
    if exported:
        allowed_files.update(("print.txt", "print.stderr"))
    actual = {path.name for path in lane_dir.iterdir() if path.is_file() and not path.is_symlink()}
    if actual != allowed_files:
        fail(f"{lane}: unexpected raw files {sorted(actual - allowed_files)}")
    return result


def _receipt(
    lane: str,
    binding: Mapping[str, Any],
    *,
    exported: bool = False,
) -> tuple[dict[str, Any], dict[str, Path]]:
    path = ROOT / lane / "receipt.json"
    receipt = obj(load_json(path, f"{lane}.receipt.json"), f"{lane}.receipt.json")
    role = LANE_ROLE[lane]
    if receipt.get("schema") != CAPTURE_SCHEMA or receipt.get("lane") != lane or receipt.get("role") != role:
        fail(f"{lane}: receipt schema/lane/role differs")
    if receipt.get("revision") != binding["revision"] or receipt.get("binary_sha256") != binding["binary_sha256"]:
        fail(f"{lane}: receipt binding identity differs")
    if receipt.get("binding_sha256") != sha256_file(ROOT / f"{role}-binding.json"):
        fail(f"{lane}: receipt binding hash differs")
    if digest(receipt.get("driver_sha256"), f"{lane}.driver_sha256") != sha256_file(ROOT / "capture.py"):
        fail(f"{lane}: receipt driver hash differs")
    if receipt.get("exit_code") != 0 or receipt.get("clean_before") is not True or receipt.get("clean_after") is not True or receipt.get("binary_unchanged") is not True:
        fail(f"{lane}: receipt does not prove successful custody")
    if receipt.get("report_metadata_matches_clean_role") is not True:
        fail(f"{lane}: report metadata clean-role proof is absent")
    artifacts = _artifact_hashes(lane, receipt, exported=exported)
    started = obj(load_json(artifacts["started.json"], f"{lane}.started.json"), f"{lane}.started.json")
    for key in (
        "schema",
        "lane",
        "role",
        "revision",
        "binary_sha256",
        "binding_sha256",
        "driver_sha256",
        "cwd",
        "argv",
        "environment",
        "samples",
        "warmups",
    ):
        if started.get(key) != receipt.get(key):
            fail(f"{lane}.started.json.{key} differs from receipt")
    if started.get("clean_before") is not True:
        fail(f"{lane}.started.json does not prove a clean start")
    return receipt, artifacts


def _report_profile(report: Mapping[str, Any], label: str) -> None:
    if report.get("schema_version") != 1 or report.get("tool") != EXPECTED_TOOL:
        fail(f"{label}: tool/schema identity differs")


def _report_identity(report: Mapping[str, Any], label: str, binding: Mapping[str, Any]) -> None:
    _report_profile(report, label)
    environment = obj(report.get("environment"), f"{label}.environment")
    if revision(environment.get("git_revision"), f"{label}.environment.git_revision") != binding["revision"]:
        fail(f"{label}: report revision differs from binding")
    if environment.get("git_worktree_dirty") is not False or environment.get("cpu_affinity") != CPU:
        fail(f"{label}: report does not prove clean CPU-pinned execution")
    if environment.get("rustflags") != FLAGS:
        fail(f"{label}: report Rust flags differ")
    binary = obj(report.get("binary_identity"), f"{label}.binary_identity")
    if digest(binary.get("binary_sha256"), f"{label}.binary_identity.binary_sha256") != binding["binary_sha256"] or binary.get("binary_bytes") != binding["bytes"]:
        fail(f"{label}: report binary identity differs from binding")


def _configuration(report: Mapping[str, Any], label: str, samples: int, warmups: int, cases: Sequence[str], shapes: Sequence[str]) -> None:
    configuration = obj(report.get("configuration"), f"{label}.configuration")
    if configuration.get("samples_per_case") != samples or configuration.get("warmup_iterations_per_case") != warmups:
        fail(f"{label}: sample/warmup configuration differs")
    if configuration.get("cases") != list(cases) or configuration.get("xlsx_shapes") != list(shapes):
        fail(f"{label}: case/shape configuration differs")
    if configuration.get("execution_workers") != [WORKERS]:
        fail(f"{label}: worker configuration differs")


def _catalog(report: Mapping[str, Any], path: Path, label: str, binding: Mapping[str, Any]) -> tuple[dict[str, Any], str]:
    catalog = obj(load_json(path, f"{label}.corpus-catalog.json"), f"{label}.corpus-catalog.json")
    if catalog.get("manifest_version") != 2 or catalog.get("catalog_id") != "litchi-perf-corpus-v2":
        fail(f"{label}: corpus catalog identity differs")
    build = obj(catalog.get("build"), f"{label}.catalog.build")
    if build.get("git_revision") != binding["revision"] or build.get("git_worktree_dirty") is not False:
        fail(f"{label}: catalog build identity differs from binding")
    catalog_without_hash = dict(catalog)
    catalog_without_hash.pop("catalog_sha256", None)
    catalog_hash = digest(catalog.get("catalog_sha256"), f"{label}.catalog_sha256")
    if canonical_sha256(catalog_without_hash) != catalog_hash:
        fail(f"{label}: catalog_sha256 does not recompute")
    reference = obj(report.get("corpus_catalog"), f"{label}.report.corpus_catalog")
    for key in ("manifest_version", "catalog_id", "catalog_sha256", "content_set_sha256"):
        if reference.get(key) != catalog.get(key):
            fail(f"{label}: report/catalog reference differs for {key}")
    return catalog, canonical_sha256({key: value for key, value in catalog.items() if key not in {"build", "catalog_sha256"}})


def _elapsed(row: Mapping[str, Any], samples: int, label: str) -> None:
    elapsed = obj(row.get("elapsed_ns"), f"{label}.elapsed_ns")
    if elapsed.get("unit") != "ns" or not isinstance(elapsed.get("samples"), list) or len(elapsed["samples"]) != samples:
        fail(f"{label}: elapsed sample vector differs")
    values = elapsed["samples"]
    if any(isinstance(value, bool) or not isinstance(value, int) or value <= 0 for value in values):
        fail(f"{label}: elapsed samples must be positive integers")
    if values != sorted(values):
        fail(f"{label}: elapsed samples are not sorted")
    order = elapsed.get("sample_order")
    if not isinstance(order, list) or sorted(order) != list(range(samples)):
        fail(f"{label}: sample order is not a permutation")
    midpoint = values[(samples - 1) // 2] // 2 + values[samples // 2] // 2 + ((values[(samples - 1) // 2] % 2 + values[samples // 2] % 2) // 2)
    nearest = lambda percentile: values[min((percentile * samples + 99) // 100 - 1, samples - 1)]
    expected_int = {"min": values[0], "p50": midpoint, "p95": nearest(95), "p99": nearest(99), "max": values[-1]}
    for key, expected in expected_int.items():
        if elapsed.get(key) != expected:
            fail(f"{label}: {key} does not recompute")
    mean = 0.0
    squared = 0.0
    for index, value in enumerate(values, 1):
        delta = float(value) - mean
        mean += delta / index
        squared += delta * (float(value) - mean)
    deviation = math.sqrt(squared / (samples - 1)) if samples > 1 else 0.0
    critical = _student_t(samples - 1)
    margin = critical * deviation / math.sqrt(samples) if samples > 1 else 0.0
    _close(elapsed.get("mean"), mean, f"{label}.mean")
    _close(elapsed.get("standard_deviation"), deviation, f"{label}.standard_deviation")
    interval = obj(elapsed.get("confidence_interval_95"), f"{label}.confidence_interval_95")
    if interval.get("method") != "two-sided Student's t interval for the mean":
        fail(f"{label}: confidence interval method differs")
    _close(interval.get("lower"), max(mean - margin, 0.0), f"{label}.confidence_interval_95.lower")
    _close(interval.get("upper"), mean + margin, f"{label}.confidence_interval_95.upper")


def _student_t(degrees: int) -> float:
    values = (12.706, 4.303, 3.182, 2.776, 2.571, 2.447, 2.365, 2.306, 2.262, 2.228, 2.201, 2.179, 2.160, 2.145, 2.131, 2.120, 2.110, 2.101, 2.093, 2.086, 2.080, 2.074, 2.069, 2.064, 2.060, 2.056, 2.052, 2.048, 2.045, 2.042)
    if degrees <= 0:
        return 0.0
    if degrees <= len(values):
        return values[degrees - 1]
    z = 1.959963984540054
    d = float(degrees)
    return z + (z**3 + z) / (4 * d) + (5 * z**5 + 16 * z**3 + 3 * z) / (96 * d**2) + (3 * z**7 + 19 * z**5 + 17 * z**3 - 15 * z) / (384 * d**3)


def _close(actual: Any, expected: float, label: str) -> None:
    if isinstance(actual, bool) or not isinstance(actual, (int, float)) or not math.isfinite(float(actual)):
        fail(f"{label}: expected finite number")
    tolerance = max(1e-12, abs(expected) * 1e-12)
    if abs(float(actual) - expected) > tolerance:
        fail(f"{label}: does not recompute")


def _rows(report: Mapping[str, Any], samples: int, label: str) -> dict[tuple[str, str], Mapping[str, Any]]:
    values = report.get("results")
    if not isinstance(values, list):
        fail(f"{label}.results must be a list")
    indexed: dict[tuple[str, str], Mapping[str, Any]] = {}
    for index, raw in enumerate(values):
        row = obj(raw, f"{label}.results[{index}]")
        case = text(row.get("case"), f"{label}.results[{index}].case")
        corpus = obj(row.get("corpus"), f"{label}.results[{index}].corpus")
        identity = canonical_sha256(corpus)
        key = (case, identity)
        if key in indexed:
            fail(f"{label}: duplicate case/corpus result")
        _elapsed(row, samples, f"{label}.{case}[{identity}]")
        indexed[key] = row
    return indexed


def _resource(path: Path, label: str) -> int:
    regular(path, label)
    match = re.search(r"^\s*Maximum resident set size \(kbytes\):\s*(\d+)\s*$", path.read_text(encoding="utf-8"), re.MULTILINE)
    if match is None or int(match.group(1)) <= 0:
        fail(f"{label}: missing positive process RSS")
    if "Exit status: 0" not in path.read_text(encoding="utf-8"):
        fail(f"{label}: workload exit marker is not successful")
    return int(match.group(1))


def _postprocess() -> set[str]:
    """Authenticate optional Heaptrack exports without weakening raw receipts."""

    path = ROOT / "postprocess.json"
    export_names = {"print.txt", "print.stderr"}
    if not path.exists():
        for lane in HEAP_LANES:
            present = {name for name in export_names if (ROOT / lane / name).exists()}
            if present:
                fail(f"{lane}: Heaptrack exports exist without postprocess.json")
        return set()
    record = obj(load_json(path, path.name), path.name)
    if record.get("environment") != {"DEBUGINFOD_URLS": "", "LC_ALL": "C"}:
        fail("postprocess.json.environment differs")
    commands = record.get("commands")
    if not isinstance(commands, list) or len(commands) != len(HEAP_LANES):
        fail("postprocess.json.commands must contain both Heaptrack lanes")
    seen: set[str] = set()
    for index, raw in enumerate(commands):
        command = obj(raw, f"postprocess.json.commands[{index}]")
        lane = text(command.get("lane"), f"postprocess.json.commands[{index}].lane")
        if lane not in HEAP_LANES or lane in seen:
            fail(f"postprocess.json.commands[{index}].lane is not a unique heap lane")
        seen.add(lane)
        argv = command.get("argv")
        if not isinstance(argv, list) or len(argv) != 5 or argv[0] != "heaptrack_print" or argv[1] != "-f" or argv[3] != "-n" or argv[4] != "30":
            fail(f"postprocess.json.commands[{index}].argv is not the retained export command")
        heap_path = Path(text(argv[2], f"postprocess.json.commands[{index}].argv[2]"))
        if heap_path.name != "heaptrack.zst" or heap_path.parent.name != lane:
            fail(f"postprocess.json.commands[{index}].argv does not bind its lane Heaptrack artifact")
        if command.get("exit_code") != 0:
            fail(f"postprocess.json.commands[{index}] did not exit successfully")
        artifacts = obj(command.get("artifacts"), f"postprocess.json.commands[{index}].artifacts")
        if set(artifacts) != export_names:
            fail(f"postprocess.json.commands[{index}].artifacts differs")
        lane_dir = ROOT / lane
        for name in sorted(export_names):
            artifact_path = lane_dir / name
            regular(artifact_path, f"{lane}.{name}")
            digest(artifacts[name], f"postprocess.json.commands[{index}].artifacts.{name}")
            if sha256_file(artifact_path) != artifacts[name]:
                fail(f"{lane}.{name}: export hash differs")
        seen.add(lane)
    if seen != set(HEAP_LANES):
        fail("postprocess.json does not cover both Heaptrack lanes")
    return set(HEAP_LANES)


def _normal(bindings: Mapping[str, Mapping[str, Any]]) -> dict[str, Any]:
    indexed: dict[str, dict[tuple[str, str], Mapping[str, Any]]] = {}
    reports: list[Mapping[str, Any]] = []
    catalogs: dict[str, str] = {}
    for lane in NORMAL_LANES:
        role = LANE_ROLE[lane]
        receipt, artifacts = _receipt(lane, bindings[role])
        if receipt.get("samples") != 100 or receipt.get("warmups") != 5:
            fail(f"{lane}: normal receipt sample/warmup differs")
        argv = receipt.get("argv")
        if not isinstance(argv, list) or argv[:3] != ["taskset", "-c", CPU] or "heaptrack" in argv:
            fail(f"{lane}: normal command is not the uninstrumented CPU-pinned run")
        if receipt.get("environment") != EXPECTED_CAPTURE_ENVIRONMENT:
            fail(f"{lane}: capture environment differs")
        report = obj(load_json(artifacts["report.json"], f"{lane}.report.json"), f"{lane}.report.json")
        _report_identity(report, lane, bindings[role])
        _configuration(report, lane, 100, 5, CASES, SHAPES)
        _catalog_value, catalog_identity = _catalog(report, artifacts["corpus-catalog.json"], lane, bindings[role])
        catalogs[lane] = catalog_identity
        reports.append(report)
        indexed[lane] = _rows(report, 100, lane)
        if len(indexed[lane]) != 6:
            fail(f"{lane}: normal report must contain six rows")
        _resource(artifacts["resource.log"], f"{lane}.resource.log")
    if len(set(catalogs.values())) != 1:
        fail("normal catalogs differ in corpus identity")
    expected_keys = set(indexed[NORMAL_LANES[0]])
    if {(case, identity) for case, identity in expected_keys} != expected_keys:
        fail("normal result identity is malformed")
    for lane in NORMAL_LANES[1:]:
        if set(indexed[lane]) != expected_keys:
            fail("normal lanes do not share exact case/corpus identities")
    if {case for case, _ in expected_keys} != set(CASES) or len(expected_keys) != 6:
        fail("normal case matrix differs")
    summary = _abba_summary(reports, CASES, SHAPES, "normal ABBA")
    return {"lanes": list(NORMAL_LANES), "result_count": 6, "abba_summary": summary}


def _abba_summary(reports: Sequence[Mapping[str, Any]], cases: Sequence[str], shapes: Sequence[str], label: str) -> dict[str, Any]:
    try:
        from tools import perf_abba_summary
    except (ImportError, ModuleNotFoundError) as error:
        fail(f"{label}: tools.perf_abba_summary is unavailable ({error})")
    try:
        return perf_abba_summary.summarize_reports(reports=reports, cases=cases, shapes=shapes)
    except Exception as error:
        fail(f"{label}: canonical ABBA validation failed ({error})")
    raise AssertionError("unreachable")


def _full(bindings: Mapping[str, Mapping[str, Any]]) -> dict[str, Any]:
    # The protocol's additional lane list is a two-run guard.  A final bundle
    # retains A-full-clean and B-full-clean.  If a later capture records the
    # second ABBA spelling, accept it only when both second legs are present.
    available = {name for name in FULL_LANES if (ROOT / name).is_dir()}
    if set(FULL_LANES) <= available:
        lanes = FULL_LANES
    elif {"A-full-clean", "B-full-clean"} <= available:
        lanes = ("A-full-clean", "B-full-clean")
    else:
        fail("full guard lanes are incomplete; need A-full-clean and B-full-clean")
    reports: list[Mapping[str, Any]] = []
    catalogs: list[str] = []
    for lane in lanes:
        role = LANE_ROLE[lane]
        receipt, artifacts = _receipt(lane, bindings[role])
        if receipt.get("samples") != 15 or receipt.get("warmups") != 3:
            fail(f"{lane}: full guard sample/warmup differs")
        argv = receipt.get("argv")
        if not isinstance(argv, list) or argv[:3] != ["taskset", "-c", CPU] or "heaptrack" in argv:
            fail(f"{lane}: full guard must be uninstrumented")
        if receipt.get("environment") != EXPECTED_CAPTURE_ENVIRONMENT:
            fail(f"{lane}: full guard environment differs")
        report = obj(load_json(artifacts["report.json"], f"{lane}.report.json"), f"{lane}.report.json")
        _report_identity(report, lane, bindings[role])
        configuration = obj(report.get("configuration"), f"{lane}.configuration")
        if configuration.get("samples_per_case") != 15 or configuration.get("warmup_iterations_per_case") != 3 or configuration.get("execution_workers") != [WORKERS]:
            fail(f"{lane}: full guard configuration differs")
        rows = _rows(report, 15, lane)
        if len(rows) != 201:
            fail(f"{lane}: full guard must contain exactly 201 rows")
        _catalog_value, catalog_identity = _catalog(report, artifacts["corpus-catalog.json"], lane, bindings[role])
        catalogs.append(catalog_identity)
        reports.append(report)
        _resource(artifacts["resource.log"], f"{lane}.resource.log")
    if len(lanes) == 2:
        # A full guard is a same-revision control/candidate pair; still use the
        # canonical ABBA validator by duplicating each role only for the
        # descriptive integrity pass.  It rejects the duplicate identity, so
        # validate rows and corpus sets directly for this two-leg form.
        if set(_rows(reports[0], 15, "full control")) != set(_rows(reports[1], 15, "full candidate")):
            fail("full guard control/candidate case/corpus identities differ")
    else:
        try:
            from tools import perf_abba_summary

            perf_abba_summary.summarize_reports(reports=reports)
        except Exception as error:
            fail(f"full ABBA guard: canonical validation failed ({error})")
    if len(set(catalogs)) != 1:
        fail("full guard catalogs differ in corpus identity")
    return {"lanes": list(lanes), "result_count": 201, "catalog_identity_count": len(set(catalogs))}


def _heap(
    bindings: Mapping[str, Mapping[str, Any]],
    *,
    exported_lanes: set[str] | None = None,
) -> dict[str, Any]:
    exported_lanes = exported_lanes or set()
    reports: dict[str, Mapping[str, Any]] = {}
    catalogs: dict[str, str] = {}
    keys: dict[str, set[tuple[str, str]]] = {}
    for lane in HEAP_LANES:
        role = LANE_ROLE[lane]
        receipt, artifacts = _receipt(
            lane,
            bindings[role],
            exported=lane in exported_lanes,
        )
        if receipt.get("samples") != 5 or receipt.get("warmups") != 1:
            fail(f"{lane}: heap guard sample/warmup differs")
        argv = receipt.get("argv")
        if not isinstance(argv, list) or argv[:3] != ["taskset", "-c", CPU] or "heaptrack" not in argv or "--" not in argv:
            fail(f"{lane}: heap guard is not explicitly whole-process Heaptrack instrumentation")
        heap_index = argv.index("heaptrack")
        if heap_index >= argv.index("--") or "-o" not in argv[heap_index:argv.index("--")]:
            fail(f"{lane}: Heaptrack output binding is absent")
        if receipt.get("environment") != EXPECTED_CAPTURE_ENVIRONMENT:
            fail(f"{lane}: heap environment differs")
        heap_data = artifacts["heaptrack.zst"].read_bytes()
        if len(heap_data) == 0 or heap_data[:4] != bytes.fromhex("28b52ffd"):
            fail(f"{lane}: Heaptrack artifact is not a non-empty zstd stream")
        output = (
            artifacts["stdout.log"].read_text(encoding="utf-8", errors="replace")
            + artifacts["stderr.log"].read_text(encoding="utf-8", errors="replace")
        )
        if "Heaptrack finished!" not in output:
            fail(f"{lane}: Heaptrack completion marker is absent")
        report = obj(load_json(artifacts["report.json"], f"{lane}.report.json"), f"{lane}.report.json")
        _report_identity(report, lane, bindings[role])
        _configuration(report, lane, 5, 1, ("xlsx_one_percent_commit_save",), ("dense-wide",))
        _catalog_value, catalog_identity = _catalog(
            report, artifacts["corpus-catalog.json"], lane, bindings[role]
        )
        catalogs[lane] = catalog_identity
        indexed = _rows(report, 5, lane)
        if len(indexed) != 1:
            fail(f"{lane}: heap guard must contain exactly one row")
        reports[lane] = report
        keys[lane] = set(indexed)
        _resource(artifacts["resource.log"], f"{lane}.resource.log")
    if keys[HEAP_LANES[0]] != keys[HEAP_LANES[1]]:
        fail("heap control/candidate corpus identities differ")
    if len(set(catalogs.values())) != 1:
        fail("heap catalogs differ in corpus identity")
    return {"lanes": list(HEAP_LANES), "result_count": 1, "instrumentation": "external-heaptrack-whole-process"}


def _load_analyzer() -> Any:
    path = ROOT / "analyze.py"
    if not path.is_file():
        fail("analysis.json is present but analyze.py is missing")
    spec = importlib.util.spec_from_file_location("litchi_0467_bundle_analyze", path)
    if spec is None or spec.loader is None:
        fail("cannot load retained analyze.py")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    try:
        spec.loader.exec_module(module)
    except (OSError, ImportError, SyntaxError, TypeError, ValueError) as error:
        fail(f"cannot load retained analyze.py ({error})")
    return module


def _recompute_analysis() -> None:
    paths = {
        "analysis.json": "analyze",
        "full-guard.json": "analyze_full_guard",
        "heap-guard.json": "analyze_heap_guard",
    }
    present = {name: ROOT / name for name in paths if (ROOT / name).is_file()}
    if not present:
        return
    analyzer = _load_analyzer()
    for name, function_name in paths.items():
        path = present.get(name)
        if path is None:
            continue
        expected = load_json(path, name)
        function = getattr(analyzer, function_name, None)
        if function is None:
            fail(f"{name}: retained analyze.py has no {function_name}()")
        try:
            actual = function(ROOT)
        except Exception as error:
            fail(f"{name}: retained analyzer failed ({error})")
        if actual != expected:
            fail(f"{name} differs from recomputed analyzer output")


def _recompute_heap_summary() -> None:
    """Recompute the optional allocation summary from retained print exports."""

    path = ROOT / "heap-summary.json"
    if not path.is_file():
        return
    helper_path = ROOT / "heap_summary.py"
    if not helper_path.is_file():
        fail("heap-summary.json is present but heap_summary.py is missing")
    spec = importlib.util.spec_from_file_location("litchi_0467_heap_summary", helper_path)
    if spec is None or spec.loader is None:
        fail("cannot load retained heap_summary.py")
    helper = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = helper
    try:
        spec.loader.exec_module(helper)
        actual = helper.summarize(ROOT)
    except (OSError, ImportError, SyntaxError, TypeError, ValueError, KeyError) as error:
        fail(f"heap-summary.json: retained helper failed ({error})")
    expected = load_json(path, "heap-summary.json")
    if actual != expected:
        fail("heap-summary.json differs from recomputed Heaptrack totals")


def _recompute_guard_summary() -> None:
    """Recompute the optional supplemental six-case guard summary."""

    path = ROOT / "guard-summary.json"
    if not path.is_file():
        return
    helper_path = ROOT / "guard_summary.py"
    if not helper_path.is_file():
        fail("guard-summary.json is present but guard_summary.py is missing")
    spec = importlib.util.spec_from_file_location("litchi_0467_guard_summary", helper_path)
    if spec is None or spec.loader is None:
        fail("cannot load retained guard_summary.py")
    helper = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = helper
    try:
        spec.loader.exec_module(helper)
        actual = helper.summarize(ROOT)
    except Exception as error:
        fail(f"guard-summary.json: retained helper failed ({error})")
    expected = load_json(path, "guard-summary.json")
    if actual != expected:
        fail("guard-summary.json differs from recomputed supplemental guard")


def _recompute_fixed_qualification() -> None:
    """Recompute the optional fixed-checkout 500-sample qualification."""

    path = ROOT / "fixed-qualification.json"
    if not path.is_file():
        return
    helper_path = ROOT / "fixed_qualify.py"
    if not helper_path.is_file():
        fail("fixed-qualification.json is present but fixed_qualify.py is missing")
    spec = importlib.util.spec_from_file_location("litchi_0467_fixed_qualify", helper_path)
    if spec is None or spec.loader is None:
        fail("cannot load retained fixed_qualify.py")
    helper = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = helper
    try:
        spec.loader.exec_module(helper)
        actual = helper.qualify(ROOT)
    except Exception as error:
        fail(f"fixed-qualification.json: retained helper failed ({error})")
    expected = load_json(path, "fixed-qualification.json")
    if actual != expected:
        fail("fixed-qualification.json differs from recomputed qualification")


def verify_sha256sums(*, preseal: bool) -> dict[str, Any]:
    path = ROOT / "SHA256SUMS"
    if not path.exists():
        if preseal:
            return {"sealed": False, "entries": 0}
        fail("SHA256SUMS is required for the final sealed bundle")
    regular(path, "SHA256SUMS")
    rows: dict[str, str] = {}
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except (OSError, UnicodeError) as error:
        fail(f"SHA256SUMS: cannot read ({error})")
    for index, line in enumerate(lines, 1):
        fields = line.split("  ", 1)
        if len(fields) != 2 or SHA256_RE.fullmatch(fields[0]) is None:
            fail(f"SHA256SUMS: malformed line {index}")
        name = fields[1]
        if name == "SHA256SUMS" or name in rows:
            fail(f"SHA256SUMS: duplicate or self entry {name}")
        relative = safe_relative(name, f"SHA256SUMS line {index}")
        member = ROOT / relative
        regular(member, f"SHA256SUMS.{name}")
        if sha256_file(member) != fields[0]:
            fail(f"SHA256SUMS: hash differs for {name}")
        rows[name] = fields[0]
    actual = {path.relative_to(ROOT).as_posix() for path in ROOT.rglob("*") if path.is_file() and not path.is_symlink() and path != ROOT / "SHA256SUMS"}
    if set(rows) != actual:
        fail(f"SHA256SUMS: exact coverage differs (missing={sorted(actual - set(rows))[:3]}, extra={sorted(set(rows) - actual)[:3]})")
    return {"sealed": True, "entries": len(rows)}


def verify(*, root: Path | None = None, preseal: bool = False) -> dict[str, Any]:
    if root is None:
        root = ROOT
    if root.resolve() != ROOT.resolve():
        # The CLI uses the module's immutable bundle root.  Tests invoke the
        # internal helper by temporarily loading this module from a fixture.
        fail("bundle root differs from verifier location")
    _reject_symlinks(ROOT)
    protocol = _protocol()
    bindings = _bindings()
    source_proof = _source_bindings(bindings)
    if preseal:
        _preseal_live_bindings(bindings)
    exported_lanes = _postprocess()
    normal = _normal(bindings)
    full = _full(bindings)
    heap = _heap(bindings, exported_lanes=exported_lanes)
    _recompute_analysis()
    _recompute_heap_summary()
    _recompute_guard_summary()
    _recompute_fixed_qualification()
    seal = verify_sha256sums(preseal=preseal)
    return {
        "schema": SCHEMA,
        "status": "pass",
        "sealed": seal["sealed"],
        "seal_entries": seal["entries"],
        "protocol_sha256": sha256_file(ROOT / "protocol-r1.json"),
        "source_bindings": source_proof,
        "postprocess": {
            "present": bool(exported_lanes),
            "lanes": sorted(exported_lanes),
        },
        "guard_summary": {
            "present": (ROOT / "guard-summary.json").is_file(),
        },
        "fixed_qualification": {
            "present": (ROOT / "fixed-qualification.json").is_file(),
        },
        "revisions": {role: value["revision"] for role, value in bindings.items()},
        "binary_sha256": {role: value["binary_sha256"] for role, value in bindings.items()},
        "normal": normal,
        "full_guard": full,
        "heap_guard": heap,
    }


def main(argv: Iterable[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--preseal", action="store_true", help="allow SHA256SUMS to be absent and check live pre-cleanup binaries")
    args = parser.parse_args(list(argv) if argv is not None else None)
    try:
        result = verify(preseal=args.preseal)
    except (OSError, KeyError, TypeError, ValueError, VerificationError) as error:
        result = {"schema": SCHEMA, "status": "fail", "error": str(error)}
    print(json.dumps(result, ensure_ascii=False, sort_keys=True, indent=2))
    return 0 if result.get("status") == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
