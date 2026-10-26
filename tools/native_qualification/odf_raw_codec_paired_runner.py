#!/usr/bin/env python3
"""Run the paired ODF raw-codec qualification protocol.

This runner is deliberately lane-specific.  It shares only the reviewed
process, source-closure, and resource helpers with the existing native
runner; route and endpoint identities are validated here against the ODF
raw-codec manifest.  The build phase creates four fixed role/variant
artifacts (baseline/candidate × normal/allocator).  The ABBA phase consumes
only those hashed artifacts and runs the ten raw-codec routes in the exact
baseline, candidate, candidate, baseline order.  Receipts are protocol
evidence only and make no acceptance claim.
"""

from __future__ import annotations

import argparse
import datetime as _dt
import hashlib
import json
import os
from pathlib import Path
import resource
import signal
import shutil
import subprocess
import sys
import time
import tomllib
from typing import Any, Mapping, Sequence


TOOLS_ROOT = Path(__file__).resolve().parents[1]
if str(TOOLS_ROOT) not in sys.path:
    sys.path.insert(0, str(TOOLS_ROOT))

import native_qualification_runner as generic  # noqa: E402
from native_qualification import odf_shared_triad_paired_runner as reviewed  # noqa: E402


SCHEMA = "litchi.odf.raw-codec-qualification.paired-runner.v1"
PREPARATION_SCHEMA = "litchi.odf.raw-codec-qualification.preparation.v2"
ROLES = ("baseline", "candidate")
VARIANTS = ("normal", "allocator")
ABBA_ORDER = ("baseline", "candidate", "candidate", "baseline")
PHASES = ("normal", "allocator")
WARMUPS = 3
SAMPLES = 30
PROBE_ITERATIONS = 128
TARGET_PREFLIGHT_BYTES = 19_200_000_000
TARGET_HARD_LIMIT_BYTES = 20_000_000_000
MEMORY_HARD_LIMIT_BYTES = 20_000_000_000
PROBE_MAX_BYTES = 1_048_576
PROBE_PATH = Path("tools/perf-baseline/examples/odf_raw_codec_qualification.rs")
CONSUMER_MANIFEST = Path("tools/perf-baseline/Cargo.toml")
PRODUCTION_MANIFEST = Path("crates/litchi-ops/Cargo.toml")
LOCK_PATH = Path("tools/perf-baseline/Cargo.lock")
CANDIDATE_COMMIT = "33a0a63d3553712c1276da592f7846a0c7dcabd6"
PROBE_NAME = PROBE_PATH.stem
ROUTE_IDS = (
    "archive_catalog_and_raw_split",
    "manifest_xml_stream_and_parse",
    "metadata_xml_and_namespace_resolution",
    "writer_memory_publication",
    "writer_stream_publication",
    "writer_bounded_policy",
    "raw_split_merge_and_opaque_preservation",
    "raw_xml_splice_merge_publication",
    "unified_facade_raw_parts_and_publication",
    "consumer_reopen_matrix",
)
EXPECTED_ROUTES = (
    {
        "id": "archive_catalog_and_raw_split",
        "family": "raw",
        "operations": ["Archive::new", "file_names", "contains", "metadata", "is_stored", "read"],
        "boundary": "Archive construction and catalog/read operation; package bytes and expected content are prepared outside the clock",
    },
    {
        "id": "manifest_xml_stream_and_parse",
        "family": "manifest",
        "operations": ["Archive::read_manifest_xml", "read_manifest", "parse_manifest"],
        "boundary": "prebuilt Archive manifest materialization/read/parse only; package bytes, Archive and limits are prepared outside the clock",
    },
    {
        "id": "metadata_xml_and_namespace_resolution",
        "family": "xml",
        "operations": ["Metadata::from_xml", "ResolvedReader::from_xml baseline/from_xml_document current", "read_resolved_event_into"],
        "boundary": "bounded metadata/namespace construction and event consumption; XML constants, limits and proof serialization are outside the clock",
    },
    {
        "id": "writer_memory_publication",
        "family": "writer",
        "operations": ["PackageWriter::new", "add_authored_xml", "add_file_reader_with_media_type_and_compression", "finish"],
        "boundary": "in-memory writer construction/publication; deterministic payload values are prepared outside and output proof is outside the clock",
    },
    {
        "id": "writer_stream_publication",
        "family": "writer",
        "operations": ["PackageWriter::with_writer", "add_file_reader_with_media_type_and_compression", "finish_to_writer"],
        "boundary": "stream writer/sink construction and reader-to-sink publication; deterministic payload is prepared outside the clock",
    },
    {
        "id": "writer_bounded_policy",
        "family": "writer",
        "operations": ["PackageWriter::new_bounded", "add_file", "finish_to_bounded_bytes"],
        "boundary": "bounded writer construction/publication and refusal; payload and output/error proof are outside the clock",
    },
    {
        "id": "raw_split_merge_and_opaque_preservation",
        "family": "merge",
        "operations": ["OwnedPackage::from_bytes", "replace_content_xml", "raw_identical_members"],
        "boundary": "source reopen and raw replacement/publication; expected payload, member audit and output proof are outside the clock",
    },
    {
        "id": "raw_xml_splice_merge_publication",
        "family": "splice",
        "operations": ["XmlSplicePublication::new", "XmlSplicePublication::replace", "rebuild_package_with_xml_splices"],
        "boundary": "prebuilt source part/range and fragment publication/rebuild only; source reopen/index/proof and output audit are outside the clock",
    },
    {
        "id": "unified_facade_raw_parts_and_publication",
        "family": "facade",
        "operations": ["open_bytes_with", "File::parts", "File::raw", "File::to_bytes", "File::metadata"],
        "boundary": "facade open, raw/parts/metadata access and byte-identical no-op publication; input bytes are intentionally cloned inside the measured API boundary",
    },
    {
        "id": "consumer_reopen_matrix",
        "family": "consumers",
        "operations": ["Archive::new", "Archive::read", "open_bytes_with", "File::parts", "File::raw"],
        "boundary": "each prebuilt ODF family package is reopened through raw and facade consumers inside the clock; package generation is outside",
    },
)
FAMILY_IDS = ("odt", "ods", "odp", "odg", "odc", "odi", "odb", "odf-formula", "oth", "odm")
ALLOCATOR_COUNTERS = (
    "allocations",
    "deallocations",
    "allocated_bytes",
    "deallocated_bytes",
    "reallocations",
)
NON_ACCEPTANCE = "no native acceptance claim"


class RunnerError(RuntimeError):
    """A fail-closed manifest, source, process, or receipt error."""


def _now() -> str:
    return _dt.datetime.now(_dt.timezone.utc).isoformat(timespec="milliseconds")


def _sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _sha256_file(path: Path) -> tuple[str, int]:
    if path.is_symlink() or not path.is_file():
        raise RunnerError(f"expected a regular file: {path}")
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        while block := stream.read(1024 * 1024):
            digest.update(block)
            size += len(block)
    return digest.hexdigest(), size


def _resolve(value: str | os.PathLike[str], *, base: Path | None = None) -> Path:
    path = Path(value)
    if not path.is_absolute():
        if base is None:
            raise RunnerError(f"path must be absolute: {value!r}")
        path = base / path
    current = Path(path.anchor)
    for part in path.parts[1:]:
        current /= part
        if current.is_symlink():
            raise RunnerError(f"symlink path component is forbidden: {path}")
    path = path.expanduser().resolve(strict=False)
    if "3rd" + "party" in path.parts:
        raise RunnerError(f"untracked upstream checkout path is forbidden: {path}")
    return path


def _under(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
    except ValueError:
        return False
    return True


def _disjoint(paths: Mapping[str, Path]) -> None:
    entries = list(paths.items())
    for index, (left_name, left) in enumerate(entries):
        for right_name, right in entries[index + 1 :]:
            if left == right or _under(left, right) or _under(right, left):
                raise RunnerError(f"paths must be disjoint: {left_name}={left}, {right_name}={right}")


def _reject_symlink_ancestors(path: Path, label: str) -> None:
    ancestors: list[Path] = []
    current = path.parent
    while True:
        ancestors.append(current)
        if current == current.parent:
            break
        current = current.parent
    for ancestor in reversed(ancestors):
        if ancestor.is_symlink():
            raise RunnerError(f"{label} ancestor may not be a symlink: {ancestor}")


def _load_json(path: Path, label: str) -> tuple[dict[str, Any], str, int]:
    if path.is_symlink() or not path.is_file():
        raise RunnerError(f"{label} is missing or symlinked: {path}")
    try:
        raw = path.read_bytes()
        data = json.loads(raw)
    except (OSError, json.JSONDecodeError) as error:
        raise RunnerError(f"cannot read {label}: {error}") from error
    if not isinstance(data, dict):
        raise RunnerError(f"{label} root must be an object")
    return data, _sha256_bytes(raw), len(raw)


def _require_sha(value: Any, label: str, length: int) -> str:
    if not isinstance(value, str) or len(value) != length or value != value.lower():
        raise RunnerError(f"{label} must be a lowercase {length}-character digest")
    alphabet = "0123456789abcdef"
    if any(char not in alphabet for char in value):
        raise RunnerError(f"{label} contains non-hex characters")
    return value


def _status_paths(status: str) -> list[str]:
    result: list[str] = []
    for line in status.splitlines():
        if len(line) < 4:
            raise RunnerError(f"malformed git status entry: {line!r}")
        path = line[3:]
        if " -> " in path:
            path = path.split(" -> ", 1)[1]
        result.append(path)
    return result


def _git(root: Path, *args: str) -> str:
    try:
        return reviewed._git(root, *args)
    except reviewed.RunnerError as error:
        raise RunnerError(str(error)) from error


def _source_state(root: Path) -> dict[str, Any]:
    """Take the expensive full closure snapshot only at protocol boundaries."""

    try:
        state = reviewed._source_state(root)
    except reviewed.RunnerError as error:
        raise RunnerError(str(error)) from error
    state["status_paths"] = _status_paths(str(state.get("status", "")))
    return state


def _production_source_tree(root: Path) -> str:
    return _sha256_bytes(
        _git(
            root,
            "ls-tree",
            "-r",
            "--full-tree",
            "HEAD",
            "--",
            "Cargo.toml",
            "Cargo.lock",
            ".cargo",
            "rust-toolchain.toml",
            "rust-toolchain",
            "crates",
        ).encode()
    )


def _cheap_snapshot(root: Path) -> dict[str, Any]:
    status = _git(root, "status", "--porcelain=v1", "--untracked-files=all")
    tracked_flags = _git(root, "ls-files", "-v", "-z")
    tracked_index = _git(root, "ls-files", "--stage", "-z")
    config_stats: list[dict[str, Any]] = []
    for path in reviewed._cargo_config_paths(root):
        digest, size = _sha256_file(path)
        stat = path.stat()
        config_stats.append({"path": str(path), "sha256": digest, "bytes": size, "mtime_ns": stat.st_mtime_ns})
    file_stats: list[dict[str, Any]] = []
    for relative in (CONSUMER_MANIFEST, PRODUCTION_MANIFEST, PROBE_PATH, LOCK_PATH):
        path = _resolve(relative, base=root)
        digest, size = _sha256_file(path)
        stat = path.stat()
        file_stats.append({"path": str(path), "sha256": digest, "bytes": size, "mtime_ns": stat.st_mtime_ns})
    return {
        "head": _git(root, "rev-parse", "HEAD").strip(),
        "tree": _git(root, "rev-parse", "HEAD^{tree}").strip(),
        "status_sha256": _sha256_bytes(status.encode()),
        "status_paths": _status_paths(status),
        "tracked_flags_sha256": _sha256_bytes(tracked_flags.encode()),
        "tracked_index_sha256": _sha256_bytes(tracked_index.encode()),
        "config_stats": config_stats,
        "file_stats": file_stats,
    }


def _file_identity(root: Path, relative: Path, label: str) -> dict[str, Any]:
    path = _resolve(relative, base=root)
    digest, size = _sha256_file(path)
    return {"path": str(path), "relative": str(relative), "label": label, "sha256": digest, "bytes": size}


def _package_identity(root: Path, relative: Path, expected_name: str) -> dict[str, Any]:
    path = _resolve(relative, base=root)
    try:
        package = tomllib.loads(path.read_text(encoding="utf-8"))["package"]
    except (OSError, KeyError, tomllib.TOMLDecodeError) as error:
        raise RunnerError(f"cannot parse package manifest {path}: {error}") from error
    if package.get("name") != expected_name:
        raise RunnerError(f"{path} is not package {expected_name!r}")
    digest, size = _sha256_file(path)
    return {"path": str(path), "relative": str(relative), "name": expected_name, "sha256": digest, "bytes": size}


def _validate_route_descriptors(routes: Any) -> list[dict[str, Any]]:
    if routes != list(EXPECTED_ROUTES):
        raise RunnerError("manifest route protocol differs from the reviewed ODF raw-codec contract")
    return [dict(route) for route in routes]


EXPECTED_SETUP_OUTSIDE = [
    "deterministic package and ODF family packages prepared before timing",
    "Archive/OwnedPackage/File opening for correctness proofs",
    "XML, metadata, namespace, reader/sink and expected-hash construction",
    "route-specific limit and refusal inputs",
]
EXPECTED_OUTPUT_PROOF = [
    "operation_ok",
    "changed_ok",
    "no_op_ok",
    "reopen_ok",
    "source_bytes_ok",
    "facade_bytes_ok",
    "untouched_members_ok",
    "manifest_ok",
    "metadata_ok",
    "namespace_ok",
    "bounded_error_ok",
    "negative_cases_ok",
]


def validate_manifest(data: Mapping[str, Any]) -> dict[str, Any]:
    if data.get("schema") != "litchi.odf.raw-codec-qualification.manifest.v1":
        raise RunnerError("manifest schema is not the ODF raw-codec schema")
    if data.get("lane") != "odf-raw-codec":
        raise RunnerError("manifest lane is not odf-raw-codec")
    _require_sha(data.get("baseline_commit"), "baseline_commit", 40)
    candidate = _require_sha(data.get("candidate_commit"), "candidate_commit", 40)
    if candidate == "15ab9324ef4f8c927acd6ecfc19d9707c3f886ee":
        raise RunnerError("historical 15ab candidate is not admissible")
    if candidate != CANDIDATE_COMMIT:
        raise RunnerError(f"candidate production pin must remain {CANDIDATE_COMMIT}")
    source = data.get("source")
    if not isinstance(source, Mapping):
        raise RunnerError("manifest source identity is missing")
    expected_source = {
        "consumer_manifest": str(CONSUMER_MANIFEST),
        "production_manifest": str(PRODUCTION_MANIFEST),
        "lock_path": str(LOCK_PATH),
        "baseline_overlay": [str(CONSUMER_MANIFEST), str(LOCK_PATH), str(PROBE_PATH)],
        "candidate_overlay": [str(CONSUMER_MANIFEST), str(PROBE_PATH)],
    }
    for key, expected in expected_source.items():
        if source.get(key) != expected:
            raise RunnerError(f"manifest source.{key} differs from the canonical path")
    consumer_sha = _require_sha(source.get("candidate_consumer_manifest_sha256"), "source.candidate_consumer_manifest_sha256", 64)
    lock_sha = _require_sha(source.get("candidate_lock_sha256"), "source.candidate_lock_sha256", 64)
    probe = data.get("probe")
    if not isinstance(probe, Mapping) or probe.get("path") != str(PROBE_PATH):
        raise RunnerError("manifest probe path differs")
    probe_sha = _require_sha(probe.get("sha256"), "probe.sha256", 64)
    if probe.get("features") != ["raw-codec-original", "raw-codec-current", "raw-codec-allocator"]:
        raise RunnerError("manifest feature adapter set differs")
    if probe.get("modes") != ["validate", "timed", "allocator"] or probe.get("max_iterations") != PROBE_ITERATIONS:
        raise RunnerError("manifest probe mode/iteration contract differs")
    routes = _validate_route_descriptors(data.get("routes"))
    if data.get("families") != list(FAMILY_IDS):
        raise RunnerError("manifest ODF family order differs")
    phases = data.get("phases")
    if not isinstance(phases, Mapping) or phases.get("setup_outside_clock") != EXPECTED_SETUP_OUTSIDE or phases.get("inside_clock") != "the route operation named by the exact descriptor boundary only" or phases.get("output_proof") != EXPECTED_OUTPUT_PROOF:
        raise RunnerError("manifest timing phase boundary is incomplete")
    if phases.get("negative_cases") != [
        "duplicate_manifest_path",
        "manifest_limit",
        "malformed_xml",
        "undeclared_namespace",
        "stale_xml_range",
        "foreign_xml_source",
        "bounded_output",
        "failing_reader_or_sink",
        "corrupt_stored_payload_without_verification",
    ]:
        raise RunnerError("manifest negative-case contract differs")
    paired = data.get("paired")
    if not isinstance(paired, Mapping):
        raise RunnerError("manifest paired contract is missing")
    if paired.get("abba_order") != list(ABBA_ORDER) or paired.get("warmup_iterations_per_leg") != WARMUPS or paired.get("samples_per_leg") != SAMPLES or paired.get("probe_iterations") != PROBE_ITERATIONS:
        raise RunnerError("manifest ABBA counts/order differ")
    if paired.get("roles") != list(ROLES) or paired.get("variants") != list(VARIANTS):
        raise RunnerError("manifest role/variant contract differs")
    calibration = paired.get("calibration")
    if calibration != {"mode": "validate", "iterations": 1}:
        raise RunnerError("manifest calibration must be one validate iteration")
    allocator = paired.get("allocator")
    if not isinstance(allocator, Mapping) or allocator.get("scope") != "operation_global_system_allocator" or allocator.get("elapsed_claimable") is not False or allocator.get("counters") != list(ALLOCATOR_COUNTERS):
        raise RunnerError("manifest allocator observer contract differs")
    quiet = paired.get("quiet")
    if not isinstance(quiet, Mapping) or quiet.get("process_guard") != "tools.native_qualification_runner.ProcessGuard" or quiet.get("cpu_affinity_required") is not True:
        raise RunnerError("manifest quiet-host guard differs")
    guards = data.get("guards")
    if not isinstance(guards, Mapping) or guards.get("cargo") != {"offline": True, "locked": True, "jobs": 1, "incremental": False}:
        raise RunnerError("manifest Cargo guard differs")
    if guards.get("target_preflight_bytes") != TARGET_PREFLIGHT_BYTES or guards.get("target_hard_limit_bytes") != TARGET_HARD_LIMIT_BYTES or guards.get("memory_hard_limit_bytes") != MEMORY_HARD_LIMIT_BYTES or guards.get("max_probe_output_bytes") != PROBE_MAX_BYTES:
        raise RunnerError("manifest resource limits differ")
    if guards.get("fresh_no_follow_outputs") is not True or guards.get("full_source_closure") is not True or guards.get("quiet_host") is not True:
        raise RunnerError("manifest source/output guard is incomplete")
    claim = data.get("claim_scope")
    if not isinstance(claim, str) or NON_ACCEPTANCE not in claim:
        raise RunnerError("manifest must be explicitly non-acceptance")
    normalized = dict(data)
    normalized["routes"] = routes
    normalized["candidate_lock_sha256"] = lock_sha
    normalized["candidate_consumer_manifest_sha256"] = consumer_sha
    normalized["probe_sha256"] = probe_sha
    normalized["candidate_commit"] = candidate
    return normalized


def _validate_preparation(data: Mapping[str, Any], digest: str, *, manifest: Mapping[str, Any], baseline_root: Path, candidate_root: Path, adapter_root: Path, candidate_commit: str, candidate_lock_sha256: str) -> dict[str, Any]:
    if data.get("schema") != PREPARATION_SCHEMA or data.get("lane") != "odf-raw-codec":
        raise RunnerError("preparation schema/lane differs")
    baseline = data.get("baseline")
    candidate = data.get("candidate")
    if not isinstance(baseline, Mapping) or not isinstance(candidate, Mapping):
        raise RunnerError("preparation endpoints are missing")
    before = baseline.get("before")
    after = baseline.get("after_overlay")
    if not isinstance(before, Mapping) or not isinstance(after, Mapping):
        raise RunnerError("preparation source states are missing")
    if _resolve(str(before.get("root"))) != baseline_root or before.get("head") != manifest["baseline_commit"] or before.get("clean") is not True:
        raise RunnerError("preparation baseline before-state differs")
    if after.get("tree") != before.get("tree"):
        raise RunnerError("preparation changed baseline production tree")
    if baseline.get("commit") != manifest["baseline_commit"] or baseline.get("production_tree_sha256") != reviewed._production_tree(baseline_root):
        raise RunnerError("preparation baseline identity differs")
    overlay = baseline.get("overlay_paths")
    if sorted(overlay or []) != sorted([str(CONSUMER_MANIFEST), str(LOCK_PATH), str(PROBE_PATH)]):
        raise RunnerError("preparation baseline overlay is not exactly consumer, lock plus probe")
    if _resolve(str(candidate.get("root"))) != candidate_root or candidate.get("commit") != candidate_commit:
        raise RunnerError("preparation candidate identity differs")
    candidate_before = candidate.get("before")
    candidate_after = candidate.get("after_overlay")
    if not isinstance(candidate_before, Mapping) or not isinstance(candidate_after, Mapping):
        raise RunnerError("preparation candidate source states are missing")
    if _resolve(str(candidate_before.get("root"))) != candidate_root or candidate_before.get("head") != manifest["candidate_commit"] or candidate_before.get("clean") is not True:
        raise RunnerError("preparation candidate production before-state differs")
    if _resolve(str(candidate_after.get("root"))) != candidate_root or candidate_after.get("tree") != candidate_before.get("tree") or sorted(_status_paths(str(candidate_after.get("status", "")))) != sorted([str(CONSUMER_MANIFEST), str(PROBE_PATH)]) or sorted(candidate.get("overlay_paths") or []) != sorted([str(CONSUMER_MANIFEST), str(PROBE_PATH)]):
        raise RunnerError("preparation candidate overlay differs")
    if candidate.get("production_tree_sha256") != reviewed._production_tree(candidate_root):
        raise RunnerError("preparation candidate production tree differs")
    adapter = data.get("adapter")
    if not isinstance(adapter, Mapping) or _resolve(str(adapter.get("root"))) != adapter_root:
        raise RunnerError("preparation adapter endpoint is missing or overlaps a source endpoint")
    if adapter_root in {baseline_root, candidate_root} or _under(adapter_root, baseline_root) or _under(adapter_root, candidate_root) or _under(baseline_root, adapter_root) or _under(candidate_root, adapter_root):
        raise RunnerError("preparation adapter endpoint overlaps a source endpoint")
    adapter_state = _source_state(adapter_root)
    if adapter_state.get("commit") != adapter.get("head") or adapter_state.get("tree") != adapter.get("tree") or adapter_state.get("status_paths"):
        raise RunnerError("consumer adapter endpoint changed or is dirty")
    if adapter.get("production_source_tree_sha256") != candidate.get("production_source_tree_sha256"):
        raise RunnerError("adapter production source does not match the pinned candidate")
    if adapter.get("production_source_tree_sha256") != _production_source_tree(adapter_root) or candidate.get("production_source_tree_sha256") != _production_source_tree(candidate_root):
        raise RunnerError("adapter or candidate production source tree is not bound")
    if candidate_commit != manifest["candidate_commit"] or candidate_commit == "15ab9324ef4f8c927acd6ecfc19d9707c3f886ee":
        raise RunnerError("preparation candidate is stale or differs from manifest")
    candidate_consumer = candidate.get("consumer_manifest")
    candidate_probe = candidate.get("probe")
    candidate_lock = candidate.get("lock")
    baseline_probe = baseline.get("probe")
    baseline_lock = baseline.get("lock")
    baseline_consumer = baseline.get("consumer_manifest")
    for value, label in ((candidate_consumer, "candidate consumer manifest"), (candidate_probe, "candidate probe"), (candidate_lock, "candidate lock"), (baseline_consumer, "baseline consumer manifest"), (baseline_probe, "baseline probe"), (baseline_lock, "baseline lock")):
        if not isinstance(value, Mapping):
            raise RunnerError(f"preparation {label} identity is missing")
        _require_sha(value.get("sha256"), f"preparation {label} SHA", 64)
        if not isinstance(value.get("bytes"), int) or value["bytes"] <= 0:
            raise RunnerError(f"preparation {label} size is invalid")
    if candidate_consumer["sha256"] != manifest["candidate_consumer_manifest_sha256"] or candidate_probe["sha256"] != manifest["probe_sha256"] or candidate_lock["sha256"] != candidate_lock_sha256 or baseline_consumer["sha256"] != manifest["candidate_consumer_manifest_sha256"] or baseline_probe["sha256"] != manifest["probe_sha256"]:
        raise RunnerError("preparation adapter/lock does not match selected manifest")
    if _resolve(str(candidate_consumer.get("path"))) != _resolve(CONSUMER_MANIFEST, base=candidate_root):
        raise RunnerError("preparation candidate consumer manifest path is not canonical")
    if _resolve(str(baseline_consumer.get("source_path"))) != _resolve(str(adapter["consumer_manifest"]["path"])):
        raise RunnerError("preparation baseline consumer source is not the selected adapter manifest")
    if _resolve(str(candidate_probe.get("path"))) != _resolve(PROBE_PATH, base=candidate_root):
        raise RunnerError("preparation candidate probe path is not canonical")
    if _resolve(str(candidate_lock.get("path"))) != _resolve(LOCK_PATH, base=candidate_root):
        raise RunnerError("preparation candidate lock path is not canonical")
    if _resolve(str(baseline_probe.get("source_path"))) != _resolve(str(adapter["probe"]["path"])):
        raise RunnerError("preparation baseline probe source is not the selected adapter probe")
    adapter_consumer = adapter.get("consumer_manifest")
    adapter_probe = adapter.get("probe")
    for value, label in ((adapter_consumer, "adapter consumer manifest"), (adapter_probe, "adapter probe")):
        if not isinstance(value, Mapping) or _resolve(str(value.get("path"))).is_symlink():
            raise RunnerError(f"preparation {label} identity is missing or symlinked")
        observed, size = _sha256_file(_resolve(str(value["path"])))
        if observed != value.get("sha256") or size != value.get("bytes"):
            raise RunnerError(f"preparation {label} bytes differ")
    if adapter_consumer["sha256"] != manifest["candidate_consumer_manifest_sha256"] or adapter_probe["sha256"] != manifest["probe_sha256"]:
        raise RunnerError("adapter overlay bytes differ from the manifest")
    if _resolve(str(candidate_consumer.get("source_path"))) != _resolve(str(adapter_consumer["path"])) or _resolve(str(candidate_probe.get("source_path"))) != _resolve(str(adapter_probe["path"])):
        raise RunnerError("candidate overlay source is not the selected adapter")
    baseline_lock_source = _resolve(str(baseline_lock.get("source_path")))
    if _under(baseline_lock_source, baseline_root) or _under(baseline_lock_source, candidate_root) or _under(baseline_lock_source, adapter_root):
        raise RunnerError("preparation baseline lock source must be external to every source endpoint")
    source_lock_sha, source_lock_bytes = _sha256_file(baseline_lock_source)
    if source_lock_sha != baseline_lock["sha256"] or source_lock_bytes != baseline_lock["bytes"]:
        raise RunnerError("preparation baseline lock source bytes differ")
    for value, relative, root, label in (
        (baseline_consumer, CONSUMER_MANIFEST, baseline_root, "baseline consumer manifest"),
        (baseline_probe, PROBE_PATH, baseline_root, "baseline probe"),
        (baseline_lock, LOCK_PATH, baseline_root, "baseline lock"),
    ):
        destination = _resolve(str(value.get("destination_path")))
        expected = _resolve(relative, base=root)
        if destination != expected or not destination.is_file():
            raise RunnerError(f"preparation {label} destination differs")
        observed, size = _sha256_file(destination)
        if observed != value["sha256"] or size != value["bytes"]:
            raise RunnerError(f"preparation {label} bytes differ")
    claim = data.get("claim_scope")
    if not isinstance(claim, str) or NON_ACCEPTANCE not in claim:
        raise RunnerError("preparation is not explicitly non-acceptance")
    return {"schema": data["schema"], "sha256": digest, "baseline": baseline, "candidate": candidate, "adapter": adapter}


def _endpoint(manifest: Mapping[str, Any], *, role: str, root: Path, preparation: Mapping[str, Any], candidate_commit: str, candidate_lock_sha256: str) -> dict[str, Any]:
    state = _source_state(root)
    package = _package_identity(root, PRODUCTION_MANIFEST, "litchi-ops")
    consumer = _package_identity(root, CONSUMER_MANIFEST, "litchi-perf-baseline")
    probe = _file_identity(root, PROBE_PATH, "probe")
    lock = _file_identity(root, LOCK_PATH, "Cargo.lock")
    if role == "baseline":
        if state.get("commit") != manifest["baseline_commit"]:
            raise RunnerError("baseline commit differs")
        if state.get("production_tree_sha256") != preparation["baseline"]["production_tree_sha256"]:
            raise RunnerError("baseline production tree differs")
        if sorted(state["status_paths"]) != sorted(preparation["baseline"]["overlay_paths"]):
            raise RunnerError("baseline has paths outside the reviewed overlay")
        if consumer["sha256"] != preparation["baseline"]["consumer_manifest"]["sha256"] or probe["sha256"] != manifest["probe_sha256"] or lock["sha256"] != preparation["baseline"]["lock"]["sha256"]:
            raise RunnerError("baseline overlay file identity differs")
    elif role == "candidate":
        if state.get("commit") != candidate_commit or sorted(state["status_paths"]) != sorted(preparation["candidate"]["overlay_paths"]):
            raise RunnerError("candidate source is stale or has an unexpected overlay")
        if state.get("production_tree_sha256") != preparation["candidate"]["production_tree_sha256"] or _production_source_tree(root) != preparation["candidate"]["production_source_tree_sha256"]:
            raise RunnerError("candidate production source identity differs")
        if consumer["sha256"] != manifest["candidate_consumer_manifest_sha256"] or probe["sha256"] != manifest["probe_sha256"] or lock["sha256"] != candidate_lock_sha256:
            raise RunnerError("candidate probe/lock identity differs")
    else:
        raise RunnerError(f"unknown endpoint role {role!r}")
    return {
        "role": role,
        "root": str(root),
        "commit": state["commit"],
        "tree": state["tree"],
        "production_tree_sha256": state["production_tree_sha256"],
        "production_source_tree_sha256": _production_source_tree(root),
        "source_tree_sha256": state["source_tree_sha256"],
        "status_sha256": state["status_sha256"],
        "status_paths": state["status_paths"],
        "tracked_closure": state["tracked_closure"],
        "cargo_configs": state["cargo_configs"],
        "consumer_manifest": consumer,
        "production_manifest": package,
        "probe": probe,
        "lock": lock,
        "features": ["raw-codec-odf", "raw-codec-original", "raw-codec-current", "raw-codec-allocator"],
        "routes": list(manifest["routes"]),
        "families": list(FAMILY_IDS),
    }


def _build_keys() -> tuple[str, ...]:
    return tuple(f"{role}/{variant}" for variant in VARIANTS for role in ROLES)


def _journal(events: list[Mapping[str, Any]]) -> dict[str, Any]:
    expected = [f"build-{key.replace('/', '-') }" for key in _build_keys()]
    actual = [event.get("event_id") for event in events]
    if actual != expected:
        raise RunnerError(f"build journal order differs: {actual!r}")
    encoded = json.dumps(events, sort_keys=True, separators=(",", ":")).encode()
    return {"count": len(events), "sha256": _sha256_bytes(encoded), "events": events}


def _journal_events(value: Any) -> dict[str, Mapping[str, Any]]:
    if not isinstance(value, Mapping) or value.get("count") != len(_build_keys()) or not isinstance(value.get("events"), list):
        raise RunnerError("build journal is incomplete")
    events = value["events"]
    expected = [f"build-{key.replace('/', '-') }" for key in _build_keys()]
    if [event.get("event_id") for event in events if isinstance(event, Mapping)] != expected:
        raise RunnerError("build journal order differs")
    digest = _sha256_bytes(json.dumps(events, sort_keys=True, separators=(",", ":")).encode())
    if value.get("sha256") != digest:
        raise RunnerError("build journal hash differs")
    return {event["event_id"]: event for event in events if isinstance(event, Mapping)}


def _resolve_executable(name: str, cwd: Path, env: Mapping[str, str]) -> Path:
    for value in env.get("PATH", os.defpath).split(os.pathsep):
        directory = Path(value or ".")
        if not directory.is_absolute():
            directory = cwd / directory
        candidate = directory / name
        if candidate.is_file() and os.access(candidate, os.X_OK):
            return candidate.resolve(strict=True)
    raise RunnerError(f"{name} is not executable in the build PATH")


class PairedRunner:
    def __init__(self, *, manifest_path: Path, preparation_path: Path, baseline_root: Path, candidate_root: Path, adapter_root: Path, candidate_commit: str, candidate_lock_sha256: str, output: Path, target_root: Path, tmp_root: Path, cpu: int, phase: str) -> None:
        if phase not in {"build", "abba"}:
            raise RunnerError("phase must be build or abba")
        if not isinstance(cpu, int) or cpu < 0 or cpu >= (os.cpu_count() or 1):
            raise RunnerError("cpu must be an available CPU index")
        self.phase = phase
        self.manifest_path = _resolve(manifest_path)
        self.preparation_path = _resolve(preparation_path)
        self.baseline_root = _resolve(baseline_root)
        self.candidate_root = _resolve(candidate_root)
        self.adapter_root = _resolve(adapter_root)
        self.candidate_commit = _require_sha(candidate_commit, "candidate_commit", 40)
        self.candidate_lock_sha256 = _require_sha(candidate_lock_sha256, "candidate_lock_sha256", 64)
        self.output = _resolve(output)
        self.target_root = _resolve(target_root)
        self.tmp_root = _resolve(tmp_root)
        self.cpu = cpu
        _disjoint({"baseline": self.baseline_root, "candidate": self.candidate_root, "adapter": self.adapter_root, "preparation": self.preparation_path, "output": self.output, "target": self.target_root, "tmp": self.tmp_root})
        self.manifest, self.manifest_sha256, _ = _load_json(self.manifest_path, "manifest")
        self.manifest = validate_manifest(self.manifest)
        preparation, prep_sha, _ = _load_json(self.preparation_path, "preparation")
        self.preparation = _validate_preparation(preparation, prep_sha, manifest=self.manifest, baseline_root=self.baseline_root, candidate_root=self.candidate_root, adapter_root=self.adapter_root, candidate_commit=self.candidate_commit, candidate_lock_sha256=self.candidate_lock_sha256)
        self.runner_sha256, _ = _sha256_file(Path(__file__).resolve())
        self.guard = generic.ProcessGuard()
        self.endpoint_cache: dict[str, dict[str, Any]] = {}
        if phase == "build":
            if self.output.exists() and (self.output.is_symlink() or not self.output.is_dir() or any(self.output.iterdir())):
                raise RunnerError("build output must be absent or empty")
            _reject_symlink_ancestors(self.output, "build output")
            self.output.mkdir(parents=True, exist_ok=True)
        else:
            if self.output.is_symlink() or not self.output.is_dir() or not (self.output / "build-receipt.json").is_file():
                raise RunnerError("ABBA requires an existing passed build output")
        for root, label in ((self.target_root, "target root"), (self.tmp_root, "tmp root")):
            _reject_symlink_ancestors(root, label)
            if root.is_symlink():
                raise RunnerError(f"{label} may not be a symlink")
            if phase == "build" and root.exists() and any(root.iterdir()):
                raise RunnerError(f"{label} must be fresh or empty for build")
            root.mkdir(parents=True, exist_ok=True)

    def _fresh_file(self, path: Path, label: str) -> None:
        adapter_root = getattr(self, "adapter_root", None)
        if _under(path, self.baseline_root) or _under(path, self.candidate_root) or (adapter_root is not None and _under(path, adapter_root)):
            raise RunnerError(f"{label} must be external to source")
        _reject_symlink_ancestors(path, label)
        if path.exists() or path.is_symlink():
            raise RunnerError(f"{label} must not overwrite an existing path: {path}")
        path.parent.mkdir(parents=True, exist_ok=True)

    def _fresh_dir(self, path: Path, label: str) -> None:
        adapter_root = getattr(self, "adapter_root", None)
        if _under(path, self.baseline_root) or _under(path, self.candidate_root) or (adapter_root is not None and _under(path, adapter_root)):
            raise RunnerError(f"{label} must be external to source")
        _reject_symlink_ancestors(path, label)
        if path.exists() and (path.is_symlink() or not path.is_dir() or any(path.iterdir())):
            raise RunnerError(f"{label} must be absent or empty: {path}")
        path.mkdir(parents=True, exist_ok=True)

    def _target(self, role: str, variant: str) -> Path:
        return self.target_root / f"{role}-{variant}"

    def _tmp(self, role: str, variant: str) -> Path:
        return self.tmp_root / f"{role}-{variant}"

    def _endpoint(self, role: str, *, full: bool = False) -> dict[str, Any]:
        root = self.baseline_root if role == "baseline" else self.candidate_root
        if not full and role in self.endpoint_cache:
            cached = self.endpoint_cache[role]
            if _cheap_snapshot(root) != cached["cheap"]:
                raise RunnerError(f"{role} source changed between operations")
            return cached["endpoint"]
        endpoint = _endpoint(self.manifest, role=role, root=root, preparation=self.preparation, candidate_commit=self.candidate_commit, candidate_lock_sha256=self.candidate_lock_sha256)
        self.endpoint_cache[role] = {"endpoint": endpoint, "cheap": _cheap_snapshot(root)}
        return endpoint

    def _assert_endpoint(self, role: str, expected: Mapping[str, Any], *, full: bool = False) -> None:
        actual = self._endpoint(role, full=full)
        for key in ("commit", "tree", "production_tree_sha256", "source_tree_sha256", "status_sha256", "status_paths", "tracked_closure", "cargo_configs", "consumer_manifest", "production_manifest", "probe", "lock", "features", "routes", "families"):
            if actual.get(key) != expected.get(key):
                raise RunnerError(f"{role} source identity changed: {key}")

    def _environment(self, role: str, variant: str, target: Path, tmpdir: Path) -> dict[str, str]:
        allowed = {"AR", "CC", "CFLAGS", "CXX", "CXXFLAGS", "HOME", "LANG", "LD_LIBRARY_PATH", "LOGNAME", "MAKEFLAGS", "PATH", "PKG_CONFIG_PATH", "PKG_CONFIG_SYSROOT_DIR", "RANLIB", "RUSTUP_HOME", "RUSTUP_TOOLCHAIN", "SOURCE_DATE_EPOCH", "TEMP", "TMP", "USER"}
        env = {key: value for key, value in os.environ.items() if key in allowed or key.startswith("LC_")}
        env.update({"CARGO_TARGET_DIR": str(target), "CARGO_BUILD_JOBS": "1", "CARGO_INCREMENTAL": "0", "TMPDIR": str(tmpdir), "LITCHI_MEM_MAX": "20G", "LITCHI_QUALIFICATION_ROLE": role, "LITCHI_QUALIFICATION_VARIANT": variant})
        return env

    @staticmethod
    def _recorded_environment(env: Mapping[str, str]) -> dict[str, str]:
        return {key: env[key] for key in sorted(env)}

    def _preexec(self) -> Any:
        cpu = self.cpu

        def configure() -> None:
            os.sched_setaffinity(0, {cpu})
            resource.setrlimit(resource.RLIMIT_AS, (MEMORY_HARD_LIMIT_BYTES, MEMORY_HARD_LIMIT_BYTES))

        return configure

    def _toolchain(self, cwd: Path, env: Mapping[str, str]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for argv, label in ((["rustc", "-Vv"], "rustc"), (["cargo", "-V"], "cargo")):
            executable = _resolve_executable(argv[0], cwd, env)
            digest, size = _sha256_file(executable)
            completed = subprocess.run(argv, cwd=str(cwd), env=dict(env), check=False, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            if completed.returncode != 0:
                raise RunnerError(f"{label} identity command failed: {completed.stderr.strip()}")
            result[label] = {"argv": argv, "cwd": str(cwd), "environment": self._recorded_environment(env), "returncode": 0, "stdout": completed.stdout, "stderr": completed.stderr, "executable": {"path": str(executable), "sha256": digest, "bytes": size}}
        return result

    def _run_process(self, *, event_id: str, role: str, variant: str, argv: Sequence[str], cwd: Path, env: Mapping[str, str], target: Path, stdout_path: Path, stderr_path: Path) -> dict[str, Any]:
        self.guard.require_quiet()
        self._fresh_file(stdout_path, f"{event_id} stdout")
        self._fresh_file(stderr_path, f"{event_id} stderr")
        before = generic._directory_size(target)
        if before > TARGET_PREFLIGHT_BYTES:
            raise RunnerError(f"target exceeds preflight limit before {event_id}")
        process: subprocess.Popen[Any] | None = None
        monitor: list[dict[str, Any]] = []
        violation: list[dict[str, Any]] = []
        error: str | None = None
        started = time.monotonic_ns()
        with stdout_path.open("xb") as stdout, stderr_path.open("xb") as stderr:
            try:
                process = subprocess.Popen(list(argv), cwd=str(cwd), env=dict(env), stdin=subprocess.DEVNULL, stdout=stdout, stderr=stderr, start_new_session=True, preexec_fn=self._preexec())
                while True:
                    returncode = process.poll()
                    if returncode is not None or not monitor or time.monotonic() - monitor[-1]["monotonic"] >= 0.25:
                        target_bytes = generic._directory_size(target)
                        rss = None
                        monitor_error: str | None = None
                        try:
                            status = (Path("/proc") / str(process.pid) / "status").read_text(encoding="utf-8", errors="replace")
                            rss = {key: int(line.split()[1]) * 1024 for line in status.splitlines() if line.split(":", 1)[0] in {"VmRSS", "VmHWM", "VmSize"}}
                        except FileNotFoundError:
                            if returncode is None:
                                monitor_error = "live RSS monitor lost the child process"
                        except (OSError, ValueError, IndexError) as exc:
                            monitor_error = f"live RSS monitor failed: {exc}"
                        monitor.append({"utc": _now(), "monotonic": time.monotonic(), "loadavg": self.guard.loadavg(), "target_bytes": target_bytes, "rss": rss})
                        if monitor_error:
                            error = monitor_error
                            if process.poll() is None:
                                self._terminate(process)
                            break
                        if target_bytes > TARGET_HARD_LIMIT_BYTES:
                            error = f"target exceeded hard limit during {event_id}"
                            if process.poll() is None:
                                self._terminate(process)
                            break
                    if returncode is not None:
                        break
                    offenders = self.guard.offenders(process.pid)
                    if offenders:
                        violation.extend(offenders)
                        error = "unrelated Litchi process appeared during child execution"
                        self._terminate(process)
                        break
                    time.sleep(0.05)
            except BaseException as exc:
                error = str(exc)
                if process is not None and process.poll() is None:
                    self._terminate(process)
        if process is None:
            raise RunnerError(f"process did not start for {event_id}: {error}")
        returncode = process.returncode
        descendants = reviewed._descendants_gone(process.pid)
        offenders = self.guard.offenders()
        after = generic._directory_size(target)
        if descendants:
            error = error or f"child descendants survived: {descendants}"
        if offenders:
            error = error or "unrelated Litchi process remained after child exit"
        if after > TARGET_HARD_LIMIT_BYTES:
            error = error or f"target exceeded hard limit after {event_id}"
        stdout_sha, stdout_bytes = _sha256_file(stdout_path)
        stderr_sha, stderr_bytes = _sha256_file(stderr_path)
        event = {"schema": SCHEMA, "event_id": event_id, "role": role, "variant": variant, "argv": list(argv), "cwd": str(cwd), "environment": self._recorded_environment(env), "cpu_affinity": self.cpu, "start_utc": _now(), "duration_ns": time.monotonic_ns() - started, "returncode": returncode, "stdout": {"path": str(stdout_path), "sha256": stdout_sha, "bytes": stdout_bytes}, "stderr": {"path": str(stderr_path), "sha256": stderr_sha, "bytes": stderr_bytes}, "target": {"path": str(target), "bytes_before": before, "bytes_after": after}, "monitor": [{key: value for key, value in sample.items() if key != "monotonic"} for sample in monitor], "process_guard_violations": violation, "post_exit_descendants": descendants, "post_exit_offenders": offenders, "error": error}
        if error:
            raise RunnerError(f"{error}; stderr={stderr_path}")
        if returncode != 0:
            raise RunnerError(f"child failed during {event_id} with exit {returncode}; stderr={stderr_path}")
        return event

    @staticmethod
    def _terminate(process: subprocess.Popen[Any]) -> None:
        try:
            os.killpg(process.pid, signal.SIGTERM)
        except (ProcessLookupError, PermissionError):
            process.terminate()
        try:
            process.wait(timeout=5)
        except subprocess.TimeoutExpired:
            try:
                os.killpg(process.pid, signal.SIGKILL)
            except (ProcessLookupError, PermissionError):
                process.kill()
            process.wait()

    @staticmethod
    def _build_command(role: str, variant: str) -> list[str]:
        adapter = "raw-codec-original" if role == "baseline" else "raw-codec-current"
        features = [adapter]
        if variant == "allocator":
            features.append("raw-codec-allocator")
        return ["cargo", "build", "--offline", "--locked", "--jobs", "1", "--manifest-path", str(CONSUMER_MANIFEST), "-p", "litchi-perf-baseline", "--example", PROBE_NAME, "--no-default-features", "--features", ",".join(features)]

    def _binary(self, role: str, variant: str) -> Path:
        return self._target(role, variant) / "debug" / "examples" / PROBE_NAME

    def _build_one(self, role: str, variant: str) -> dict[str, Any]:
        root = self.baseline_root if role == "baseline" else self.candidate_root
        target = self._target(role, variant)
        tmpdir = self._tmp(role, variant)
        self._fresh_dir(target, f"{role}/{variant} target")
        self._fresh_dir(tmpdir, f"{role}/{variant} tmp")
        endpoint = self._endpoint(role, full=True)
        env = self._environment(role, variant, target, tmpdir)
        toolchain = self._toolchain(root, env)
        command = self._build_command(role, variant)
        stdout = self.output / "build" / f"{role}-{variant}.stdout"
        stderr = self.output / "build" / f"{role}-{variant}.stderr"
        event = self._run_process(event_id=f"build-{role}-{variant}", role=role, variant=variant, argv=command, cwd=root, env=env, target=target, stdout_path=stdout, stderr_path=stderr)
        binary = self._binary(role, variant)
        if binary.is_symlink() or not binary.is_file() or not os.access(binary, os.X_OK):
            raise RunnerError(f"built probe binary is missing or not executable: {binary}")
        binary_sha, binary_bytes = _sha256_file(binary)
        event["binary"] = {"path": str(binary), "sha256": binary_sha, "bytes": binary_bytes}
        event["toolchain"] = toolchain
        self._assert_endpoint(role, endpoint, full=True)
        target_bytes = generic._directory_size(target)
        if target_bytes > TARGET_PREFLIGHT_BYTES:
            raise RunnerError(f"{role}/{variant} target exceeded preflight limit")
        return {"role": role, "variant": variant, "command": command, "environment": self._recorded_environment(env), "toolchain": toolchain, "source": endpoint, "binary": {"path": str(binary), "sha256": binary_sha, "bytes": binary_bytes}, "target": {"path": str(target), "bytes": target_bytes}, "tmpdir": str(tmpdir), "event": event}

    def _write_fresh(self, path: Path, value: Mapping[str, Any], label: str) -> None:
        self._fresh_file(path, label)
        with path.open("x", encoding="utf-8") as stream:
            json.dump(value, stream, indent=2, sort_keys=True)
            stream.write("\n")
            stream.flush()
            os.fsync(stream.fileno())

    def build(self) -> None:
        builds: dict[str, Any] = {}
        endpoints = {role: self._endpoint(role, full=True) for role in ROLES}
        for variant in VARIANTS:
            for role in ROLES:
                builds[f"{role}/{variant}"] = self._build_one(role, variant)
        receipt = {"schema": SCHEMA, "phase": "build", "status": "passed", "manifest": {"path": str(self.manifest_path), "sha256": self.manifest_sha256}, "preparation": {"path": str(self.preparation_path), **self.preparation}, "runner": {"path": str(Path(__file__).resolve()), "sha256": self.runner_sha256}, "cpu_affinity": self.cpu, "features": ["raw-codec-odf", "raw-codec-original", "raw-codec-current", "raw-codec-allocator"], "routes": self.manifest["routes"], "families": list(FAMILY_IDS), "paired": {key: self.manifest["paired"][key] for key in ("abba_order", "warmup_iterations_per_leg", "samples_per_leg", "probe_iterations", "roles", "variants")}, "endpoints": endpoints, "builds": builds, "event_journal": _journal([builds[key]["event"] for key in _build_keys()]), "claim_scope": "build identity only; no native acceptance claim"}
        self._write_fresh(self.output / "build-receipt.json", receipt, "build receipt")

    def _load_build_receipt(self) -> dict[str, Any]:
        receipt, receipt_sha, _ = _load_json(self.output / "build-receipt.json", "build receipt")
        if receipt.get("schema") != SCHEMA or receipt.get("phase") != "build" or receipt.get("status") != "passed" or NON_ACCEPTANCE not in str(receipt.get("claim_scope")):
            raise RunnerError("build receipt is not a passed non-acceptance receipt")
        if receipt.get("manifest", {}).get("sha256") != self.manifest_sha256 or receipt.get("preparation", {}).get("sha256") != self.preparation["sha256"] or receipt.get("runner", {}).get("sha256") != self.runner_sha256:
            raise RunnerError("build handoff identity differs")
        if receipt.get("cpu_affinity") != self.cpu or receipt.get("routes") != self.manifest["routes"] or receipt.get("families") != list(FAMILY_IDS):
            raise RunnerError("build handoff route/cpu identity differs")
        events = _journal_events(receipt.get("event_journal"))
        builds = receipt.get("builds")
        endpoints = receipt.get("endpoints")
        if not isinstance(builds, Mapping) or set(builds) != set(_build_keys()) or not isinstance(endpoints, Mapping) or set(endpoints) != set(ROLES):
            raise RunnerError("build handoff role/variant entries are incomplete")
        for role in ROLES:
            current = self._endpoint(role, full=True)
            if endpoints[role] != current:
                raise RunnerError(f"build source endpoint changed for {role}")
            for variant in VARIANTS:
                key = f"{role}/{variant}"
                build = builds[key]
                if not isinstance(build, Mapping) or build.get("role") != role or build.get("variant") != variant or build.get("source") != current:
                    raise RunnerError(f"build handoff source differs for {key}")
                target = self._target(role, variant)
                tmpdir = self._tmp(role, variant)
                binary = self._binary(role, variant)
                env = self._environment(role, variant, target, tmpdir)
                expected_command = self._build_command(role, variant)
                if build.get("command") != expected_command or build.get("environment") != self._recorded_environment(env) or _resolve(str(build.get("tmpdir"))) != tmpdir:
                    raise RunnerError(f"build handoff command/environment differs for {key}")
                current_toolchain = self._toolchain(self.baseline_root if role == "baseline" else self.candidate_root, env)
                if build.get("toolchain") != current_toolchain:
                    raise RunnerError(f"build handoff toolchain differs for {key}")
                binary_record = build.get("binary")
                if not isinstance(binary_record, Mapping) or _resolve(str(binary_record.get("path"))) != binary or binary.is_symlink() or not binary.is_file() or not os.access(binary, os.X_OK):
                    raise RunnerError(f"build handoff binary path differs for {key}")
                binary_sha, binary_bytes = _sha256_file(binary)
                if binary_record.get("sha256") != binary_sha or binary_record.get("bytes") != binary_bytes:
                    raise RunnerError(f"build handoff binary hash differs for {key}")
                target_record = build.get("target")
                if not isinstance(target_record, Mapping) or _resolve(str(target_record.get("path"))) != target or target_record.get("bytes") != generic._directory_size(target) or target_record.get("bytes", 0) > TARGET_PREFLIGHT_BYTES:
                    raise RunnerError(f"build handoff target identity differs for {key}")
                event = build.get("event")
                if not isinstance(event, Mapping) or events.get(f"build-{role}-{variant}") != event:
                    raise RunnerError(f"build event is not journal-bound for {key}")
                if event.get("schema") != SCHEMA or event.get("event_id") != f"build-{role}-{variant}" or event.get("role") != role or event.get("variant") != variant or event.get("argv") != expected_command or _resolve(str(event.get("cwd"))) != (self.baseline_root if role == "baseline" else self.candidate_root) or event.get("environment") != self._recorded_environment(env) or event.get("cpu_affinity") != self.cpu or event.get("returncode") != 0 or event.get("error") is not None:
                    raise RunnerError(f"build event identity differs for {key}")
                if event.get("toolchain") != current_toolchain:
                    raise RunnerError(f"build event toolchain differs for {key}")
                event_binary = event.get("binary")
                if not isinstance(event_binary, Mapping) or event_binary.get("sha256") != binary_sha or event_binary.get("bytes") != binary_bytes or _resolve(str(event_binary.get("path"))) != binary:
                    raise RunnerError(f"build event binary differs for {key}")
                target_record = event.get("target")
                if not isinstance(target_record, Mapping) or _resolve(str(target_record.get("path"))) != target or target_record.get("bytes_before") != 0 or target_record.get("bytes_after") != generic._directory_size(target):
                    raise RunnerError(f"build target freshness differs for {key}")
                if not isinstance(event.get("monitor"), list) or not event["monitor"] or event.get("process_guard_violations") or event.get("post_exit_descendants") or event.get("post_exit_offenders"):
                    raise RunnerError(f"build resource/quiet proof differs for {key}")
                for stream_name in ("stdout", "stderr"):
                    stream = event.get(stream_name)
                    expected = self.output / "build" / f"{role}-{variant}.{stream_name}"
                    if not isinstance(stream, Mapping) or _resolve(str(stream.get("path"))) != expected:
                        raise RunnerError(f"build {stream_name} path differs for {key}")
                    stream_sha, stream_bytes = _sha256_file(expected)
                    if stream.get("sha256") != stream_sha or stream.get("bytes") != stream_bytes:
                        raise RunnerError(f"build {stream_name} hash differs for {key}")
        receipt["receipt_sha256"] = receipt_sha
        return receipt

    def _validate_probe_output(self, path: Path, mode: str, iterations: int) -> dict[str, Any]:
        raw = path.read_bytes()
        if not raw or len(raw) > PROBE_MAX_BYTES:
            raise RunnerError(f"probe output is empty or too large: {path}")
        try:
            payload = json.loads(raw)
        except json.JSONDecodeError as error:
            raise RunnerError(f"probe output is not JSON: {error}") from error
        if not isinstance(payload, Mapping) or payload.get("schema") != "litchi.odf.raw-codec-qualification.v1" or payload.get("mode") != mode or payload.get("iterations") != iterations or NON_ACCEPTANCE not in str(payload.get("claim_scope")):
            raise RunnerError("probe output schema/mode/claim identity differs")
        if payload.get("families") != list(FAMILY_IDS):
            raise RunnerError("probe family order differs")
        negative = payload.get("negative_cases")
        if not isinstance(negative, Mapping) or any(negative.get(key) is not True for key in self.manifest["phases"]["negative_cases"]):
            raise RunnerError("probe negative-case proof is incomplete")
        routes = payload.get("routes")
        if not isinstance(routes, list) or [row.get("route") for row in routes if isinstance(row, Mapping)] != list(ROUTE_IDS):
            raise RunnerError("probe route order differs")
        expected_descriptors = {route["id"]: route for route in self.manifest["routes"]}
        for row in routes:
            if not isinstance(row, Mapping) or row.get("descriptor") != expected_descriptors.get(row.get("route")):
                raise RunnerError("probe route descriptor differs")
            details = row.get("details")
            proof = details.get("proof") if isinstance(details, Mapping) else None
            if not isinstance(proof, Mapping) or any(proof.get(key) is not True for key in ("operation_ok", "changed_ok", "no_op_ok", "reopen_ok", "source_bytes_ok", "facade_bytes_ok", "untouched_members_ok", "manifest_ok", "metadata_ok", "namespace_ok", "bounded_error_ok", "negative_cases_ok")):
                raise RunnerError(f"probe route proof is incomplete: {row.get('route')}")
            timing = row.get("timing")
            if not isinstance(timing, Mapping):
                raise RunnerError("probe route timing is missing")
            expected_allocator = mode == "allocator"
            if timing.get("allocator_enabled") is not expected_allocator or timing.get("elapsed_claimable") is not (mode == "timed"):
                raise RunnerError("probe timing/allocator claim differs")
            if mode == "validate" and timing.get("status") != "not_run":
                raise RunnerError("validate output contains a timing claim")
        return dict(payload)

    def _run_probe(self, builds: Mapping[str, Any], *, role: str, phase: str, kind: str, index: int, iterations: int) -> dict[str, Any]:
        variant = "normal" if phase in {"calibration", "normal"} else "allocator"
        mode = "validate" if phase == "calibration" else ("timed" if phase == "normal" else "allocator")
        build = builds[f"{role}/{variant}"]
        binary = self._binary(role, variant)
        binary_record = build["binary"]
        binary_sha, binary_bytes = _sha256_file(binary)
        if _resolve(str(binary_record.get("path"))) != binary or binary_record.get("sha256") != binary_sha or binary_record.get("bytes") != binary_bytes:
            raise RunnerError(f"{role}/{phase} binary handoff changed")
        target = self._target(role, variant)
        tmpdir = self._tmp(role, variant)
        root = self.baseline_root if role == "baseline" else self.candidate_root
        env = self._environment(role, variant, target, tmpdir)
        prefix = self.output / "probe" / phase / role / f"{kind}-{index:03d}"
        output = prefix.with_suffix(".json")
        stderr = prefix.with_suffix(".stderr")
        command = [str(binary), "--mode", mode, "--route", "all", "--iterations", str(iterations)]
        event = self._run_process(event_id=f"probe-{phase}-{role}-{kind}-{index:03d}", role=role, variant=variant, argv=command, cwd=root, env=env, target=target, stdout_path=output, stderr_path=stderr)
        payload = self._validate_probe_output(output, mode, iterations)
        self._assert_endpoint(role, build["source"])
        output_sha, output_bytes = _sha256_file(output)
        return {"role": role, "phase": phase, "kind": kind, "index": index, "mode": mode, "iterations": iterations, "command": command, "source": build["source"], "binary": {"path": str(binary), "sha256": binary_sha, "bytes": binary_bytes}, "output": {"path": str(output), "sha256": output_sha, "bytes": output_bytes}, "stderr": event["stderr"], "monitor": event["monitor"], "payload_identity": {"schema": payload["schema"], "mode": payload["mode"], "iterations": payload["iterations"], "routes": [row["route"] for row in payload["routes"]]}, "claim_scope": "paired correctness/protocol evidence only; no native acceptance claim"}

    def abba(self) -> None:
        build_receipt = self._load_build_receipt()
        builds = build_receipt["builds"]
        calibration = {role: self._run_probe(builds, role=role, phase="calibration", kind="validate", index=0, iterations=1) for role in ROLES}
        phases: dict[str, Any] = {}
        counters: dict[str, int] = {role: 0 for role in ROLES}
        for phase in PHASES:
            records: list[dict[str, Any]] = []
            for leg, role in enumerate(ABBA_ORDER, start=1):
                for kind, count in (("warmup", WARMUPS), ("sample", SAMPLES)):
                    for _ in range(count):
                        counters[role] += 1
                        records.append(self._run_probe(builds, role=role, phase=phase, kind=f"{kind}-leg{leg}", index=counters[role], iterations=PROBE_ITERATIONS))
            if len(records) != (WARMUPS + SAMPLES) * len(ABBA_ORDER):
                raise RunnerError(f"{phase} ABBA record count differs")
            for role in ROLES:
                self._assert_endpoint(role, builds[f"{role}/{phase}" if phase == "normal" else f"{role}/allocator"]["source"], full=True)
            phases[phase] = {"mode": "timed" if phase == "normal" else "allocator", "abba_order": list(ABBA_ORDER), "warmup_iterations_per_leg": WARMUPS, "samples_per_leg": SAMPLES, "records": records, "allocator_elapsed_claimable": phase == "normal"}
        receipt = {"schema": SCHEMA, "phase": "abba", "status": "passed", "manifest": {"path": str(self.manifest_path), "sha256": self.manifest_sha256}, "preparation": {"path": str(self.preparation_path), **self.preparation}, "build_receipt": {"path": str(self.output / "build-receipt.json"), "sha256": build_receipt["receipt_sha256"]}, "runner": {"path": str(Path(__file__).resolve()), "sha256": self.runner_sha256}, "cpu_affinity": self.cpu, "features": ["raw-codec-odf", "raw-codec-original", "raw-codec-current", "raw-codec-allocator"], "routes": self.manifest["routes"], "families": list(FAMILY_IDS), "calibration": calibration, "abba": phases, "claim_scope": "paired correctness/protocol evidence only; no native acceptance claim"}
        self._write_fresh(self.output / "abba-receipt.json", receipt, "ABBA receipt")


def _parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    for name in ("manifest", "preparation", "baseline-root", "candidate-root", "adapter-root", "candidate-commit", "candidate-lock-sha256", "output", "target-root", "tmp-root"):
        parser.add_argument(f"--{name}", required=True, type=Path if name.endswith(("manifest", "preparation", "root", "output")) else str)
    parser.add_argument("--cpu", required=True, type=int)
    parser.add_argument("--phase", required=True, choices=("build", "abba"))
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = _parse_args(argv or sys.argv[1:])
    try:
        runner = PairedRunner(manifest_path=args.manifest, preparation_path=args.preparation, baseline_root=args.baseline_root, candidate_root=args.candidate_root, adapter_root=args.adapter_root, candidate_commit=args.candidate_commit, candidate_lock_sha256=args.candidate_lock_sha256, output=args.output, target_root=args.target_root, tmp_root=args.tmp_root, cpu=args.cpu, phase=args.phase)
        if args.phase == "build":
            runner.build()
        else:
            runner.abba()
        return 0
    except (RunnerError, OSError, json.JSONDecodeError) as error:
        print(f"ODF raw-codec paired qualification refused: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
