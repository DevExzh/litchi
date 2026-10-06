#!/usr/bin/env python3
"""Freeze, capture, and replay the prospective 0925 provider cohort.

The default operation is ``freeze`` and writes only an external manifest. A
formal capture additionally requires an externally stored manifest marked by
the root coordinator and ``LITCHI_PERF_0925_CAPTURE_APPROVED=1``. The runner
uses the existing ``schema_version=1`` report, the
``LITCHI_DOCX_INGRESS_ROUTE`` selector, and never restarts a child after an
observation timeout or parser failure.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
import re
import signal
import shutil
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


HERE = Path(__file__).resolve().parent
try:
    _REPO_ROOT_CANDIDATE = HERE.parents[4]
except IndexError:
    _REPO_ROOT_CANDIDATE = Path("/")
# Copied packet scripts live under an external temporary directory.  Their
# relative parent chain is not the checkout, so do not mistake /tmp for the
# repository when replay validates a moved packet.
REPO_ROOT = (_REPO_ROOT_CANDIDATE
             if (_REPO_ROOT_CANDIDATE / ".git").exists()
             else Path("/__litchi_repo_unavailable__"))
BASE_REVISION = "aa077c030819c3019288e350f34959666ab7204f"
CASE = "docx_source_backed_existing_section_layout_edit_save"
MAX_ARTIFACT_BYTES = 1 << 20
OBSERVE_INTERVAL_SECONDS = 1.0
DEFAULT_PROCESS_DEADLINE_SECONDS = 900.0
DEFAULT_PACKET_ROOT = Path("/tmp/litchi-performance-0925/packet")


class CaptureError(RuntimeError):
    """A custody, capture, or replay contract failed closed."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise CaptureError(message)


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def canonical(value: Any) -> bytes:
    return json.dumps(value, sort_keys=True, separators=(",", ":")).encode()


def sha256_path(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def descriptor(path: Path, *, max_bytes: int | None = None) -> dict[str, Any]:
    path = path.resolve()
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    size = path.stat().st_size
    if max_bytes is not None:
        require(size < max_bytes, f"retained artifact is at least 1 MiB: {path}")
    return {"path": str(path), "bytes": size, "sha256": sha256_path(path)}


def external(path: Path) -> Path:
    path = path.resolve()
    try:
        path.relative_to(REPO_ROOT)
    except ValueError:
        return path
    raise CaptureError(f"custody/packet path must be external to checkout: {path}")


def load_json(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, ValueError) as error:
        raise CaptureError(f"cannot read JSON {path}: {error}") from error
    require(isinstance(value, dict), f"expected JSON object: {path}")
    return value


def write_new(path: Path, value: Any, *, max_bytes: int = MAX_ARTIFACT_BYTES) -> None:
    path = external(path)
    require(not path.exists(), f"refusing to replace retained evidence: {path}")
    data = value if isinstance(value, bytes) else (json.dumps(value, indent=2, sort_keys=True) + "\n").encode()
    require(len(data) < max_bytes, f"receipt is at least 1 MiB: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)


def safe_relative(value: str, label: str) -> Path:
    path = Path(value)
    require(not path.is_absolute() and ".." not in path.parts and str(path) not in ("", "."),
            f"{label}: unsafe relative path {value!r}")
    return path


def packet_descriptor(path: Path, packet_root: Path, label: str) -> dict[str, Any]:
    """Describe a packet file with a move-safe path relative to that packet."""
    path = path.resolve()
    packet_root = packet_root.resolve()
    try:
        relative = path.relative_to(packet_root)
    except ValueError as error:
        raise CaptureError(f"{label}: file is outside packet: {path}") from error
    return {
        "path": str(relative),
        "bytes": path.stat().st_size,
        "sha256": sha256_path(path),
    }


def verify_local_descriptor(packet_root: Path, value: Any, label: str) -> Path:
    require(isinstance(value, dict), f"{label}: descriptor missing")
    relative = safe_relative(str(value.get("path")), label)
    path = (packet_root / relative).resolve()
    try:
        path.relative_to(packet_root.resolve())
    except ValueError as error:
        raise CaptureError(f"{label}: descriptor escapes packet") from error
    require(path.is_file() and not path.is_symlink(), f"{label}: local file missing")
    require(path.stat().st_size == value.get("bytes")
            and sha256_path(path) == value.get("sha256"),
            f"{label}: local file hash or size changed")
    return path


def load_comparison(path: Path = HERE / "comparison.py") -> Any:
    spec = importlib.util.spec_from_file_location("litchi_perf_0925_comparison_capture", path)
    require(spec is not None and spec.loader is not None, f"cannot load validator: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def validate_source_tree(source_root: Path, source_receipt: Path, listing_path: Path) -> dict[str, Any]:
    source_root = external(source_root)
    require(source_root.is_dir(), f"source root is not a directory: {source_root}")
    receipt = load_json(source_receipt)
    revision = receipt.get("base_revision") or receipt.get("source_revision")
    require(revision == BASE_REVISION, f"source receipt base revision changed: {revision}")
    claimed = receipt.get("source_files") or receipt.get("candidate_source_files")
    require(isinstance(claimed, list) and claimed, "source receipt has no exact source_files list")
    claimed_paths = {str(safe_relative(str(value), "source_files")) for value in claimed}
    require(len(claimed_paths) == len(claimed), "source receipt has duplicate source paths")
    listing_path = external(listing_path)
    listing_rows: list[dict[str, Any]] = []
    with listing_path.open(encoding="utf-8") as stream:
        for line_number, line in enumerate(stream, start=1):
            if not line.strip():
                continue
            try:
                row = json.loads(line)
            except ValueError as error:
                raise CaptureError(f"source listing line {line_number} is invalid JSON") from error
            require(isinstance(row, dict), f"source listing line {line_number} is not an object")
            relative = str(safe_relative(str(row.get("path")), "source listing"))
            size = row.get("bytes")
            digest = row.get("sha256")
            require(type(size) is int and size >= 0 and isinstance(digest, str)
                    and len(digest) == 64, f"source listing row malformed: {relative}")
            listing_rows.append({"path": relative, "bytes": size, "sha256": digest})
    listed_paths = {row["path"] for row in listing_rows}
    require(len(listing_rows) == len(listed_paths), "source listing contains duplicate paths")
    require(listed_paths == claimed_paths,
            f"source closure differs from receipt: listing={len(listed_paths)} claimed={len(claimed_paths)}")
    for row in listing_rows:
        path = source_root / row["path"]
        require(path.is_file() and not path.is_symlink(), f"source file missing: {path}")
        require(path.stat().st_size == row["bytes"] and sha256_path(path) == row["sha256"],
                f"source file hash changed: {row['path']}")
    actual_paths = {
        str(path.relative_to(source_root))
        for path in source_root.rglob("*")
        if path.is_file() and not path.is_symlink()
        and "target" not in path.parts and ".git" not in path.parts
    }
    require(actual_paths == listed_paths,
            f"source tree has unbound files: {len(actual_paths - listed_paths)} extras, "
            f"{len(listed_paths - actual_paths)} missing")
    candidate_file = receipt.get("candidate_file")
    candidate_file_sha256 = receipt.get("candidate_file_sha256")
    candidate_patch = receipt.get("candidate_patch")
    candidate_patch_sha256 = receipt.get("candidate_patch_sha256")
    require(isinstance(candidate_file, str) and candidate_file in listed_paths,
            "source receipt candidate_file is not in the exact source closure")
    require(isinstance(candidate_file_sha256, str) and len(candidate_file_sha256) == 64,
            "source receipt candidate_file_sha256 is missing")
    candidate_row = next(row for row in listing_rows if row["path"] == candidate_file)
    require(candidate_row["sha256"] == candidate_file_sha256,
            "candidate source file hash differs from the source listing")
    candidate_path = source_root / candidate_file
    require(sha256_path(candidate_path) == candidate_file_sha256,
            "candidate source file hash changed")
    require(isinstance(candidate_patch, str) and candidate_patch,
            "source receipt candidate_patch is missing")
    candidate_patch_path = external(Path(candidate_patch))
    require(candidate_patch_path.is_file() and not candidate_patch_path.is_symlink(),
            f"candidate patch is missing: {candidate_patch_path}")
    candidate_patch_descriptor = descriptor(candidate_patch_path, max_bytes=MAX_ARTIFACT_BYTES)
    require(isinstance(candidate_patch_sha256, str) and len(candidate_patch_sha256) == 64
            and candidate_patch_descriptor["sha256"] == candidate_patch_sha256,
            "candidate patch hash changed")
    return {
        "revision": revision,
        "source_receipt": descriptor(source_receipt, max_bytes=MAX_ARTIFACT_BYTES),
        "source_root": str(source_root),
        "source_files": len(listing_rows),
        "source_bytes": sum(row["bytes"] for row in listing_rows),
        "full_source_listing": descriptor(listing_path),
        "candidate_file": candidate_file,
        "candidate_file_sha256": candidate_file_sha256,
        "candidate_patch": candidate_patch_descriptor,
        "candidate_patch_sha256": candidate_patch_sha256,
    }


def _binary_from_receipt(receipt: dict[str, Any], label: str) -> dict[str, Any]:
    value = receipt.get("binary") or receipt.get("binary_identity")
    require(isinstance(value, dict), f"{label}: binary descriptor missing")
    if "binary_sha256" in value:
        normalized = {
            "path": value.get("path"), "bytes": value.get("binary_bytes"),
            "sha256": value.get("binary_sha256"),
        }
    else:
        normalized = {"path": value.get("path"), "bytes": value.get("bytes"),
                      "sha256": value.get("sha256")}
    require(isinstance(normalized["path"], str) and type(normalized["bytes"]) is int
            and normalized["bytes"] > 0 and isinstance(normalized["sha256"], str)
            and len(normalized["sha256"]) == 64, f"{label}: binary descriptor malformed")
    binary_path = Path(normalized["path"])
    require(binary_path.is_file() and not binary_path.is_symlink(),
            f"{label}: binary path missing: {binary_path}")
    actual = descriptor(binary_path)
    require(actual["bytes"] == normalized["bytes"] and actual["sha256"] == normalized["sha256"],
            f"{label}: binary bytes or hash changed")
    return normalized


def validate_build_receipt(build_receipt: Path) -> dict[str, Any]:
    receipt = load_json(build_receipt)
    revision = receipt.get("base_revision") or receipt.get("source_revision")
    require(revision == BASE_REVISION, f"build receipt base revision changed: {revision}")
    profile = receipt.get("profile") or receipt.get("effective_profile")
    require(isinstance(profile, str) and profile == "release", "build profile is not release")
    command = receipt.get("command") or receipt.get("build_command")
    require(isinstance(command, str) and "--locked" in command and "--release" in command
            and "--offline" in command and "-j2" in command,
            "build receipt does not bind the required standalone offline command")
    toolchain = receipt.get("toolchain") or receipt.get("versions")
    require(isinstance(toolchain, dict), "build receipt toolchain record missing")
    require(any("1.95.0" in str(value) for value in toolchain.values()),
            "build receipt does not bind rust/cargo 1.95.0")
    locks = receipt.get("locks") or receipt.get("lock_hashes")
    require(isinstance(locks, dict) and locks, "build receipt lock identities missing")
    source_custody = receipt.get("source_custody")
    require(isinstance(source_custody, dict), "build receipt source custody missing")
    for name in ("source_receipt", "full_source_listing"):
        value = source_custody.get(name)
        require(isinstance(value, dict), f"build receipt source custody field missing: {name}")
        path = Path(str(value.get("path")))
        require(path.is_file() and not path.is_symlink()
                and path.stat().st_size == value.get("bytes")
                and sha256_path(path) == value.get("sha256"),
                f"build receipt source custody changed: {name}")
    binary = _binary_from_receipt(receipt, "build receipt")
    return {
        "receipt": descriptor(build_receipt, max_bytes=MAX_ARTIFACT_BYTES),
        "base_revision": revision,
        "profile": profile,
        "command": command,
        "toolchain": toolchain,
        "locks": locks,
        "source_custody": source_custody,
        "binary": binary,
    }


def _script_descriptors() -> dict[str, Any]:
    paths = {
        "comparison": HERE / "comparison.py",
        "controls": HERE / "comparison-controls.py",
        "capture": HERE / "comparison-capture.py",
        "reduce": HERE / "comparison-reduce.py",
    }
    return {name: descriptor(path, max_bytes=MAX_ARTIFACT_BYTES) for name, path in paths.items()}


def validate_protocol_document(path: Path, comparison: Any) -> dict[str, Any]:
    """Bind the human-readable protocol to the executable schedule and report shape."""
    document = load_json(path)
    require(document.get("schema") == "litchi.performance.0925.docx-provider-protocol.v1",
            "protocol schema changed")
    base_source = document.get("base_source")
    require(isinstance(base_source, dict)
            and base_source.get("revision") == BASE_REVISION,
            "protocol base revision changed")
    require(base_source.get("harness_changed_files_allowlist")
            == ["tools/perf-baseline/src/lib.rs"],
            "protocol harness changed-file allowlist is not final")
    require(base_source.get("harness_patch_sha256")
            == "9675346c827a4d546a47257f9c7cbcbfb37e76dae8055b2c9a38667e7d82fc2b",
            "protocol harness patch binding changed")
    require(base_source.get("harness_changed_file_sha256")
            == {"tools/perf-baseline/src/lib.rs":
                "7904a6161e7b40a675d61acd9867b31b81a47f049b1ebb5a9979778bf14a373a"},
            "protocol harness source hash binding changed")
    build_custody = document.get("build_custody")
    require(isinstance(build_custody, dict), "protocol build custody is missing")
    require(build_custody.get("profile_effective_receipt") == {
        "path": "/tmp/litchi-performance-0925/measurement-build-v4-final.json",
        "bytes": 3137,
        "sha256": "2b923bb232a2a433d03e97ee26991b48e4caf548924846ba15060fc7b639bf99",
    }, "protocol effective build receipt binding changed")
    require(build_custody.get("actual_binary") == {
        "path": "/tmp/litchi-performance-0925/measurement-target-v4-final/release/litchi-perf-baseline",
        "bytes": 70187352,
        "sha256": "7138c6c458dc13f1677deba3f413120ffc1af155d07848a243a43ebcf335feb7",
    }, "protocol binary binding changed")
    require(build_custody.get("actual_rustc_info") == {
        "path": "/tmp/litchi-performance-0925/measurement-target-v4-final/.rustc_info.json",
        "bytes": 1040,
        "sha256": "8fcff6af6f9610c25626b31588689eca31edd31111abb8968b9b772c70be4dcb",
    }, "protocol rustc identity binding changed")
    require(build_custody.get("actual_source_custody") == {
        "source_receipt_sha256": "fe35ae7a2926304915be3c48e3554143313e1acf902f31233a44fa668e0d44b2",
        "full_source_listing_sha256": "95fa57adadba7f24a483090adcaee609aeb3ac72e585f2479477442f81817e68",
    }, "protocol source custody binding changed")
    require(build_custody.get("actual_toolchain") == {
        "rustc": "rustc 1.95.0 (59807616e 2026-04-14)",
        "cargo": "cargo 1.95.0 (f2d3ce0bd 2026-03-21)",
        "environment_overrides": {},
    }, "protocol build toolchain binding changed")
    require(document.get("release_preflight") == {
        "schema": "litchi.performance.0925.comparison-preflight.v1",
        "receipt": {
            "path": "comparison-preflight-final.json",
            "bytes": 21156,
            "sha256": "16c67fa43a19f00db2d6d3f98f65d0c47835755ae7a74575f12329f137753802",
        },
        "routes": list(comparison.ROUTES),
        "children": 7,
        "samples": 3,
        "warmups": 1,
        "measured_samples": 21,
        "processes_started": 7,
        "binary_sha256": "7138c6c458dc13f1677deba3f413120ffc1af155d07848a243a43ebcf335feb7",
        "build_receipt_sha256": "2b923bb232a2a433d03e97ee26991b48e4caf548924846ba15060fc7b639bf99",
        "candidate_file_sha256": "7904a6161e7b40a675d61acd9867b31b81a47f049b1ebb5a9979778bf14a373a",
        "status": "passed; release schema/custody preflight only, no comparative timing claim",
    }, "protocol release preflight binding changed")
    require(document.get("claim_scope", {}).get("performance_claim")
            == "descriptive_provider_baseline_only",
            "protocol claim scope changed")
    require(document.get("filesystem_staging") == {
        "root": "/tmp",
        "filesystem_type": "tmpfs",
        "scope": "FileSource is staged and warmed on tmpfs; this does not claim disk or physical filesystem I/O",
    }, "protocol filesystem staging scope changed")
    require(document.get("analysis_tooling", {}).get("reducer_script")
            == "comparison-reduce.py",
            "protocol reducer binding changed")
    routes = document.get("routes")
    require(isinstance(routes, list)
            and [route.get("name") for route in routes] == list(comparison.ROUTES),
            "protocol route order changed")
    fixed = document.get("fixed_cohort")
    require(isinstance(fixed, dict), "protocol fixed cohort missing")
    require(fixed.get("children") == 42 and fixed.get("route_count") == 7
            and fixed.get("samples_per_child") == 50 and fixed.get("measured_samples_per_route") == 300
            and fixed.get("process_deadline_seconds") == 900,
            "protocol fixed cohort changed")
    require(fixed.get("schedule") == comparison.schedule_rows(),
            "protocol fixed schedule changed")
    argv = fixed.get("argv_template")
    require(argv == ["<binary>", "--case", CASE, "--samples", "50", "--warmup", "5", "--json", "<report.json>"],
            "protocol argv template changed")
    schedule = document.get("schedule_validation")
    require(isinstance(schedule, dict) and schedule.get("children") == 42
            and schedule.get("measured_samples") == 2100
            and schedule.get("routes") == list(comparison.ROUTES),
            "protocol schedule validation changed")
    require(schedule.get("route_counts") == {route: 6 for route in comparison.ROUTES},
            "protocol route counts changed")
    corpus = document.get("corpus")
    require(isinstance(corpus, dict)
            and corpus.get("archive_bytes") == comparison.SOURCE_BYTES
            and corpus.get("archive_sha256") == comparison.SOURCE_SHA256
            and corpus.get("expected_output_bytes") == comparison.OUTPUT_BYTES
            and corpus.get("expected_output_sha256") == comparison.OUTPUT_SHA256,
            "protocol corpus identity changed")
    timing = document.get("timing_boundaries")
    require(isinstance(timing, dict) and timing.get("phase_names") == list(comparison.PHASES),
            "protocol timing phase names changed")
    report_contract = document.get("report_contract")
    require(isinstance(report_contract, dict)
            and report_contract.get("schema")
            == "existing litchi-perf-baseline report schema_version=1; provider fields are nested in results[0].source.docx_section_layout",
            "protocol report schema changed")
    require(document.get("source_mutation_by_this_tool") is False,
            "protocol source mutation claim changed")
    return {
        "schema": document["schema"],
        "routes": list(comparison.ROUTES),
        "children": 42,
        "process_deadline_seconds": 900,
    }


def _copy_packet_custody(manifest: dict[str, Any], packet: Path) -> dict[str, Any]:
    """Copy small custody receipts so replay can move without external paths."""
    custody_sources = {
        "source_receipt": Path(str(manifest["source"]["source_receipt"]["path"])),
        "build_receipt": Path(str(manifest["build"]["receipt"]["path"])),
        "protocol": Path(str(manifest["protocol"]["path"])),
    }
    custody_root = packet / "custody"
    custody_root.mkdir()
    copied: dict[str, Any] = {}
    for name, source in custody_sources.items():
        source_descriptor = (
            manifest["source"]["source_receipt"] if name == "source_receipt"
            else manifest["build"]["receipt"] if name == "build_receipt"
            else manifest["protocol"]
        )
        require(source.is_file() and not source.is_symlink(),
                f"packet custody source missing: {name}")
        require(source.stat().st_size == source_descriptor.get("bytes")
                and sha256_path(source) == source_descriptor.get("sha256"),
                f"packet custody source changed: {name}")
        require(source.stat().st_size < MAX_ARTIFACT_BYTES,
                f"packet custody receipt is at least 1 MiB: {name}")
        target = custody_root / f"{name}.json"
        shutil.copy2(source, target)
        copied[name] = packet_descriptor(target, packet, f"packet custody {name}")
        require(copied[name]["bytes"] == source_descriptor.get("bytes")
                and copied[name]["sha256"] == source_descriptor.get("sha256"),
                f"packet custody copy changed: {name}")
    return copied


def _verify_packet_custody(manifest: dict[str, Any], packet: Path) -> None:
    local = manifest.get("packet_custody")
    require(isinstance(local, dict), "copied packet custody descriptors are missing")
    expected = {
        "source_receipt": manifest.get("source", {}).get("source_receipt"),
        "build_receipt": manifest.get("build", {}).get("receipt"),
        "protocol": manifest.get("protocol"),
    }
    require(set(local) == set(expected), "copied packet custody keys changed")
    for name, original in expected.items():
        require(isinstance(original, dict), f"manifest custody descriptor missing: {name}")
        value = local[name]
        path = verify_local_descriptor(packet, value, f"packet custody {name}")
        require(value.get("bytes") == original.get("bytes")
                and value.get("sha256") == original.get("sha256"),
                f"packet custody descriptor changed: {name}")
        if name in ("source_receipt", "build_receipt", "protocol"):
            loaded = load_json(path)
            require(isinstance(loaded, dict), f"packet custody {name} is not a JSON object")


def _receipt_binary_descriptor(receipt: dict[str, Any]) -> dict[str, Any]:
    value = receipt.get("binary") or receipt.get("binary_identity")
    require(isinstance(value, dict), "packet build receipt binary descriptor missing")
    if "binary_sha256" in value:
        return {
            "path": value.get("path"), "bytes": value.get("binary_bytes"),
            "sha256": value.get("binary_sha256"),
        }
    return {"path": value.get("path"), "bytes": value.get("bytes"),
            "sha256": value.get("sha256")}


def _validate_local_custody_semantics(manifest: dict[str, Any], packet: Path,
                                      comparison: Any) -> None:
    """Check copied receipt meaning without depending on moved absolute paths."""
    local = manifest["packet_custody"]
    source_receipt = load_json(verify_local_descriptor(
        packet, local["source_receipt"], "packet source receipt"))
    source_revision = source_receipt.get("base_revision") or source_receipt.get("source_revision")
    require(source_revision == BASE_REVISION, "packet source receipt revision changed")
    claimed = source_receipt.get("source_files") or source_receipt.get("candidate_source_files")
    require(isinstance(claimed, list) and claimed,
            "packet source receipt lacks exact source file closure")
    expected_source = manifest.get("source", {})
    if "candidate_file" in expected_source:
        require(source_receipt.get("candidate_file") == expected_source.get("candidate_file"),
                "packet source candidate file binding changed")
        require(source_receipt.get("candidate_file_sha256")
                == expected_source.get("candidate_file_sha256"),
                "packet source candidate file hash binding changed")
        require(source_receipt.get("candidate_patch_sha256")
                == expected_source.get("candidate_patch_sha256"),
                "packet source candidate patch hash binding changed")

    build_receipt = load_json(verify_local_descriptor(
        packet, local["build_receipt"], "packet build receipt"))
    build_revision = build_receipt.get("base_revision") or build_receipt.get("source_revision")
    require(build_revision == BASE_REVISION, "packet build receipt revision changed")
    require(build_receipt.get("profile") == "release", "packet build profile changed")
    command = build_receipt.get("command") or build_receipt.get("build_command")
    require(isinstance(command, str) and all(flag in command for flag in (
        "--offline", "-j2", "--locked", "--release"
    )), "packet build command changed")
    toolchain = build_receipt.get("toolchain") or build_receipt.get("versions")
    require(isinstance(toolchain, dict)
            and any("1.95.0" in str(value) for value in toolchain.values()),
            "packet build toolchain changed")
    require(build_receipt.get("locks") or build_receipt.get("lock_hashes"),
            "packet build lock identities missing")
    source_custody = build_receipt.get("source_custody")
    require(isinstance(source_custody, dict), "packet build source custody missing")
    for name in ("source_receipt", "full_source_listing"):
        value = source_custody.get(name)
        expected = manifest["source"][name]
        require(isinstance(value, dict)
                and value.get("bytes") == expected.get("bytes")
                and value.get("sha256") == expected.get("sha256"),
                f"packet build source custody differs: {name}")
    require(_receipt_binary_descriptor(build_receipt) == manifest["binary"],
            "packet build binary descriptor differs from manifest")

    protocol_path = verify_local_descriptor(packet, local["protocol"], "packet protocol")
    validate_protocol_document(protocol_path, comparison)


def freeze(args: argparse.Namespace) -> dict[str, Any]:
    source_receipt = args.source_receipt.resolve()
    build_receipt = args.build_receipt.resolve()
    source = validate_source_tree(args.source_root, source_receipt, args.full_source_listing)
    build = validate_build_receipt(build_receipt)
    for name in ("source_receipt", "full_source_listing"):
        require(build["source_custody"][name] == source[name],
                f"build/source custody binding differs: {name}")
    protocol = descriptor(args.protocol.resolve(), max_bytes=MAX_ARTIFACT_BYTES)
    scripts = _script_descriptors()
    comparison = load_comparison()
    validate_protocol_document(args.protocol.resolve(), comparison)
    schedule = comparison.schedule_rows()
    schedule_hash = hashlib.sha256(canonical(schedule)).hexdigest()
    require(build["binary"]["sha256"], "binary hash is empty")
    manifest = {
        "schema": "litchi.performance.0925.provider-capture-manifest.v1",
        "created_utc": utc_now(),
        "capture_authorized": False,
        "formal_capture_children_run": 0,
        "base_revision": BASE_REVISION,
        "source": source,
        "build": build,
        "protocol": protocol,
        "scripts": scripts,
        "binary": build["binary"],
        "schedule": {
            "routes": list(comparison.ROUTES), "blocks": 6, "children": 42,
            "children_per_route": 6, "warmups": 5, "samples_per_child": 50,
            "cpu": comparison.CPU, "sha256": schedule_hash,
            "rows": schedule,
        },
        "process_deadline_seconds": args.process_deadline,
        "selector": {"name": "LITCHI_DOCX_INGRESS_ROUTE", "values": list(comparison.ROUTES)},
        "claim_scope": "descriptive provider baseline only; no ranking, optimization, cold filesystem, physical I/O, allocator, or RSS claim",
        "failure_policy": "retain every child; any terminal failure invalidates the cohort; never rerun or selectively recapture",
    }
    write_new(args.manifest, manifest)
    return {"manifest": str(args.manifest.resolve()), "bytes": args.manifest.stat().st_size,
            "sha256": sha256_path(args.manifest), "capture_authorized": False,
            "children": 42, "performance_children_run": 0}


def _verify_manifest(manifest_path: Path, *, require_authorized: bool,
                     packet_root: Path | None = None,
                     revalidate_external: bool = False) -> tuple[dict[str, Any], Any]:
    manifest_path = external(manifest_path)
    manifest = load_json(manifest_path)
    require(manifest.get("schema") == "litchi.performance.0925.provider-capture-manifest.v1",
            "manifest schema changed")
    require(manifest.get("base_revision") == BASE_REVISION, "manifest base revision changed")
    if require_authorized:
        require(manifest.get("capture_authorized") is True, "root has not authorized capture")
        require(os.environ.get("LITCHI_PERF_0925_CAPTURE_APPROVED") == "1",
                "LITCHI_PERF_0925_CAPTURE_APPROVED=1 is required")
    comparison = load_comparison()
    schedule = comparison.schedule_rows()
    recorded = manifest.get("schedule")
    require(isinstance(recorded, dict) and recorded.get("rows") == schedule,
            "manifest schedule differs from the fixed six-block schedule")
    require(recorded.get("sha256") == hashlib.sha256(canonical(schedule)).hexdigest(),
            "manifest schedule hash changed")
    deadline = manifest.get("process_deadline_seconds")
    require(isinstance(deadline, (int, float)) and not isinstance(deadline, bool)
            and deadline > 0, "manifest process deadline is missing or invalid")
    require(manifest.get("selector") == {"name": "LITCHI_DOCX_INGRESS_ROUTE",
                                          "values": list(comparison.ROUTES)},
            "manifest route selector changed")
    original_script_values = manifest.get("scripts") or {}
    script_values = original_script_values
    if packet_root is not None:
        local_values = manifest.get("packet_scripts")
        require(isinstance(local_values, dict) and set(local_values) == set(script_values),
                "copied packet script descriptors are missing")
        for name, value in local_values.items():
            original = original_script_values[name]
            require(value.get("bytes") == original.get("bytes")
                    and value.get("sha256") == original.get("sha256"),
                    f"copied packet script descriptor differs: {name}")
        script_values = local_values
    for name, value in script_values.items():
        require(isinstance(value, dict), f"manifest script descriptor malformed: {name}")
        path = (packet_root / safe_relative(str(value.get("path")), f"script {name}")
                if packet_root is not None else Path(str(value.get("path"))))
        require(path.is_file() and sha256_path(path) == value.get("sha256")
                and path.stat().st_size == value.get("bytes"),
                f"manifest script changed: {name}")
    if packet_root is not None:
        _verify_packet_custody(manifest, packet_root)
        _validate_local_custody_semantics(manifest, packet_root, comparison)
    if revalidate_external:
        source = manifest.get("source")
        build = manifest.get("build")
        protocol = manifest.get("protocol")
        require(isinstance(source, dict) and isinstance(build, dict)
                and isinstance(protocol, dict), "manifest custody descriptors are missing")
        source_now = validate_source_tree(
            Path(str(source["source_root"])),
            Path(str(source["source_receipt"]["path"])),
            Path(str(source["full_source_listing"]["path"])),
        )
        require(source_now["source_files"] == source.get("source_files")
                and source_now["source_bytes"] == source.get("source_bytes")
                and source_now["full_source_listing"]["sha256"] == source["full_source_listing"].get("sha256"),
                "source closure changed since freeze")
        for name in ("candidate_file", "candidate_file_sha256", "candidate_patch_sha256"):
            require(source_now.get(name) == source.get(name),
                    f"source candidate binding changed: {name}")
        require(source_now.get("candidate_patch") == source.get("candidate_patch"),
                "source candidate patch descriptor changed")
        build_now = validate_build_receipt(Path(str(build["receipt"]["path"])))
        require(build_now["binary"] == manifest.get("binary"),
                "binary custody changed since freeze")
        protocol_path = Path(str(protocol.get("path")))
        require(protocol_path.is_file() and protocol_path.stat().st_size == protocol.get("bytes")
                and sha256_path(protocol_path) == protocol.get("sha256"),
                "protocol changed since freeze")
        validate_protocol_document(protocol_path, comparison)
    return manifest, comparison


def _rss_bytes(time_path: Path) -> int | None:
    text = time_path.read_text(encoding="utf-8", errors="replace")
    match = re.search(r"Maximum resident set size \(kbytes\):\s*(\d+)", text)
    return int(match.group(1)) * 1024 if match else None


def _child_command(row: dict[str, Any], manifest: dict[str, Any], report_path: Path) -> list[str]:
    binary_path = Path(str(manifest["binary"]["path"]))
    return ["/usr/bin/time", "-v", "taskset", "-c", str(row["cpu"]), str(binary_path),
            "--case", CASE, "--samples", str(row["samples"]), "--warmup",
            str(row["warmups"]), "--json", str(report_path)]


def _compact_validation_receipt(summary: dict[str, Any], report: dict[str, Any],
                                comparison: Any) -> dict[str, Any]:
    """Keep only scalar validation custody in the terminal child receipt.

    ``comparison.validate_report`` returns the phase/provider vectors needed by
    the offline reducer.  Those vectors remain in the raw report identified by
    this descriptor; duplicating them in ``capture-state.json`` can push the
    aggregate receipt over the one-MiB custody limit.
    """
    require(isinstance(summary, dict), "validated report summary is not an object")
    require(isinstance(report, dict), "validated report descriptor is not an object")
    route = summary.get("route")
    samples = summary.get("samples")
    warmups = summary.get("warmups")
    require(isinstance(route, str) and route in comparison.ROUTES,
            "validated report summary route is missing")
    require(type(samples) is int and samples > 0,
            "validated report summary sample count is missing")
    require(type(warmups) is int and warmups >= 0,
            "validated report summary warmup count is missing")
    report_sha256 = report.get("sha256")
    report_bytes = report.get("bytes")
    require(isinstance(report_sha256, str) and len(report_sha256) == 64,
            "validated report descriptor hash is missing")
    require(type(report_bytes) is int and report_bytes >= 0,
            "validated report descriptor size is missing")
    return {
        "schema": "litchi.performance.0925.compact-validation.v1",
        "status": "passed",
        "route": route,
        "samples": samples,
        "warmups": warmups,
        "report_bytes": report_bytes,
        "report_sha256": report_sha256,
    }


def _run_child(row: dict[str, Any], manifest: dict[str, Any], comparison: Any,
               packet: Path, process_deadline: float) -> dict[str, Any]:
    child = packet / "children" / str(row["id"])
    child.mkdir(parents=True, exist_ok=False)
    report_path = child / "report.json"
    stdout_path = child / "stdout.log"
    stderr_path = child / "stderr.log"
    time_path = child / "time.log"
    command = _child_command(row, manifest, report_path)
    env = os.environ.copy()
    env["LITCHI_DOCX_INGRESS_ROUTE"] = str(row["route"])
    source_root = Path(str(manifest["source"]["source_root"]))
    require(source_root.is_dir(), f"measurement source root is missing: {source_root}")
    started = utc_now()
    monotonic_start = time.monotonic()
    with stdout_path.open("wb") as stdout, stderr_path.open("wb") as stderr:
        process = subprocess.Popen(command, stdout=stdout, stderr=stderr, env=env,
                                   cwd=str(source_root), start_new_session=True)
        observation_timeouts = 0
        deadline_exceeded = False
        while True:
            try:
                exit_code = process.wait(timeout=OBSERVE_INTERVAL_SECONDS)
                break
            except subprocess.TimeoutExpired:
                observation_timeouts += 1
                if time.monotonic() - monotonic_start > process_deadline:
                    deadline_exceeded = True
                    try:
                        os.killpg(process.pid, signal.SIGTERM)
                    except ProcessLookupError:
                        pass
                    try:
                        exit_code = process.wait(timeout=10)
                    except subprocess.TimeoutExpired:
                        try:
                            os.killpg(process.pid, signal.SIGKILL)
                        except ProcessLookupError:
                            pass
                        exit_code = process.wait(timeout=10)
                    break
    ended = utc_now()
    # /usr/bin/time -v writes its diagnostics to stderr. Retain the complete
    # stream; this summary is deliberately not reconstructed from a partial
    # line when RSS is absent.
    time_path.write_bytes(stderr_path.read_bytes())
    row_receipt: dict[str, Any] = {
        "schema": "litchi.performance.0925.provider-child.v1",
        "schedule": row,
        "argv": command,
        "environment": {"LITCHI_DOCX_INGRESS_ROUTE": row["route"]},
        "cwd": str(source_root),
        "started_utc": started,
        "ended_utc": ended,
        "exit_code": exit_code,
        "observation_timeout_count": observation_timeouts,
        "deadline_exceeded": deadline_exceeded,
        "stdout": descriptor(stdout_path),
        "stderr_and_time": descriptor(stderr_path),
        "time": descriptor(time_path),
        "rss_bytes": _rss_bytes(time_path),
        "report": None,
    }
    valid = exit_code == 0 and not deadline_exceeded and row_receipt["rss_bytes"] is not None
    if report_path.exists():
        # Keep an oversized or malformed report for audit, but make the
        # complete cohort invalid.  The raw packet is external custody; the
        # one-MiB rule applies to retained repository evidence.
        row_receipt["report"] = packet_descriptor(report_path, packet, "child report")
        valid = valid and report_path.stat().st_size < MAX_ARTIFACT_BYTES
        try:
            report = json.loads(report_path.read_text(encoding="utf-8"))
            summary = comparison.validate_report(report, row["route"],
                                                 warmups=row["warmups"], samples=row["samples"],
                                                 label=str(report_path), binary=manifest["binary"])
            row_receipt["validation"] = _compact_validation_receipt(
                summary, row_receipt["report"], comparison
            )
            valid = valid and True
        except Exception as error:
            row_receipt["validation_error"] = f"{type(error).__name__}: {error}"
            valid = False
    else:
        row_receipt["validation_error"] = "report file is missing"
        valid = False
    row_receipt["status"] = "passed" if valid else "failed"
    for field in ("stdout", "stderr_and_time", "time"):
        row_receipt[field] = packet_descriptor(
            Path(str(row_receipt[field]["path"])), packet, f"child {field}"
        )
    return row_receipt


def _failed_child_receipt(row: dict[str, Any], manifest: dict[str, Any], packet: Path,
                          error: BaseException) -> dict[str, Any]:
    """Retain an unexpected runner failure as a terminal invalid child."""
    child = packet / "children" / str(row["id"])
    child.mkdir(parents=True, exist_ok=True)
    error_path = child / "runner-error.log"
    error_path.write_text(f"{type(error).__name__}: {error}\n", encoding="utf-8")
    source_root = Path(str(manifest["source"]["source_root"]))
    report_path = child / "report.json"
    return {
        "schema": "litchi.performance.0925.provider-child.v1",
        "schedule": row,
        "argv": _child_command(row, manifest, report_path),
        "environment": {"LITCHI_DOCX_INGRESS_ROUTE": row["route"]},
        "cwd": str(source_root),
        "started_utc": utc_now(),
        "ended_utc": utc_now(),
        "exit_code": None,
        "deadline_exceeded": False,
        "status": "failed",
        "runner_error": descriptor(error_path),
        "validation_error": f"{type(error).__name__}: {error}",
    }


def capture(args: argparse.Namespace) -> dict[str, Any]:
    manifest, comparison = _verify_manifest(args.manifest, require_authorized=True,
                                             revalidate_external=True)
    require(args.process_deadline == manifest["process_deadline_seconds"],
            "capture process deadline differs from frozen manifest")
    packet = external(args.packet)
    require(not packet.exists(), f"refusing to replace existing packet: {packet}")
    packet.mkdir(parents=True)
    (packet / "children").mkdir()
    # Copy only small scripts/receipts into the packet. Full source listings,
    # binaries, and build logs remain external custody artifacts.
    packet_scripts = packet / "scripts"
    packet_scripts.mkdir()
    for name, value in manifest["scripts"].items():
        source = Path(str(value["path"]))
        target = packet_scripts / source.name
        shutil.copy2(source, target)
        require(sha256_path(target) == value["sha256"], f"packet script copy changed: {name}")
    packet_manifest = dict(manifest)
    packet_manifest["packet_custody"] = _copy_packet_custody(manifest, packet)
    packet_manifest["packet_scripts"] = {
        name: {"path": f"scripts/{Path(str(value['path'])).name}",
                "bytes": (packet_scripts / Path(str(value["path"])).name).stat().st_size,
                "sha256": sha256_path(packet_scripts / Path(str(value["path"])).name)}
        for name, value in manifest["scripts"].items()
    }
    (packet / "manifest.json").write_text(json.dumps(packet_manifest, indent=2, sort_keys=True) + "\n",
                                           encoding="utf-8")
    rows = manifest["schedule"]["rows"]
    receipts = []
    runner_error: str | None = None
    for row in rows:
        try:
            receipts.append(_run_child(row, manifest, comparison, packet, args.process_deadline))
        except Exception as error:
            receipts.append(_failed_child_receipt(row, manifest, packet, error))
            runner_error = f"{type(error).__name__}: {error}"
            break
    passed = len(receipts) == len(rows) and all(row.get("status") == "passed" for row in receipts)
    try:
        _verify_manifest(args.manifest, require_authorized=True, revalidate_external=True)
        custody_after = {"status": "verified"}
    except Exception as error:
        custody_after = {"status": "failed", "error": f"{type(error).__name__}: {error}"}
        passed = False
    state = {
        "schema": "litchi.performance.0925.provider-capture-state.v1",
        "started_utc": receipts[0]["started_utc"] if receipts else utc_now(),
        "ended_utc": receipts[-1]["ended_utc"] if receipts else utc_now(),
        "children": receipts,
        "children_run": len(receipts),
        "process_deadline_seconds": args.process_deadline,
        "status": "completed" if passed else "invalid",
        "custody_revalidation_after": custody_after,
        "failure_policy": "invalid cohorts are retained and cannot be recaptured",
    }
    if runner_error is not None:
        state["runner_error"] = runner_error
    write_new(packet / "capture-state.json", state)
    return {"status": state["status"], "packet": str(packet),
            "children_run": len(receipts), "passed_children": sum(row["status"] == "passed" for row in receipts)}


def replay(args: argparse.Namespace) -> dict[str, Any]:
    packet = external(args.packet)
    state = load_json(packet / "capture-state.json")
    require(state.get("schema") == "litchi.performance.0925.provider-capture-state.v1",
            "capture state schema changed")
    require(state.get("status") == "completed", "copied packet state is not completed")
    require(state.get("children_run") == 42, "copied packet child count changed")
    require(state.get("custody_revalidation_after", {}).get("status") == "verified",
            "copied packet lacks successful post-capture custody revalidation")
    manifest_path = packet / "manifest.json"
    require(manifest_path.is_file(), "copied packet manifest is missing")
    manifest, comparison = _verify_manifest(manifest_path, require_authorized=False,
                                            packet_root=packet)
    require(state.get("process_deadline_seconds") == manifest.get("process_deadline_seconds"),
            "copied packet process deadline changed")
    rows = state.get("children")
    require(isinstance(rows, list) and len(rows) == 42, "copied packet child count changed")
    by_id = {str(row.get("schedule", {}).get("id")): row for row in rows}
    require(len(by_id) == 42, "copied packet has duplicate or missing child IDs")
    expected = {str(row["id"]): row for row in manifest["schedule"]["rows"]}
    require(set(by_id) == set(expected), "copied packet schedule IDs changed")
    validated = 0
    for child_id, receipt in by_id.items():
        schedule = expected[child_id]
        require(receipt.get("schedule") == schedule, f"{child_id}: schedule changed")
        require(receipt.get("status") == "passed", f"{child_id}: child was not passed")
        require(receipt.get("exit_code") == 0 and receipt.get("deadline_exceeded") is False,
                f"{child_id}: child terminal status changed")
        argv = receipt.get("argv")
        expected_argv = _child_command(
            schedule, manifest, packet / "children" / child_id / "report.json"
        )
        require(isinstance(argv, list) and len(argv) == len(expected_argv)
                and argv[:-1] == expected_argv[:-1]
                and Path(str(argv[-1])).name == "report.json"
                and Path(str(argv[-1])).parent.name == child_id,
                f"{child_id}: child argv changed")
        require(receipt.get("environment") == {
            "LITCHI_DOCX_INGRESS_ROUTE": schedule["route"]
        }, f"{child_id}: child route environment changed")
        require(receipt.get("cwd") == manifest["source"]["source_root"],
                f"{child_id}: child working directory changed")
        for field in ("stdout", "stderr_and_time", "time"):
            verify_local_descriptor(packet, receipt.get(field), f"{child_id}: {field}")
        time_path = verify_local_descriptor(packet, receipt.get("time"), f"{child_id}: time")
        require(receipt.get("rss_bytes") == _rss_bytes(time_path),
                f"{child_id}: RSS receipt changed or is unavailable")
        report_descriptor = receipt.get("report")
        require(isinstance(report_descriptor, dict), f"{child_id}: report descriptor missing")
        report_path = verify_local_descriptor(packet, report_descriptor, f"{child_id}: report")
        require(report_path == (packet / "children" / child_id / "report.json").resolve(),
                f"{child_id}: report path changed")
        report = load_json(report_path)
        comparison.validate_report(report, schedule["route"], warmups=schedule["warmups"],
                                   samples=schedule["samples"], label=child_id,
                                   binary=manifest["binary"])
        validated += 1
    return {"schema": "litchi.performance.0925.provider-capture-replay.v1",
            "status": "verified", "packet": str(packet), "reports": validated,
            "performance_children_run": 0,
            "claim_scope": "offline integrity replay only; no provider ranking or adoption"}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    freeze_parser = sub.add_parser("freeze")
    freeze_parser.add_argument("--source-root", type=Path, required=True)
    freeze_parser.add_argument("--source-receipt", type=Path, required=True)
    freeze_parser.add_argument("--full-source-listing", type=Path, required=True)
    freeze_parser.add_argument("--build-receipt", type=Path, required=True)
    freeze_parser.add_argument("--protocol", type=Path, required=True)
    freeze_parser.add_argument("--manifest", type=Path, required=True)
    freeze_parser.add_argument("--process-deadline", type=float,
                               default=DEFAULT_PROCESS_DEADLINE_SECONDS)
    capture_parser = sub.add_parser("capture")
    capture_parser.add_argument("--manifest", type=Path, required=True)
    capture_parser.add_argument("--packet", type=Path, default=DEFAULT_PACKET_ROOT)
    capture_parser.add_argument("--process-deadline", type=float, default=DEFAULT_PROCESS_DEADLINE_SECONDS)
    replay_parser = sub.add_parser("replay")
    replay_parser.add_argument("--packet", type=Path, required=True)
    args = parser.parse_args()
    try:
        result = freeze(args) if args.command == "freeze" else capture(args) if args.command == "capture" else replay(args)
        print(json.dumps(result, sort_keys=True))
    except Exception as error:
        print(f"0925 capture tooling failed: {type(error).__name__}: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
