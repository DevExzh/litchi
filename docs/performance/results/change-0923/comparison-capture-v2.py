#!/usr/bin/env python3
"""Capture, freeze, reduce, and replay the prospective 0923 comparison.

The packet root is deliberately external to the checkout.  ``capture`` is
fail-closed: it requires an externally frozen manifest and an explicit
``LITCHI_PERF_0923_CAPTURE_APPROVED=1`` environment guard.  A failed or
interrupted batch leaves a state marker and cannot be selectively recaptured.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import importlib.util
import json
import os
import re
import shutil
import statistics
import subprocess
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable


REPO_ROOT = Path(__file__).resolve().parents[4]
HERE = Path(__file__).resolve().parent
PACKET_ROOT = Path(os.environ.get("LITCHI_PERF_0923_PACKET_ROOT", "/tmp/litchi-performance-0923/packet-v2")).resolve()
PROTOCOL_PATH = HERE / "comparison-protocol-v2.json"
COMPARISON_PATH = HERE / "comparison-v2.py"
CONTROLS_PATH = HERE / "comparison-controls-v2.py"
RECEIPT_SCHEMA_PATH = HERE / "comparison-receipt-schema-v3.json"
CASE = "docx_source_backed_existing_section_layout_edit_save"
ROUTES = ("generic_read_at", "owned_vec", "owned_reader")
PAIR_ORDERS = ("baseline,candidate", "candidate,baseline") * 4
CPU = 12
WARMUPS = 30
SAMPLES = 200
CHILDREN = 48
ONE_MIB = 1 << 20
BASE_REVISION = "4e2532a0ce958f5928496b0ad9b356d8636916e1"
PINNED_ARCHIVE = Path(f"/tmp/litchi-performance-0923/base-{BASE_REVISION}.tar")
PINNED_LISTING = Path(f"/tmp/litchi-performance-0923/custody/source-listing-{BASE_REVISION}.jsonl")
HARNESS_PATCH = Path("/tmp/litchi-performance-0923/harness-0923.patch")
BASELINE_SOURCE_RECEIPT = Path("/tmp/litchi-performance-0923/receipts/baseline-source-final.json")
BASELINE_BUILD_RECEIPT = Path("/tmp/litchi-performance-0923/receipts/baseline-build-final.json")
BASELINE_BINARY = Path("/tmp/litchi-performance-0923/frozen-baseline-native/litchi-perf-baseline")
BUILD_PARITY_RECEIPT = Path("/tmp/litchi-performance-0923/custody/candidate-build-parity-v3.json")
BUILD_PROFILE_AUDIT_RECEIPT = Path("/tmp/litchi-performance-0923/receipts/candidate-build-profile-audit-v1.json")
SOURCE_DELTA_REVIEW = HERE / "root-source-delta.json"
SHARED_OWNER_QUALITY_REVIEW = HERE / "shared-owner-quality-v3.json"
BASELINE_REPORTS = {
    "generic_read_at": Path("/tmp/litchi-performance-0923/baseline-generic-final.json"),
    "owned_vec": Path("/tmp/litchi-performance-0923/baseline-owned-vec-final.json"),
    "owned_reader": Path("/tmp/litchi-performance-0923/baseline-owned-reader-final.json"),
}


class CaptureError(RuntimeError):
    """A fail-closed freeze, capture, reduction, or replay error."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise CaptureError(message)


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def canonical(value: Any) -> bytes:
    return json.dumps(value, sort_keys=True, separators=(",", ":")).encode("utf-8")


def sha256_path(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def descriptor(path: Path, *, limit: int | None = None) -> dict[str, Any]:
    path = path.resolve()
    require(path.exists() and path.is_file(), f"missing file: {path}")
    size = path.stat().st_size
    if limit is not None:
        require(size < limit, f"file is too large for retained packet evidence: {path} ({size} bytes)")
    return {"path": str(path), "bytes": size, "sha256": sha256_path(path)}


def ensure_external(path: Path) -> Path:
    path = path.resolve()
    try:
        path.relative_to(REPO_ROOT)
    except ValueError:
        return path
    raise CaptureError(f"packet/evidence path must be outside the checkout: {path}")


def write_new(path: Path, value: Any, *, max_bytes: int = ONE_MIB) -> None:
    path = ensure_external(path)
    require(not path.exists(), f"refusing to replace retained evidence: {path}")
    data = value if isinstance(value, bytes) else json.dumps(value, indent=2, sort_keys=True).encode("utf-8") + b"\n"
    require(len(data) < max_bytes, f"retained receipt is at least 1 MiB: {path} ({len(data)} bytes)")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)


def load_json(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as error:
        raise CaptureError(f"cannot read JSON {path}: {error}") from error
    require(isinstance(value, dict), f"expected JSON object: {path}")
    return value


def load_comparison() -> Any:
    spec = importlib.util.spec_from_file_location("litchi_perf_0923_comparison_v2_capture", COMPARISON_PATH)
    require(spec is not None and spec.loader is not None, f"cannot import {COMPARISON_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def schedule() -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for route in ROUTES:
        for pair_number, order in enumerate(PAIR_ORDERS, start=1):
            variants = order.split(",")
            for leg, variant in enumerate(variants, start=1):
                rows.append({
                    "index": len(rows) + 1,
                    "route": route,
                    "pair": pair_number,
                    "order": order,
                    "leg": leg,
                    "variant": variant,
                    "id": f"{route}-pair-{pair_number:02d}-leg-{leg}-{variant}",
                })
    require(len(rows) == CHILDREN, f"schedule cardinality changed: {len(rows)}")
    require(sum(row["variant"] == "baseline" for row in rows) == 24, "baseline child cardinality changed")
    require(sum(row["variant"] == "candidate" for row in rows) == 24, "candidate child cardinality changed")
    for route in ROUTES:
        require(sum(row["route"] == route for row in rows) == 16, f"route child cardinality changed: {route}")
    return rows


def schedule_descriptor() -> dict[str, Any]:
    rows = schedule()
    return {
        "routes": list(ROUTES),
        "pairs": 8,
        "pair_orders": list(PAIR_ORDERS),
        "children_per_route": 16,
        "children_per_version": 24,
        "children_total": CHILDREN,
        "rows": rows,
        "sha256": hashlib.sha256(canonical(rows)).hexdigest(),
    }


def _receipt_descriptor(path: Path) -> dict[str, Any]:
    return descriptor(path, limit=ONE_MIB)


def _embedded_descriptor(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}: descriptor missing")
    require(isinstance(value.get("path"), str), f"{label}: descriptor path missing")
    require(type(value.get("bytes")) is int and value["bytes"] >= 0, f"{label}: descriptor size invalid")
    require(isinstance(value.get("sha256"), str) and len(value["sha256"]) == 64,
            f"{label}: descriptor hash invalid")
    return value


def _source_file_claim(receipt: dict[str, Any], *, candidate: bool) -> Any:
    value = receipt.get("source_files")
    if candidate:
        require(isinstance(value, list) and value,
                "candidate source receipt must contain the v3 source_files list")
        require(all(isinstance(item, str) and item and not Path(item).is_absolute()
                    and ".." not in Path(item).parts for item in value),
                "candidate v3 source_files must contain safe relative paths")
    return value


def _receipt_source_binding(receipt_path: Path, *, candidate: bool) -> dict[str, Any] | None:
    if not receipt_path.exists():
        return None
    receipt = load_json(receipt_path)
    revision = (receipt.get("base_revision") if candidate
                else (receipt.get("source_revision") or BASE_REVISION))
    require(revision == BASE_REVISION, f"{receipt_path}: unexpected source base {revision}")
    # The baseline receipt historically names base-head.tar and an unmatched
    # listing.  Bind the prospective packet to the pinned archive/listing
    # descriptors while preserving the old receipt as an immutable record.
    archive_value = receipt.get("source_archive") if candidate else None
    listing_value = receipt.get("source_listing") if candidate else None
    patch_value = receipt.get("harness_patch")
    lock_value = receipt.get("cargo_lock")
    probe_value = receipt.get("probe_lock")
    tree_value = receipt.get("source_tree") if candidate else None
    production_patch_value = receipt.get("production_patch") if candidate else None
    if candidate:
        # Candidate custody must be self-contained.  In particular, it may
        # not silently borrow the baseline archive, listing, lock, or harness
        # patch when a candidate receipt omits one of them.
        for name, value in (("source_archive", archive_value), ("source_listing", listing_value),
                            ("harness_patch", patch_value), ("cargo_lock", lock_value),
                            ("probe_lock", probe_value), ("source_tree", tree_value),
                            ("production_patch", production_patch_value)):
            require(isinstance(value, dict) and isinstance(value.get("path"), str),
                    f"candidate v3 source receipt missing actual {name}")
        archive = Path(archive_value["path"])
        listing = Path(listing_value["path"])
        patch = Path(patch_value["path"])
        cargo_lock = Path(lock_value["path"])
        probe_lock = Path(probe_value["path"])
        source_tree = Path(tree_value["path"])
        production_patch = Path(production_patch_value["path"])
    else:
        archive = PINNED_ARCHIVE
        listing = PINNED_LISTING
        patch = HARNESS_PATCH
        cargo_lock = Path(receipt.get("cargo_lock", {}).get("path", "/tmp/litchi-performance-0923/baseline-source/tree/Cargo.lock"))
        probe_lock = Path("/tmp/litchi-performance-0923/baseline-source/tree/tools/perf-baseline/Cargo.lock")
        source_tree_value = receipt.get("source_tree", "/tmp/litchi-performance-0923/baseline-source/tree")
        if isinstance(source_tree_value, dict):
            source_tree_value = source_tree_value.get("path")
        require(isinstance(source_tree_value, str), f"{receipt_path}: baseline source tree path is missing")
        source_tree = Path(source_tree_value)
    binding = {
        "kind": "candidate" if candidate else "baseline",
        "revision": revision,
        "receipt": _receipt_descriptor(receipt_path),
        # Archives/listings are deliberately external and may exceed the
        # repository's one-MiB evidence threshold.  Only their compact
        # descriptors enter the manifest.
        "archive": descriptor(archive),
        "full_source_listing": descriptor(listing),
        "harness_patch": _receipt_descriptor(patch),
        "cargo_lock": _receipt_descriptor(cargo_lock),
        "probe_lock": _receipt_descriptor(probe_lock),
        "source_tree": str(source_tree.resolve()),
        "source_files": _source_file_claim(receipt, candidate=candidate),
        "legacy_receipt_listing_claim": {
            "bytes": receipt.get("source_listing_bytes"),
            "sha256": receipt.get("source_listing_sha256"),
            "status": "preserved_unmatched_legacy_claim",
        },
    }
    if candidate:
        binding["production_patch"] = _receipt_descriptor(production_patch)
        binding["production_patch_claim"] = production_patch_value
    if candidate:
        binding["candidate_base_revision"] = revision
    return binding


def _binary_binding(binary_path: Path, build_receipt_path: Path, *, candidate: bool) -> dict[str, Any] | None:
    if not binary_path.exists():
        return None
    require(build_receipt_path.exists(),
            f"{'candidate' if candidate else 'baseline'} build receipt is required: {build_receipt_path}")
    binary = descriptor(binary_path)
    binding: dict[str, Any] = {
        "kind": "candidate" if candidate else "baseline",
        "binary": binary,
    }
    receipt = load_json(build_receipt_path)
    expected = receipt.get("binary", {})
    require(expected.get("sha256") == binary["sha256"] and expected.get("bytes") == binary["bytes"],
            f"{build_receipt_path}: binary does not match the selected executable")
    binding["build_receipt"] = _receipt_descriptor(build_receipt_path)
    binding["build_record"] = {
        "source_revision": receipt.get("source_revision"),
        "profile": receipt.get("profile"),
        "argv": receipt.get("argv") or receipt.get("regeneration_command"),
        "rustc_vv": receipt.get("rustc_vv"),
        "source_receipt": receipt.get("source_receipt"),
        "features": receipt.get("features"),
        # Retain these fields for review, but do not infer parity from a
        # missing value in the historical baseline receipt.
        "environment": receipt.get("environment"),
        "profile_overrides": receipt.get("profile_overrides"),
        "cargo_v": receipt.get("cargo_v"),
    }
    require(isinstance(binding["build_record"]["rustc_vv"], str)
            and binding["build_record"]["rustc_vv"]
            and isinstance(binding["build_record"]["profile"], str)
            and binding["build_record"]["profile"],
            f"{build_receipt_path}: compiler/profile provenance is required")
    require(isinstance(binding["build_record"]["source_receipt"], dict),
            f"{build_receipt_path}: source receipt binding is required")
    return binding


def _env_path(name: str, default: Path) -> Path:
    return Path(os.environ.get(name, str(default))).resolve()


def _listing_records(path: Path) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    with path.open("r", encoding="utf-8") as stream:
        for line_number, line in enumerate(stream, start=1):
            if not line.strip():
                continue
            try:
                value = json.loads(line)
            except ValueError as error:
                raise CaptureError(f"source listing is not JSONL at {path}:{line_number}: {error}") from error
            require(isinstance(value, dict), f"source listing record is not an object at {path}:{line_number}")
            require(isinstance(value.get("path"), str) and not Path(value["path"]).is_absolute()
                    and ".." not in Path(value["path"]).parts,
                    f"source listing path is unsafe at {path}:{line_number}")
            require(type(value.get("bytes")) is int and value["bytes"] >= 0,
                    f"source listing size is invalid at {path}:{line_number}")
            require(isinstance(value.get("sha256"), str) and len(value["sha256"]) == 64,
                    f"source listing hash is invalid at {path}:{line_number}")
            records.append(value)
    require(records, f"source listing is empty: {path}")
    return records


def verify_source_closure(binding: dict[str, Any], label: str) -> None:
    """Rehash every file in a bound source tree against its external listing."""
    tree = Path(binding["source_tree"])
    listing = Path(binding["full_source_listing"]["path"])
    require(tree.exists() and tree.is_dir(), f"{label}: source tree is missing: {tree}")
    records = _listing_records(listing)
    listed: set[str] = set()
    for record in records:
        relative = record["path"]
        require(relative not in listed, f"{label}: duplicate source listing path: {relative}")
        listed.add(relative)
        path = tree / relative
        require(path.exists() and path.is_file(), f"{label}: source listing file is missing: {path}")
        current = descriptor(path)
        require(current["bytes"] == record["bytes"] and current["sha256"] == record["sha256"],
                f"{label}: source file changed: {relative}")
    actual = {str(path.relative_to(tree)) for path in tree.rglob("*") if path.is_file()}
    require(actual == listed,
            f"{label}: source closure differs (listed={len(listed)}, actual={len(actual)})")
    source_claim = binding.get("source_files")
    if label.startswith("candidate"):
        require(source_claim not in (None, [], {}), f"{label}: source file claim is missing")
    if isinstance(source_claim, list):
        claim_paths = {item if isinstance(item, str) else item.get("path")
                       for item in source_claim}
        require(None not in claim_paths and claim_paths <= listed,
                f"{label}: candidate source-file claim is outside the verified listing")
    elif isinstance(source_claim, dict):
        if {"path", "bytes", "sha256"} <= set(source_claim):
            claim_path = Path(source_claim["path"])
            verify_descriptor(source_claim, f"{label}.source-file-claim", external=True)
            claim_paths = {record["path"] for record in _listing_records(claim_path)}
        else:
            nested = source_claim.get("files")
            if isinstance(nested, list):
                claim_paths = {item if isinstance(item, str) else item.get("path")
                               for item in nested}
            elif isinstance(nested, dict):
                claim_paths = set(nested) | {
                    item.get("path") for item in nested.values()
                    if isinstance(item, dict) and isinstance(item.get("path"), str)
                }
            else:
                claim_paths = set(source_claim) | {
                    item.get("path") for item in source_claim.values()
                    if isinstance(item, dict) and isinstance(item.get("path"), str)
                }
        require(claim_paths <= listed,
                f"{label}: candidate source-file map is outside the verified listing")


def verify_candidate_patch(candidate: dict[str, Any], baseline: dict[str, Any]) -> None:
    """Apply the reviewed production patch to changed baseline files in scratch.

    The full source trees remain immutable.  Only files named by the patch are
    copied into a temporary external directory; the resulting bytes must match
    the candidate tree and the receipt's exact ``source_files`` claim.
    """
    patch_path = Path(candidate["production_patch"]["path"])
    patch_text = patch_path.read_text(encoding="utf-8")
    patch_paths: set[str] = set()
    for line in patch_text.splitlines():
        if line.startswith("--- a/"):
            path = line[6:]
        elif line.startswith("+++ b/"):
            path = line[6:]
        else:
            continue
        if path == "/dev/null":
            continue
        require(path and not Path(path).is_absolute() and ".." not in Path(path).parts,
                f"candidate production patch contains an unsafe path: {path!r}")
        patch_paths.add(path)
    source_claim = candidate.get("source_files")
    require(isinstance(source_claim, list) and all(isinstance(item, str) for item in source_claim),
            "candidate v3 source_files must be a list of relative paths")
    declared = set(source_claim)
    require(patch_paths == declared,
            f"candidate production patch paths differ from source_files: patch={sorted(patch_paths)} claim={sorted(declared)}")
    baseline_tree = Path(baseline["source_tree"])
    candidate_tree = Path(candidate["source_tree"])
    baseline_listing = {
        record["path"]: (record["bytes"], record["sha256"])
        for record in _listing_records(Path(baseline["full_source_listing"]["path"]))
    }
    candidate_listing = {
        record["path"]: (record["bytes"], record["sha256"])
        for record in _listing_records(Path(candidate["full_source_listing"]["path"]))
    }
    changed_listing_paths = {
        relative for relative in set(baseline_listing) | set(candidate_listing)
        if baseline_listing.get(relative) != candidate_listing.get(relative)
    }
    require(changed_listing_paths == patch_paths,
            "candidate listing changes differ from the reviewed production patch: "
            f"listing={sorted(changed_listing_paths)} patch={sorted(patch_paths)}")
    with tempfile.TemporaryDirectory(prefix="comparison-0923-patch-", dir="/tmp/litchi-performance-0923") as scratch_name:
        scratch = Path(scratch_name)
        for relative in patch_paths:
            source = baseline_tree / relative
            target = scratch / relative
            if source.exists():
                target.parent.mkdir(parents=True, exist_ok=True)
                shutil.copyfile(source, target)
        result = subprocess.run(
            ["patch", "--batch", "--forward", "--fuzz=0", "-p1"],
            cwd=scratch, input=patch_text.encode("utf-8"), stdout=subprocess.PIPE,
            stderr=subprocess.PIPE, check=False,
        )
        require(result.returncode == 0,
                f"candidate production patch did not apply to the retained base: {result.stderr.decode(errors='replace')[-1000:]}")
        for relative in patch_paths:
            expected = candidate_tree / relative
            actual = scratch / relative
            require(expected.exists() == actual.exists(),
                    f"candidate patched file presence differs: {relative}")
            if expected.exists():
                require(expected.read_bytes() == actual.read_bytes(),
                        f"candidate bytes differ from base plus production patch: {relative}")


def verify_review_bindings(manifest: dict[str, Any], *, external: bool) -> None:
    """Check the independent source, quality, and build-profile evidence.

    The historical baseline build receipt predates the retained environment
    and profile metadata.  Its ``release`` label is therefore only paired
    with the independently retained fingerprint graph and profile audit; a
    missing field is never treated as equal to another missing field.
    """
    review = manifest.get("review_bindings")
    require(isinstance(review, dict), "independent review bindings are missing")
    required = ("source_delta", "shared_owner_quality", "build_profile_audit",
                "build_fingerprint_parity")
    for name in required:
        value = review.get(name)
        require(isinstance(value, dict), f"review binding is missing: {name}")
        verify_descriptor(value, f"review.{name}", external=external)
    if not external:
        return

    source_delta = load_json(Path(review["source_delta"]["path"]))
    require(source_delta.get("schema") == "litchi.performance.0923.root-source-delta.v1"
            and source_delta.get("status") == "pass",
            "root source delta review is not passing")
    require(source_delta.get("changed_files") == ["crates/litchi-opc/src/source_backed.rs"],
            "root source delta changed-file set differs")

    quality = load_json(Path(review["shared_owner_quality"]["path"]))
    require(quality.get("schema") == "litchi.performance.0923.shared-owner-quality.v3",
            "shared-owner quality receipt schema changed")
    results = quality.get("results")
    require(isinstance(results, list) and results
            and all(item.get("exit") == 0 and item.get("result") == "pass" for item in results),
            "shared-owner quality receipt contains a failed or incomplete result")

    audit = load_json(Path(review["build_profile_audit"]["path"]))
    require(audit.get("schema") == "litchi.performance.0923.build-profile-audit.v1",
            "build profile audit schema changed")
    parity = audit.get("parity_receipt", {})
    parity_path = Path(parity.get("path", ""))
    parity_binding = review["build_fingerprint_parity"]
    require(str(parity_path.resolve()) == str(Path(parity_binding["path"]).resolve()),
            "build profile audit points at a different fingerprint parity receipt")
    require(parity.get("sha256") == parity_binding.get("sha256"),
            "build profile audit fingerprint parity hash differs")
    parity_receipt = load_json(parity_path)
    require(parity_receipt.get("field_mismatch_records") == 0
            and parity_receipt.get("baseline_only") == 0
            and parity_receipt.get("candidate_only") == 0,
            "build fingerprint parity receipt reports a mismatch")
    require(parity.get("profile_hashes_equal") is True,
            "build profile fingerprint hashes are not equal")
    comparison = parity.get("record_comparison", "")
    require("0 field mismatches" in comparison and "238 shared" in comparison,
            "build profile fingerprint comparison is incomplete")
    environment = audit.get("actual_command_and_environment", {})
    require(environment.get("profile_environment_variables_present") == [],
            "build profile audit reports unexpected profile environment overrides")
    binaries = audit.get("binary_section_evidence", {})
    for variant, key in (("baseline", "baseline"), ("candidate", "candidate")):
        binding = manifest.get("binaries", {}).get(variant)
        if binding is None:
            continue
        expected = binding.get("binary", {})
        observed = binaries.get(key, {})
        require(observed.get("sha256") == expected.get("sha256"),
                f"build profile audit {variant} binary differs from manifest")


def build_manifest(*, dry_run: bool) -> dict[str, Any]:
    protocol = descriptor(PROTOCOL_PATH, limit=ONE_MIB)
    comparison = descriptor(COMPARISON_PATH, limit=ONE_MIB)
    controls = descriptor(CONTROLS_PATH, limit=ONE_MIB)
    capture_script = descriptor(Path(__file__), limit=ONE_MIB)
    receipt_schema = descriptor(RECEIPT_SCHEMA_PATH, limit=ONE_MIB)
    baseline_binary = _env_path("LITCHI_PERF_0923_BASELINE_BINARY", BASELINE_BINARY)
    candidate_binary = _env_path("LITCHI_PERF_0923_CANDIDATE_BINARY", Path("/nonexistent/candidate"))
    baseline_build = _env_path("LITCHI_PERF_0923_BASELINE_BUILD_RECEIPT", BASELINE_BUILD_RECEIPT)
    candidate_build = _env_path("LITCHI_PERF_0923_CANDIDATE_BUILD_RECEIPT", Path("/nonexistent/candidate-build.json"))
    baseline_source = _env_path("LITCHI_PERF_0923_BASELINE_SOURCE_RECEIPT", BASELINE_SOURCE_RECEIPT)
    candidate_source = _env_path("LITCHI_PERF_0923_CANDIDATE_SOURCE_RECEIPT", Path("/nonexistent/candidate-source.json"))
    baseline_source_binding = _receipt_source_binding(baseline_source, candidate=False)
    candidate_source_binding = None
    if candidate_source.exists():
        candidate_source_binding = _receipt_source_binding(candidate_source, candidate=True)
    baseline = _binary_binding(baseline_binary, baseline_build, candidate=False)
    candidate = _binary_binding(candidate_binary, candidate_build, candidate=True) if candidate_binary.exists() else None
    require(baseline_source_binding is not None, f"baseline source receipt missing: {baseline_source}")
    require(baseline is not None, f"baseline executable missing: {baseline_binary}")
    review_bindings = {
        "source_delta": descriptor(SOURCE_DELTA_REVIEW, limit=ONE_MIB),
        "shared_owner_quality": descriptor(SHARED_OWNER_QUALITY_REVIEW, limit=ONE_MIB),
        "build_profile_audit": descriptor(BUILD_PROFILE_AUDIT_RECEIPT, limit=ONE_MIB),
        "build_fingerprint_parity": descriptor(BUILD_PARITY_RECEIPT, limit=ONE_MIB),
    }
    now = utc_now()
    schedule_info = schedule_descriptor()
    manifest = {
        "schema": "litchi.performance.0923.comparison-manifest.v2",
        "status": "frozen" if candidate is not None and candidate_source_binding is not None else "prospective",
        "created_utc": now,
        "created_timestamp_provenance": "runner wall clock at manifest construction",
        "protocol": protocol,
        "comparison_runner": comparison,
        "controls_runner": controls,
        "capture_runner": capture_script,
        "receipt_schema": receipt_schema,
        "review_bindings": review_bindings,
        "baseline_source": baseline_source_binding,
        "candidate_source": candidate_source_binding,
        "binaries": {"baseline": baseline, "candidate": candidate},
        "case": CASE,
        "routes": list(ROUTES),
        "native": {
            "cpu": CPU,
            "warmups_per_child": WARMUPS,
            "samples_per_child": SAMPLES,
            "schedule": schedule_info["rows"],
            "schedule_sha256": schedule_info["sha256"],
            "capture_approval_env": "LITCHI_PERF_0923_CAPTURE_APPROVED=1",
            "process_wrapper": ["taskset", "-c", str(CPU), "/usr/bin/time", "-v"],
        },
        "allocation_evidence": {"lane": None, "measured": False},
        "performance_children_run": 0,
        "freeze_mode": "dry_run" if dry_run else "write_external_manifest",
    }
    return manifest


def verify_descriptor(value: dict[str, Any], label: str, *, external: bool = True) -> None:
    path = Path(_embedded_descriptor(value, label)["path"])
    if external:
        current = descriptor(path)
        require(current["bytes"] == value["bytes"] and current["sha256"] == value["sha256"],
                f"{label}: descriptor changed")


def verify_manifest(manifest: dict[str, Any], *, external: bool = True) -> None:
    require(manifest.get("schema") == "litchi.performance.0923.comparison-manifest.v2", "manifest schema changed")
    require(manifest.get("case") == CASE, "manifest case changed")
    require(manifest.get("routes") == list(ROUTES), "manifest route set changed")
    native = manifest.get("native", {})
    require(native.get("cpu") == CPU and native.get("warmups_per_child") == WARMUPS
            and native.get("samples_per_child") == SAMPLES, "manifest native configuration changed")
    rows = native.get("schedule")
    require(rows == schedule_descriptor()["rows"], "manifest schedule changed")
    require(manifest.get("allocation_evidence", {}).get("measured") is False,
            "manifest must state that allocation evidence is unmeasured")
    for label, value in (("protocol", manifest.get("protocol")),
                         ("comparison_runner", manifest.get("comparison_runner")),
                         ("controls_runner", manifest.get("controls_runner")),
                         ("capture_runner", manifest.get("capture_runner")),
                         ("receipt_schema", manifest.get("receipt_schema"))):
        verify_descriptor(value, label, external=external)
    verify_review_bindings(manifest, external=external)
    for variant in ("baseline", "candidate"):
        binding = manifest.get("binaries", {}).get(variant)
        if binding is None:
            require(variant == "candidate" and manifest.get("status") == "prospective",
                    f"{variant}: binary binding missing")
            continue
        verify_descriptor(binding["binary"], f"{variant}.binary", external=external)
        if binding.get("build_receipt") is not None:
            verify_descriptor(binding["build_receipt"], f"{variant}.build_receipt", external=external)
            source_receipt_claim = binding.get("build_record", {}).get("source_receipt")
            source_binding = manifest.get("baseline_source" if variant == "baseline" else "candidate_source")
            if isinstance(source_receipt_claim, dict):
                require(source_receipt_claim.get("sha256") == source_binding.get("receipt", {}).get("sha256"),
                        f"{variant}: build receipt is bound to a different source receipt")
    for variant, binding in (("baseline", manifest.get("baseline_source")),
                             ("candidate", manifest.get("candidate_source"))):
        if binding is None:
            require(variant == "candidate" and manifest.get("status") == "prospective",
                    f"{variant}: source binding missing")
            continue
        for key in ("receipt", "archive", "full_source_listing", "harness_patch", "cargo_lock", "probe_lock"):
            verify_descriptor(binding[key], f"{variant}.source.{key}", external=external)
        if variant == "candidate":
            verify_descriptor(binding["production_patch"], "candidate.source.production_patch", external=external)
            patch_claim = binding.get("production_patch_claim")
            if patch_claim is not None:
                claimed_hash = patch_claim.get("sha256") if isinstance(patch_claim, dict) else patch_claim
                require(claimed_hash == binding["production_patch"]["sha256"],
                        "candidate production patch claim does not match the bound patch")
        if external:
            verify_source_closure(binding, f"{variant}.source")
    baseline_binding = manifest.get("binaries", {}).get("baseline") or {}
    candidate_binding = manifest.get("binaries", {}).get("candidate") or {}
    baseline_build = baseline_binding.get("build_record", {})
    candidate_build = candidate_binding.get("build_record")
    candidate_source_binding = manifest.get("candidate_source")
    baseline_source_binding = manifest.get("baseline_source")
    if external and candidate_source_binding is not None:
        verify_candidate_patch(candidate_source_binding, baseline_source_binding)
    for variant, build in (("baseline", baseline_build), ("candidate", candidate_build)):
        if build is None:
            continue
        argv_text = " ".join(build.get("argv", [])) if isinstance(build.get("argv"), list) else str(build.get("argv", ""))
        for flag in ("--locked", "--offline", "--release"):
            require(flag in argv_text, f"{variant}: build provenance is missing required flag {flag}")
    if candidate_build is not None:
        require(candidate_build.get("rustc_vv") == baseline_build.get("rustc_vv"),
                "candidate compiler fingerprint differs from baseline")
        require(candidate_build.get("profile") == baseline_build.get("profile"),
                "candidate build profile differs from baseline")
        baseline_source = manifest.get("baseline_source", {})
        candidate_source = manifest.get("candidate_source", {})
        require(baseline_source.get("cargo_lock", {}).get("sha256") == candidate_source.get("cargo_lock", {}).get("sha256"),
                "candidate Cargo.lock differs from baseline")
        require(baseline_source.get("probe_lock", {}).get("sha256") == candidate_source.get("probe_lock", {}).get("sha256"),
                "candidate probe lock differs from baseline")
        require(baseline_source.get("harness_patch", {}).get("sha256") == candidate_source.get("harness_patch", {}).get("sha256"),
                "candidate harness patch differs from baseline")
        require(candidate_source.get("candidate_base_revision", BASE_REVISION) == BASE_REVISION,
                "candidate base revision differs from the retained baseline")


def packet_path(name: str) -> Path:
    return ensure_external(PACKET_ROOT / name)


def action_schedule(args: argparse.Namespace) -> dict[str, Any]:
    value = schedule_descriptor()
    if args.output:
        write_new(ensure_external(args.output), value)
    else:
        print(json.dumps(value, indent=2, sort_keys=True))
    return value


def action_freeze(args: argparse.Namespace) -> dict[str, Any]:
    value = build_manifest(dry_run=args.dry_run)
    if args.dry_run:
        print(json.dumps(value, indent=2, sort_keys=True))
    else:
        path = packet_path("comparison-manifest-v2.json")
        # Keep the small source/build receipts beside an eventual packet so a
        # moved replay does not depend on the original external receipt paths.
        # Large source listings, archives, and binaries remain descriptor-only.
        custody: dict[str, Any] = {}
        for variant in ("baseline", "candidate"):
            source_binding = value.get("baseline_source" if variant == "baseline" else "candidate_source")
            binary_binding = value.get("binaries", {}).get(variant)
            if source_binding is not None:
                relative = Path("custody") / f"{variant}-source-receipt.json"
                copied = copy_new(Path(source_binding["receipt"]["path"]), path.parent / relative)
                custody[str(relative)] = copied["copy"]
            if binary_binding is not None:
                relative = Path("custody") / f"{variant}-build-receipt.json"
                copied = copy_new(Path(binary_binding["build_receipt"]["path"]), path.parent / relative)
                custody[str(relative)] = copied["copy"]
        value["local_custody"] = {
            "scope": "small source/build receipts copied for independent replay; all large artifacts remain external descriptors",
            "files": custody,
        }
        write_new(path, value)
        print(json.dumps({"manifest": descriptor(path), "status": value["status"]}, sort_keys=True))
    return value


def copy_new(source: Path, destination: Path, *, limit: int = ONE_MIB) -> dict[str, Any]:
    source = source.resolve()
    destination = ensure_external(destination)
    require(source.exists() and source.is_file(), f"missing packet input: {source}")
    value = descriptor(source, limit=limit)
    require(not destination.exists(), f"refusing to replace retained packet file: {destination}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(source, destination)
    require(descriptor(destination)["sha256"] == value["sha256"], f"copy hash mismatch: {destination}")
    return {"source": value, "copy": descriptor(destination)}


def action_seed_baseline(args: argparse.Namespace) -> dict[str, Any]:
    root = ensure_external(args.packet_root or PACKET_ROOT)
    require(not root.exists(), f"refusing to replace retained packet root: {root}")
    root.mkdir(parents=True)
    files: dict[str, Any] = {}
    for route, source in BASELINE_REPORTS.items():
        files[f"baseline/{route}.json"] = copy_new(source, root / "baseline" / f"{route}.json")
    controls = Path("/tmp/litchi-performance-0923/receipts/comparison-controls-v2.json")
    files["tooling/comparison-controls-v2.json"] = copy_new(controls, root / "tooling" / controls.name)
    for source in (PROTOCOL_PATH, COMPARISON_PATH, CONTROLS_PATH, Path(__file__), RECEIPT_SCHEMA_PATH):
        files[f"tooling/{source.name}"] = copy_new(source, root / "tooling" / source.name)
    for source in (SOURCE_DELTA_REVIEW, SHARED_OWNER_QUALITY_REVIEW,
                   BUILD_PROFILE_AUDIT_RECEIPT, BUILD_PARITY_RECEIPT):
        files[f"tooling/{source.name}"] = copy_new(source, root / "tooling" / source.name)
    seed = {
        "schema": "litchi.performance.0923.baseline-packet-seed.v2",
        "created_utc": utc_now(),
        "performance_children_run": 0,
        "files": files,
        "reduction_status": "not available: candidate packet and comparative children are intentionally absent",
    }
    write_new(root / "packet-seed.json", seed)
    return seed


def parse_rss(path: Path) -> int | None:
    if not path.exists():
        return None
    text = path.read_text(encoding="utf-8", errors="replace")
    match = re.search(r"Maximum resident set size \(kbytes\):\s*(\d+)", text)
    return int(match.group(1)) if match else None


def receipt_payload(receipt: dict[str, Any]) -> bytes:
    value = dict(receipt)
    value.pop("receipt_payload_sha256", None)
    value.pop("receipt_payload_bytes", None)
    return canonical(value)


def child_paths(root: Path, row: dict[str, Any]) -> dict[str, Path]:
    child = root / "children" / row["id"]
    return {
        "root": child,
        "report": child / "report.json",
        "stdout": child / "stdout.log",
        "stderr": child / "stderr.log",
        "time": child / "process-time.txt",
        "receipt": child / "child.json",
        "receipt_hash": child / "child.json.sha256",
    }


def safe_relative(root: Path, relative: str, label: str) -> Path:
    candidate = Path(relative)
    require(not candidate.is_absolute() and ".." not in candidate.parts,
            f"{label}: relative path is unsafe")
    path = (root / candidate).resolve()
    try:
        path.relative_to(root.resolve())
    except ValueError as error:
        raise CaptureError(f"{label}: path escapes packet root") from error
    return path


def preflight_capture(manifest: dict[str, Any], root: Path) -> None:
    require(manifest.get("status") == "frozen", "capture requires a complete frozen candidate manifest")
    require(os.environ.get("LITCHI_PERF_0923_CAPTURE_APPROVED") == "1",
            "capture requires LITCHI_PERF_0923_CAPTURE_APPROVED=1")
    state = root / "capture-state.json"
    require(not state.exists(), f"capture state already exists; recapture is prohibited: {state}")
    rows = manifest["native"]["schedule"]
    paths = [child_paths(root, row) for row in rows]
    for values in paths:
        for key, path in values.items():
            if key != "root":
                require(not path.exists(), f"capture output already exists: {path}")
    require(manifest["binaries"].get("candidate") is not None, "candidate binary binding is missing")
    verify_manifest(manifest, external=True)


def action_capture(args: argparse.Namespace) -> dict[str, Any]:
    root = ensure_external(args.packet_root or PACKET_ROOT)
    manifest_path = root / "comparison-manifest-v2.json"
    require(manifest_path.exists(), f"frozen manifest is missing: {manifest_path}")
    manifest = load_json(manifest_path)
    preflight_capture(manifest, root)
    state_path = root / "capture-state.json"
    state = {"schema": "litchi.performance.0923.capture-state.v2", "status": "running",
             "started_utc": utc_now(), "children_started": 0, "children_completed": 0}
    write_new(state_path, state)
    failures: list[str] = []
    rows = manifest["native"]["schedule"]
    try:
        for row in rows:
            paths = child_paths(root, row)
            paths["root"].mkdir(parents=True, exist_ok=False)
            binary = manifest["binaries"][row["variant"]]["binary"]["path"]
            env = os.environ.copy()
            env.update({"LITCHI_DOCX_INGRESS_ROUTE": row["route"], "LC_ALL": "C"})
            argv = ["taskset", "-c", str(CPU), "/usr/bin/time", "-v", "-o", str(paths["time"]),
                    binary, "--case", CASE, "--samples", str(SAMPLES), "--warmup", str(WARMUPS),
                    "--json", str(paths["report"])]
            started = utc_now()
            with paths["stdout"].open("wb") as stdout, paths["stderr"].open("wb") as stderr:
                process = subprocess.Popen(argv, cwd=root, env=env, stdout=stdout, stderr=stderr)
                state["children_started"] += 1
                state_path.write_text(json.dumps(state, indent=2, sort_keys=True) + "\n")
                exit_code = process.wait()
            finished = utc_now()
            report_descriptor = None
            if paths["report"].exists():
                report_descriptor = descriptor(paths["report"], limit=ONE_MIB)
            rss_value = parse_rss(paths["time"])
            require(isinstance(rss_value, int) and rss_value > 0,
                    f"{row['id']}: /usr/bin/time did not emit a positive max RSS")
            receipt = {
                "schema": "litchi.performance.0923.comparison-child.v2",
                "schedule": row,
                "argv": argv,
                "cwd": str(root),
                "env_overrides": {"LITCHI_DOCX_INGRESS_ROUTE": row["route"], "LC_ALL": "C"},
                "pid": process.pid,
                "started_utc": started,
                "finished_utc": finished,
                "exit_code": exit_code,
                "binary": manifest["binaries"][row["variant"]]["binary"],
                "stdout": descriptor(paths["stdout"]),
                "stderr": descriptor(paths["stderr"]),
                "process_time": descriptor(paths["time"]),
                "max_rss_kib": rss_value,
                "report": report_descriptor,
                "report_relative": str(paths["report"].relative_to(root)),
                "stdout_relative": str(paths["stdout"].relative_to(root)),
                "stderr_relative": str(paths["stderr"].relative_to(root)),
                "process_time_relative": str(paths["time"].relative_to(root)),
            }
            payload = receipt_payload(receipt)
            receipt["receipt_payload_sha256"] = hashlib.sha256(payload).hexdigest()
            receipt["receipt_payload_bytes"] = len(payload)
            write_new(paths["receipt"], receipt)
            receipt_info = descriptor(paths["receipt"])
            write_new(paths["receipt_hash"], {
                "schema": "litchi.performance.0923.child-receipt-hash.v1",
                "bytes": receipt_info["bytes"],
                "sha256": receipt_info["sha256"],
            })
            state["children_completed"] += 1
            state_path.write_text(json.dumps(state, indent=2, sort_keys=True) + "\n")
            if exit_code != 0 or report_descriptor is None:
                failures.append(row["id"])
        # Revalidate the complete source closure and all frozen descriptors
        # after the last child, before the batch can be marked complete.
        verify_manifest(manifest, external=True)
        state["source_closure_revalidated"] = True
        state["status"] = "failed" if failures else "completed"
        state["finished_utc"] = utc_now()
        state["failures"] = failures
        state_path.write_text(json.dumps(state, indent=2, sort_keys=True) + "\n")
    except BaseException:
        state["status"] = "interrupted"
        state["finished_utc"] = utc_now()
        state_path.write_text(json.dumps(state, indent=2, sort_keys=True) + "\n")
        raise
    require(not failures, f"capture batch failed; selective recapture is prohibited: {failures}")
    return state


def _read_child(root: Path, row: dict[str, Any], *, external: bool) -> tuple[dict[str, Any], dict[str, Any]]:
    paths = child_paths(root, row)
    receipt = load_json(paths["receipt"])
    require(paths["receipt_hash"].is_file(), f"{row['id']}: receipt hash sidecar is missing")
    receipt_info = descriptor(paths["receipt"])
    receipt_hash = load_json(paths["receipt_hash"])
    require(receipt_hash.get("bytes") == receipt_info["bytes"]
            and receipt_hash.get("sha256") == receipt_info["sha256"],
            f"{row['id']}: receipt file hash changed")
    payload = receipt_payload(receipt)
    require(receipt.get("receipt_payload_sha256") == hashlib.sha256(payload).hexdigest()
            and receipt.get("receipt_payload_bytes") == len(payload),
            f"{row['id']}: child receipt payload hash changed")
    # Hash-check every retained process log and the time wrapper output.  A
    # copied replay may omit a log that exceeded the packet's one-MiB receipt
    # limit; the raw capture still required and recorded it.
    for key, relative_key in (("stdout", "stdout_relative"), ("stderr", "stderr_relative"),
                              ("process_time", "process_time_relative")):
        info = receipt.get(key)
        require(isinstance(info, dict), f"{row['id']}: {key} descriptor missing")
        if external:
            verify_descriptor(info, f"{row['id']}.{key}", external=True)
        else:
            local = safe_relative(root, str(receipt.get(relative_key, "")), f"{row['id']}.{key}")
            if local.exists():
                current = descriptor(local)
                require(current["bytes"] == info.get("bytes") and current["sha256"] == info.get("sha256"),
                        f"{row['id']}: copied {key} descriptor changed")
    require(receipt.get("schedule") == row, f"{row['id']}: schedule receipt differs")
    require(receipt.get("exit_code") == 0, f"{row['id']}: child exit was not zero")
    report_info = receipt.get("report")
    require(isinstance(report_info, dict), f"{row['id']}: report descriptor missing")
    if external:
        verify_descriptor(report_info, f"{row['id']}.report", external=True)
    report_path = paths["report"]
    if not external:
        relative = receipt.get("report_relative")
        require(isinstance(relative, str), f"{row['id']}: replay report path missing")
        report_path = safe_relative(root, relative, f"{row['id']}.report")
        verify_descriptor(report_info, f"{row['id']}.report.replay-index", external=False)
        current = descriptor(report_path)
        require(current["bytes"] == report_info["bytes"] and current["sha256"] == report_info["sha256"],
                f"{row['id']}: copied report descriptor changed")
    time_path = Path(receipt["process_time"]["path"]) if external else safe_relative(
        root, str(receipt.get("process_time_relative", "")), f"{row['id']}.process_time")
    require(time_path.is_file(), f"{row['id']}: process-time log is missing")
    rss = parse_rss(time_path)
    require(isinstance(rss, int) and rss > 0, f"{row['id']}: parsed max RSS is missing")
    require(receipt.get("max_rss_kib") == rss, f"{row['id']}: max RSS value changed")
    return receipt, load_json(report_path)


def percentile(values: Iterable[float], fraction: float) -> float:
    ordered = sorted(float(value) for value in values)
    require(ordered, "cannot calculate a percentile of an empty vector")
    index = min(max(int((len(ordered) * fraction) + 0.999999999) - 1, 0), len(ordered) - 1)
    return ordered[index]


def reduce_packet(manifest: dict[str, Any], root: Path, *, external: bool) -> dict[str, Any]:
    verify_manifest(manifest, external=external)
    module = load_comparison()
    cells: dict[str, dict[int, dict[str, tuple[dict[str, Any], dict[str, Any]]]]] = {
        route: {} for route in ROUTES
    }
    rss: dict[str, dict[str, list[int]]] = {
        route: {"baseline": [], "candidate": []} for route in ROUTES
    }
    latency: dict[str, dict[str, list[float]]] = {
        route: {"baseline": [], "candidate": []} for route in ROUTES
    }
    phase_values: dict[str, dict[str, dict[str, list[float]]]] = {
        route: {"baseline": {key: [] for key in load_comparison().PHASES},
                "candidate": {key: [] for key in load_comparison().PHASES}}
        for route in ROUTES
    }
    child_summaries: list[dict[str, Any]] = []
    for row in manifest["native"]["schedule"]:
        receipt, report = _read_child(root, row, external=external)
        binary = manifest["binaries"][row["variant"]]["binary"]
        validated = module.validate_report(report, row["route"], warmups=WARMUPS, samples=SAMPLES,
                                           label=row["id"], binary=binary)
        p50 = percentile(validated["elapsed_ns"], 0.50)
        p95 = percentile(validated["elapsed_ns"], 0.95)
        p99 = percentile(validated["elapsed_ns"], 0.99)
        child_summaries.append({"id": row["id"], "route": row["route"], "pair": row["pair"],
                                "variant": row["variant"], "p50_ns": p50, "p95_ns": p95, "p99_ns": p99,
                                "phase_p50_ns": {key: percentile(values, 0.50)
                                                 for key, values in validated["phase_ns"].items()},
                                "max_rss_kib": receipt.get("max_rss_kib")})
        latency[row["route"]][row["variant"]].extend([p50, p95, p99])
        for key, values in validated["phase_ns"].items():
            phase_values[row["route"]][row["variant"]][key].append(percentile(values, 0.50))
        if isinstance(receipt.get("max_rss_kib"), int):
            rss[row["route"]][row["variant"]].append(receipt["max_rss_kib"])
        cells[row["route"]].setdefault(row["pair"], {})[row["variant"]] = (report, binary)
    reductions: dict[str, Any] = {}
    for route in ROUTES:
        pairs = []
        for pair in range(1, 9):
            cell = cells[route].get(pair, {})
            require(set(cell) == {"baseline", "candidate"}, f"{route} pair {pair}: incomplete pair")
            pairs.append((cell["baseline"][0], cell["candidate"][0]))
        reductions[route] = module.compare_paired_cells(
            pairs, route, f"{route}-reduction", gate="generic_control" if route == "generic_read_at" else "owned_benefit",
            warmups=WARMUPS, samples=SAMPLES,
            baseline_binary=manifest["binaries"]["baseline"]["binary"],
            candidate_binary=manifest["binaries"]["candidate"]["binary"],
        )
    rss_review: dict[str, Any] = {}
    for route, variants in rss.items():
        if all(variants[name] for name in ("baseline", "candidate")):
            baseline_median = statistics.median(variants["baseline"])
            candidate_median = statistics.median(variants["candidate"])
            rss_review[route] = {
                "baseline_median_kib": baseline_median,
                "candidate_median_kib": candidate_median,
                "candidate_over_baseline_ratio": candidate_median / baseline_median if baseline_median else None,
                "increase_fraction": (candidate_median / baseline_median - 1.0) if baseline_median else None,
            }
        else:
            rss_review[route] = {"status": "unavailable", "baseline_samples": len(variants["baseline"]),
                                 "candidate_samples": len(variants["candidate"])}
    pooled_latency_review: dict[str, Any] = {}
    phase_review: dict[str, Any] = {}
    for route in ROUTES:
        pooled_latency_review[route] = {}
        phase_review[route] = {}
        for variant in ("baseline", "candidate"):
            values = latency[route][variant]
            pooled_latency_review[route][variant] = {
                "p50_ns": percentile(values, 0.50),
                "p95_ns": percentile(values, 0.95),
                "p99_ns": percentile(values, 0.99),
                "source": "pooled child p50/p95/p99 statistics; descriptive review flag",
            }
            phase_review[route][variant] = {
                phase: {"p50_ns": percentile(values, 0.50), "p95_ns": percentile(values, 0.95),
                        "p99_ns": percentile(values, 0.99)}
                for phase, values in phase_values[route][variant].items()
            }
    return {
        "schema": "litchi.performance.0923.comparison-reduction.v2",
        "created_utc": utc_now(),
        "manifest_status": manifest.get("status"),
        "children": len(child_summaries),
        "child_summaries": child_summaries,
        "routes": reductions,
        "pooled_latency_review": pooled_latency_review,
        "phase_summaries": phase_review,
        "whole_process_rss_review": rss_review,
        "performance_claim": "timing eligibility only; no production adoption",
        "recapture": "none; every child and pair was retained and replayable",
    }


def action_reduce(args: argparse.Namespace, *, replay: bool = False) -> dict[str, Any]:
    root = ensure_external(args.packet_root or PACKET_ROOT)
    manifest_path = root / "comparison-manifest-v2.json"
    if replay:
        manifest_path = root / "comparison-manifest-v2.json"
    require(manifest_path.exists(), f"manifest is missing: {manifest_path}")
    manifest = load_json(manifest_path)
    require(manifest.get("status") == "frozen", "reduction requires a frozen complete manifest")
    value = reduce_packet(manifest, root, external=not replay)
    if replay:
        value["replay"] = {"root": str(root), "external_bindings_revalidated": False,
                           "copied_reports_and_receipts_revalidated": True}
    output = ensure_external(args.output or (root / ("replay-reduction-v2.json" if replay else "comparison-reduction-v2.json")))
    write_new(output, value)
    return value


def action_copy(args: argparse.Namespace) -> dict[str, Any]:
    source = ensure_external(args.source_packet)
    destination = ensure_external(args.replay_root)
    require(source.exists() and source.is_dir(), f"source packet is missing: {source}")
    require(not destination.exists(), f"refusing to replace retained replay packet: {destination}")
    destination.mkdir(parents=True)
    copied: dict[str, Any] = {}
    for path in sorted(source.rglob("*")):
        if not path.is_file():
            continue
        relative = path.relative_to(source)
        # Replays need only immutable metadata, receipts, and raw reports.  A
        # process log is retained only when it fits the packet receipt bound;
        # large logs remain at the original external custody path.
        if (path.name.endswith(".log") or path.name == "process-time.txt") and path.stat().st_size >= ONE_MIB:
            continue
        copied[str(relative)] = copy_new(path, destination / relative)
    index = {"schema": "litchi.performance.0923.replay-index.v2", "created_utc": utc_now(),
             "source_packet": str(source), "files": copied,
             "excluded": "process logs are intentionally not needed to recompute report validation"}
    write_new(destination / "replay-index.json", index)
    return index


def verify_replay_index(root: Path) -> dict[str, Any]:
    index = load_json(root / "replay-index.json")
    require(index.get("schema") == "litchi.performance.0923.replay-index.v2", "replay index schema changed")
    files = index.get("files")
    require(isinstance(files, dict), "replay index file map missing")
    for relative, item in files.items():
        copy_info = item.get("copy") if isinstance(item, dict) else None
        require(isinstance(copy_info, dict), f"replay index copy descriptor missing: {relative}")
        copy_path = root / relative
        try:
            copy_path.resolve().relative_to(root.resolve())
        except ValueError as error:
            raise CaptureError(f"replay index path escapes packet root: {relative}") from error
        current = descriptor(copy_path)
        require(current["bytes"] == copy_info.get("bytes") and current["sha256"] == copy_info.get("sha256"),
                f"replay copy {relative}: local bytes or hash changed")
    manifest_path = root / "comparison-manifest-v2.json"
    require(manifest_path.is_file(), "replay packet has no local frozen manifest")
    manifest = load_json(manifest_path)
    manifest_info = files.get("comparison-manifest-v2.json")
    if manifest_info is not None:
        current = descriptor(manifest_path)
        expected = manifest_info["copy"]
        require(current["bytes"] == expected.get("bytes") and current["sha256"] == expected.get("sha256"),
                "replay manifest: local bytes or hash changed")
    local_custody = manifest.get("local_custody", {}).get("files")
    require(isinstance(local_custody, dict) and local_custody,
            "replay manifest has no local source/build receipt custody")
    for relative, expected in local_custody.items():
        local_info = files.get(relative)
        require(isinstance(local_info, dict), f"replay packet missing local custody file: {relative}")
        expected_copy = local_info.get("copy", {})
        require(expected_copy.get("bytes") == expected.get("bytes")
                and expected_copy.get("sha256") == expected.get("sha256"),
                f"replay custody descriptor differs from frozen manifest: {relative}")
        local_path = safe_relative(root, relative, f"replay custody {relative}")
        current = descriptor(local_path)
        require(current["bytes"] == expected.get("bytes") and current["sha256"] == expected.get("sha256"),
                f"replay custody file changed: {relative}")
    # The replay must use packet-local copies of the validator/protocol and
    # capture tooling.  Their hashes are compared to the frozen manifest; no
    # source or script path under the original /tmp packet is trusted.
    for field, filename in (("protocol", "comparison-protocol-v2.json"),
                            ("comparison_runner", "comparison-v2.py"),
                            ("controls_runner", "comparison-controls-v2.py"),
                            ("capture_runner", "comparison-capture-v2.py"),
                            ("receipt_schema", "comparison-receipt-schema-v3.json")):
        claim = manifest.get(field)
        require(isinstance(claim, dict), f"replay manifest missing {field}")
        local_info = files.get(f"tooling/{filename}")
        require(isinstance(local_info, dict), f"replay packet missing local {field}")
        require(local_info["copy"].get("sha256") == claim.get("sha256"),
                f"replay local {field} does not match frozen descriptor")
        local_path = safe_relative(root, f"tooling/{filename}", f"replay local {field}")
        current = descriptor(local_path)
        require(current["bytes"] == local_info["copy"].get("bytes")
                and current["sha256"] == local_info["copy"].get("sha256"),
                f"replay local {field}: bytes or hash changed")
    review_files = {
        "source_delta": "root-source-delta.json",
        "shared_owner_quality": "shared-owner-quality-v3.json",
        "build_profile_audit": "candidate-build-profile-audit-v1.json",
        "build_fingerprint_parity": "candidate-build-parity-v3.json",
    }
    for field, filename in review_files.items():
        claim = manifest.get("review_bindings", {}).get(field)
        require(isinstance(claim, dict), f"replay manifest missing review binding {field}")
        local_info = files.get(f"tooling/{filename}")
        require(isinstance(local_info, dict), f"replay packet missing review binding {field}")
        require(local_info["copy"].get("sha256") == claim.get("sha256"),
                f"replay local review {field} does not match frozen descriptor")
        local_path = safe_relative(root, f"tooling/{filename}", f"replay local review {field}")
        current = descriptor(local_path)
        require(current["bytes"] == local_info["copy"].get("bytes")
                and current["sha256"] == local_info["copy"].get("sha256"),
                f"replay local review {field}: bytes or hash changed")
    return manifest


def main(argv: Iterable[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=("schedule", "freeze", "seed-baseline", "capture", "reduce", "copy", "replay"))
    parser.add_argument("--output", type=Path)
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--packet-root", type=Path)
    parser.add_argument("--source-packet", type=Path)
    parser.add_argument("--replay-root", type=Path)
    args = parser.parse_args(list(argv) if argv is not None else None)
    try:
        if args.action == "schedule":
            action_schedule(args)
        elif args.action == "freeze":
            action_freeze(args)
        elif args.action == "seed-baseline":
            print(json.dumps(action_seed_baseline(args), indent=2, sort_keys=True))
        elif args.action == "capture":
            print(json.dumps(action_capture(args), indent=2, sort_keys=True))
        elif args.action == "reduce":
            print(json.dumps(action_reduce(args), indent=2, sort_keys=True))
        elif args.action == "copy":
            require(args.source_packet is not None and args.replay_root is not None,
                    "copy requires --source-packet and --replay-root")
            print(json.dumps(action_copy(args), indent=2, sort_keys=True))
        elif args.action == "replay":
            require(args.replay_root is not None, "replay requires --replay-root")
            args.packet_root = args.replay_root
            verify_replay_index(ensure_external(args.replay_root))
            print(json.dumps(action_reduce(args, replay=True), indent=2, sort_keys=True))
        return 0
    except Exception as error:
        print(f"0923 capture tooling failed: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
