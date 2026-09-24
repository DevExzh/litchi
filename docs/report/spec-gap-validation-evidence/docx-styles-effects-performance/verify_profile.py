#!/usr/bin/env python3
"""Fail-closed verifier for the gated stylesWithEffects performance scaffold.

This verifier only accepts receipts produced by the reviewed profile runner. It
does not infer a speedup, combine process RSS with allocator live bytes, or
treat a missing disposable Cargo target as a missing retained receipt.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import sys
import zipfile
from pathlib import Path
from typing import Any

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from verify import (  # noqa: E402
    MANIFEST_FORMAT,
    NATIVE_FIXTURES,
    PROFILE_SOURCE_COMMIT,
    verify_manifest,
)

NATIVE_EFFECTS_MEMBERS = {
    "Bug54849.docx": {
        "main": (19883, "799de1f7a4ce43f0ca101dc750e8a8bd6e75bcb721f7744d4787dad576cda3b1"),
        "glossary": (16138, "d27f6ced340ffa173b3b861b08a4006e76673a411687b4145f5e46dc6dae13eb"),
    },
    "ms-office-2010-signed.docx": {
        "main": (15710, "00c5cda7671bf545a8c97312f14b2b8bc0ee7fa469b36c25ee158c8a5c1c1568"),
    },
    "ComplexNumberedLists.docx": {
        "main": (15955, "b4bf5d355a45daf0a1085e73fe27041b5db22bfa23820f18ffaa7f9c8cb70f18"),
    },
    "testGlossary.docx": {
        "main": (20117, "e72df38e71a351ebaaf7b102e8ab7862b7e5eb04cc86e187f8510746b70d3f54"),
        "glossary": (16244, "78112f02ff4e0b94a6c99f8688d86c6fa91d54d6c57c504001453b9c73e3dad1"),
    },
}

# The profile prerequisite is the separately reviewed current-source smoke.
# The historical d1/clean-46 smoke remains owned by verify.py/run_smoke.sh and
# is intentionally not accepted as this profile's correctness gate.
CURRENT_SMOKE_COMMIT = "c4ac516353591f88a9b797349002d05737614e78"
CURRENT_SMOKE_RESULT_DIR = "results/smoke-after-d000d977b"
CURRENT_SMOKE_SOURCE_COMMIT = PROFILE_SOURCE_COMMIT

# Compatibility name for profile receipt builders.  Keep the profile source
# pin explicit so a historical smoke source cannot satisfy this gate by
# accident.
SOURCE_COMMIT = PROFILE_SOURCE_COMMIT

PROFILE_SCHEMA = "docx-styles-effects-profile-scaffold-v2"
GENERATED_SCHEMA = "docx-styles-effects-generated-fixtures-v1"
HOST_SCHEMA = "docx-styles-effects-host-v1"
U64_MAX = (1 << 64) - 1
HASH_RE = re.compile(r"^[0-9a-f]{64}$")
PHASES = (
    "capture_ns",
    "snapshot_ns",
    "stage_ns",
    "commit_ns",
    "publish_ns",
    "reopen_ns",
    "inverse_reopen_ns",
    "inverse_ns",
    "projection_ns",
    "opaque_ns",
    "graph_ns",
    "readback_ns",
    "validation_ns",
)
SUBPHASES = (
    "apply_ns",
    "serialize_ns",
    "inverse_apply_ns",
    "inverse_serialize_ns",
)
ALLOCATION_FIELDS = {
    "direct_allocated_bytes",
    "realloc_old_bytes",
    "realloc_new_bytes",
    "deallocated_bytes",
    "requested_alloc_bytes",
    "live_before",
    "live_after",
    "peak_live_delta",
    "allocation_calls",
    "reallocation_calls",
    "deallocation_calls",
    "allocation_failed",
    "alloc_balance_ok",
    "alloc_invalid",
}
METRICS = (
    "parts",
    "total_part_bytes",
    "total_relationships",
    "relationship_parts",
    "relationship_graph_nodes",
    "relationship_xml_bytes",
    "relationship_xml_events",
)
SAMPLE_COUNTERS = (
    "elapsed_ns",
    "output_bytes",
    "direct_allocated_bytes",
    "realloc_old_bytes",
    "realloc_new_bytes",
    "deallocated_bytes",
    "requested_alloc_bytes",
    "live_before",
    "live_after",
    "peak_live_delta",
    "allocation_calls",
    "reallocation_calls",
    "deallocation_calls",
    "allocation_failed",
)
BOOLS = (
    "actual_success",
    "semantic_ok",
    "opaque_ok",
    "exact_inverse_ok",
    "alloc_balance_ok",
)
SAMPLE_FIELDS = {
    "schema",
    "source_commit",
    "lane",
    "smoke_lane",
    "scale",
    "fixture",
    "fixture_native",
    "fixture_signed",
    "fixture_main_present",
    "fixture_glossary_present",
    "fixture_expected_package_sha256",
    "source_backed_api",
    "process_id",
    "process_index",
    "elapsed_ns",
    "phases",
    "apply_allocation",
    "serialize_allocation",
    "inverse_apply_allocation",
    "inverse_serialize_allocation",
    "actual_success",
    "semantic_ok",
    "opaque_ok",
    "exact_inverse_ok",
    "output_bytes",
    "input_sha256",
    "output_sha256",
    "input_member_digest",
    "output_member_digest",
    "input_resource_present",
    "output_resource_present",
    "input_resource",
    "output_resource",
    "input_effects_member",
    "output_effects_member",
    "input_metrics",
    "output_metrics",
    "direct_allocated_bytes",
    "realloc_old_bytes",
    "realloc_new_bytes",
    "deallocated_bytes",
    "requested_alloc_bytes",
    "live_before",
    "live_after",
    "peak_live_delta",
    "allocation_calls",
    "reallocation_calls",
    "deallocation_calls",
    "allocation_failed",
    "alloc_balance_ok",
    "alloc_invalid",
}
SPECS = (
    ("capture_bug_main", "native"),
    ("capture_bug_glossary", "native"),
    ("capture_signed_main", "native"),
    ("capture_complex_main", "native"),
    ("capture_glossary_main", "native"),
    ("capture_glossary_glossary", "native"),
    ("capture_synthetic_main", "64k"),
    ("capture_synthetic_main", "1m"),
    ("capture_synthetic_glossary", "64k"),
    ("capture_synthetic_glossary", "1m"),
    ("projection_main", "native"),
    ("projection_main", "64k"),
    ("projection_main", "1m"),
    ("projection_glossary", "native"),
    ("projection_glossary", "64k"),
    ("projection_glossary", "1m"),
    ("noop_main", "native"),
    ("noop_glossary", "native"),
    ("replace_main", "native"),
    ("replace_main", "64k"),
    ("replace_main", "1m"),
    ("replace_glossary", "native"),
    ("replace_glossary", "64k"),
    ("replace_glossary", "1m"),
    ("remove_main", "native"),
    ("remove_glossary", "native"),
    ("add_main_absent", "native"),
    ("inverse_replace_main", "native"),
    ("inverse_remove_main", "native"),
    ("independent_main", "native"),
    ("independent_glossary", "native"),
)
SPLIT_LANES = {
    "noop_main",
    "noop_glossary",
    "replace_main",
    "replace_glossary",
    "remove_main",
    "remove_glossary",
    "add_main_absent",
    "inverse_replace_main",
    "inverse_remove_main",
    "independent_main",
    "independent_glossary",
}


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def unsigned(value: Any, label: str, *, positive: bool = False) -> int:
    require(type(value) is int, f"{label} must be an integer")
    require(0 <= value <= U64_MAX, f"{label} is outside u64")
    require(not positive or value > 0, f"{label} must be positive")
    return value


def boolean(value: Any, label: str) -> bool:
    require(type(value) is bool, f"{label} must be boolean")
    return value


def text(value: Any, label: str) -> str:
    require(type(value) is str, f"{label} must be text")
    return value


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def hash_value(value: Any, label: str) -> str:
    value = text(value, label)
    require(HASH_RE.fullmatch(value) is not None, f"{label} is not a SHA-256 hex digest")
    return value


def parse_receipt(path: Path) -> dict[str, Any]:
    require(path.is_file(), f"missing profile receipt: {path}")
    value = json.loads(path.read_text())
    require(type(value) is dict, f"profile receipt is not an object: {path}")
    return value


def git_blob_sha256(root: Path, revision: str, relative: str) -> str:
    blob = subprocess.check_output(
        ["git", "-C", str(root), "show", f"{revision}:{relative}"],
    )
    return hashlib.sha256(blob).hexdigest()


def verify_smoke_retention(root: Path, evidence: Path) -> None:
    """Verify the reviewed correctness capture remains replayable in Git.

    The profile is allowed to run only from a descendant of the reviewed
    current-source smoke commit, with its retained 52-lane result subtree
    still bound to the current Git blobs.  The historical smoke has a separate
    verifier contract and cannot silently replace this prerequisite.
    """
    head = subprocess.check_output(["git", "-C", str(root), "rev-parse", "HEAD"], text=True).strip()
    require(
        subprocess.run(
            ["git", "-C", str(root), "merge-base", "--is-ancestor", CURRENT_SMOKE_COMMIT, head],
            check=False,
        ).returncode
        == 0,
        f"profile checkout does not descend from reviewed current-source smoke commit {CURRENT_SMOKE_COMMIT}",
    )
    smoke = evidence / CURRENT_SMOKE_RESULT_DIR
    require(smoke.is_dir(), f"retained smoke result directory is missing: {smoke}")
    verification = parse_receipt(smoke / "root-verification.json")
    require(verification.get("passed") is True, "retained smoke verification did not pass")
    require(
        verification.get("source_commit") == CURRENT_SMOKE_SOURCE_COMMIT,
        "retained current-source smoke source changed",
    )
    receipts = sorted(smoke.glob("smoke-*-p1.json"))
    require(len(receipts) == 52, f"retained smoke receipt count changed: {len(receipts)}")
    for receipt in receipts:
        value = parse_receipt(receipt)
        require(value.get("schema") == "docx-styles-effects-smoke-v1", f"retained smoke schema changed: {receipt.name}")
        require(
            value.get("source_commit") == CURRENT_SMOKE_SOURCE_COMMIT,
            f"retained current-source smoke receipt source changed: {receipt.name}",
        )
        require(isinstance(value.get("lane"), str) and value["lane"], f"retained smoke lane missing: {receipt.name}")
    relative_root = smoke.relative_to(root).as_posix()
    approved_files = subprocess.check_output(
        [
            "git",
            "-C",
            str(root),
            "ls-tree",
            "-r",
            "--name-only",
            "-z",
            CURRENT_SMOKE_COMMIT,
            "--",
            relative_root,
        ],
    ).split(b"\0")
    approved_files = {entry.decode() for entry in approved_files if entry}
    require(approved_files, "approved current-source smoke subtree is empty")
    tracked = subprocess.check_output(
        ["git", "-C", str(root), "ls-files", "-z", "--", relative_root],
    ).split(b"\0")
    tracked = {entry.decode() for entry in tracked if entry}
    require(tracked, "retained smoke subtree is not tracked")
    require(tracked == approved_files, "retained smoke file set changed from approved smoke commit")
    all_paths = list(smoke.rglob("*"))
    require(all(not path.is_symlink() for path in all_paths), "retained smoke subtree contains a symlink")
    disk_files = {path.relative_to(root).as_posix() for path in all_paths if path.is_file()}
    require(disk_files == tracked, "retained smoke file set changed")
    for relative in sorted(approved_files):
        path = root / relative
        require(not path.is_symlink() and path.is_file(), f"retained smoke path is missing or a symlink: {relative}")
        require(
            sha256(path) == git_blob_sha256(root, CURRENT_SMOKE_COMMIT, relative),
            f"retained smoke blob changed from approved smoke commit: {relative}",
        )


def parse_kv(path: Path, label: str) -> dict[str, str]:
    require(path.is_file(), f"missing {label}: {path}")
    values: dict[str, str] = {}
    for line in path.read_text().splitlines():
        if not line:
            continue
        if len(line.split()) == 2 and HASH_RE.fullmatch(line.split()[0]):
            continue
        require("=" in line, f"{label} has an unkeyed line: {line}")
        key, value = line.split("=", 1)
        require(key and key not in values, f"{label} has duplicate key: {key}")
        values[key] = value
    require(values, f"{label} is empty")
    return values


def verify_host(path: Path) -> None:
    values = parse_kv(path, "host probe")
    require(values.get("schema") == HOST_SCHEMA, f"host schema changed: {path}")
    unsigned(int(values["pid"]), "host pid", positive=True)
    for key in ("kernel", "os", "cpu_model", "time_version", "affinity"):
        require(values.get(key), f"host field is empty: {key}")
    if values["logical_cpus"] != "unavailable":
        unsigned(int(values["logical_cpus"]), "logical CPU count", positive=True)
    if values["memory_total_kib"] != "unavailable":
        unsigned(int(values["memory_total_kib"]), "memory total", positive=True)
    if values["memory_available_kib"] != "unavailable":
        unsigned(int(values["memory_available_kib"]), "memory available", positive=True)
    if values["load_1m"] != "unavailable":
        load = float(values["load_1m"])
        require(load >= 0.0, "host load is negative")


def verify_toolchain(path: Path) -> None:
    require(path.is_file(), f"missing toolchain receipt: {path}")
    lines = [line.strip() for line in path.read_text().splitlines() if line.strip()]
    require(any(line.startswith("rustc ") for line in lines), "rustc provenance is missing")
    require(any(line == "rustc -vV" for line in lines), "verbose rustc provenance is missing")
    require(any(line.startswith("cargo ") for line in lines), "cargo provenance is missing")
    require(any(line.startswith("active-toolchain:") for line in lines), "active toolchain provenance is missing")
    target = [line.split("=", 1)[1] for line in lines if line.startswith("target_triple=")]
    linker = [line.split("=", 1)[1] for line in lines if line.startswith("linker=")]
    require(len(target) == 1 and re.fullmatch(r"[A-Za-z0-9_.-]+", target[0]), "target triple provenance is missing")
    require(len(linker) == 1 and linker[0], "linker provenance is missing")


def verify_hash_receipt(path: Path, label: str) -> tuple[str, Path]:
    lines = [line.split() for line in path.read_text().splitlines() if line.strip()]
    require(len(lines) == 1 and len(lines[0]) == 2, f"{label} shape changed")
    digest, shown = lines[0]
    require(HASH_RE.fullmatch(digest) is not None, f"{label} digest is malformed")
    binary = Path(shown)
    require(binary.is_absolute(), f"{label} path is not absolute")
    return digest, binary


def verify_build_receipts(results: Path, target: Path, current_head: str) -> None:
    build_log = results / "build.log"
    require(build_log.is_file() and "Finished " in build_log.read_text(), "build log is incomplete")
    before_digest, before_binary = verify_hash_receipt(results / "binary.sha256", "binary")
    after_digest, after_binary = verify_hash_receipt(results / "binary-after.sha256", "post-build binary")
    require((before_digest, before_binary) == (after_digest, after_binary), "binary changed after profile")
    provenance = parse_kv(results / "build-provenance.txt", "build provenance")
    raw = [line.split() for line in (results / "build-provenance.txt").read_text().splitlines() if len(line.split()) == 2 and HASH_RE.fullmatch(line.split()[0])]
    require(raw == [[before_digest, str(before_binary)]], "build provenance binary digest changed")
    require(
        set(provenance) == {
            "binary",
            "source_commit",
            "git_head",
            "git_status_before_sha256",
            "git_status_after_sha256",
            "rustc",
            "cargo",
            "target",
            "target_triple",
            "linker",
            "toolchain_sha256",
            "environment_sha256",
            "allocator",
            "rss",
            "mode",
        },
        "build provenance keys changed",
    )
    require(provenance["binary"] == str(before_binary), "provenance binary path changed")
    require(provenance["source_commit"] == PROFILE_SOURCE_COMMIT, "provenance source changed")
    require(provenance["git_head"] == current_head, "provenance Git head changed")
    require(provenance["rustc"] and provenance["cargo"], "toolchain provenance is empty")
    target_value = Path(provenance["target"])
    require(target_value.is_absolute() and target_value == target, "target provenance changed")
    require(before_binary.is_relative_to(target_value), "binary is outside the disposable target")
    require(not before_binary.is_symlink(), "profile binary receipt points through a symlink")
    require("CountingAllocator" in provenance["allocator"], "allocator provenance missing")
    require(provenance["rss"] == "/usr/bin/time -v process maximum resident set size", "RSS provenance changed")
    require(provenance["mode"] == "profile processes=3 warmup=2 samples=20", "profile settings changed")
    for name, digest_key in (("git-status-before.txt", "git_status_before_sha256"), ("git-status-after.txt", "git_status_after_sha256"), ("toolchain-before.txt", "toolchain_sha256"), ("environment-before.txt", "environment_sha256")):
        receipt = results / name
        require(receipt.is_file(), f"missing provenance receipt: {receipt}")
        require(sha256(receipt) == provenance[digest_key], f"provenance receipt changed: {name}")
    environment = (results / "environment-before.txt").read_text()
    for marker in ("CARGO_TARGET_DIR=", "CARGO_INCREMENTAL=0", "LC_ALL=C", "DOCX_PROFILE_GENERATED_FIXTURES="):
        require(marker in environment, f"captured environment is missing {marker}")
    require(provenance["git_status_before_sha256"] == provenance["git_status_after_sha256"], "Git status changed during profile")
    require(re.fullmatch(r"[A-Za-z0-9_.-]+", provenance["target_triple"]), "target triple provenance changed")
    require(provenance["linker"], "linker provenance is empty")
    if before_binary.is_file():
        require(sha256(before_binary) == before_digest, "retained binary hash changed")
    else:
        require(not target_value.exists(), "disposable target disappeared only partially")


def verify_time_sidecar(time_path: Path, stderr_path: Path) -> None:
    require(time_path.is_file(), f"missing RSS sidecar: {time_path}")
    require(stderr_path.is_file() and stderr_path.read_bytes() == b"", f"unexpected profile stderr: {stderr_path}")
    content = time_path.read_text()
    statuses = [line.strip() for line in content.splitlines() if line.strip().startswith("Exit status:")]
    require(statuses == ["Exit status: 0"], f"nonzero process status: {time_path}")
    marker = "Maximum resident set size (kbytes):"
    rss = [line.split(":", 1)[1].strip() for line in content.splitlines() if line.lstrip().startswith(marker)]
    require(len(rss) == 1 and rss[0].isdigit() and int(rss[0]) > 0, f"RSS marker invalid: {time_path}")
    user = [line.split(":", 1)[1].strip() for line in content.splitlines() if line.lstrip().startswith("User time (seconds):")]
    system = [line.split(":", 1)[1].strip() for line in content.splitlines() if line.lstrip().startswith("System time (seconds):")]
    elapsed = [line.split("): ", 1)[1].strip() for line in content.splitlines() if line.lstrip().startswith("Elapsed (wall clock) time (h:mm:ss or m:ss):")]
    require(len(user) == 1 and re.fullmatch(r"(?:0|[0-9]+(?:\.[0-9]+)?)", user[0]), f"user-time marker invalid: {time_path}")
    require(len(system) == 1 and re.fullmatch(r"(?:0|[0-9]+(?:\.[0-9]+)?)", system[0]), f"system-time marker invalid: {time_path}")
    require(len(elapsed) == 1 and re.fullmatch(r"(?:[0-9]+:)?[0-9]+:[0-9]+(?:\.[0-9]+)?", elapsed[0]), f"elapsed-time marker invalid: {time_path}")


def verify_metrics(value: Any, label: str) -> None:
    require(type(value) is dict, f"{label} must be an object")
    require(set(value) == set(METRICS), f"{label} fields changed")
    for field in METRICS:
        unsigned(value[field], f"{label}.{field}")


def verify_resource(value: Any, label: str, expected_owner: str) -> None:
    require(type(value) is dict, f"{label} must be an object")
    require(
        set(value)
        == {
            "owner",
            "member",
            "resource_bytes",
            "resource_sha256",
            "conformance",
            "style_count",
            "xml_events",
            "xml_depth",
        },
        f"{label} fields changed",
    )
    require(value["owner"] == expected_owner, f"{label} owner changed")
    expected_member = (
        "word/glossary/stylesWithEffects.xml"
        if expected_owner == "glossary"
        else "word/stylesWithEffects.xml"
    )
    require(value["member"] == expected_member, f"{label} member changed")
    unsigned(value["resource_bytes"], f"{label}.resource_bytes", positive=True)
    hash_value(value["resource_sha256"], f"{label}.resource_sha256")
    require(value["conformance"] in {"transitional", "strict"}, f"{label} conformance is unknown")
    unsigned(value["style_count"], f"{label}.style_count", positive=True)
    unsigned(value["xml_events"], f"{label}.xml_events", positive=True)
    unsigned(value["xml_depth"], f"{label}.xml_depth", positive=True)


def verify_member_digest(value: Any, label: str, *, required: bool) -> None:
    if value is None:
        require(not required, f"{label} is missing")
        return
    require(type(value) is dict and set(value) == {"bytes", "sha256"}, f"{label} fields changed")
    unsigned(value["bytes"], f"{label}.bytes", positive=True)
    hash_value(value["sha256"], f"{label}.sha256")


def verify_allocation_delta(value: Any, label: str) -> None:
    require(type(value) is dict and set(value) == ALLOCATION_FIELDS, f"{label} fields changed")
    counters = (
        "direct_allocated_bytes",
        "realloc_old_bytes",
        "realloc_new_bytes",
        "deallocated_bytes",
        "requested_alloc_bytes",
        "live_before",
        "live_after",
        "peak_live_delta",
        "allocation_calls",
        "reallocation_calls",
        "deallocation_calls",
        "allocation_failed",
    )
    for field in counters:
        unsigned(value[field], f"{label}.{field}")
    require(value["allocation_failed"] == 0, f"{label} allocator reported failure")
    require(boolean(value["alloc_balance_ok"], f"{label}.alloc_balance_ok"), f"{label} allocation equation failed")
    require(not boolean(value["alloc_invalid"], f"{label}.alloc_invalid"), f"{label} allocator underflow flag set")
    require(
        value["live_before"] + value["direct_allocated_bytes"] + value["realloc_new_bytes"]
        == value["live_after"] + value["realloc_old_bytes"] + value["deallocated_bytes"],
        f"{label} live-byte equation does not balance",
    )
    require(
        value["requested_alloc_bytes"]
        == value["direct_allocated_bytes"] + value["realloc_new_bytes"],
        f"{label} requested allocation total changed",
    )
    require(
        value["peak_live_delta"] >= max(0, value["live_after"] - value["live_before"]),
        f"{label} peak live delta is impossible",
    )


def verify_sample(sample: dict[str, Any], lane: str, scale: str, expected_process: int, generated: dict[tuple[str, str], dict[str, Any]]) -> None:
    require(set(sample) == SAMPLE_FIELDS, "sample fields changed")
    require(sample.get("schema") == PROFILE_SCHEMA, "sample schema changed")
    require(sample.get("source_commit") == PROFILE_SOURCE_COMMIT, "sample source changed")
    require(sample.get("lane") == lane and sample.get("scale") == scale, "sample lane/scale changed")
    for field in ("fixture_native", "fixture_signed", "fixture_main_present", "fixture_glossary_present"):
        boolean(sample.get(field), field)
    unsigned(sample.get("process_id"), "sample process id", positive=True)
    process_index = unsigned(sample.get("process_index"), "sample process index", positive=True)
    require(process_index == expected_process, "sample process index changed")
    require(boolean(sample.get("source_backed_api"), "source_backed_api"), "sample is not source-backed")
    for field in BOOLS:
        require(boolean(sample.get(field), field), f"{field} is false")
    for field in SAMPLE_COUNTERS:
        unsigned(sample.get(field), field, positive=field == "elapsed_ns")
    require(sample["allocation_failed"] == 0, "allocator reported failure")
    require(boolean(sample["alloc_invalid"], "alloc_invalid") is False, "allocator underflow flag set")
    phases = sample.get("phases")
    require(type(phases) is dict and set(phases) == set(PHASES) | set(SUBPHASES), "phase fields changed")
    phase_values = [unsigned(phases[field], f"phases.{field}") for field in PHASES]
    require(sum(phase_values) <= sample["elapsed_ns"], "disjoint phases exceed elapsed")
    for field in SUBPHASES:
        unsigned(phases[field], f"phases.{field}")
    require(
        phases["apply_ns"] + phases["serialize_ns"] <= phases["publish_ns"],
        "forward attribution slices exceed publication clock",
    )
    require(
        phases["inverse_apply_ns"] + phases["inverse_serialize_ns"] <= phases["inverse_ns"],
        "inverse attribution slices exceed inverse clock",
    )
    require(
        sample["live_before"] + sample["direct_allocated_bytes"] + sample["realloc_new_bytes"]
        == sample["live_after"] + sample["realloc_old_bytes"] + sample["deallocated_bytes"],
        "allocator live-byte equation does not balance",
    )
    require(sample["requested_alloc_bytes"] == sample["direct_allocated_bytes"] + sample["realloc_new_bytes"], "requested allocation total changed")
    require(sample["peak_live_delta"] >= max(0, sample["live_after"] - sample["live_before"]), "peak live delta is impossible")
    for field in (
        "apply_allocation",
        "serialize_allocation",
        "inverse_apply_allocation",
        "inverse_serialize_allocation",
    ):
        value = sample.get(field)
        if value is not None:
            verify_allocation_delta(value, field)
    if lane in SPLIT_LANES:
        require(sample["apply_allocation"] is not None, f"apply attribution missing for {lane}")
        require(sample["serialize_allocation"] is not None, f"serialize attribution missing for {lane}")
    else:
        require(sample["apply_allocation"] is None, f"unexpected apply attribution for {lane}")
        require(sample["serialize_allocation"] is None, f"unexpected serialize attribution for {lane}")
    if lane.startswith("inverse_"):
        require(sample["inverse_apply_allocation"] is not None, f"inverse apply attribution missing for {lane}")
        require(sample["inverse_serialize_allocation"] is not None, f"inverse serialize attribution missing for {lane}")
    else:
        require(sample["inverse_apply_allocation"] is None, f"unexpected inverse apply attribution for {lane}")
        require(sample["inverse_serialize_allocation"] is None, f"unexpected inverse serialize attribution for {lane}")
    for field in ("input_sha256", "input_member_digest", "output_member_digest"):
        hash_value(sample.get(field), field)
    hash_value(sample.get("output_sha256"), "output_sha256")
    verify_metrics(sample.get("input_metrics"), "input_metrics")
    verify_metrics(sample.get("output_metrics"), "output_metrics")
    require(sample["output_bytes"] > 0, "output package is empty")
    owner = "glossary" if lane.endswith("glossary") else "main"
    boolean(sample["input_resource_present"], "input_resource_present")
    boolean(sample["output_resource_present"], "output_resource_present")
    require(sample["input_resource_present"] == (sample["input_resource"] is not None), "input resource presence disagrees with metadata")
    require(sample["output_resource_present"] == (sample["output_resource"] is not None), "output resource presence disagrees with metadata")
    if lane == "add_main_absent":
        require(not sample["input_resource_present"], "absent-owner input unexpectedly has a resource")
    else:
        verify_resource(sample["input_resource"], "input_resource", owner)
    if lane in {"remove_main", "remove_glossary"}:
        require(not sample["output_resource_present"], "removed owner still has output metadata")
    else:
        require(sample["output_resource_present"], f"successful output metadata is missing for {lane}")
    if sample["output_resource_present"]:
        verify_resource(sample["output_resource"], "output_resource", owner)
    verify_member_digest(sample["input_effects_member"], "input_effects_member", required=lane != "add_main_absent")
    verify_member_digest(sample["output_effects_member"], "output_effects_member", required=lane not in {"remove_main", "remove_glossary"})
    if sample["input_resource"] is not None:
        require(sample["input_resource"]["resource_bytes"] == sample["input_effects_member"]["bytes"], f"input resource/member size changed for {lane}")
        require(sample["input_resource"]["resource_sha256"] == sample["input_effects_member"]["sha256"], f"input resource/member hash changed for {lane}")
    if sample["output_resource_present"]:
        require(sample["output_resource"]["resource_bytes"] == sample["output_effects_member"]["bytes"], f"output resource/member size changed for {lane}")
        require(sample["output_resource"]["resource_sha256"] == sample["output_effects_member"]["sha256"], f"output resource/member hash changed for {lane}")
    if lane in {"capture_bug_main", "capture_bug_glossary", "capture_signed_main", "capture_complex_main", "capture_glossary_main", "capture_glossary_glossary", "capture_synthetic_main", "capture_synthetic_glossary", "projection_main", "projection_glossary", "noop_main", "noop_glossary", "inverse_replace_main", "inverse_remove_main"}:
        require(sample["output_sha256"] == sample["input_sha256"], f"unchanged package hash changed for {lane}")
        require(sample["output_resource_present"], f"unchanged output resource missing for {lane}")
        require(sample["output_resource"] == sample["input_resource"], f"unchanged output resource changed for {lane}")
    if lane.startswith("inverse_"):
        require(sample["exact_inverse_ok"], "inverse lane lost exact inverse")
        require(sample["output_sha256"] == sample["input_sha256"], "inverse package hash changed")
    if scale == "native" and sample["fixture_native"]:
        fixture = text(sample["fixture"], "fixture")
        require(fixture in NATIVE_FIXTURES, f"unknown native fixture: {fixture}")
        require(sample["input_sha256"] == NATIVE_FIXTURES[fixture]["sha256"], f"native input hash changed for {lane}")
        expected_package = sample["fixture_expected_package_sha256"]
        require(expected_package == NATIVE_FIXTURES[fixture]["sha256"], f"native expected package hash changed for {lane}")
        expected_member = NATIVE_EFFECTS_MEMBERS[fixture].get(owner)
        require(expected_member is not None, f"native effects member is absent for {lane}")
        require(sample["input_resource"]["resource_bytes"] == expected_member[0], f"native effects size changed for {lane}")
        require(sample["input_resource"]["resource_sha256"] == expected_member[1], f"native effects hash changed for {lane}")
    elif scale == "native":
        require(lane == "add_main_absent", f"native lane is not bound to a native fixture: {lane}")
        require(sample["fixture_expected_package_sha256"] is None, "derived absent-owner fixture unexpectedly has a package hash")
    else:
        require(not sample["fixture_native"], f"synthetic lane unexpectedly uses a native fixture: {lane}")
        require(sample["fixture_expected_package_sha256"] is None, f"synthetic expected package hash is not null for {lane}")
        generated_row = generated.get((owner, scale))
        require(generated_row is not None, f"generated fixture missing for {owner}/{scale}")
        require(sample["input_sha256"] == generated_row["package_sha256"], f"synthetic input hash changed for {lane}/{scale}")
        input_resource = sample["input_resource"]
        require(input_resource is not None, f"synthetic resource missing for {lane}/{scale}")
        require(input_resource["resource_sha256"] == generated_row["resource_sha256"], f"synthetic resource hash changed for {lane}/{scale}")
        require(input_resource["resource_bytes"] == generated_row["resource_bytes"], f"synthetic resource size changed for {lane}/{scale}")


def verify_generated(path: Path) -> dict[tuple[str, str], dict[str, Any]]:
    value = parse_receipt(path)
    require(value.get("schema") == GENERATED_SCHEMA, "generated fixture schema changed")
    require(value.get("source_commit") == PROFILE_SOURCE_COMMIT, "generated fixture source changed")
    rows = value.get("fixtures")
    require(type(rows) is list and len(rows) == 4, "generated fixture count changed")
    result: dict[tuple[str, str], dict[str, Any]] = {}
    for row in rows:
        require(type(row) is dict, "generated fixture row is not an object")
        require(row.get("schema") == "docx-styles-effects-generated-fixture-v1", "generated fixture row schema changed")
        owner = text(row.get("owner"), "generated owner")
        scale = text(row.get("scale"), "generated scale")
        require(owner in {"main", "glossary"} and scale in {"64k", "1m"}, "generated owner/scale changed")
        key = (owner, scale)
        require(key not in result, f"duplicate generated fixture: {key}")
        result[key] = row
        resource_bytes = unsigned(row.get("resource_bytes"), "generated resource bytes", positive=True)
        target = 64 * 1024 if scale == "64k" else 1024 * 1024
        require(abs(resource_bytes - target) <= 1024, f"generated resource scale drifted: {key}")
        unsigned(row.get("package_bytes"), "generated package bytes", positive=True)
        for field in ("resource_sha256", "package_sha256"):
            hash_value(row.get(field), f"generated {field}")
        unsigned(row.get("xml_events"), "generated XML events", positive=True)
        unsigned(row.get("xml_depth"), "generated XML depth", positive=True)
        unsigned(row.get("style_count"), "generated style count", positive=True)
        unsigned(row.get("opaque_marker_bytes"), "generated opaque marker bytes", positive=True)
        require(text(row.get("member"), "generated member").endswith("stylesWithEffects.xml"), "generated member changed")
        generated_dir_raw = path.parent / "generated-fixtures"
        require(not generated_dir_raw.is_symlink(), "generated fixture directory is a symlink")
        generated_dir = generated_dir_raw.resolve()
        resource_raw = Path(text(row.get("resource_path"), "generated resource path"))
        package_raw = Path(text(row.get("package_path"), "generated package path"))
        resource_path = resource_raw.resolve()
        package_path = package_raw.resolve()
        require(resource_path.is_file() and package_path.is_file(), f"generated fixture files are missing: {key}")
        require(not resource_raw.is_symlink() and not package_raw.is_symlink(), f"generated fixture path is a symlink: {key}")
        require(resource_path.is_relative_to(generated_dir), f"generated resource escaped its result directory: {key}")
        require(package_path.is_relative_to(generated_dir), f"generated package escaped its result directory: {key}")
        require(resource_path != package_path, f"generated resource/package paths overlap: {key}")
        require(sha256(resource_path) == row["resource_sha256"], f"generated resource bytes changed: {key}")
        require(sha256(package_path) == row["package_sha256"], f"generated package bytes changed: {key}")
        require(resource_path.stat().st_size == resource_bytes, f"generated resource size changed: {key}")
        with zipfile.ZipFile(package_path) as archive:
            member = text(row.get("member"), "generated member")
            require(archive.read(member) == resource_path.read_bytes(), f"generated package/member mismatch: {key}")
    require(set(result) == {("main", "64k"), ("main", "1m"), ("glossary", "64k"), ("glossary", "1m")}, "generated fixture matrix changed")
    return result


def verify_commands(path: Path) -> None:
    lines = path.read_text().splitlines()
    for prefix in (
        "command-verify-smoke=",
        "command-host-before=",
        "command-toolchain-before=",
        "command-cargo-metadata-before=",
        "command-source-manifest-before=",
        "command-cargo-build=",
        "command-fixture-manifest=",
        "command-cargo-metadata-after=",
        "command-source-manifest-after=",
        "command-host-after=",
        "command-verify-profile=",
    ):
        require(sum(line.startswith(prefix) for line in lines) == 1, f"missing or duplicate command receipt: {prefix}")
    environments = [line for line in lines if line.startswith("environment=")]
    require(len(environments) == 1, "profile environment receipt is missing")
    for marker in (
        "CARGO_TARGET_DIR=",
        "CARGO_INCREMENTAL=0",
        "LC_ALL=C",
        "DOCX_PROFILE_GENERATED_FIXTURES=",
        "RUSTFLAGS=unset",
        "CARGO_ENCODED_RUSTFLAGS=unset",
        "RUSTC_BOOTSTRAP=unset",
        "RUSTDOCFLAGS=unset",
    ):
        require(marker in environments[0], f"profile environment is missing {marker}")
    fixture_commands = [line for line in lines if line.startswith("command-fixture-manifest=")]
    require(len(fixture_commands) == 1 and "--emit-fixture-manifest" in fixture_commands[0] and "--output-dir" in fixture_commands[0], "generated fixture command is incomplete")
    runs = [line for line in lines if line.startswith("run=")]
    require(len(runs) == len(SPECS) * 3, "profile command count changed")
    expected = {(lane, scale, process) for process in range(1, 4) for lane, scale in SPECS}
    observed: set[tuple[str, str, int]] = set()
    for line in runs:
        lane_match = re.search(r"--lane\s+([A-Za-z0-9_]+)", line)
        scale_match = re.search(r"--scale\s+([A-Za-z0-9]+)", line)
        process_match = re.search(r"process_index=(\d+)", line)
        require(lane_match and scale_match and process_match, f"profile command is incomplete: {line}")
        lane, scale, process = lane_match.group(1), scale_match.group(1), int(process_match.group(1))
        require((lane, scale, process) in expected, f"unexpected profile command: {line}")
        require("--warmup 2 --samples 20" in line and "fresh_process=1" in line, f"profile settings missing: {line}")
        key = (lane, scale, process)
        require(key not in observed, f"duplicate profile command: {key}")
        observed.add(key)
    require(observed == expected, "profile command matrix changed")


def verify_external_paths(root: Path, evidence: Path, results: Path, target: Path) -> None:
    root = root.resolve()
    evidence = evidence.resolve()
    results = results.resolve()
    target = target.resolve()
    require(results.is_absolute() and target.is_absolute(), "profile replay paths must be absolute")
    require(not results.is_relative_to(root) and not target.is_relative_to(root), "profile replay path is inside checkout")
    require(not results.is_relative_to(evidence) and not target.is_relative_to(evidence), "profile replay path is inside evidence tree")
    if target.exists():
        require(not target.is_symlink(), "profile target replay path is a symlink")
    require(results != target, "profile results and target paths overlap")
    require(not results.is_relative_to(target) and not target.is_relative_to(results), "profile results and target paths are nested")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", type=Path, required=True)
    parser.add_argument("--evidence", type=Path, required=True)
    parser.add_argument("--smoke-only", action="store_true")
    parser.add_argument("--results", type=Path)
    parser.add_argument("--manifest-before", type=Path)
    parser.add_argument("--manifest-after", type=Path)
    parser.add_argument("--metadata-before", type=Path)
    parser.add_argument("--metadata-after", type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    root = args.root.resolve()
    evidence = args.evidence.resolve()
    if args.smoke_only:
        verify_smoke_retention(root, evidence)
        return
    for argument, label in (
        (args.results, "--results"),
        (args.manifest_before, "--manifest-before"),
        (args.manifest_after, "--manifest-after"),
        (args.metadata_before, "--metadata-before"),
        (args.metadata_after, "--metadata-after"),
        (args.output, "--output"),
    ):
        require(argument is not None, f"{label} is required unless --smoke-only is used")
    assert args.results is not None
    assert args.manifest_before is not None
    assert args.manifest_after is not None
    assert args.metadata_before is not None
    assert args.metadata_after is not None
    assert args.output is not None
    require(not args.results.is_symlink(), "profile results replay path is a symlink")
    results = args.results.resolve()
    for receipt_path, label in (
        (args.manifest_before, "manifest-before"),
        (args.manifest_after, "manifest-after"),
        (args.metadata_before, "metadata-before"),
        (args.metadata_after, "metadata-after"),
        (args.output, "verification output"),
    ):
        require(not receipt_path.is_symlink(), f"{label} is a symlink")
        require(receipt_path.resolve().is_relative_to(results), f"{label} escaped profile results")
    require(results.is_dir(), "profile results directory is missing")
    current_head = subprocess.check_output(["git", "-C", str(root), "rev-parse", "HEAD"], text=True).strip()
    require(
        subprocess.run(
            ["git", "-C", str(root), "merge-base", "--is-ancestor", PROFILE_SOURCE_COMMIT, current_head],
            check=False,
        ).returncode
        == 0,
        "checkout does not descend from approved source",
    )
    verify_smoke_retention(root, evidence)
    status = subprocess.check_output(["git", "-C", str(root), "status", "--porcelain", "--untracked-files=all"], text=True)
    require(status == (results / "git-status-before.txt").read_text(), "current Git status differs from captured before status")
    require(status == (results / "git-status-after.txt").read_text(), "current Git status differs from captured after status")
    manifest_summary = verify_manifest(
        args.manifest_before,
        root,
        args.metadata_before,
        evidence,
        source_commit=PROFILE_SOURCE_COMMIT,
    )
    require(args.manifest_after.read_text() == args.manifest_before.read_text(), "source closure changed")
    require(sha256(args.metadata_before) == sha256(args.metadata_after), "Cargo metadata changed")
    source = parse_kv(results / "source-provenance.txt", "source provenance")
    require(set(source) == {"source_manifest_sha256", "metadata_before_sha256", "metadata_after_sha256", "git_head", "git_status_before_sha256", "git_status_after_sha256"}, "source provenance keys changed")
    require(source["source_manifest_sha256"] == sha256(args.manifest_before), "source manifest provenance changed")
    require(source["metadata_before_sha256"] == sha256(args.metadata_before), "metadata-before provenance changed")
    require(source["metadata_after_sha256"] == sha256(args.metadata_after), "metadata-after provenance changed")
    require(source["git_head"] == current_head, "source provenance Git head changed")
    require(source["git_status_before_sha256"] == sha256(results / "git-status-before.txt"), "source before-status provenance changed")
    require(source["git_status_after_sha256"] == sha256(results / "git-status-after.txt"), "source after-status provenance changed")
    generated = verify_generated(results / "generated-fixtures.json")
    target = Path(parse_kv(results / "build-provenance.txt", "build provenance")["target"])
    verify_external_paths(root, evidence, results, target)
    verify_build_receipts(results, target, current_head)
    verify_host(results / "host-before.txt")
    verify_host(results / "host-after.txt")
    verify_toolchain(results / "toolchain-before.txt")
    verify_commands(results / "commands.txt")
    for process in range(1, 4):
        for lane, scale in SPECS:
            stem = f"profile-{lane}-{scale}-p{process}"
            path = results / f"{stem}.json"
            value = parse_receipt(path)
            require(value.get("schema") == PROFILE_SCHEMA, f"receipt schema changed: {path}")
            require(value.get("source_commit") == PROFILE_SOURCE_COMMIT, f"receipt source changed: {path}")
            require(value.get("lane") == lane and value.get("scale") == scale, f"receipt identity changed: {path}")
            require(unsigned(value.get("warmup"), "receipt warmup") == 2, f"receipt warmup changed: {path}")
            require(unsigned(value.get("sample_count"), "receipt sample count") == 20, f"receipt sample count changed: {path}")
            samples = value.get("samples")
            require(type(samples) is list and len(samples) == 20, f"sample count changed: {path}")
            for sample in samples:
                require(type(sample) is dict, f"sample is not an object: {path}")
                verify_sample(sample, lane, scale, process, generated)
            verify_time_sidecar(results / f"{stem}.time.txt", results / f"{stem}.stderr.log")
    output = {
        "schema": "docx-styles-effects-profile-verification-v1",
        "passed": True,
        "timed_claim": "bounded absolute observations only; no before-after or speedup claim",
        "source_commit": PROFILE_SOURCE_COMMIT,
        "git_head": current_head,
        "manifest": manifest_summary,
        "processes": 3,
        "warmup": 2,
        "samples_per_process": 20,
        "specs": len(SPECS),
        "total_samples": len(SPECS) * 3 * 20,
        "rss": "one /usr/bin/time -v process maximum marker per fresh process",
    }
    args.output.write_text(json.dumps(output, indent=2, sort_keys=True) + "\n")


if __name__ == "__main__":
    try:
        main()
    except (AssertionError, KeyError, ValueError, OSError, json.JSONDecodeError) as error:
        print(f"profile verification failed: {error}", file=sys.stderr)
        raise SystemExit(1)
