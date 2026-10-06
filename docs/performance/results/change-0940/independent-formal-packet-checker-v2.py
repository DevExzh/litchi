#!/usr/bin/env python3
"""Independent terminal-first checker for the frozen 0940 formal packet.

This checker validates the authorized manifest, source closure, release binaries,
outer capture terminal, all 80 child terminals and reports, and a moved packet
copy.  Recorded child paths remain bound to the original packet; only physical
reads are mapped to the moved packet root.  It performs no workload execution.
"""
from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Mapping

ROOT = Path("/home/zhuhe/litchi-performance-0940")
REPO = Path("/home/zhuhe/code/litchi")
MANIFEST_PATH = ROOT / "manifests/formal-manifest-authorized-v2.json"
ORIGINAL_PACKET = ROOT / "formal-packet-v1"
MOVED_PACKET = ROOT / "formal-packet-v1-moved"
OUTER_TERMINAL = ROOT / "receipts/formal-outer-terminal-v2.json"
VALIDATOR_PATH = ROOT / "docx_owned_attribution_validator_v3.py"
RUNNER_PATH = REPO / "docs/performance/results/change-0940/comparison-runner.py"
HELPER_PATH = REPO / "tools/perf_child_validation.py"
COMPARISON_PATH = REPO / "docs/performance/results/change-0940/comparison.py"
TASKSET_PATH = Path("/usr/bin/taskset")
AUTHORIZED_MANIFEST_SHA256 = "14c6c32810f03017711e2e4d86e30a984c35834d3eb649cba94224da92d250f4"
RUNNER_SHA256 = "142c896d22290c1325548e6c46728de12de834a8c88d43e4cb536d15f4cc264b"
VALIDATOR_SHA256 = "7592d51124edd46a15d664594823cf3fbd2f6a480989e3462cb695b3a88c486f"
INDEPENDENT_PREFLIGHT_SHA256 = "fe030a34b41eb922c6237ec4620062fc94430a2329a2d39216fa41f4dfda361b"
HELPER_SHA256 = "198223e4231fe62c503f60010b7867724aa9d8311f8067a14b36ed56eb69667a"
COMPARISON_SHA256 = "c332d90b6e2a542adc24ea97942d0d2b43705b34f9f08d0c28511c5ed8d2745a"
TASKSET_SHA256 = "f3ae0ca36e0a8a892fee6ed3f72f6058fc07991e49b03fe6d1f796989d8777c2"
PLAN_SHA256 = "a653a45f4f15110df6b1ff2fd175a428df677fc9702b02750813df4227fa2417"
SCHEDULE_SHA256 = "9ba2e96533a271ff89d38351c321990d914cdc8387ecb333589e65a039560b8a"
SOURCE_REVISION = "3ae856d76424f2f9b58d26dcd20fd074b8e97541"
SOURCE_ROOT = ROOT / "source-candidate-v1"
SOURCE_LISTING = ROOT / "source-candidate-manifest-v1.jsonl"
SOURCE_LISTING_SHA256 = "8eb16940342230e36c4c121a71a49675e30b77b67a52f73e18b977fa1a100f1f"
SOURCE_RECEIPT = ROOT / "receipts/source-candidate-v1.json"
SOURCE_RECEIPT_SHA256 = "83317ba38fcd14a407bff85671624e353ab3a8ccbc70ffadecad4ef1537fe74e"
SOURCE_FILE_COUNT = 10780
SOURCE_BYTES = 260009611
SOURCE_CORPUS_BYTES = 2101427
SOURCE_CORPUS_SHA256 = "183204a832a3518bf1d4474e9b7ff347ad7d0dc95ad2e9b1dbdde405837a29e8"
OUTPUT_BYTES = 2101446
OUTPUT_SHA256 = "5e87e9ca9fd6b9a98933c36d9aee1e848bc97eb7693229f7949514830121c4ec"
TIMING_SCOPE = (
    "operation_elapsed_ns includes prep, open, and workload phases; "
    "wall_elapsed_ns is the measured outer bracket from before-prep observation "
    "through the selected body end (publication for workload, post-open observation for calibration)"
)
OBSERVER_SCOPE = (
    "observer_gap_ns starts after prep's first Region::split and ends after the "
    "second split; allocation observer_gap contains both procfs reads, allocation "
    "combined ends before after-open observation, process combined includes after-open "
    "observation, and wall includes the full outer bracket"
)
MAX_REPORT_BYTES = 1 << 20
ENVIRONMENT = {"PATH": "/usr/bin:/bin", "LC_ALL": "C", "LANG": "C", "TZ": "UTC", "RUST_BACKTRACE": "0"}
ARM_SHORT = {
    "normal_probes_off_reference": "normal_off",
    "normal_probes_on_diagnostic": "normal_on",
    "allocator_probes_on_diagnostic": "allocator_on",
}
EXPECTED_BINARIES = {
    "normal": {
        "path": str(ROOT / "target-normal-v1/release/litchi-perf-baseline"),
        "bytes": 70076392,
        "sha256": "c735995c902d410b68d43348d169535728d01df6efbf1169263c2345b8666e92",
        "identity": "litchi-perf-baseline",
    },
    "allocator": {
        "path": str(ROOT / "target-allocator-v1/release/litchi-perf-baseline-alloc"),
        "bytes": 70036128,
        "sha256": "769b9bf2df27ed27f1b509c1c8abfca1d80633fccb418420e832cc052f946b4e",
        "identity": "litchi-perf-baseline-alloc",
    },
}
ARMS = (
    "normal_probes_off_reference",
    "normal_probes_on_diagnostic",
    "allocator_probes_on_diagnostic",
)
ARM_PERMUTATIONS = (
    ARMS,
    (ARMS[0], ARMS[2], ARMS[1]),
    (ARMS[1], ARMS[0], ARMS[2]),
    (ARMS[1], ARMS[2], ARMS[0]),
    (ARMS[2], ARMS[0], ARMS[1]),
    (ARMS[2], ARMS[1], ARMS[0]),
)
ROUTES = ("owned_vec", "owned_reader")
HEX = set("0123456789abcdef")


class CheckError(RuntimeError):
    pass


def need(condition: bool, message: str) -> None:
    if not condition:
        raise CheckError(message)


def digest_file(path: Path) -> dict[str, Any]:
    need(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            size += len(block)
            digest.update(block)
    return {"path": str(path.resolve()), "bytes": size, "sha256": digest.hexdigest()}


def digest_raw(path: Path) -> str:
    return digest_file(path)["sha256"]


def canonical(value: Any) -> str:
    return hashlib.sha256(json.dumps(value, sort_keys=True, separators=(",", ":")).encode()).hexdigest()


def read_json(path: Path, role: str, *, max_bytes: int | None = None) -> tuple[dict[str, Any], dict[str, Any]]:
    desc = digest_file(path)
    if max_bytes is not None:
        need(desc["bytes"] < max_bytes, f"{role} is not below bound")
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except Exception as error:
        raise CheckError(f"{role} is not valid JSON: {error}") from error
    need(isinstance(value, dict), f"{role} root is not an object")
    return value, desc


def exact_desc(value: Any, expected_path: Path, actual_path: Path, role: str, *, max_bytes: int | None = None) -> dict[str, Any]:
    need(isinstance(value, Mapping), f"{role} descriptor missing")
    recorded = value.get("path")
    need(isinstance(recorded, str), f"{role} recorded path missing")
    # The runner records absolute original-packet paths. A moved packet may only
    # change the physical read root; its recorded path must remain original.
    need(recorded == str(expected_path), f"{role} recorded path changed")
    actual = digest_file(actual_path)
    need(value.get("bytes") == actual["bytes"], f"{role} byte count changed")
    need(value.get("sha256") == actual["sha256"], f"{role} digest changed")
    if max_bytes is not None:
        need(actual["bytes"] < max_bytes, f"{role} exceeds size bound")
    return {"path": expected_path.relative_to(ORIGINAL_PACKET).as_posix(), "bytes": actual["bytes"], "sha256": actual["sha256"]}


def external_desc(value: Any, expected_path: Path, role: str, expected_sha: str | None = None) -> dict[str, Any]:
    need(isinstance(value, Mapping), f"{role} descriptor missing")
    need(Path(str(value.get("path", ""))).resolve() == expected_path.resolve(), f"{role} path changed")
    actual = digest_file(expected_path)
    need(value.get("bytes") == actual["bytes"] and value.get("sha256") == actual["sha256"], f"{role} digest changed")
    if expected_sha is not None:
        need(actual["sha256"] == expected_sha, f"{role} is not frozen")
    return actual


def expected_schedule() -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for block in range(1, 13):
        if block == 1:
            for position, arm in enumerate(("normal_probes_on_diagnostic", "allocator_probes_on_diagnostic"), 1):
                rows.append({
                    "arm": arm, "boundary_after_block": 0, "cpu": 12,
                    "deadline_seconds": 900, "id": f"calibration-after-00-{arm}",
                    "kind": "calibration", "position": position, "route": "route_independent",
                    "samples": 50, "warmups": 5,
                })
        routes = ROUTES if block % 2 else tuple(reversed(ROUTES))
        for arm in ARM_PERMUTATIONS[(block - 1) % len(ARM_PERMUTATIONS)]:
            for position, route in enumerate(routes, 1):
                rows.append({
                    "arm": arm, "block": block, "cpu": 12, "deadline_seconds": 900,
                    "id": f"workload-block-{block:02d}-{arm}-{position:02d}-{route}",
                    "kind": "workload", "position": position, "route": route,
                    "samples": 50, "warmups": 5,
                })
        if block in (4, 8, 12):
            for position, arm in enumerate(("normal_probes_on_diagnostic", "allocator_probes_on_diagnostic"), 1):
                rows.append({
                    "arm": arm, "boundary_after_block": block, "cpu": 12,
                    "deadline_seconds": 900, "id": f"calibration-after-{block:02d}-{arm}",
                    "kind": "calibration", "position": position, "route": "route_independent",
                    "samples": 50, "warmups": 5,
                })
    need(len(rows) == 80, "internal schedule cardinality changed")
    return rows


def expected_report(arm: str) -> dict[str, Any]:
    if arm == "normal_probes_off_reference":
        binary, enabled, allocator, instrumentation, revision = (
            "litchi-perf-baseline", False, "Rust system allocator", "none", None
        )
    elif arm == "normal_probes_on_diagnostic":
        binary, enabled, allocator, instrumentation, revision = (
            "litchi-perf-baseline", True, "Rust system allocator", "none", None
        )
    else:
        binary, enabled, allocator, instrumentation, revision = (
            "litchi-perf-baseline-alloc", True,
            "CountingSystemAllocator(std::alloc::System)",
            "system_allocator_operation_scoped", "serialized_region_peak_v3"
        )
    return {
        "source_revision": SOURCE_REVISION,
        "binary_identity": binary,
        "diagnostic_arm": ARM_SHORT[arm],
        "diagnostic_enabled": enabled,
        "allocator_identity": allocator,
        "instrumentation_identity": instrumentation,
        "allocator_counter_revision": revision,
        "timing_scope": TIMING_SCOPE,
        "observer_gap_scope": OBSERVER_SCOPE,
        "expected_allocation_present": enabled,
        "expected_process_present": enabled,
    }


def report_fields(arm: str) -> dict[str, Any]:
    return {
        key: value
        for key, value in expected_report(arm).items()
        if key not in ("expected_allocation_present", "expected_process_present")
    }


def load_validator():
    need(digest_raw(VALIDATOR_PATH) == VALIDATOR_SHA256, "validator source hash changed")
    spec = importlib.util.spec_from_file_location("frozen_docx_owned_validator", VALIDATOR_PATH)
    need(spec is not None and spec.loader is not None, "validator import failed")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def validate_source(manifest: Mapping[str, Any]) -> dict[str, Any]:
    source = manifest.get("source")
    need(isinstance(source, Mapping), "manifest source missing")
    need(source.get("source_root") == str(SOURCE_ROOT), "source root changed")
    need(source.get("source_revision") == SOURCE_REVISION, "source revision changed")
    external_desc(source.get("full_source_listing"), SOURCE_LISTING, "source listing", SOURCE_LISTING_SHA256)
    need(source.get("full_source_listing_sha256") == SOURCE_LISTING_SHA256, "source listing binding changed")
    external_desc(source.get("source_receipt"), SOURCE_RECEIPT, "source receipt", SOURCE_RECEIPT_SHA256)
    listing: dict[str, tuple[int, str]] = {}
    total = 0
    with SOURCE_LISTING.open(encoding="utf-8") as stream:
        for number, line in enumerate(stream, 1):
            value = json.loads(line)
            need(isinstance(value, dict) and isinstance(value.get("path"), str), f"source listing line {number} malformed")
            relative = value["path"]
            need(not Path(relative).is_absolute() and ".." not in Path(relative).parts, f"unsafe source listing path {relative}")
            size = value.get("bytes")
            digest = value.get("sha256")
            need(type(size) is int and size >= 0 and isinstance(digest, str) and len(digest) == 64 and set(digest) <= HEX, f"source listing line {number} digest malformed")
            need(relative not in listing, f"duplicate source listing path {relative}")
            listing[relative] = (size, digest)
            total += size
    actual: dict[str, Path] = {}
    for directory, dirs, files in os.walk(SOURCE_ROOT, followlinks=False):
        directory_path = Path(directory)
        for name in dirs + files:
            path = directory_path / name
            need(not path.is_symlink(), f"source tree contains symlink: {path}")
        for name in files:
            path = directory_path / name
            need(path.is_file(), f"source tree entry is not regular: {path}")
            actual[path.relative_to(SOURCE_ROOT).as_posix()] = path
    need(set(actual) == set(listing), f"source closure path set changed ({len(actual)} vs {len(listing)})")
    for relative, (size, digest) in listing.items():
        path = actual[relative]
        actual_desc = digest_file(path)
        need(actual_desc["bytes"] == size and actual_desc["sha256"] == digest, f"source file changed: {relative}")
    need(len(listing) == SOURCE_FILE_COUNT and total == SOURCE_BYTES, "source closure count/bytes changed")
    receipt, _ = read_json(SOURCE_RECEIPT, "source receipt")
    need(receipt.get("source_root") == str(SOURCE_ROOT) and receipt.get("source_revision") == SOURCE_REVISION, "source receipt identity changed")
    need(receipt.get("source_file_count") == SOURCE_FILE_COUNT and receipt.get("source_bytes") == SOURCE_BYTES, "source receipt totals changed")
    return {"file_count": len(listing), "bytes": total, "listing_sha256": SOURCE_LISTING_SHA256}


def validate_manifest(manifest: dict[str, Any]) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    manifest_desc = external_desc({"path": str(MANIFEST_PATH), "bytes": MANIFEST_PATH.stat().st_size, "sha256": AUTHORIZED_MANIFEST_SHA256}, MANIFEST_PATH, "authorized manifest", AUTHORIZED_MANIFEST_SHA256)
    need(manifest.get("schema") == "litchi.performance.0940.capture-manifest.v1", "manifest schema changed")
    need(manifest.get("status") == "frozen_authorized", "manifest status changed")
    need(manifest.get("plan_sha256") == PLAN_SHA256 and manifest.get("schedule_sha256") == SCHEDULE_SHA256, "manifest plan/schedule binding changed")
    need(manifest.get("authorization", {}).get("capture_authorized") is True, "manifest is not capture-authorized")
    rows = expected_schedule()
    need(manifest.get("schedule") == rows, "manifest schedule differs from frozen schedule")
    need(canonical(rows) == SCHEDULE_SHA256, "frozen schedule hash mismatch")
    source_summary = validate_source(manifest)
    scripts = manifest.get("scripts")
    need(isinstance(scripts, Mapping), "manifest scripts missing")
    expected_scripts = {
        "child_validation": (HELPER_PATH, HELPER_SHA256),
        "comparison": (COMPARISON_PATH, COMPARISON_SHA256),
        "runner": (RUNNER_PATH, RUNNER_SHA256),
        "independent_validator": (VALIDATOR_PATH, VALIDATOR_SHA256),
        "taskset": (TASKSET_PATH, TASKSET_SHA256),
    }
    for name, (path, digest) in expected_scripts.items():
        value = scripts.get(name)
        need(isinstance(value, Mapping), f"manifest script {name} missing")
        external_desc(value, path, f"manifest script {name}", digest)
    builds = manifest.get("build")
    need(isinstance(builds, Mapping), "manifest build mapping missing")
    for lane, expected_binary in EXPECTED_BINARIES.items():
        entry = builds.get(lane)
        need(isinstance(entry, Mapping), f"manifest build {lane} missing")
        binary = entry.get("binary")
        need(isinstance(binary, Mapping), f"manifest {lane} binary missing")
        need(binary.get("path") == expected_binary["path"] and binary.get("bytes") == expected_binary["bytes"] and binary.get("sha256") == expected_binary["sha256"], f"manifest {lane} binary binding changed")
        external_desc(binary, Path(expected_binary["path"]), f"{lane} binary", expected_binary["sha256"])
        expected_keys = ("normal_off", "normal_on") if lane == "normal" else ("allocator_on",)
        reports = entry.get("expected_reports")
        need(isinstance(reports, Mapping) and set(reports) == set(expected_keys), f"manifest {lane} expected report keys changed")
        for key in expected_keys:
            expected = expected_report({"normal_off": "normal_probes_off_reference", "normal_on": "normal_probes_on_diagnostic", "allocator_on": "allocator_probes_on_diagnostic"}[key])
            need(reports[key] == expected, f"manifest {lane} expected report {key} changed")
    return rows, {"manifest": manifest_desc, "source": source_summary}


def validate_build_and_preflight(manifest: Mapping[str, Any]) -> dict[str, Any]:
    builds = manifest.get("build")
    need(isinstance(builds, Mapping), "manifest build mapping missing")
    build_receipts: dict[str, Any] = {}
    for lane, expected_binary in EXPECTED_BINARIES.items():
        entry = builds.get(lane)
        need(isinstance(entry, Mapping), f"manifest build {lane} missing")
        need(entry.get("source_root") == str(SOURCE_ROOT)
             and entry.get("source_revision") == SOURCE_REVISION
             and entry.get("source_receipt_sha256") == SOURCE_RECEIPT_SHA256
             and entry.get("full_source_listing_sha256") == SOURCE_LISTING_SHA256,
             f"manifest build {lane} source binding changed")
        binary_desc = entry.get("binary")
        external_desc(binary_desc, Path(expected_binary["path"]), f"{lane} binary", expected_binary["sha256"])
        paths: dict[str, Path] = {}
        values: dict[str, dict[str, Any]] = {}
        for key in ("start", "terminal", "binary_receipt"):
            value = entry.get(key)
            need(isinstance(value, Mapping) and isinstance(value.get("path"), str), f"{lane} build {key} descriptor missing")
            path = Path(value["path"]).resolve()
            actual = external_desc(value, path, f"{lane} build {key}")
            parsed, _ = read_json(path, f"{lane} build {key}")
            paths[key] = path
            values[key] = parsed
        start = values["start"]
        terminal = values["terminal"]
        binary_receipt = values["binary_receipt"]
        need(start.get("schema") == "litchi.performance.0940.build-start.v1"
             and terminal.get("schema") == "litchi.performance.0940.build-terminal.v1",
             f"{lane} build receipt schema changed")
        need(start.get("source_root") == str(SOURCE_ROOT)
             and terminal.get("source_root") == str(SOURCE_ROOT)
             and start.get("cwd") == str(SOURCE_ROOT)
             and terminal.get("cwd") == str(SOURCE_ROOT),
             f"{lane} build source/cwd changed")
        need(start.get("argv") == terminal.get("argv") and terminal.get("exit_code") == 0
             and terminal.get("terminal_written_before_binary_descriptor") is True,
             f"{lane} build terminal/start ordering or status changed")
        need(binary_receipt == {"bytes": expected_binary["bytes"], "path": expected_binary["path"], "sha256": expected_binary["sha256"]},
             f"{lane} binary receipt does not match frozen binary")
        need(terminal.get("binary_name") == expected_binary["identity"], f"{lane} build binary identity changed")
        build_receipts[lane] = {
            "start": {"path": str(paths["start"]), "bytes": entry["start"]["bytes"], "sha256": entry["start"]["sha256"]},
            "terminal": {"path": str(paths["terminal"]), "bytes": entry["terminal"]["bytes"], "sha256": entry["terminal"]["sha256"]},
            "binary_receipt": {"path": str(paths["binary_receipt"]), "bytes": entry["binary_receipt"]["bytes"], "sha256": entry["binary_receipt"]["sha256"]},
        }
    preflight_value = manifest.get("independent_real_preflight")
    preflight_path = ROOT / "receipts/independent-real-preflight-v3-independent-v2.json"
    preflight_desc = external_desc(preflight_value, preflight_path, "independent real preflight", INDEPENDENT_PREFLIGHT_SHA256)
    preflight, _ = read_json(preflight_path, "independent real preflight")
    need(preflight.get("schema") == "litchi.performance.0940.independent-real-preflight-v3.v2"
         and preflight.get("status") == "passed" and preflight.get("child_count") == 8,
         "independent real preflight status/cardinality changed")
    need(preflight.get("source") == {"bytes": SOURCE_CORPUS_BYTES, "revision": SOURCE_REVISION, "root": str(SOURCE_ROOT), "sha256": SOURCE_CORPUS_SHA256},
         "independent real preflight source binding changed")
    expected_preflight_binaries = {
        "normal_binary": {"bytes": EXPECTED_BINARIES["normal"]["bytes"], "path": EXPECTED_BINARIES["normal"]["path"], "sha256": EXPECTED_BINARIES["normal"]["sha256"]},
        "allocator_binary": {"bytes": EXPECTED_BINARIES["allocator"]["bytes"], "path": EXPECTED_BINARIES["allocator"]["path"], "sha256": EXPECTED_BINARIES["allocator"]["sha256"]},
    }
    need(preflight.get("binaries") == expected_preflight_binaries, "independent real preflight binary binding changed")
    need(preflight.get("report_size_bound_bytes") == MAX_REPORT_BYTES and preflight.get("report_size_max_observed", MAX_REPORT_BYTES) < MAX_REPORT_BYTES,
         "independent real preflight report bound changed")
    return {"build_receipts": build_receipts, "independent_real_preflight": preflight_desc}


def validate_outer_terminal(outer_path: Path, manifest: dict[str, Any]) -> dict[str, Any]:
    need(outer_path.resolve() == OUTER_TERMINAL.resolve(), "noncanonical outer terminal path rejected")
    terminal, terminal_desc = read_json(outer_path, "outer capture terminal")
    expected_command = [
        "python3", "-B", str(RUNNER_PATH), "--capture", "--manifest", str(MANIFEST_PATH),
        "--packet", str(ORIGINAL_PACKET),
    ]
    need(terminal.get("schema") == "litchi.performance.0940.outer-capture-terminal.v2", "outer terminal schema changed")
    need(terminal.get("exit_code") == 0 and terminal.get("terminal_written_before_analysis") is True, "outer capture did not terminate successfully")
    need(terminal.get("command") == expected_command and terminal.get("cwd") == str(REPO), "outer capture command/cwd changed")
    need(terminal.get("packet") == str(ORIGINAL_PACKET), "outer capture packet changed")
    need(terminal.get("manifest", {}).get("path") == str(MANIFEST_PATH) and terminal.get("manifest", {}).get("sha256") == AUTHORIZED_MANIFEST_SHA256, "outer terminal manifest binding changed")
    need(terminal.get("runner", {}).get("path") == str(RUNNER_PATH) and terminal.get("runner", {}).get("sha256") == RUNNER_SHA256, "outer terminal runner binding changed")
    need(terminal.get("environment") == {"PATH": "/usr/bin:/bin", "LC_ALL": "C", "LANG": "C", "TZ": "UTC", "LITCHI_PERF_0940_CAPTURE_APPROVED": "1", "PYTHONDONTWRITEBYTECODE": "1"}, "outer environment changed")
    for name in ("stdout", "stderr"):
        value = terminal.get(name)
        need(isinstance(value, Mapping), f"outer {name} descriptor missing")
        path = Path(value["path"])
        actual = digest_file(path)
        need(value.get("bytes") == actual["bytes"] and value.get("sha256") == actual["sha256"], f"outer {name} descriptor changed")
    return terminal_desc


def binary_binding(manifest: Mapping[str, Any], arm: str) -> tuple[str, dict[str, Any]]:
    lane = "allocator" if arm == "allocator_probes_on_diagnostic" else "normal"
    short = ARM_SHORT[arm]
    build = manifest["build"][lane]
    binary = dict(build["binary"])
    binary["binary_identity"] = EXPECTED_BINARIES[lane]["identity"]
    binary["source_revision"] = SOURCE_REVISION
    expected = expected_report(arm)
    binary["expected_report"] = expected
    binary["expected_allocation_present"] = expected["expected_allocation_present"]
    binary["expected_process_present"] = expected["expected_process_present"]
    return lane, binary


def packet_json(packet: Path, relative: str, role: str) -> tuple[dict[str, Any], dict[str, Any]]:
    path = packet / relative
    return read_json(path, role)


def validate_child(packet: Path, row: Mapping[str, Any], index: int, validator: Any) -> dict[str, Any]:
    child_id = row["id"]
    child = packet / "children" / child_id
    need(child.is_dir() and not child.is_symlink(), f"missing child directory {child_id}")
    expected_files = {"terminal.json", "validation.json", "report.json", "stdout.log", "stderr.log"}
    actual_files = {path.name for path in child.iterdir()}
    need(actual_files == expected_files, f"{child_id}: child file set changed")
    # Parse only the terminal first; report semantics are unreachable until terminal
    # status, command, binary metadata, and custody descriptors pass.
    terminal, terminal_desc = packet_json(packet, f"children/{child_id}/terminal.json", f"{child_id} terminal")
    need(terminal.get("schema") == "litchi.performance.child-validation.v1", f"{child_id}: terminal schema changed")
    need(terminal.get("status") == "terminal" and terminal.get("exit_code") == 0, f"{child_id}: terminal unsuccessful")
    need(terminal.get("deadline_exceeded") is False and terminal.get("process_group_closed") is True and terminal.get("process_group_forced_cleanup") is False, f"{child_id}: terminal process/deadline state invalid")
    need(terminal.get("runner_error") is None and terminal.get("report_error") is None, f"{child_id}: terminal has runner/report error")
    need(type(terminal.get("observation_timeout_count")) is int and terminal["observation_timeout_count"] >= 0, f"{child_id}: observation timeout count invalid")
    need(terminal.get("cwd") == str(SOURCE_ROOT) and terminal.get("label") == child_id and terminal.get("route") == row["route"], f"{child_id}: terminal identity changed")
    report_contract = "docx-owned-calibration" if row["kind"] == "calibration" else "docx-owned-workload"
    need(terminal.get("report_contract") == report_contract and terminal.get("expected_output_sha256") == (None if row["kind"] == "calibration" else OUTPUT_SHA256), f"{child_id}: terminal contract/output binding changed")
    need(terminal.get("deadline_seconds") == 900.0, f"{child_id}: deadline changed")
    need(terminal.get("environment") == ENVIRONMENT, f"{child_id}: child environment changed")
    lane, binary = binary_binding(manifest_global, row["arm"])
    need(terminal.get("binary") == binary, f"{child_id}: terminal binary metadata differs from frozen build binding")
    argv = terminal.get("argv")
    expected_report_original = ORIGINAL_PACKET / "children" / child_id / "report.json"
    expected_argv = [
        str(TASKSET_PATH), "--cpu-list", "12", binary["path"], "docx-owned-attribution",
        "--case", row["kind"], "--diagnostic", "off" if row["arm"] == "normal_probes_off_reference" else "on",
        "--warmups", "5", "--samples", "50", "--source-revision", SOURCE_REVISION,
        "--output", str(expected_report_original),
    ]
    if row["kind"] == "workload":
        expected_argv.extend(("--route", row["route"]))
    need(argv == expected_argv, f"{child_id}: argv differs from frozen schedule")
    report_actual = packet / "children" / child_id / "report.json"
    report_record = exact_desc(terminal.get("report"), expected_report_original, report_actual, f"{child_id} terminal report", max_bytes=MAX_REPORT_BYTES)
    report_desc = digest_file(report_actual)
    stdout_original = ORIGINAL_PACKET / "children" / child_id / "stdout.log"
    stderr_original = ORIGINAL_PACKET / "children" / child_id / "stderr.log"
    stdout_record = exact_desc(terminal.get("stdout"), stdout_original, packet / "children" / child_id / "stdout.log", f"{child_id} stdout")
    stderr_record = exact_desc(terminal.get("stderr"), stderr_original, packet / "children" / child_id / "stderr.log", f"{child_id} stderr")
    need(report_record["bytes"] == report_desc["bytes"] and report_record["sha256"] == report_desc["sha256"], f"{child_id}: report terminal descriptor mismatch")
    # Only now parse validation/report JSON and invoke semantic validation.
    validation, validation_desc = packet_json(packet, f"children/{child_id}/validation.json", f"{child_id} validation")
    report, parsed_report_desc = packet_json(packet, f"children/{child_id}/report.json", f"{child_id} report")
    need(parsed_report_desc["bytes"] == report_desc["bytes"] and parsed_report_desc["sha256"] == report_desc["sha256"], f"{child_id}: report changed during parse")
    expected_meta = report_fields(row["arm"])
    expected_binding = expected_report(row["arm"])
    for key, expected_value in expected_meta.items():
        need(report.get(key) == expected_value and type(report.get(key)) is type(expected_value), f"{child_id}: report {key} differs from frozen metadata")
    need(report.get("schema") == "docx_owned_ingress_attribution_v1" and report.get("scenario") == "docx-owned-ingress-attribution", f"{child_id}: report schema changed")
    need(report.get("case") == row["kind"] and report.get("route") == row["route"], f"{child_id}: report case/route changed")
    need(report.get("source_bytes") == SOURCE_CORPUS_BYTES and report.get("source_sha256") == SOURCE_CORPUS_SHA256, f"{child_id}: source corpus identity changed")
    bound = {"expected_report": expected_meta, "expected_allocation_present": expected_binding["expected_allocation_present"], "expected_process_present": expected_binding["expected_process_present"]}
    validator.validate_report(report, row["route"], warmups=5, samples=50, label=child_id, binary=bound)
    need(validation.get("schema") == "litchi.performance.child-validation.v1" and validation.get("status") == "passed", f"{child_id}: validation status/schema changed")
    need(validation.get("label") == child_id and validation.get("route") == row["route"] and validation.get("report_contract") == report_contract and validation.get("sample_count") == 50, f"{child_id}: validation identity changed")
    expected_validation_report = {"path": str(expected_report_original), "bytes": report_desc["bytes"], "sha256": report_desc["sha256"]}
    need(validation.get("report") == expected_validation_report, f"{child_id}: validation report descriptor changed")
    if row["kind"] == "workload":
        need(validation.get("output_sha256") == OUTPUT_SHA256, f"{child_id}: validation output digest changed")
    else:
        need("output_sha256" not in validation, f"{child_id}: calibration publishes output digest")
    return {
        "id": child_id,
        "arm": ARM_SHORT[row["arm"]],
        "case": row["kind"],
        "route": row["route"],
        "block": row.get("block", row.get("boundary_after_block")),
        "position": row["position"],
        "report": report_record,
        "terminal": {"path": f"children/{child_id}/terminal.json", "bytes": terminal_desc["bytes"], "sha256": terminal_desc["sha256"]},
        "validation": {"path": f"children/{child_id}/validation.json", "bytes": validation_desc["bytes"], "sha256": validation_desc["sha256"]},
        "stdout": stdout_record,
        "stderr": stderr_record,
        "observation_timeout_count": terminal["observation_timeout_count"],
    }


def validate_packet(packet: Path, rows: list[dict[str, Any]], validator: Any, role: str) -> tuple[dict[str, Any], dict[str, Any]]:
    need(packet.is_dir() and not packet.is_symlink(), f"{role} packet missing")
    state, state_desc = read_json(packet / "capture-state.json", f"{role} capture state")
    need(state.get("schema") == "litchi.performance.0940.capture-state.v1" and state.get("status") == "completed", f"{role}: capture state incomplete")
    need(state.get("children") == 80 and state.get("passed") == 80 and state.get("failed") == 0 and state.get("schedule_sha256") == SCHEDULE_SHA256, f"{role}: capture state counts/schedule changed")
    closure = state.get("source_closure_after")
    need(closure == {"bytes": SOURCE_BYTES, "file_count": SOURCE_FILE_COUNT, "listing_sha256": SOURCE_LISTING_SHA256}, f"{role}: source closure receipt changed")
    expected_results = [{"id": row["id"], "index": index, "kind": row["kind"], "status": "passed"} for index, row in enumerate(rows, 1)]
    need(state.get("results") == expected_results, f"{role}: capture results differ from frozen schedule")
    for index, row in enumerate(rows, 1):
        progress, _ = read_json(packet / f"capture-state-{index:03d}.json", f"{role} progress {index}")
        need(progress == {"completed": index, "last": {"id": row["id"], "index": index, "kind": row["kind"], "status": "passed"}, "schema": "litchi.performance.0940.capture-progress.v1", "total": 80}, f"{role}: progress {index} changed")
    expected_child_dirs = {row["id"] for row in rows}
    children_dir = packet / "children"
    need(children_dir.is_dir() and {path.name for path in children_dir.iterdir()} == expected_child_dirs, f"{role}: child directory set changed")
    children = [validate_child(packet, row, index, validator) for index, row in enumerate(rows, 1)]
    index = {"schema": "litchi.performance.0940.independent-formal-children-index.v1", "children": children}
    return index, {"state": state_desc, "children_index_sha256": canonical(children), "max_report_bytes": max(item["report"]["bytes"] for item in children)}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--packet", type=Path, required=True)
    parser.add_argument("--peer-packet", type=Path, required=True)
    parser.add_argument("--outer-terminal", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--index", type=Path, required=True)
    args = parser.parse_args()
    packet = args.packet.resolve()
    peer = args.peer_packet.resolve()
    outer = args.outer_terminal.resolve()
    need(packet == ORIGINAL_PACKET.resolve(), "noncanonical original packet path rejected")
    need(peer == MOVED_PACKET.resolve() and peer != packet, "peer packet must be the distinct canonical moved packet")
    need(outer == OUTER_TERMINAL.resolve(), "noncanonical outer terminal path rejected")
    global manifest_global
    manifest_global, _ = read_json(MANIFEST_PATH, "authorized manifest")
    rows, manifest_info = validate_manifest(manifest_global)
    build_info = validate_build_and_preflight(manifest_global)
    # The outer terminal is checked before opening any child packet JSON.
    outer_desc = validate_outer_terminal(outer, manifest_global)
    validator = load_validator()
    original_index, original_meta = validate_packet(packet, rows, validator, "original")
    peer_index, peer_meta = validate_packet(peer, rows, validator, "moved")
    need(original_index["children"] == peer_index["children"], "moved packet child index differs from original")
    index = original_index
    index_data = json.dumps(index, indent=2, sort_keys=True) + "\n"
    args.index.parent.mkdir(parents=True, exist_ok=True)
    args.index.write_text(index_data, encoding="utf-8")
    index_desc = digest_file(args.index)
    children_hash = canonical(index["children"])
    result = {
        "schema": "litchi.performance.0940.independent-formal-audit.v2",
        "created_utc": datetime.now(timezone.utc).isoformat(),
        "status": "passed",
        "scope": "terminal-first direct validation of authorized 80-child original and moved packets; no workloads executed",
        "manifest_sha256": AUTHORIZED_MANIFEST_SHA256,
        "outer_terminal": outer_desc,
        "validator": {"path": str(VALIDATOR_PATH), "bytes": VALIDATOR_PATH.stat().st_size, "sha256": VALIDATOR_SHA256},
        "source_closure": manifest_info["source"],
        "build_custody": build_info["build_receipts"],
        "independent_real_preflight": build_info["independent_real_preflight"],
        "child_count": 80,
        "children_index_sha256": children_hash,
        "children_index": index_desc,
        "original": {"packet": str(packet), "capture_state": original_meta["state"], "max_report_bytes": original_meta["max_report_bytes"]},
        "moved": {"packet": str(peer), "capture_state": peer_meta["state"], "max_report_bytes": peer_meta["max_report_bytes"]},
        "peer_equivalence": True,
        "report_size_bound_bytes": MAX_REPORT_BYTES,
        "max_report_bytes": max(original_meta["max_report_bytes"], peer_meta["max_report_bytes"]),
        "review": {"status": "passed", "child_count": 80, "manifest_sha256": AUTHORIZED_MANIFEST_SHA256, "children_index_sha256": children_hash},
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps({"status": "passed", "child_count": 80, "manifest_sha256": AUTHORIZED_MANIFEST_SHA256, "children_index_sha256": children_hash, "max_report_bytes": result["max_report_bytes"], "output": str(args.output), "index": str(args.index)}, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except CheckError as error:
        print(f"independent formal audit failed: {error}")
        raise SystemExit(1)
