#!/usr/bin/env python3
"""Read-only, fail-closed validator for the prospective 0941 XLSX controls.

This validator never starts Cargo, the probe, or a child process. It checks
the frozen matrix, exact fixture custody, source receipts, build terminals,
binary witnesses, and the old 0779 published-output witnesses. Capture
remains blocked until the coordinator changes the plan status to
``authorized``.
"""

from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parent
PLAN = ROOT / "xlsx-plan-v4.json"
SCRIPT = ROOT / "xlsx-capture-v4.py"
RUNTIME_ENV = {
    "PATH": "/usr/bin:/bin",
    "LANG": "C",
    "LC_ALL": "C",
    "TZ": "UTC",
    "RUST_BACKTRACE": "0",
    "unset": ["RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "CARGO_HOME", "RUSTUP_TOOLCHAIN", "CARGO_INCREMENTAL"],
}
FIXTURE_FILES = {
    "generated-xlsx": "generated-xlsx-medium-source.xlsx",
    "conditional-formatting": "conditional-formatting-source.xlsx",
    "small-xlsx": "small-xlsx-source.xlsx",
    "boundary-xlsx": "boundary-xlsx-source.xlsx",
}
FIXTURE_WITNESSES = {
    "generated-xlsx": {"bytes": 4226429, "sha256": "dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036", "sheet": "Sheet1"},
    "conditional-formatting": {"bytes": 654688, "sha256": "8dfa997a642c00a1da8c0750c371f204553a98c537ca046028b052f9b424a074", "sheet": "Home"},
    "small-xlsx": {"bytes": 3555, "sha256": "a6d7f8a056622201862c354bd508820f7e57a84c5e2531e4e4da20180615f312", "sheet": "Sheet1"},
    "boundary-xlsx": {"bytes": 8224, "sha256": "9d57919aa4f6036fb922daba5d1dc913a8847581fd9a3dd3c4d79bca7d28bdb8", "sheet": "Sheet1"},
}
PROBE_SOURCE_FILES = {
    "src/main.rs": "c601ce2b369b51f3921ac24b972b0fa573c02dec356289a5c5ccbd4995d35695",
    "src/allocation_metrics.rs": "68b1c02c3a576cd6e0dde23c7c8442f1ff9a034d586cae3e960fec5cbec320d1",
    "src/counting_allocator.rs": "bd89fa4d5bdd1ac9d02921dbdf373dd7b484f358e5ac3c34455259a4cd7cfe10",
}
PROBE_LOCK_SHA = "9f5e7da95e9af1ec0a35c21fd1c2daa6081d14c54031553cdca56a8804d26d3f"
CANDIDATE_PHYS_PKG_SHA = "899ce1567755d8b4d2c3128a016ccc85c968874750e426d0ffd7da1ff5dc147c"


def fail(message: str) -> None:
    raise SystemExit(f"plan validation failed: {message}")


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def read(path: Path) -> dict:
    try:
        value = json.loads(path.read_text())
    except (OSError, json.JSONDecodeError) as exc:
        fail(f"cannot read JSON {path}: {exc}")
    require(isinstance(value, dict), f"{path} is not a JSON object")
    return value


def artifact(path: Path) -> dict:
    require(path.is_file(), f"artifact missing: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha(path)}


def check_artifact(path: Path, expected: dict, label: str) -> None:
    require(artifact(path) == expected, f"{label} artifact changed")


def validate_source_receipt(plan: dict, leg: str, leg_data: dict, receipt: dict) -> None:
    source = plan["source_and_probe"]
    root = Path(leg_data["source_root"])
    require(receipt.get("schema") == "litchi.performance.0941.xlsx-probe-source-receipt.v1", f"{leg} source receipt schema changed")
    require(receipt.get("leg") == leg, f"{leg} receipt leg mismatch")
    require(receipt.get("source_root") == str(root), f"{leg} source root binding changed")
    require(receipt.get("source_revision") == "3ae856d76424f2f9b58d26dcd20fd074b8e97541", f"{leg} source revision changed")

    listing = receipt.get("source_listing", {})
    listing_path = Path(listing.get("path", ""))
    require(listing_path.is_file(), f"{leg} source listing missing")
    require(listing.get("bytes") == listing_path.stat().st_size and listing.get("sha256") == sha(listing_path), f"{leg} source listing binding changed")

    source_lock = receipt.get("source_lock", {})
    source_lock_path = Path(source_lock.get("path", ""))
    require(source_lock_path == root / "Cargo.lock", f"{leg} source lock path changed")
    require(source_lock.get("bytes") == source_lock_path.stat().st_size and source_lock.get("sha256") == sha(source_lock_path), f"{leg} source lock binding changed")

    source_toolchain = receipt.get("source_toolchain", {})
    toolchain_path = Path(source_toolchain.get("path", ""))
    require(toolchain_path == root / "rust-toolchain.toml", f"{leg} toolchain path changed")
    require(source_toolchain.get("sha256") == sha(toolchain_path) and source_toolchain.get("channel") == "1.95.0", f"{leg} toolchain binding changed")

    probe = receipt.get("probe", {})
    probe_root = Path(leg_data["probe_root"])
    manifest_path = probe_root / probe.get("manifest", "")
    require(manifest_path == probe_root / "Cargo.toml" and manifest_path.is_file(), f"{leg} probe manifest binding changed")
    require(probe.get("manifest_sha256") == sha(manifest_path), f"{leg} probe manifest hash changed")
    lock_path = probe_root / probe.get("lock", "")
    require(lock_path == probe_root / "Cargo.lock" and lock_path.is_file(), f"{leg} probe lock path changed")
    require(probe.get("lock_sha256") == PROBE_LOCK_SHA and sha(lock_path) == PROBE_LOCK_SHA, f"{leg} probe lock binding changed")
    require(probe.get("source_sha256") == PROBE_SOURCE_FILES, f"{leg} probe source manifest changed")
    for filename, digest in PROBE_SOURCE_FILES.items():
        path = probe_root / filename
        require(path.is_file() and sha(path) == digest, f"{leg} probe source changed: {filename}")

    if leg == "candidate":
        overlay = receipt.get("candidate_overlay", {})
        overlay_path = root / overlay.get("path", "")
        require(overlay_path == root / "crates/litchi-opc/src/phys_pkg.rs", "candidate overlay path changed")
        require(overlay.get("bytes") == overlay_path.stat().st_size and overlay.get("sha256") == CANDIDATE_PHYS_PKG_SHA and sha(overlay_path) == CANDIDATE_PHYS_PKG_SHA, "candidate overlay binding changed")
    else:
        require("candidate_overlay" not in receipt, "baseline has candidate overlay")

    expected_fixtures = []
    for fixture_id, expected in FIXTURE_WITNESSES.items():
        expected_fixtures.append({"id": fixture_id, "path": f"fixtures/{FIXTURE_FILES[fixture_id]}", "bytes": expected["bytes"], "sha256": expected["sha256"], "sheet": expected["sheet"]})
    require(receipt.get("fixtures") == expected_fixtures, f"{leg} fixture receipt changed")
    for fixture in expected_fixtures:
        path = probe_root / fixture["path"]
        require(path.is_file() and path.stat().st_size == fixture["bytes"] and sha(path) == fixture["sha256"], f"{leg} fixture changed: {fixture['id']}")
    require(receipt.get("status") == "immutable_source_and_probe_inputs_before_build", f"{leg} source receipt status changed")


def validate_build(plan: dict, leg: str, lane: str, leg_data: dict) -> None:
    build = plan["build_witnesses"][leg][lane]
    probe_root = Path(leg_data["probe_root"])
    source_root = Path(leg_data["source_root"])
    manifest = probe_root / "Cargo.toml"
    lock = probe_root / "Cargo.lock"
    binary_key = "normal" if lane == "native" else "allocator"
    binary = plan["binary_witnesses"][leg][binary_key]
    start_path = Path(build["start"]["path"])
    terminal_path = Path(build["terminal"]["path"])
    start_artifact = artifact(start_path)
    terminal_artifact = artifact(terminal_path)
    require(start_artifact["path"] == build["start"]["path"] and start_artifact["sha256"] == build["start"]["sha256"], f"{leg} {lane} build start artifact changed")
    require(terminal_artifact["path"] == build["terminal"]["path"] and terminal_artifact["sha256"] == build["terminal"]["sha256"], f"{leg} {lane} build terminal artifact changed")
    start = read(start_path)
    terminal = read(terminal_path)
    require(start.get("schema") == "litchi.performance.0941.xlsx-probe-build-start.v1", f"{leg} {lane} start schema changed")
    require(start.get("leg") == leg and start.get("variant") == lane and start.get("status") == "launching", f"{leg} {lane} start identity changed")
    require(start.get("cwd") == str(probe_root), f"{leg} {lane} start cwd changed")
    require(start.get("inputs", {}).get("source_receipt") == str(Path(leg_data["source_receipt"])), f"{leg} {lane} start source receipt binding changed")
    require(start.get("inputs", {}).get("manifest") == str(manifest), f"{leg} {lane} start manifest binding changed")
    require(start.get("inputs", {}).get("lock") == str(lock) and start.get("inputs", {}).get("lock_sha256") == PROBE_LOCK_SHA, f"{leg} {lane} start lock binding changed")
    policy = start.get("environment_policy", {})
    require(policy.get("toolchain") == "1.95.0" and policy.get("offline") is True and policy.get("locked") is True and policy.get("jobs") == 2, f"{leg} {lane} start environment policy changed")
    argv = start.get("argv", [])
    require(len(argv) >= 3 and argv[1:3] == ["+1.95.0", "build"] and "--offline" in argv and "--locked" in argv and "-j2" in argv, f"{leg} {lane} build argv changed")
    require("--manifest-path" in argv and argv[argv.index("--manifest-path") + 1] == str(manifest), f"{leg} {lane} argv manifest binding changed")
    require("--target-dir" in argv and argv[argv.index("--target-dir") + 1] == f"/home/zhuhe/litchi-performance-0941/target-xlsx-{leg}", f"{leg} {lane} target binding changed")
    if lane == "allocator":
        require("--features" in argv and argv[argv.index("--features") + 1] == "allocator-metrics", f"{leg} allocator feature binding changed")
    else:
        require("--features" not in argv, f"{leg} native feature binding changed")

    require(terminal.get("schema") == "litchi.performance.0941.xlsx-probe-build-terminal.v1", f"{leg} {lane} terminal schema changed")
    require(terminal.get("leg") == leg and terminal.get("variant") == lane and terminal.get("status") == "succeeded" and terminal.get("exit_code") == 0, f"{leg} {lane} terminal status changed")
    require(terminal.get("cwd") == str(probe_root) and terminal.get("argv") == argv, f"{leg} {lane} terminal command binding changed")
    inputs = terminal.get("inputs", {})
    require(inputs.get("source_root") == str(source_root), f"{leg} {lane} terminal source root binding changed")
    require(inputs.get("source_receipt") == str(Path(leg_data["source_receipt"])), f"{leg} {lane} terminal source receipt binding changed")
    require(inputs.get("manifest") == str(manifest) and inputs.get("manifest_sha256") == sha(manifest), f"{leg} {lane} terminal manifest binding changed")
    require(inputs.get("lock") == str(lock) and inputs.get("lock_sha256") == PROBE_LOCK_SHA, f"{leg} {lane} terminal lock binding changed")
    expected_env = {"RUSTFLAGS": "unset", "CARGO_ENCODED_RUSTFLAGS": "unset", "CARGO_HOME": "unset", "RUSTUP_TOOLCHAIN": "unset", "CARGO_INCREMENTAL": "unset", "toolchain": "1.95.0", "offline": True, "locked": True, "jobs": 2}
    if lane == "allocator":
        expected_env["features"] = ["allocator-metrics"]
    require(terminal.get("environment") == expected_env, f"{leg} {lane} terminal environment changed")
    require(terminal.get("binary") == binary, f"{leg} {lane} terminal binary witness changed")
    check_artifact(Path(binary["path"]), binary, f"{leg} {lane} binary")
    for stream in ("stdout", "stderr"):
        recorded = terminal.get(stream, {})
        recorded_path = Path(recorded.get("path", ""))
        stream_path = recorded_path if recorded_path.is_absolute() else terminal_path.parent / recorded_path
        actual = artifact(stream_path)
        require(actual["bytes"] == recorded.get("bytes") and actual["sha256"] == recorded.get("sha256"), f"{leg} {lane} build {stream} artifact changed")


def validate(require_binaries: bool = False) -> dict:
    plan = read(PLAN)
    require(plan.get("schema") == "litchi.performance.0941.xlsx-controls-plan.v4", "schema changed")
    require(plan.get("status") in {"prospective_only_pending_root_approval", "authorized"}, "unexpected plan status")
    scope = plan.get("scope", {})
    require(scope.get("iwork") == "excluded" and scope.get("cpu") == 12 and scope.get("deadline_seconds") == 900, "scope changed")
    require(scope.get("no_retry_or_replacement") is True and scope.get("terminal_first") is True, "capture policy changed")
    native = scope["native"]
    require(native == {"blocks": 6, "warmup": 3, "samples": 30, "orders": [["baseline", "candidate"], ["candidate", "baseline"], ["baseline", "candidate"], ["candidate", "baseline"], ["candidate", "baseline"], ["baseline", "candidate"]], "children_per_cell": 12, "total_children": 72}, "native matrix changed")
    allocator = scope["allocator"]
    require(allocator == {"blocks": 2, "warmup": 0, "samples": 3, "orders": [["baseline", "candidate"], ["candidate", "baseline"]], "children_per_cell": 4, "total_children": 24}, "allocator matrix changed")
    expected_cells = [
        ("generated-xlsx/open", "generated-xlsx", "open", "Sheet1", "primary"),
        ("generated-xlsx/lifecycle", "generated-xlsx", "lifecycle", "Sheet1", "primary"),
        ("conditional-formatting/open", "conditional-formatting", "open", "Home", "primary"),
        ("conditional-formatting/lifecycle", "conditional-formatting", "lifecycle", "Home", "primary"),
        ("small-xlsx/open", "small-xlsx", "open", "Sheet1", "boundary-control"),
        ("boundary-xlsx/open", "boundary-xlsx", "open", "Sheet1", "boundary-control"),
    ]
    cells = plan.get("cells")
    require(isinstance(cells, list) and len(cells) == len(expected_cells), "cell count changed")
    seen: set[str] = set()
    for actual, expected in zip(cells, expected_cells):
        require(tuple(actual.get(key) for key in ("id", "fixture", "phase", "sheet", "group")) == expected, f"cell changed: {actual}")
        require(actual["id"] not in seen and "iwork" not in actual["id"].lower(), f"cell identity changed: {actual}")
        seen.add(actual["id"])
    require(len(seen) == len(expected_cells), "cell identity is not unique")
    require(plan.get("fixtures") == {key: {"bytes": value["bytes"], "sha256": value["sha256"]} for key, value in FIXTURE_WITNESSES.items()}, "fixture witness changed")

    source = plan["source_and_probe"]
    require(source.get("probe_lock_sha256") == PROBE_LOCK_SHA and source.get("toolchain") == "1.95.0", "probe/toolchain witness changed")
    require(plan.get("runtime_environment") == RUNTIME_ENV, "runtime environment changed")
    require(plan["script_witness"]["path"] == str(SCRIPT) and SCRIPT.stat().st_size == plan["script_witness"]["bytes"] and sha(SCRIPT) == plan["script_witness"]["sha256"], "runner script witness changed")
    validator_path = Path(__file__).resolve()
    require(plan["validator_witness"]["path"] == str(validator_path) and validator_path.stat().st_size == plan["validator_witness"]["bytes"] and sha(validator_path) == plan["validator_witness"]["sha256"], "validator witness changed")
    for name in ("time", "taskset"):
        tool = plan["tool_witnesses"][name]
        path = Path(tool["path"])
        require(path.is_file() and path.stat().st_size == tool["bytes"] and sha(path) == tool["sha256"], f"{name} witness changed")

    for leg in ("baseline", "candidate"):
        leg_data = source["legs"][leg]
        receipt_path = Path(leg_data["source_receipt"])
        require(receipt_path.is_file(), f"{leg} source receipt missing")
        validate_source_receipt(plan, leg, leg_data, read(receipt_path))
        for key in ("normal_binary", "allocator_binary"):
            binary = Path(leg_data[key])
            require("iwork" not in str(binary).lower(), "iWork binary path present")
            if require_binaries:
                require(binary.is_file() and os.access(binary, os.X_OK), f"{leg} binary missing: {binary}")
            witness = plan["binary_witnesses"][leg]["normal" if key == "normal_binary" else "allocator"]
            require(witness["path"] == str(binary), f"{leg} binary path witness changed")
            if binary.is_file():
                check_artifact(binary, witness, f"{leg} {key} binary")
        for lane in ("native", "allocator"):
            validate_build(plan, leg, lane, leg_data)

    outputs = plan["expected_published_outputs_from_0779"]
    require(outputs.get("generated-xlsx/lifecycle") == {"bytes": 4226568, "marker": "litchi-perf-0638-ordinary-save", "sha256": "20335c4480405e051be2f4fea0e2a5aa8c5912b729cb1df980112ea7c5056687"}, "generated output witness changed")
    require(outputs.get("conditional-formatting/lifecycle") == {"bytes": 652357, "marker": "litchi-perf-0638-ordinary-save", "sha256": "c7d26c4a406ce6b1044acee33301212db08f68b39250bf7c7804aec5d1e474bb"}, "conditional output witness changed")
    require(plan["output_layout"]["root"].startswith("/home/zhuhe/litchi-performance-0941/"), "output left task scratch")
    return {"status": plan["status"], "cells": len(cells), "native_children": native["total_children"], "allocator_children": allocator["total_children"], "fixture_ids": sorted(plan["fixtures"]), "binaries_checked": require_binaries}


if __name__ == "__main__":
    print(json.dumps(validate("--require-binaries" in sys.argv[1:]), indent=2, sort_keys=True))
