#!/usr/bin/env python3
"""Independently verify the owned 0457 ODP append control capture."""

from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
import sys
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parent
BUNDLE = ROOT.parent
CONTROL_RELATIVE = Path("docs/performance/results/change-0457/control")
PHASES = {"R1": (0, 6), "R2": (6, 12)}


def load(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def fail(message: str) -> None:
    raise AssertionError(message)


def text(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        fail(f"{label}: expected non-empty text")
    return value


def artifact(base: Path, item: dict[str, Any]) -> Path:
    path = base / text(item["path"], "artifact.path")
    if not path.is_file() or path.is_symlink() or not path.resolve().is_relative_to(base.resolve()):
        fail(f"invalid artifact path: {path}")
    if sha(path) != item["sha256"] or path.stat().st_size != item["bytes"]:
        fail(f"artifact identity differs: {path}")
    return path


def expected_lanes(protocol: dict[str, Any], phase: str) -> list[dict[str, Any]]:
    first, last = PHASES[phase]
    lanes = protocol["order"][first:last]
    expected = (
        [("normal", "tiny"), ("normal", "medium"), ("normal", "large"),
         ("allocator", "tiny"), ("allocator", "medium"), ("allocator", "large")]
        if phase == "R1" else
        [("allocator", "large"), ("allocator", "medium"), ("allocator", "tiny"),
         ("normal", "large"), ("normal", "medium"), ("normal", "tiny")]
    )
    if len(lanes) != 6 or [(row["mode"], row["shape"]) for row in lanes] != expected:
        fail(f"{phase}: lane order differs from protocol")
    if any(row["phase"] != phase or row["repeat"] != phase for row in lanes):
        fail(f"{phase}: lane phase/repeat identity differs")
    return lanes


def oracle_command(
    protocol: dict[str, Any],
    report: Path,
    mode: str,
    shape: str,
    interpreter: str,
    command_root: Path,
) -> list[str]:
    verifier = command_root / protocol["oracle"]["verifier_path"]
    values = {"python": interpreter, "verifier": str(verifier), "report": str(report), "mode": mode, "shape": shape}
    return [item.format(**values) for item in protocol["oracle"]["argv"]]


def check_report(report: Path, binary: dict[str, Any], lane: dict[str, Any], protocol: dict[str, Any], revision: str) -> None:
    value = load(report)
    identity = value["binary_identity"]
    if identity["path"] != binary["copy_path"] or identity["binary_sha256"] != binary["sha256"] or identity["binary_bytes"] != binary["bytes"] or identity["profile"] != "release" or identity["executable"] is not True:
        fail(f"{report}: executable identity differs")
    if value["environment"]["git_revision"] != revision:
        fail(f"{report}: build revision differs")
    configuration = value["configuration"]
    if configuration["samples_per_case"] != protocol["samples"] or configuration["warmup_iterations_per_case"] != protocol["warmups"] or configuration["execution_workers"] != [protocol["workers"]]:
        fail(f"{report}: samples, warmups, or workers differ")
    results = value["results"]
    if len(results) != 1 or results[0]["case"] != protocol["selector"] or results[0]["corpus"]["shape"] != lane["shape"]:
        fail(f"{report}: selector or shape differs")
    if protocol.get("amendment") is not None and value["environment"].get("rustflags") is not None:
        fail(f"{report}: amended control must record null rustflags and make no frame-pointer claim")


def check_ambient_allocator(value: dict[str, Any], label: str) -> None:
    if value.get("allocator_environment") != {} or value.get("allocator_environment_after") != {}:
        fail(f"{label}: ambient allocator environment changed or was not empty")


def check_r1_oracle_amendment(amendment: dict[str, Any], amended_protocol: dict[str, Any]) -> Path:
    base_path = ROOT / text(amendment["base_protocol"], "amendment.base_protocol")
    if sha(base_path) != amendment["base_protocol_sha256"]:
        fail("amendment base protocol hash differs")
    base_protocol = load(base_path)
    base_oracle = base_protocol["oracle"]
    amended_oracle = amended_protocol["oracle"]
    if amended_oracle["protocol_path"] != base_oracle["protocol_path"] or amended_oracle["protocol_sha256"] != base_oracle["protocol_sha256"]:
        fail("R1 amendment changed the semantic oracle protocol")
    base_verifier_path = ROOT / text(base_oracle["verifier_path"], "base oracle verifier path")
    amended_verifier_path = ROOT / text(amended_oracle["verifier_path"], "amended oracle verifier path")
    if base_verifier_path == amended_verifier_path or not base_verifier_path.is_file() or not amended_verifier_path.is_file():
        fail("R1 amendment oracle paths are not distinct")
    base_text = base_verifier_path.read_text(encoding="utf-8")
    expected_text = base_text.replace(
        "\"\"\"Independent report/fixture oracle for the 0439 ODP append baseline.\n\n"
        "This verifier intentionally validates the serialized report and the corpus\n",
        "\"\"\"Independent report/fixture oracle for the 0439 ODP append baseline.\n\n"
        "This R1 amendment accepts the preserved baseline binary's null rustflags\n"
        "field and makes no frame-pointer build claim.\n\n"
        "This verifier intentionally validates the serialized report and the corpus\n",
        1,
    )
    expected_text = expected_text.replace(
        "    if environment.get('rustc_version') != 'rustc 1.98.1 (48a229cea 2026-09-01)' or environment.get('rustflags') != '-Cforce-frame-pointers=yes':\n"
        "        _fail('toolchain or frame-pointer build flags differ')\n",
        "    if environment.get('rustc_version') != 'rustc 1.98.1 (48a229cea 2026-09-01)' or environment.get('rustflags') is not None:\n"
        "        _fail('toolchain or unexpected build flags differ; R1 makes no frame-pointer build claim')\n",
        1,
    )
    if amended_verifier_path.read_text(encoding="utf-8") != expected_text:
        fail("R1 amended oracle differs outside the documented rustflags/doc/error change")
    if sha(amended_verifier_path) != amended_oracle["verifier_sha256"]:
        fail("R1 amended oracle hash differs from protocol")
    if amendment.get("scope") != "R1 retry and R2 current-build control":
        fail("R1 amendment scope differs")
    return base_path


def check_bindings(protocol: dict[str, Any], require_binaries: bool) -> tuple[dict[str, Any], dict[str, Any], Path, Path]:
    if protocol["schema"] != "litchi-0457-odp-existing-append-control-v1" or protocol["change"] != 457:
        fail("not the frozen 0457 control protocol")
    if protocol["capture_driver_sha256"] != sha(ROOT / "capture.py"):
        fail("capture driver hash differs from protocol")
    if protocol["cpu"] != 2 or protocol["workers"] != 1 or protocol["samples"] != 30 or protocol["warmups"] != 3:
        fail("protocol execution dimensions differ")
    if protocol["selector"] != "odp_existing_append_lifecycle" or protocol["shapes"] != {"tiny": 64, "medium": 4096, "large": 8192}:
        fail("protocol selector or shapes differ")
    role = protocol["roles"]["control"]
    build_path = ROOT / role["build_receipt"]["path"]
    source_path = ROOT / role["source_manifest"]["path"]
    binding_path = ROOT / role["binary_binding"]["path"]
    if sha(build_path) != role["build_receipt"]["sha256"] or sha(source_path) != role["source_manifest"]["sha256"] or sha(binding_path) != role["binary_binding"]["sha256"]:
        fail("control binding file hash differs")
    if build_path.read_bytes() != (BUNDLE / "checks" / "baseline-build.json").read_bytes():
        fail("copied build receipt differs from parent receipt")
    parent_source = BUNDLE / "sources" / (role["source_manifest"]["sha256"] + ".json")
    if not parent_source.is_file() or source_path.read_bytes() != parent_source.read_bytes():
        fail("copied source manifest differs from parent custody manifest")
    build = load(build_path)
    if build["change"] != 457 or build["revision"] != role["revision"]:
        fail("build receipt identity differs")
    build_cwd_text = text(build.get("cwd"), "build receipt cwd")
    build_cwd = Path(build_cwd_text)
    if not build_cwd.is_absolute() or build_cwd != build_cwd.resolve():
        fail("build receipt cwd must be an absolute canonical path")
    original_control_root = build_cwd / CONTROL_RELATIVE
    source = load(source_path)
    if len(source) != role["source_manifest"]["files"]:
        fail("source manifest file count differs")
    custody_path = (ROOT / protocol["custody"]["driver_path"]).resolve()
    if sha(custody_path) != protocol["custody"]["driver_sha256"]:
        fail("parent custody driver hash differs")
    if sha(ROOT / protocol["oracle"]["verifier_path"]) != protocol["oracle"]["verifier_sha256"] or sha(ROOT / protocol["oracle"]["protocol_path"]) != protocol["oracle"]["protocol_sha256"]:
        fail("copied oracle hash differs")
    bindings = load(binding_path)
    for mode in ("normal", "allocator"):
        binary = bindings["binaries"][mode]
        path = Path(binary["copy_path"])
        if require_binaries and (not path.is_file() or path.is_symlink() or path.stat().st_size != binary["bytes"] or sha(path) != binary["sha256"]):
            fail(f"{mode} control binary identity differs in precleanup mode")
    return role, bindings, build_cwd, original_control_root


def verify_phase(
    protocol: dict[str, Any],
    protocol_sha256: str,
    role: dict[str, Any],
    bindings: dict[str, Any],
    build_cwd: Path,
    original_control_root: Path,
    phase: str,
    attempt: str,
) -> int:
    directory = ROOT / "runs" / phase / attempt
    state = load(directory / "capture-state.json")
    expected_source = protocol["custody"]["expected_source_manifest"]
    if state["status"] != "pass" or state["completed_lanes"] != 6 or state["protocol_sha256"] != protocol_sha256 or state["driver_sha256"] != sha(ROOT / "capture.py"):
        fail(f"{phase}: capture state is not a successful bound capture")
    if state["source_before"] != expected_source or state["source_after"] != expected_source or not state["source_unchanged"] or not state["outside_bundle_status_unchanged"]:
        fail(f"{phase}: source or outside-bundle custody changed")
    check_ambient_allocator(state, f"{phase} capture state")
    lanes = expected_lanes(protocol, phase)
    receipt_paths = [directory / f"{phase}-{lane['mode']}-{lane['shape']}-{lane['repeat'].lower()}-receipt.json" for lane in lanes]
    if sorted(state["index"]) != sorted(str(path.relative_to(ROOT)) for path in receipt_paths):
        fail(f"{phase}: capture index differs")
    for lane, receipt_path in zip(lanes, receipt_paths):
        receipt = load(receipt_path)
        name = f"{phase}-{lane['mode']}-{lane['shape']}-{lane['repeat'].lower()}"
        if receipt["status"] != "pass" or receipt["phase"] != phase or receipt["attempt"] != attempt or receipt["name"] != name or receipt["lane"] != lane or receipt["selector"] != protocol["selector"] or receipt["revision"] != role["revision"]:
            fail(f"{name}: receipt identity differs")
        if receipt.get("cwd") != str(build_cwd):
            fail(f"{name}: captured cwd differs from authenticated build cwd")
        if receipt["protocol_sha256"] != protocol_sha256 or receipt["driver_sha256"] != sha(ROOT / "capture.py") or receipt["custody_driver_sha256"] != protocol["custody"]["driver_sha256"]:
            fail(f"{name}: receipt driver/protocol binding differs")
        if receipt["source_before"] != expected_source or receipt["source_after"] != expected_source or not receipt["source_unchanged"]:
            fail(f"{name}: source custody differs")
        check_ambient_allocator(receipt, f"{name} receipt")
        binary = bindings["binaries"][lane["mode"]]
        if receipt["binary"] != binary:
            fail(f"{name}: binary binding differs")
        report = directory / f"{name}.json"
        catalog = directory / f"{name}-catalog.json"
        workload_log = directory / f"{name}.log"
        resource_log = directory / f"{name}-resource.log"
        oracle_log = directory / f"{name}-oracle.log"
        captured_directory = original_control_root / "runs" / phase / attempt
        captured_report = captured_directory / f"{name}.json"
        captured_catalog = captured_directory / f"{name}-catalog.json"
        captured_resource_log = captured_directory / f"{name}-resource.log"
        expected_artifacts = {"report": report, "catalog": catalog, "workload_log": workload_log, "resource_log": resource_log, "oracle_log": oracle_log}
        if set(receipt["artifacts"]) != set(expected_artifacts):
            fail(f"{name}: artifact set differs")
        for key, path in expected_artifacts.items():
            if artifact(ROOT, receipt["artifacts"][key]) != path:
                fail(f"{name}: artifact path differs for {key}")
        cli = protocol["workload_cli"]
        expected_argv = ["taskset", "-c", "2", "/usr/bin/time", "-v", "-o", str(captured_resource_log), binary["copy_path"], cli["case_flag"], protocol["selector"], cli["shape_flag"], lane["shape"], cli["workers_flag"], "1", cli["samples_flag"], "30", cli["warmup_flag"], "3", cli["report_flag"], str(captured_report), cli["catalog_flag"], str(captured_catalog)]
        if receipt["argv"] != expected_argv or receipt["exit_code"] != 0 or receipt["oracle_exit_code"] != 0:
            fail(f"{name}: command or exit identity differs")
        interpreter = text(receipt["oracle_argv"][0], f"{name}.oracle_argv[0]")
        expected_oracle = oracle_command(protocol, captured_report, lane["mode"], lane["shape"], interpreter, original_control_root)
        if receipt["oracle_argv"] != expected_oracle:
            fail(f"{name}: oracle command differs")
        check_report(report, binary, lane, protocol, role["revision"])
        result = subprocess.run(oracle_command(protocol, report, lane["mode"], lane["shape"], sys.executable, ROOT), capture_output=True, text=True)
        if result.returncode != 0 or result.stdout.strip() != protocol["oracle"]["success_stdout"]:
            fail(f"{name}: independent copied oracle rejected report")
    return len(lanes) * protocol["samples"]


def authenticate_original_failure(
    amendment: dict[str, Any],
    amended_protocol: dict[str, Any],
    amended_role: dict[str, Any],
    amended_bindings: dict[str, Any],
    amended_build_cwd: Path,
    amended_control_root: Path,
    require_binaries: bool,
) -> dict[str, Any]:
    base_path = check_r1_oracle_amendment(amendment, amended_protocol)
    base_protocol = load(base_path)
    if "RUSTFLAGS" in base_protocol.get("environment", {}).get("fixed", {}) or "RUSTFLAGS" in amended_protocol.get("environment", {}).get("fixed", {}):
        fail("control protocol must not set RUSTFLAGS")
    base_role, base_bindings, base_build_cwd, base_control_root = check_bindings(base_protocol, require_binaries)
    if base_build_cwd != amended_build_cwd or base_control_root != amended_control_root:
        fail("amended and original build cwd bindings differ")
    directory = ROOT / "runs" / "R1" / "formal"
    state = load(directory / "capture-state.json")
    base_sha256 = sha(base_path)
    expected_source = base_protocol["custody"]["expected_source_manifest"]
    if (
        state["status"] != "failed"
        or state["phase"] != "R1"
        or state["attempt"] != "formal"
        or state["completed_lanes"] != 1
        or state["expected_lanes"] != 6
        or state["protocol_sha256"] != base_sha256
        or state["driver_sha256"] != sha(ROOT / "capture.py")
        or state["source_before"] != expected_source
        or state["source_after"] != expected_source
        or not state["source_unchanged"]
        or not state["outside_bundle_status_unchanged"]
    ):
        fail("original R1 failure is not an authenticated preserved capture")
    check_ambient_allocator(state, "original R1 failure capture state")
    lanes = expected_lanes(base_protocol, "R1")
    lane = lanes[0]
    name = f"R1-{lane['mode']}-{lane['shape']}-{lane['repeat'].lower()}"
    receipt_path = directory / f"{name}-receipt.json"
    expected_index = [str(receipt_path.relative_to(ROOT))]
    if state["index"] != expected_index:
        fail("original R1 failure index differs")
    receipt = load(receipt_path)
    if (
        receipt["status"] != "failed"
        or receipt["phase"] != "R1"
        or receipt["attempt"] != "formal"
        or receipt["name"] != name
        or receipt["lane"] != lane
        or receipt["selector"] != base_protocol["selector"]
        or receipt["revision"] != base_role["revision"]
        or receipt["protocol_sha256"] != base_sha256
        or receipt["driver_sha256"] != sha(ROOT / "capture.py")
        or receipt["custody_driver_sha256"] != base_protocol["custody"]["driver_sha256"]
        or receipt.get("cwd") != str(base_build_cwd)
        or receipt["source_before"] != expected_source
        or receipt["source_after"] != expected_source
        or not receipt["source_unchanged"]
        or receipt.get("exit_code") != 0
        or receipt.get("oracle_exit_code", 0) == 0
        or "oracle rejected" not in receipt.get("error", "")
    ):
        fail("original R1 failure receipt differs")
    check_ambient_allocator(receipt, "original R1 failure receipt")
    binary = base_bindings["binaries"][lane["mode"]]
    if receipt["binary"] != binary:
        fail("original R1 failure binary binding differs")
    report = directory / f"{name}.json"
    catalog = directory / f"{name}-catalog.json"
    workload_log = directory / f"{name}.log"
    resource_log = directory / f"{name}-resource.log"
    oracle_log = directory / f"{name}-oracle.log"
    captured_directory = base_control_root / "runs" / "R1" / "formal"
    captured_report = captured_directory / f"{name}.json"
    captured_catalog = captured_directory / f"{name}-catalog.json"
    captured_resource_log = captured_directory / f"{name}-resource.log"
    expected_artifacts = {"report": report, "catalog": catalog, "workload_log": workload_log, "resource_log": resource_log, "oracle_log": oracle_log}
    if set(receipt["artifacts"]) != set(expected_artifacts):
        fail("original R1 failure artifact set differs")
    for key, path in expected_artifacts.items():
        if artifact(ROOT, receipt["artifacts"][key]) != path:
            fail(f"original R1 failure artifact path differs for {key}")
    cli = base_protocol["workload_cli"]
    expected_argv = ["taskset", "-c", "2", "/usr/bin/time", "-v", "-o", str(captured_resource_log), binary["copy_path"], cli["case_flag"], base_protocol["selector"], cli["shape_flag"], lane["shape"], cli["workers_flag"], "1", cli["samples_flag"], "30", cli["warmup_flag"], "3", cli["report_flag"], str(captured_report), cli["catalog_flag"], str(captured_catalog)]
    if receipt["argv"] != expected_argv:
        fail("original R1 failure workload argv differs")
    interpreter = text(receipt["oracle_argv"][0], "original R1 failure oracle_argv[0]")
    if receipt["oracle_argv"] != oracle_command(base_protocol, captured_report, lane["mode"], lane["shape"], interpreter, base_control_root):
        fail("original R1 failure oracle argv differs")
    check_report(report, binary, lane, base_protocol, base_role["revision"])
    report_value = load(report)
    if report_value["environment"].get("rustflags") is not None:
        fail("original R1 failure report does not preserve null RUSTFLAGS")
    oracle_log_text = oracle_log.read_text(encoding="utf-8")
    if "INVALID: toolchain or frame-pointer build flags differ" not in oracle_log_text:
        fail("original R1 failure does not preserve the stale frame-pointer oracle rejection")
    original = subprocess.run(oracle_command(base_protocol, report, lane["mode"], lane["shape"], sys.executable, ROOT), capture_output=True, text=True)
    if original.returncode == 0 or "toolchain or frame-pointer build flags differ" not in original.stderr:
        fail("original R1 oracle no longer reproduces its recorded rejection")
    amended_result = subprocess.run(
        [
            sys.executable,
            "-B",
            str(ROOT / amended_protocol["oracle"]["verifier_path"]),
            "--report",
            str(report),
            "--mode",
            lane["mode"],
            "--shape",
            lane["shape"],
        ],
        capture_output=True,
        text=True,
    )
    if amended_result.returncode != 0 or amended_result.stdout.strip() != amended_protocol["oracle"]["success_stdout"]:
        fail("amended R1 oracle does not accept the preserved baseline report")
    return {"attempt": "formal", "name": name, "status": "authenticated-failure", "base_protocol_sha256": base_sha256}


def verify(attempt: str, require_binaries: bool, protocol_path: Path) -> dict[str, Any]:
    protocol_path = protocol_path.resolve()
    protocol = load(protocol_path)
    protocol_sha256 = sha(protocol_path)
    role, bindings, build_cwd, original_control_root = check_bindings(protocol, require_binaries)
    original_failure = None
    amendment = protocol.get("amendment")
    if amendment is not None:
        original_failure = authenticate_original_failure(amendment, protocol, role, bindings, build_cwd, original_control_root, require_binaries)
    samples = sum(verify_phase(protocol, protocol_sha256, role, bindings, build_cwd, original_control_root, phase, attempt) for phase in ("R1", "R2"))
    result = {"status": "pass", "change": 457, "phases": 2, "reports": 12, "samples": samples, "selector": protocol["selector"], "protocol_sha256": protocol_sha256, "claim": "owned current-revision control only"}
    if original_failure is not None:
        result["original_failure"] = original_failure
    return result


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--attempt", default="formal")
    parser.add_argument("--protocol", type=Path, default=ROOT / "protocol.json")
    parser.add_argument(
        "--precleanup",
        action="store_true",
        help="require the authenticated temporary control binaries to still exist",
    )
    args = parser.parse_args()
    print(json.dumps(verify(args.attempt, args.precleanup, args.protocol), sort_keys=True))
