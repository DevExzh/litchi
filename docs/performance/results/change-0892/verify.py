#!/usr/bin/env python3
"""Verify the retained 0892 current-source profiling packet.

The semantic report contract is owned by the 0885 producer baseline.  This
module loads that verifier through an import adapter so the current packet can
reuse the reviewed oracle without copying its large report checker.  Packet
receipts remain create-only and this module never launches a workload.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import importlib.util
import json
from pathlib import Path
import re
import stat
import sys
from typing import Any


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
BASELINE = PACKET.parent / "change-0885"
CASE = "xlsx_producer_medium_source_full_cell_scan"
PROFILE = "release"
CPU = 12
EVENT = "cycles:u"
FREQUENCY_HZ = 499
CALL_GRAPH = "dwarf,16384"
SOURCE_REVISION = "0a0957ef7249b4b16ec26ed81d1fb32b7288bb0b"
NATIVE_CHILDREN = 3
NATIVE_WARMUP = 10
NATIVE_SAMPLES = 2000
CALLGRIND_CHILDREN = 2
CALLGRIND_WARMUP = 0
CALLGRIND_SAMPLES = 10
TOGGLE_FUNCTION = "litchi_perf_baseline::producer_shape::run_xlsx_full_cell_scan"
NATIVE_SCHEMA = "litchi.xlsx-scan-cycle-profile-0892.v1"
CALLGRIND_SCHEMA = "litchi.xlsx-scan-callgrind-diagnostic-0892.v1"
CAPTURE_SCHEMA = "litchi.xlsx-scan-cycle-profile-capture-0892.v1"
CALLGRIND_CAPTURE_SCHEMA = "litchi.xlsx-scan-callgrind-capture-0892.v1"
VERIFICATION_SCHEMA = "litchi.xlsx-scan-profile-verification-0892.v1"
CHILD_ENV = {"PATH": "/usr/bin:/bin", "LC_ALL": "C", "LANG": "C", "TZ": "UTC"}


def utc_now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    path = path.resolve()
    info = path.stat()
    if not stat.S_ISREG(info.st_mode):
        raise ValueError(f"expected a regular file: {path}")
    return {
        "path": str(path),
        "bytes": info.st_size,
        "sha256": sha256(path),
        "mode": stat.S_IMODE(info.st_mode),
        "executable": bool(info.st_mode & stat.S_IXUSR),
    }


def read(path: Path) -> Any:
    with path.open(encoding="utf-8") as handle:
        return json.load(handle)


def write_once(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as handle:
        json.dump(value, handle, indent=2, sort_keys=True)
        handle.write("\n")


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def resolve(raw: Any, base: Path = PACKET) -> Path:
    require(isinstance(raw, str), "descriptor path is missing")
    value = Path(raw)
    if value.is_absolute():
        return value.resolve()
    if value.parts[:2] == ("docs", "performance"):
        return (ROOT / value).resolve()
    return (base / value).resolve()


def _load_0885_verifier() -> Any:
    """Load 0885's verifier under private module names.

    Its reviewed ``verify.py`` imports its sibling as ``capture``.  Install
    that sibling only while loading the adapter, then restore any caller
    module so ``change-0892/verify.py`` never collides with it.
    """

    capture_path = BASELINE / "capture.py"
    verify_path = BASELINE / "verify.py"
    capture_spec = importlib.util.spec_from_file_location("litchi_0885_capture_adapter", capture_path)
    require(capture_spec is not None and capture_spec.loader is not None, "cannot load 0885 capture helper")
    capture_module = importlib.util.module_from_spec(capture_spec)
    capture_spec.loader.exec_module(capture_module)
    old_capture = sys.modules.get("capture")
    sys.modules["capture"] = capture_module
    try:
        verify_spec = importlib.util.spec_from_file_location("litchi_0885_verify_adapter", verify_path)
        require(verify_spec is not None and verify_spec.loader is not None, "cannot load 0885 verify helper")
        verify_module = importlib.util.module_from_spec(verify_spec)
        sys.modules[verify_spec.name] = verify_module
        verify_spec.loader.exec_module(verify_module)
    finally:
        if old_capture is None:
            sys.modules.pop("capture", None)
        else:
            sys.modules["capture"] = old_capture
    return verify_module


_BASELINE_VERIFY = _load_0885_verifier()
if __name__ != "verify":
    sys.modules.setdefault("verify", sys.modules[__name__])


def validate_report(report: dict[str, Any], evidence: dict[str, Any], mode: str, lane: dict[str, Any]) -> dict[str, Any]:
    """Run the retained 0885 semantic oracle against one current child."""

    return _BASELINE_VERIFY._validate_report(report, evidence, mode, lane)


def binary_record(plan: dict[str, Any]) -> dict[str, Any]:
    value: Any = plan.get("binary")
    if value is None and isinstance(plan.get("binaries"), dict):
        value = plan["binaries"].get("native")
    require(isinstance(value, dict) and isinstance(value.get("path"), str), "plan binary descriptor is missing")
    return value


def plan_path(plan: dict[str, Any], packet: Path, key: str, default: str) -> Path:
    raw = plan.get(key, default)
    return resolve(raw, packet)


def validate_source_identity(plan: dict[str, Any], packet: Path) -> dict[str, Any] | None:
    raw = plan.get("source_identity")
    if raw is None:
        raw = plan.get("source_identity_receipt")
    if raw is None:
        return None
    if isinstance(raw, dict):
        path_value = raw.get("path")
        expected = raw
    else:
        path_value = raw
        expected = {}
    path = resolve(path_value, packet)
    actual = descriptor(path)
    for key in ("bytes", "sha256", "mode", "executable"):
        if key in expected:
            require(expected[key] == actual[key], f"source identity {key} mismatch")
    receipt = read(path)
    require(receipt.get("status") == "pass", "source identity receipt did not pass")
    require(receipt.get("source_revision") == SOURCE_REVISION, "source identity revision mismatch")
    return {"descriptor": actual, "receipt": receipt}


def _iter_descriptors(value: Any, prefix: str = ""):
    if isinstance(value, dict):
        if isinstance(value.get("path"), str) and isinstance(value.get("sha256"), str):
            yield prefix or "descriptor", value
        else:
            for key, child in value.items():
                name = f"{prefix}.{key}" if prefix else key
                yield from _iter_descriptors(child, name)
    elif isinstance(value, list):
        for index, child in enumerate(value):
            yield from _iter_descriptors(child, f"{prefix}[{index}]")


def validate_bindings(plan: dict[str, Any], packet: Path) -> list[dict[str, Any]]:
    bindings = plan.get("bindings")
    if bindings is None:
        return []
    require(isinstance(bindings, dict), "plan bindings must be an object")
    checked: list[dict[str, Any]] = []
    for name, item in _iter_descriptors(bindings):
        path = resolve(item["path"], packet)
        actual = descriptor(path)
        for key in ("bytes", "sha256", "mode", "executable"):
            if key in item:
                require(item[key] == actual[key], f"binding {name} {key} mismatch")
        checked.append({"name": name, **actual})
    return checked


def validate_binary(plan: dict[str, Any], packet: Path) -> dict[str, Any]:
    record = binary_record(plan)
    path = resolve(record["path"], packet)
    actual = descriptor(path)
    require(actual["executable"], "profile binary is not executable")
    for key in ("bytes", "sha256", "mode", "executable"):
        if key in record:
            require(record[key] == actual[key], f"binary {key} mismatch")
    build = record.get("build_receipt") or record.get("build")
    if isinstance(build, dict) and isinstance(build.get("path"), str):
        build_path = resolve(build["path"], packet)
        build_actual = descriptor(build_path)
        for key in ("bytes", "sha256", "mode", "executable"):
            if key in build:
                require(build[key] == build_actual[key], f"binary build receipt {key} mismatch")
        build_json = read(build_path)
        require(build_json.get("status") == "pass", "binary build receipt did not pass")
    return {"identity": actual, "build_receipt": build}


def _int(plan: dict[str, Any], *keys: str, default: int) -> int:
    for key in keys:
        if key in plan:
            value = plan[key]
            require(type(value) is int, f"plan {key} must be an integer")
            return value
    return default


def validate_plan(plan: dict[str, Any], kind: str) -> dict[str, Any]:
    require(isinstance(plan, dict), "plan must be a JSON object")
    expected_schema = NATIVE_SCHEMA if kind == "native" else CALLGRIND_SCHEMA
    require(plan.get("schema") == expected_schema, f"unexpected {kind} plan schema")
    require(plan.get("source_revision") == SOURCE_REVISION, "profile source revision mismatch")
    require(plan.get("case") == CASE, "plan case is not the producer full-cell case")
    require(plan.get("profile") == PROFILE, "profile plan requires release binary")
    require(_int(plan, "cpu", "cpu_affinity", default=CPU) == CPU, "profile must be pinned to CPU 12")
    cache = plan.get("cache", plan.get("filesystem_cache"))
    require(cache == "warm", "profile requires warm filesystem cache")
    require(plan.get("fresh_child", True) is True, "profile requires a fresh child per process")
    require(plan.get("fail_fast", True) is True, "profile capture must be fail-fast")
    elf_build_id = plan.get("elf_build_id", plan.get("build_id"))
    require(isinstance(elf_build_id, str) and re.fullmatch(r"[0-9a-fA-F]{40}", elf_build_id), "plan ELF build ID is missing or malformed")
    if kind == "native":
        require(plan.get("event") == EVENT, "native event must be cycles:u")
        require(plan.get("frequency_hz") == FREQUENCY_HZ, "native frequency must be 499 Hz")
        require(plan.get("call_graph") == CALL_GRAPH, "native call graph must be dwarf,16384")
        require(_int(plan, "children", default=NATIVE_CHILDREN) == NATIVE_CHILDREN, "native child count mismatch")
        require(_int(plan, "warmup", default=NATIVE_WARMUP) == NATIVE_WARMUP, "native warmup mismatch")
        require(_int(plan, "samples", default=NATIVE_SAMPLES) == NATIVE_SAMPLES, "native sample count mismatch")
        lane = {"children": NATIVE_CHILDREN, "warmup": NATIVE_WARMUP, "samples": NATIVE_SAMPLES}
    else:
        require(plan.get("toggle_function", TOGGLE_FUNCTION) == TOGGLE_FUNCTION, "Callgrind toggle function mismatch")
        require(_int(plan, "children", default=CALLGRIND_CHILDREN) == CALLGRIND_CHILDREN, "Callgrind child count mismatch")
        require(_int(plan, "warmup", default=CALLGRIND_WARMUP) == CALLGRIND_WARMUP, "Callgrind warmup mismatch")
        require(_int(plan, "samples", default=CALLGRIND_SAMPLES) == CALLGRIND_SAMPLES, "Callgrind sample count mismatch")
        lane = {"children": CALLGRIND_CHILDREN, "warmup": CALLGRIND_WARMUP, "samples": CALLGRIND_SAMPLES}
    packet = PACKET
    validate_bindings(plan, packet)
    source = validate_source_identity(plan, packet)
    require(source is not None, "profile plan must bind source identity")
    binary = validate_binary(plan, packet)
    source_binary = source["receipt"].get("binary")
    require(isinstance(source_binary, dict), "source identity binary binding is missing")
    for key in ("path", "bytes", "sha256"):
        expected = binary["identity"][key]
        actual = source_binary.get(key)
        if key == "path" and isinstance(actual, str):
            actual = str(Path(actual).resolve())
        require(actual == expected, f"source identity binary {key} does not match profile binary")
    return {"lane": lane, "binary": binary, "packet": packet}


def native_argv(binary: str, report: Path, evidence: Path, plan: dict[str, Any]) -> list[str]:
    return [
        "/usr/bin/perf", "record", "--no-buildid-cache", "-e", EVENT,
        "-F", str(FREQUENCY_HZ), "--call-graph", CALL_GRAPH,
        "-o", str(report.parent / "perf.data"), "--", "/usr/bin/taskset", "-c", str(CPU),
        binary, "--case", CASE, "--warmup", str(NATIVE_WARMUP), "--samples", str(NATIVE_SAMPLES),
        "--filesystem-cache", "warm", "--json", str(report), "--producer-evidence", str(evidence),
    ]


def callgrind_argv(binary: str, output: Path, report: Path, evidence: Path, plan: dict[str, Any]) -> list[str]:
    toggle = plan.get("toggle_function", TOGGLE_FUNCTION)
    return [
        "/usr/bin/taskset", "-c", str(CPU), "/usr/bin/valgrind", "--tool=callgrind",
        "--collect-atstart=no", "--toggle-collect=" + toggle,
        "--callgrind-out-file=" + str(output), binary,
        "--case", CASE, "--warmup", "0", "--samples", "10", "--filesystem-cache", "warm",
        "--json", str(report), "--producer-evidence", str(evidence),
    ]


def validate_child_semantics(folder: Path, mode: str, lane: dict[str, Any], binary: dict[str, Any] | None = None) -> dict[str, Any]:
    report_path = folder / "report.json"
    evidence_path = folder / "producer.json"
    require(report_path.is_file() and evidence_path.is_file(), f"{folder} semantic outputs missing")
    report = read(report_path)
    checked = validate_report(report, read(evidence_path), "native", lane)
    if binary is not None:
        identity = report.get("binary_identity")
        require(isinstance(identity, dict), f"{folder} report binary identity missing")
        expected_identity = {
            "path": str(Path(binary["path"]).resolve()),
            "binary_sha256": binary["sha256"],
            "binary_bytes": binary["bytes"],
            "mode_bits": binary["mode"],
            "executable": True,
            "profile": PROFILE,
        }
        for key, expected in expected_identity.items():
            actual = identity.get(key)
            if key == "path" and isinstance(actual, str):
                actual = str(Path(actual).resolve())
            require(actual == expected, f"{folder} report binary identity {key} mismatch")
    return checked


def validate_callgrind_totals(path: Path) -> int:
    require(path.is_file(), f"Callgrind output missing: {path}")
    totals = re.findall(r"^totals:\s+(\d+)", path.read_text(encoding="utf-8", errors="replace"), re.MULTILINE)
    require(totals and int(totals[-1]) > 0, "Callgrind collected no guest instructions")
    return int(totals[-1])


def validate_capture_receipt(path: Path, expected_schema: str, plan_path: Path) -> dict[str, Any]:
    receipt = read(path)
    require(receipt.get("schema") == expected_schema, f"capture receipt schema mismatch: {path.name}")
    require(receipt.get("status") == "pass", f"capture receipt did not pass: {path}")
    require(receipt.get("plan_sha256") == sha256(plan_path), f"capture receipt plan binding mismatch: {path}")
    return receipt


def check_artifacts(folder: Path, command: dict[str, Any]) -> None:
    for name, identity in command.get("artifacts", {}).items():
        path = folder / name
        require(path.is_file(), f"capture artifact missing: {path}")
        actual = descriptor(path)
        for key in ("bytes", "sha256", "mode", "executable"):
            require(identity.get(key) == actual[key], f"capture artifact {path} {key} mismatch")


def check_command(
    folder: Path,
    command: dict[str, Any],
    *,
    expected_argv: list[str],
    expected_mode: str,
    expected_index: int,
    expected_binary: dict[str, Any],
    plan_digest: str,
) -> None:
    require(command.get("mode") == expected_mode, f"{folder} command mode mismatch")
    require(command.get("child_index") == expected_index, f"{folder} command index mismatch")
    require(command.get("argv") == expected_argv, f"{folder} command argv mismatch")
    require(command.get("cwd") == str(PACKET), f"{folder} command cwd mismatch")
    require(command.get("environment") == CHILD_ENV, f"{folder} command environment mismatch")
    require(command.get("profile") == PROFILE, f"{folder} command profile mismatch")
    require(command.get("plan_sha256") == plan_digest, f"{folder} command plan binding mismatch")
    actual_binary = command.get("binary")
    require(isinstance(actual_binary, dict), f"{folder} command binary descriptor missing")
    for key in ("path", "bytes", "sha256", "mode", "executable"):
        if key in expected_binary:
            require(actual_binary.get(key) == expected_binary[key], f"{folder} command binary {key} mismatch")
    start_path = folder / "command-start.json"
    require(start_path.is_file(), f"{folder} command-start receipt missing")
    started = read(start_path)
    for key in ("mode", "child_index", "argv", "cwd", "environment", "profile", "binary", "plan_sha256"):
        require(started.get(key) == command.get(key), f"{folder} command-start {key} mismatch")
    require(isinstance(command.get("started_utc"), str) and isinstance(command.get("finished_utc"), str), f"{folder} command timestamps missing")


def verify_exports(plan: dict[str, Any], capture_root: Path, binary: dict[str, Any]) -> dict[str, Any]:
    expected_build_id = plan.get("elf_build_id", plan.get("build_id"))
    exporter = _load_local_helper("export.py", "litchi_0892_export_replay")
    expected_outputs = set(exporter.COMMANDS)
    rows: list[dict[str, Any]] = []
    for index in range(NATIVE_CHILDREN):
        folder = capture_root / f"child-{index:02}"
        receipt_path = folder / "exports.json"
        require(receipt_path.is_file(), f"native export receipt missing: {receipt_path}")
        receipt = read(receipt_path)
        require(receipt.get("status") == "pass", f"native export receipt did not pass: {receipt_path}")
        require(receipt.get("child") == folder.name, f"native export child mismatch: {folder}")
        perf_data = folder / "perf.data"
        actual_perf = descriptor(perf_data)
        declared_perf = receipt.get("perf_data")
        require(isinstance(declared_perf, dict), f"native export perf descriptor missing: {folder}")
        for key in ("bytes", "sha256", "mode", "executable"):
            require(declared_perf.get(key) == actual_perf[key], f"native export perf descriptor {key} mismatch: {folder}")
        command = read(folder / "command.json")
        artifact = command.get("artifacts", {}).get("perf.data")
        require(isinstance(artifact, dict), f"native command perf artifact missing: {folder}")
        require(artifact.get("sha256") == actual_perf["sha256"] and artifact.get("bytes") == actual_perf["bytes"], f"native perf artifact differs from export: {folder}")
        export_rows = receipt.get("records")
        require(isinstance(export_rows, list) and len(export_rows) == 5, f"native export record count mismatch: {folder}")
        output_names = [record.get("output") for record in export_rows]
        require(len(set(output_names)) == 5 and set(output_names) == expected_outputs, f"native export output inventory mismatch: {folder}")
        for record in export_rows:
            require(record.get("exit_code") == 0, f"native export command failed: {folder}/{record.get('output')}")
            output = folder / str(record.get("output"))
            stderr = folder / str(record.get("stderr"))
            require(output.is_file() and stderr.is_file(), f"native export output missing: {output}")
            expected_argv = ["/usr/bin/perf", *exporter.COMMANDS[output.name], "-i", str(perf_data)]
            require(record.get("argv") == expected_argv, f"native export argv mismatch: {output}")
            for field, path in (("output_identity", output), ("stderr_identity", stderr)):
                declared = record.get(field)
                require(isinstance(declared, dict), f"native export {field} missing: {output}")
                actual = descriptor(path)
                for key in ("bytes", "sha256", "mode", "executable"):
                    require(declared.get(key) == actual[key], f"native export {field} {key} mismatch: {output}")
            if output.name == "buildids.txt":
                text = output.read_text(encoding="utf-8", errors="replace")
                expected_path = str(Path(binary["path"]).resolve())
                require(any(expected_build_id.lower() in line.lower() and expected_path in line for line in text.splitlines()), f"native build ID/path pair is absent from {output}")
        rows.append({"child": folder.name, "exports": receipt_path.name})
    return {"status": "pass", "children": rows, "build_id": expected_build_id}

def _load_local_helper(filename: str, module_name: str) -> Any:
    spec = importlib.util.spec_from_file_location(module_name, PACKET / filename)
    require(spec is not None and spec.loader is not None, f"cannot load packet helper: {filename}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


def verify_self_analysis(capture_root: Path, path: Path, baseline_path: Path) -> dict[str, Any]:
    require(path.is_file(), f"self-analysis output missing: {path}")
    result = read(path)
    require(result.get("status") == "pass", "self-analysis did not pass")
    reducer = _load_local_helper("reduce.py", "litchi_0892_reduce_replay")
    rows = result.get("rows")
    require(isinstance(rows, list) and len(rows) == NATIVE_CHILDREN, "self-analysis row count mismatch")
    replayed: list[dict[str, Any]] = []
    for index, folder in enumerate(sorted(capture_root.glob("child-*"))):
        expected = reducer.parse_child(folder)
        actual = rows[index]
        require(actual.get("child") == expected["child"], f"self-analysis child order mismatch: {folder}")
        for key in ("sample_count", "total_period", "lost_samples", "empty_stack_percent", "any_unknown_stack_percent", "self_groups", "top_self_symbols", "input_sha256"):
            require(actual.get(key) == expected[key], f"self-analysis replay mismatch: {folder}/{key}")
        replayed.append({"child": folder.name, "total_period": expected["total_period"]})
    comparison = result.get("comparison_to_0886")
    require(isinstance(comparison, dict), "self-analysis baseline comparison missing")
    baseline = comparison.get("baseline")
    require(isinstance(baseline, dict) and Path(str(baseline.get("path"))).resolve() == baseline_path.resolve(), "self-analysis baseline path mismatch")
    require(baseline.get("sha256") == sha256(baseline_path), "self-analysis baseline hash mismatch")
    computed = reducer._comparison([reducer.parse_child(folder) for folder in sorted(capture_root.glob("child-*"))], baseline_path)
    require(comparison == computed, "self-analysis baseline comparison does not replay")
    return {"status": "pass", "rows": replayed, "path": str(path.resolve()), "sha256": sha256(path)}


def verify_callgrind_analysis(capture_root: Path, path: Path, baseline_path: Path) -> dict[str, Any]:
    require(path.is_file(), f"Callgrind analysis output missing: {path}")
    result = read(path)
    require(result.get("status") == "pass", "Callgrind analysis did not pass")
    analyzer = _load_local_helper("analyze_callgrind.py", "litchi_0892_callgrind_replay")
    rows = result.get("rows")
    require(isinstance(rows, list) and len(rows) == CALLGRIND_CHILDREN, "Callgrind analysis row count mismatch")
    replayed: list[dict[str, Any]] = []
    for index, folder in enumerate(sorted(capture_root.glob("callgrind-*"))):
        canonical = folder / "inclusive-canonical.txt"
        annotation_command = folder / "annotate-command.json"
        require(canonical.is_file() and annotation_command.is_file(), f"Callgrind annotation custody missing: {folder}")
        command = read(annotation_command)
        expected_argv = ["/usr/bin/callgrind_annotate", "--inclusive=yes", "--threshold=100", "--auto=no", str(folder / "callgrind.out")]
        require(command.get("argv") == expected_argv and command.get("exit_code") == 0, f"Callgrind annotation command mismatch: {folder}")
        for field, item in (("output", canonical), ("stderr", folder / "annotate.stderr")):
            declared = command.get(field)
            require(isinstance(declared, dict), f"Callgrind annotation {field} descriptor missing: {folder}")
            actual = descriptor(item)
            for key in ("bytes", "sha256", "mode", "executable"):
                require(declared.get(key) == actual[key], f"Callgrind annotation {field} {key} mismatch: {folder}")
        expected = analyzer.parse_child(folder)
        actual = rows[index]
        require(actual.get("child") == expected["child"], f"Callgrind analysis child order mismatch: {folder}")
        for key in ("total_ir", "selected_inclusive_rows", "parent_pid", "brk_warning", "callgrind_sha256", "annotation_sha256"):
            require(actual.get(key) == expected[key], f"Callgrind analysis replay mismatch: {folder}/{key}")
        replayed.append({"child": folder.name, "total_ir": expected["total_ir"]})
    comparison = result.get("comparison_to_0886")
    require(isinstance(comparison, dict), "Callgrind baseline comparison missing")
    baseline = comparison.get("baseline")
    require(isinstance(baseline, dict) and Path(str(baseline.get("path"))).resolve() == baseline_path.resolve(), "Callgrind baseline path mismatch")
    require(baseline.get("sha256") == sha256(baseline_path), "Callgrind baseline hash mismatch")
    computed = analyzer.compare([analyzer.parse_child(folder) for folder in sorted(capture_root.glob("callgrind-*"))], baseline_path)
    require(comparison == computed, "Callgrind baseline comparison does not replay")
    return {"status": "pass", "rows": replayed, "path": str(path.resolve()), "sha256": sha256(path)}


def verify_packet(
    plan_path: Path,
    capture_root: Path,
    callgrind_root: Path | None,
    output: Path,
    self_analysis_path: Path,
    callgrind_analysis_path: Path | None,
) -> dict[str, Any]:
    require(not output.exists(), f"verification output already exists: {output}")
    plan = read(plan_path)
    native = validate_plan(plan, "native")
    native_capture = validate_capture_receipt(capture_root / "capture.json", CAPTURE_SCHEMA, plan_path)
    rows: list[dict[str, Any]] = []
    native_records = native_capture.get("records", native_capture.get("children", []))
    require(isinstance(native_records, list) and len(native_records) == NATIVE_CHILDREN, "native child receipt count mismatch")
    for index, item in enumerate(native_records):
        relative = item.get("path") if isinstance(item, dict) else None
        folder = capture_root / (relative if isinstance(relative, str) else f"child-{index:02}")
        if folder.name in ("receipt.json", "command.json"):
            folder = folder.parent
        command_path = folder / "command.json"
        require(command_path.is_file(), f"native command receipt missing: {command_path}")
        command = read(command_path)
        require(command.get("exit_code") == 0, f"native child {index} failed")
        require(item.get("command_sha256") == sha256(command_path), f"native child command digest mismatch: {folder}")
        check_command(
            folder,
            command,
            expected_argv=native_argv(native["binary"]["identity"]["path"], folder / "report.json", folder / "producer.json", plan),
            expected_mode="native",
            expected_index=index,
            expected_binary=native["binary"]["identity"],
            plan_digest=sha256(plan_path),
        )
        check_artifacts(folder, command)
        checked = validate_child_semantics(folder, "native", native["lane"], native["binary"]["identity"])
        rows.append({"child": folder.name, "stats": checked["elapsed"]})
    call_rows: list[dict[str, Any]] = []
    if callgrind_root is not None:
        callgrind = validate_plan(read(plan_path.parent / "callgrind-plan.json"), "callgrind") if (plan_path.parent / "callgrind-plan.json").is_file() else validate_plan(plan.get("callgrind", {}), "callgrind")
        call_capture = validate_capture_receipt(callgrind_root / "capture.json", CALLGRIND_CAPTURE_SCHEMA, plan_path.parent / "callgrind-plan.json" if (plan_path.parent / "callgrind-plan.json").is_file() else plan_path)
        call_records = call_capture.get("records", call_capture.get("children", []))
        require(isinstance(call_records, list) and len(call_records) == CALLGRIND_CHILDREN, "Callgrind child receipt count mismatch")
        for index, item in enumerate(call_records):
            relative = item.get("path") if isinstance(item, dict) else None
            folder = callgrind_root / (relative if isinstance(relative, str) else f"callgrind-{index:02}")
            if folder.name in ("receipt.json", "command.json"):
                folder = folder.parent
            command_path = folder / "command.json"
            require(command_path.is_file(), f"Callgrind command receipt missing: {command_path}")
            command = read(command_path)
            require(command.get("exit_code") == 0, f"Callgrind child {index} failed")
            require(item.get("command_sha256") == sha256(command_path), f"Callgrind child command digest mismatch: {folder}")
            call_plan_path = plan_path.parent / "callgrind-plan.json" if (plan_path.parent / "callgrind-plan.json").is_file() else plan_path
            check_command(
                folder,
                command,
                expected_argv=callgrind_argv(callgrind["binary"]["identity"]["path"], folder / "callgrind.out", folder / "report.json", folder / "producer.json", read(call_plan_path)),
                expected_mode="callgrind",
                expected_index=index,
                expected_binary=callgrind["binary"]["identity"],
                plan_digest=sha256(call_plan_path),
            )
            check_artifacts(folder, command)
            checked = validate_child_semantics(folder, "callgrind", callgrind["lane"], callgrind["binary"]["identity"])
            total = validate_callgrind_totals(folder / "callgrind.out")
            call_rows.append({"child": folder.name, "total_ir": total, "stats": checked["elapsed"]})
    source = validate_source_identity(plan, plan_path.parent)
    require(source is not None, "source identity receipt is required")
    source_binary = source["receipt"].get("binary")
    require(isinstance(source_binary, dict), "source identity binary binding is missing")
    for key in ("path", "bytes", "sha256"):
        expected = native["binary"]["identity"][key]
        actual = source_binary.get(key)
        if key == "path" and isinstance(actual, str):
            actual = str(Path(actual).resolve())
        require(actual == expected, f"source identity binary {key} does not match profile binary")
    exports = verify_exports(plan, capture_root, native["binary"]["identity"])
    self_analysis = verify_self_analysis(capture_root, self_analysis_path, plan_path.parent.parent / "change-0886" / "self-analysis.json")
    callgrind_analysis = None
    if callgrind_root is not None:
        require(callgrind_analysis_path is not None, "Callgrind analysis path is required with Callgrind capture")
        callgrind_analysis = verify_callgrind_analysis(callgrind_root, callgrind_analysis_path, plan_path.parent.parent / "change-0886" / "callgrind-analysis.json")
    result = {
        "schema": VERIFICATION_SCHEMA,
        "status": "pass",
        "plan": str(plan_path.resolve()),
        "plan_sha256": sha256(plan_path),
        "source_identity": source,
        "binary": native["binary"],
        "exports": exports,
        "self_analysis": self_analysis,
        "native": {"children": rows, "measured_scans": NATIVE_CHILDREN * NATIVE_SAMPLES},
        "callgrind": {"children": call_rows, "measured_scans": CALLGRIND_CHILDREN * CALLGRIND_SAMPLES, "analysis": callgrind_analysis} if callgrind_root is not None else None,
        "production_changes": False,
        "performance_claim": False,
        "scope": "Custody and semantic replay only; native self-IP and guest-instruction diagnostics remain separate and descriptive.",
        "verified_utc": utc_now(),
    }
    write_once(output, result)
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--capture-root", type=Path, required=True)
    parser.add_argument("--callgrind-root", type=Path)
    parser.add_argument("--self-analysis", type=Path, default=PACKET / "self-analysis.json")
    parser.add_argument("--callgrind-analysis", type=Path, default=PACKET / "callgrind-analysis.json")
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    result = verify_packet(args.plan.resolve(), args.capture_root.resolve(), args.callgrind_root.resolve() if args.callgrind_root else None, args.output.resolve(), args.self_analysis.resolve(), args.callgrind_analysis.resolve() if args.callgrind_analysis else None)
    print(json.dumps(result, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
