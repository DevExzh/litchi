#!/usr/bin/env python3
"""Relocatable semantic adapter for the 0953 FP native profile.

This module validates one report against the corrected 0950 semantic oracle,
while taking every packet and external immutable input from a 0953 plan.  It
never starts a workload.  Packet-local descriptors are resolved below the
``--packet`` root; external absolute descriptors are permitted only for the
frozen oracle/source/binary chain and are checked by digest, size, mode, and
regular-file status.

The production contract is deliberately narrow: 200 measured samples, zero
warmups, the bound XLSX case and shape, and an uninstrumented release report.
A test-only fixture override is not exposed by the 0953 command line.
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import importlib.util
import json
import math
import os
from pathlib import Path
import stat
import sys
from typing import Any

DEFAULT_PACKET = Path(__file__).resolve().parents[1]
DEFAULT_PLAN_NAME = "plan-0953-v1.json"
EXPECTED_SCHEMA = "litchi.performance.0953.native-publish-sampling-plan.v1"
EXPECTED_CASE = "xlsx_source_backed_cell_values_one_percent_edit_save"
EXPECTED_SAMPLES = 200
EXPECTED_WARMUP = 0
EXPECTED_EVENT = "cycles:u"
EXPECTED_CPU = 12


class AdapterError(AssertionError):
    """A fail-closed semantic or custody mismatch."""


def check(condition: bool, message: str) -> None:
    if not condition:
        raise AdapterError(message)


def sha256(path: Path) -> str:
    check(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def _no_symlink_component(path: Path) -> None:
    current = path if path.is_absolute() else Path.cwd() / path
    parts = current.parts
    cursor = Path(parts[0]) if parts and parts[0] == "/" else Path()
    for part in parts[1:] if cursor == Path("/") else parts:
        cursor = cursor / part
        if cursor.is_symlink():
            raise AdapterError(f"symlink path component is not allowed: {cursor}")


def resolve_path(raw: Any, packet: Path, label: str) -> Path:
    check(isinstance(raw, str) and raw and "\x00" not in raw,
          f"{label} path is invalid")
    candidate = Path(raw)
    if candidate.is_absolute():
        path = candidate
    else:
        path = packet / candidate
        try:
            path.resolve().relative_to(packet.resolve())
        except ValueError as exc:
            raise AdapterError(f"{label} relative path escapes packet: {raw}") from exc
    _no_symlink_component(path)
    resolved = path.resolve()
    check(resolved.is_file() and not resolved.is_symlink(),
          f"{label} is missing or not a regular file: {resolved}")
    return resolved


def descriptor(path: Path) -> dict[str, Any]:
    path = path.resolve()
    check(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    mode = stat.S_IMODE(path.stat().st_mode)
    return {"path": str(path), "bytes": path.stat().st_size,
            "sha256": sha256(path), "mode": mode,
            "executable": bool(mode & stat.S_IXUSR)}


def _unique_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    output: dict[str, Any] = {}
    for key, value in pairs:
        check(key not in output, f"duplicate JSON key: {key}")
        output[key] = value
    return output


def read_json(path: Path) -> Any:
    check(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=_unique_pairs)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise AdapterError(f"invalid JSON {path}: {error}") from error


def load_plan(packet: Path, plan_path: Path | None = None) -> tuple[dict[str, Any], Path]:
    packet = packet.resolve()
    raw_path = plan_path if plan_path is not None else packet / DEFAULT_PLAN_NAME
    path = raw_path if raw_path.is_absolute() else packet / raw_path
    path = path.resolve()
    try:
        path.relative_to(packet)
    except ValueError as exc:
        raise AdapterError(f"plan is outside packet: {path}") from exc
    value = read_json(path)
    check(isinstance(value, dict), "plan is not an object")
    check(value.get("schema") == EXPECTED_SCHEMA,
          f"plan schema differs: {value.get('schema')!r}")
    return value, path


def verify_descriptor(value: Any, packet: Path, label: str) -> tuple[Path, dict[str, Any]]:
    check(isinstance(value, dict), f"{label} descriptor is missing")
    path = resolve_path(value.get("path"), packet, label)
    actual = descriptor(path)
    for key in ("bytes", "sha256", "mode", "executable"):
        check(value.get(key) == actual[key], f"{label} descriptor {key} differs")
    return path, actual


def _contract(plan: dict[str, Any]) -> dict[str, Any]:
    contract = plan.get("contract")
    if contract is None:
        contract = plan
    check(isinstance(contract, dict), "plan contract is missing")
    samples = contract.get("samples")
    warmup = contract.get("warmup")
    case = contract.get("case")
    cpu = contract.get("cpu")
    event = contract.get("event", EXPECTED_EVENT)
    check(samples == EXPECTED_SAMPLES and warmup == EXPECTED_WARMUP,
          "0953 semantic contract requires 200 samples and zero warmups")
    check(case == EXPECTED_CASE, "0953 case contract differs")
    check(cpu == EXPECTED_CPU, "0953 CPU contract differs")
    check(event == EXPECTED_EVENT, "0953 perf event contract differs")
    shapes = contract.get("shapes")
    if shapes is None:
        shapes = sorted({row.get("shape") for row in plan.get("schedule", [])})
    check(isinstance(shapes, list) and shapes and all(isinstance(x, str) and x for x in shapes),
          "0953 shape contract is missing")
    check(len(set(shapes)) == len(shapes), "0953 shape contract is duplicated")
    return {"samples": samples, "warmup": warmup, "case": case,
            "cpu": cpu, "event": event, "shapes": tuple(shapes)}


def _binding_descriptors(plan: dict[str, Any], packet: Path) -> dict[str, tuple[Path, dict[str, Any]]]:
    names = ("binary", "runner")
    result: dict[str, tuple[Path, dict[str, Any]]] = {}
    for name in names:
        result[name] = verify_descriptor(plan.get(name), packet, f"plan {name}")
    for name in ("sample_exporter", "mmap_exporter"):
        if name in plan:
            result[name] = verify_descriptor(plan.get(name), packet, f"plan {name}")
    semantic = plan.get("semantic_binding")
    check(isinstance(semantic, dict), "semantic_binding is missing")
    for name in ("oracle", "oracle_original", "correction", "context", "source_manifest"):
        result[f"semantic_{name}"] = verify_descriptor(
            semantic.get(name), packet, f"semantic {name}")
    source = plan.get("source_binding")
    check(isinstance(source, dict), "source_binding is missing")
    manifest = source.get("manifest")
    check(isinstance(manifest, dict), "source manifest binding is missing")
    check(manifest.get("sha256") == result["semantic_source_manifest"][1]["sha256"],
          "source and semantic manifest bindings differ")
    return result


def _load_oracle(path: Path) -> Any:
    spec = importlib.util.spec_from_file_location("semantic_0950_v2_bound_0953", path)
    check(spec is not None and spec.loader is not None, "corrected semantic oracle cannot load")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    required = ("validate_corpus", "validate_sink", "check_reported_statistics",
                "validate_source", "stats", "close", "PHASES", "ALL_PHASES")
    for name in required:
        check(hasattr(module, name), f"semantic oracle lacks {name}")
    return module


def verify_bindings(packet: Path, plan_path: Path | None = None) -> tuple[dict[str, Any], Any, dict[str, Any], Path]:
    plan, plan_file = load_plan(packet, plan_path)
    contract = _contract(plan)
    bindings = _binding_descriptors(plan, packet)
    oracle = _load_oracle(bindings["semantic_oracle"][0])
    return ({name: actual for name, (_, actual) in bindings.items()}, oracle, contract, plan_file)


def validate_terminal(path: Path, label: str) -> dict[str, Any]:
    terminal = read_json(path)
    check(isinstance(terminal, dict), f"{label} terminal is not an object")
    check(terminal.get("returncode") == 0 and terminal.get("timed_out") is False,
          f"{label} terminal did not pass")
    check(terminal.get("launch_error") is None
          and terminal.get("process_group_clear") is True
          and terminal.get("terminal_before_parse") is True,
          f"{label} terminal custody did not pass")
    return terminal


def validate_child_receipt(receipt_path: Path, report_path: Path, shape: str,
                           packet: Path, plan: dict[str, Any], binary_sha: str) -> dict[str, Any]:
    receipt = read_json(receipt_path)
    expected_schema = plan.get("child_receipt_schema", "litchi.performance.0953.child-receipt.v1")
    check(receipt.get("schema") == expected_schema, "child receipt schema differs")
    check(receipt.get("status") == "raw_captured" and receipt.get("retry") is False,
          "child receipt is not admitted raw evidence")
    check(receipt.get("row", {}).get("shape") == shape, "child receipt shape differs")
    check(receipt.get("binary", {}).get("sha256") == binary_sha,
          "child receipt binary differs")
    names = plan.get("child_artifacts", {})
    record_name = names.get("record_terminal", "record-terminal.json")
    stdout_name = names.get("record_stdout", "record.stdout")
    stderr_name = names.get("record_stderr", "record.stderr")
    for key, name in (("report", names.get("report", "report.json")),
                      ("perf_data", names.get("perf_data", "perf.data")),
                      ("record_terminal", record_name),
                      ("record_stdout", stdout_name),
                      ("record_stderr", stderr_name)):
        actual = descriptor(receipt_path.parent / name)
        given = receipt.get(key)
        if given is not None:
            for field in ("bytes", "sha256", "mode", "executable"):
                check(given.get(field) == actual[field], f"child {key} descriptor {field} differs")
    validate_terminal(receipt_path.parent / record_name, "record")
    return receipt


def validate_report(report_path: Path, shape: str, packet: Path, plan: dict[str, Any],
                    contract: dict[str, Any], oracle: Any, *,
                    receipt_path: Path | None = None) -> dict[str, Any]:
    check(shape in contract["shapes"], f"unsupported 0953 shape: {shape}")
    binary_desc = plan["binary"]
    binary_sha = binary_desc["sha256"]
    binary_bytes = binary_desc["bytes"]
    if receipt_path is not None:
        validate_child_receipt(receipt_path, report_path, shape, packet, plan, binary_sha)
    report = read_json(report_path)
    check(isinstance(report, dict), "report is not an object")
    tool = report.get("tool", {})
    check(tool.get("profile") == "release" and tool.get("instrumentation") == "none",
          "report is not uninstrumented release evidence")
    identity = report.get("binary_identity", {})
    check(identity.get("binary_sha256") == binary_sha
          and identity.get("binary_bytes") == binary_bytes,
          "report binary identity differs from plan")
    configuration = report.get("configuration", {})
    check(configuration.get("samples_per_case") == contract["samples"]
          and configuration.get("warmup_iterations_per_case") == contract["warmup"],
          "report sample/warmup configuration differs")
    check(configuration.get("cases") == [contract["case"]]
          and configuration.get("xlsx_cell_crud_shapes") == [shape],
          "report case/shape configuration differs")
    check(report.get("environment", {}).get("cpu_affinity") == str(contract["cpu"]),
          "report CPU affinity differs")
    results = report.get("results")
    check(isinstance(results, list) and len(results) == 1,
          "report must contain exactly one result")
    result = results[0]
    check(result.get("case") == contract["case"], "report result case differs")
    corpus = oracle.validate_corpus(result.get("corpus", {}), shape)
    oracle.validate_sink(result.get("sink"), str(report_path))
    elapsed = result.get("elapsed_ns")
    check(isinstance(elapsed, dict), "report elapsed summary is missing")
    measured = oracle.check_reported_statistics(elapsed, contract["samples"])
    phase_vectors, constants = oracle.validate_source(
        result, corpus, shape, contract["samples"], compatibility=False)
    phases = {phase: oracle.stats(phase_vectors[phase]) for phase in oracle.ALL_PHASES}
    denominator = sum(elapsed["samples"])
    check(denominator > 0, "report elapsed denominator is zero")
    shares = {phase: sum(phase_vectors[phase]) / denominator for phase in oracle.PHASES}
    oracle.close(sum(shares.values()), 1.0, "aggregate phase shares")
    return {
        "status": "pass", "shape": shape, "samples": contract["samples"],
        "warmup": contract["warmup"], "report": descriptor(report_path),
        "binary_sha256": binary_sha, "elapsed_ns": measured,
        "phase_statistics": phases, "aggregate_phase_shares": shares,
        "output_sha256": constants["output_sha256"],
        "semantic_sha256": constants["semantic_sha256"],
        "uncertainty": "Descriptive semantic validation only; no native timing, speedup, allocation, RSS, or adoption claim.",
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="action", required=True)
    bindings = sub.add_parser("bindings")
    bindings.add_argument("--packet", type=Path, default=DEFAULT_PACKET)
    bindings.add_argument("--plan", type=Path)
    validate = sub.add_parser("validate")
    validate.add_argument("--packet", type=Path, default=DEFAULT_PACKET)
    validate.add_argument("--plan", type=Path)
    validate.add_argument("--report", type=Path, required=True)
    validate.add_argument("--shape", required=True)
    validate.add_argument("--receipt", type=Path)
    args = parser.parse_args()
    try:
        packet = args.packet.resolve()
        bindings_value, oracle, contract, plan_file = verify_bindings(packet, args.plan)
        plan = read_json(plan_file)
        if args.action == "bindings":
            print(json.dumps({"schema": EXPECTED_SCHEMA, "plan": descriptor(plan_file),
                              "contract": contract, "bindings": bindings_value},
                             indent=2, sort_keys=True))
        else:
            report = args.report if args.report.is_absolute() else packet / args.report
            receipt = None if args.receipt is None else (args.receipt if args.receipt.is_absolute() else packet / args.receipt)
            result = validate_report(report.resolve(), args.shape, packet, plan, contract, oracle,
                                     receipt_path=None if receipt is None else receipt.resolve())
            print(json.dumps(result, indent=2, sort_keys=True))
        return 0
    except (AdapterError, AssertionError, OSError, TypeError, ValueError, ImportError) as error:
        print(f"semantic_0953_v1: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
