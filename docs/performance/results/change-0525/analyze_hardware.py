#!/usr/bin/env python3
"""Validate the 0525 whole-child ``perf stat`` diagnostic captures.

The counters are deliberately kept at whole-child scope.  This report checks
event coverage, receipt custody, and logical result parity with the matched
native child; it does not turn whole-child counters into operation-local
measurements or a speedup claim.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import importlib.util
import json
from pathlib import Path
import sys
from typing import Any


HERE = Path(__file__).resolve().parent
PLAN_PATH = HERE / "plan.json"
NUMERIC_PATH = HERE / "analyze.py"
GROUP = ("cycles", "instructions", "branches", "branch-misses")
EVENTS = GROUP + ("page-faults", "context-switches", "cpu-migrations")


class EvidenceError(ValueError):
    """A missing, malformed, or contradictory evidence artifact."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def sha256(path: Path) -> str:
    try:
        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()
    except OSError as error:
        raise EvidenceError(f"cannot hash {path}: {error}") from error


def read_json(path: Path) -> Any:
    try:
        with path.open(encoding="utf-8") as stream:
            return json.load(stream)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read JSON {path}: {error}") from error


def load_numeric() -> Any:
    require(NUMERIC_PATH.is_file(), f"missing numerical helper {NUMERIC_PATH}")
    spec = importlib.util.spec_from_file_location("xlsx_0525_numeric", NUMERIC_PATH)
    require(spec is not None and spec.loader is not None,
            f"cannot load numerical helper {NUMERIC_PATH}")
    numeric = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(numeric)
    return numeric


NUMERIC = load_numeric()
BASE = NUMERIC.BASE


def relative(path: Path) -> str:
    try:
        return str(path.relative_to(HERE))
    except ValueError as error:
        raise EvidenceError(f"path is outside evidence directory: {path}") from error


def plan_data() -> dict[str, Any]:
    plan = read_json(PLAN_PATH)
    require(isinstance(plan, dict), "plan is not an object")
    hardware = plan.get("hardware")
    require(isinstance(hardware, dict), "plan hardware lane is missing")
    require(hardware.get("shapes") == ["medium", "dense-sparse"],
            "hardware shape matrix differs")
    require(hardware.get("repeats") == 2 and hardware.get("warmup") == 0
            and hardware.get("samples") == 100,
            "hardware repeat/warmup/sample counts differ")
    expected_events = "{cycles,instructions,branches,branch-misses},page-faults,context-switches,cpu-migrations"
    require(hardware.get("events") == expected_events,
            "hardware event set differs")
    return plan


def binary_metadata(stage_dir: Path, plan: dict[str, Any]) -> dict[str, Any]:
    path = stage_dir / "binary-normal.json"
    identity = read_json(path)
    require(isinstance(identity, dict), f"{relative(path)} is not an object")
    digest = identity.get("sha256")
    require(isinstance(digest, str) and len(digest) == 64,
            f"{relative(path)} has no binary digest")
    binary_path = Path(identity.get("path", ""))
    if binary_path.is_file() and not binary_path.is_symlink():
        require(sha256(binary_path) == digest,
                f"{relative(path)} live binary digest differs")
    require(identity.get("source_manifest_sha256") ==
            sha256(stage_dir / "source-manifest.json"),
            f"{relative(path)} source manifest binding differs")
    return {"sha256": digest, "path": str(binary_path), "bytes": identity.get("bytes")}


def validate_artifacts(stage_dir: Path, receipt: dict[str, Any],
                       expected: set[str], label: str, strict: bool) -> None:
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict), f"{label}: artifacts are not an object")
    names = set(artifacts)
    if strict:
        require(names == expected, f"{label}: artifact inventory differs")
    else:
        require({f"{label.split('/', 1)[-1].removesuffix('.receipt.json')}.stdout",
                 f"{label.split('/', 1)[-1].removesuffix('.receipt.json')}.stderr"}
                <= names,
                f"{label}: failure receipt omitted stdout/stderr")
    for name, digest in artifacts.items():
        require(isinstance(name, str) and Path(name).name == name,
                f"{label}: artifact is not stage-local: {name!r}")
        path = stage_dir / name
        require(path.is_file() and not path.is_symlink(),
                f"{label}: artifact is missing or non-regular: {name}")
        require(sha256(path) == digest, f"{label}: artifact digest differs: {name}")


def validate_command(receipt: dict[str, Any], stage_dir: Path, name: str,
                     shape: str, plan: dict[str, Any], binary: dict[str, Any]) -> None:
    command = receipt.get("command")
    require(isinstance(command, list), f"{name}: command is not a list")
    require(command[:5] == ["taskset", "-c", str(plan["cpu"]), "perf", "stat"],
            f"{name}: command does not use taskset/perf")
    require("-x" in command and command[command.index("-x") + 1] == ",",
            f"{name}: perf CSV delimiter is not bound")
    require("-e" in command and command[command.index("-e") + 1] ==
            plan["hardware"]["events"], f"{name}: event set is not bound")
    require("-o" in command and command[command.index("-o") + 1] ==
            str(stage_dir / f"{name}.csv"), f"{name}: CSV path is not bound")
    require(binary["path"] in command, f"{name}: bound binary is not invoked")
    for option, value in (
        ("--warmup", plan["hardware"]["warmup"]),
        ("--samples", plan["hardware"]["samples"]),
        ("--case", plan["primary"]["case"]),
        ("--xlsx-cell-crud-shape", shape),
        ("--json", str(stage_dir / f"{name}.json")),
    ):
        require(command.count(option) == 1 and
                command[command.index(option) + 1] == str(value),
                f"{name}: command option {option} differs")


def validate_receipt(stage_dir: Path, name: str, plan: dict[str, Any],
                     shape: str, binary: dict[str, Any], manifest_sha: str) -> dict[str, Any]:
    path = stage_dir / f"{name}.receipt.json"
    receipt = read_json(path)
    require(isinstance(receipt, dict), f"{relative(path)} is not an object")
    require(receipt.get("binary_sha256") == binary["sha256"],
            f"{relative(path)} binary binding differs")
    require(receipt.get("source_manifest_sha256") == manifest_sha,
            f"{relative(path)} source binding differs")
    require(receipt.get("plan_sha256") == sha256(PLAN_PATH),
            f"{relative(path)} plan binding differs")
    require(receipt.get("script_sha256") == sha256(HERE / "run.py"),
            f"{relative(path)} run script binding differs")
    validate_command(receipt, stage_dir, name, shape, plan, binary)
    expected = {f"{name}.json", f"{name}.stdout", f"{name}.stderr", f"{name}.csv"}
    validate_artifacts(stage_dir, receipt, expected, relative(path),
                       receipt.get("exit_code") == 0)
    return receipt


def parse_events(path: Path) -> dict[str, dict[str, Any]]:
    require(path.is_file(), f"{relative(path)} CSV is missing")
    events: dict[str, dict[str, Any]] = {}
    try:
        rows = csv.reader(path.read_text(encoding="utf-8", errors="replace").splitlines())
        for fields in rows:
            if not fields or not fields[0].strip() or fields[0].strip().startswith("#"):
                continue
            require(len(fields) >= 5, f"{relative(path)} malformed CSV row: {fields!r}")
            count, unit, event, runtime, running = (item.strip() for item in fields[:5])
            require(event in EVENTS, f"{relative(path)} unexpected event {event!r}")
            require(event not in events, f"{relative(path)} duplicate event {event!r}")
            if not count or count.startswith("<") or not runtime or runtime.startswith("<"):
                events[event] = {"status": "unavailable", "raw": fields}
                continue
            try:
                value = int(count.replace(",", ""))
                duration = int(runtime.replace(",", ""))
                fraction = float(running)
            except ValueError as error:
                raise EvidenceError(f"{relative(path)} malformed measured row: {fields!r}") from error
            require(value >= 0 and duration > 0 and 0 <= fraction <= 100,
                    f"{relative(path)} measured values are invalid: {fields!r}")
            events[event] = {"status": "measured", "value": value,
                             "event_runtime_ns": duration,
                             "running_percent": fraction, "unit": unit}
    except OSError as error:
        raise EvidenceError(f"cannot read {path}: {error}") from error
    require(set(events) == set(EVENTS),
            f"{relative(path)} event coverage differs: {sorted(set(EVENTS) - set(events))}")
    return events


def hardware_plan_for_report(plan: dict[str, Any]) -> dict[str, Any]:
    copy = dict(plan)
    primary = dict(plan["primary"])
    primary["warmup"] = plan["hardware"]["warmup"]
    primary["samples"] = plan["hardware"]["samples"]
    copy["primary"] = primary
    return copy


def native_identity(stage_dir: Path, name: str, plan: dict[str, Any],
                    binary: dict[str, Any], manifest_sha: str) -> dict[str, Any]:
    # The retained 0521 validator exposes the general receipt/result helpers,
    # rather than the profile-specific convenience methods this analyzer used
    # in its draft form.  Validate the matched native row through that public
    # receipt path so its source, output, and RSS custody are checked exactly
    # as they are in the numerical comparison.
    repeat = int(name.split("-")[1][1:])
    prefix = f"native-r{repeat}-primary-"
    require(name.startswith(prefix), f"{name}: native primary name is malformed")
    shape = name[len(prefix):]
    job = {
        "name": name,
        "kind": "primary",
        "guard": None,
        "repeat": repeat,
        "case": plan["primary"]["case"],
        "shape": shape,
        "warmup": plan["primary"]["warmup"],
        "samples": plan["primary"]["samples"],
    }
    receipt, row = BASE.check_receipt(stage_dir, plan, job, binary, False)
    native_meta = {
        "receipt_sha256": sha256(stage_dir / f"{name}.receipt.json"),
        "exit_code": receipt["exit_code"],
        "source_manifest_sha256": manifest_sha,
        "binary_sha256": binary["sha256"],
        "plan_sha256": sha256(PLAN_PATH),
    }
    return {"metadata": native_meta, "identity": row["identity"],
            "report_sha256": sha256(stage_dir / f"{name}.json")}


def capture(stage: str, plan: dict[str, Any]) -> dict[str, Any]:
    stage_dir = HERE / stage
    require(stage_dir.is_dir(), f"stage directory is missing: {stage}")
    manifest_sha = sha256(stage_dir / "source-manifest.json")
    binary = binary_metadata(stage_dir, plan)
    report_plan = hardware_plan_for_report(plan)
    captures = []
    hardware = plan["hardware"]
    for repeat in range(1, hardware["repeats"] + 1):
        for shape in hardware["shapes"]:
            name = f"hardware-r{repeat}-{shape}"
            receipt = validate_receipt(stage_dir, name, plan, shape, binary, manifest_sha)
            native = native_identity(
                stage_dir, f"native-r{repeat}-primary-{shape}", plan,
                binary, manifest_sha)
            row: dict[str, Any] = {
                "name": name,
                "repeat": repeat,
                "shape": shape,
                "receipt": relative(stage_dir / f"{name}.receipt.json"),
                "receipt_sha256": sha256(stage_dir / f"{name}.receipt.json"),
                "exit_code": receipt.get("exit_code"),
                "native": native,
            }
            if receipt.get("exit_code") != 0:
                row.update(status="unavailable",
                           reason="perf child failed; raw receipt, stderr, and CSV are retained")
                captures.append(row)
                continue
            report_path = stage_dir / f"{name}.json"
            job = {
                "name": name,
                "kind": "primary",
                "guard": None,
                "repeat": repeat,
                "case": plan["primary"]["case"],
                "shape": shape,
                "warmup": report_plan["primary"]["warmup"],
                "samples": report_plan["primary"]["samples"],
            }
            parsed = BASE.validate_result(
                read_json(report_path), report_plan, job, binary, False)
            identity = parsed["identity"]
            normalized_hardware = BASE.normalize_iteration_counts(identity)
            normalized_native = BASE.normalize_iteration_counts(native["identity"])
            require(normalized_hardware == normalized_native,
                    f"{name}: hardware/native logical result identity differs")
            events = parse_events(stage_dir / f"{name}.csv")
            group = [events[event] for event in GROUP]
            group_valid = all(item["status"] == "measured" and
                              item["running_percent"] == 100 for item in group)
            group_valid = group_valid and len({item["event_runtime_ns"] for item in group}) == 1
            row.update(events=events, identity_equal=True, group_valid=group_valid,
                       report_sha256=sha256(report_path))
            if group_valid:
                values = {event: events[event]["value"] for event in GROUP}
                require(values["cycles"] > 0 and
                        values["branches"] >= values["branch-misses"],
                        f"{name}: grouped counter invariants fail")
                row.update(status="measured",
                           ipc=values["instructions"] / values["cycles"],
                           branch_miss_percent=100 * values["branch-misses"] /
                           values["branches"] if values["branches"] else 0.0)
            else:
                row.update(status="unavailable_for_group_claim",
                           reason="missing or multiplexed grouped event; raw counters retained")
            captures.append(row)
    return {
        "status": "pass",
        "stage": stage,
        "captures": captures,
        "scope": hardware["scope"],
        "source_manifest_sha256": manifest_sha,
        "plan_sha256": sha256(PLAN_PATH),
        "latency_samples_excluded_from_native": hardware["samples"] * len(captures),
        "no_operation_local_hardware_or_speedup_claim": True,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output", nargs="?", type=Path)
    parser.add_argument("--stage", choices=("baseline", "candidate"), default="baseline")
    args = parser.parse_args()
    output = args.output or HERE / args.stage / "hardware-analysis.json"
    try:
        document = capture(args.stage, plan_data())
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(document, indent=2) + "\n", encoding="utf-8")
    except (EvidenceError, OSError, json.JSONDecodeError) as error:
        print(f"analyze_hardware.py: error: {error}", file=sys.stderr)
        return 2
    print(f"0525 whole-child hardware diagnostic verified: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
