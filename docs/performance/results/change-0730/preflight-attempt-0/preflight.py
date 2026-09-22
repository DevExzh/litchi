#!/usr/bin/env python3
"""Run analyzer and independent audit on synthetic full-matrix evidence.

The temporary packet is isolated from the real captures.  Reports are copied
from the already qualified before artifacts and receive deterministic synthetic
timing values; this exercises custody, command order, schema, pair, and
allocation paths without running Cargo or a native probe and must never be
read as a measurement.
"""

from __future__ import annotations

import contextlib
import copy
import hashlib
import importlib.util
import io
import json
import shutil
import tempfile
from pathlib import Path


P = Path(__file__).resolve().parent
ROOT = P.parents[3]


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def read(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))


def write(path: Path, value) -> None:
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def clone_packet(target: Path) -> None:
    target.mkdir(exist_ok=True)
    for child in P.iterdir():
        if child.name in {"captures", "before", "__pycache__"}:
            continue
        if child.is_file():
            shutil.copy2(child, target / child.name)
        elif child.is_dir():
            shutil.copytree(child, target / child.name)
    shutil.copytree(P / "before", target / "before")
    (target / "captures").mkdir()


def report_template(packet: Path, lane: str, case: str) -> dict:
    name = f"{lane}-c0-{case}-0-baseline.json"
    return read(packet / "before" / name)


def synthetic_captures(packet: Path) -> None:
    captures = packet / "captures"
    runs = []
    cases = ("docfloat", "docnohf")
    native_slots = ("baseline", "baseline", "baseline", "candidate", "candidate", "baseline")
    allocation_slots = ("baseline", "candidate", "candidate", "baseline")
    builds = {variant: read(packet / f"{variant}-builds.json")
              for variant in ("baseline", "candidate")}

    def binary(variant: str, allocation: bool) -> str:
        suffix = "ole_format_save_probe_alloc" if allocation else "ole_format_save_probe"
        return next(row["path"] for row in builds[variant]["binaries"]
                    if Path(row["path"]).name == f"{variant}-{suffix}")

    def emit(name: str, report: dict, command: list[str], lane: str, cycle: int,
             case: str, variant: str, slot: int) -> None:
        output = captures / name
        write(output, report)
        stderr = captures / (name + ".stderr")
        stderr.write_text("", encoding="utf-8")
        runs.append({"lane": lane, "cycle": cycle, "case": case, "variant": variant,
                     "slot": slot, "command": command, "exit_code": 0,
                     "output": name, "sha256": digest(output),
                     "stderr": name + ".stderr", "stderr_sha256": digest(stderr)})

    for cycle in range(3):
        order = cases if cycle % 2 == 0 else tuple(reversed(cases))
        for case in order:
            for slot, variant in enumerate(native_slots):
                report = report_template(packet, "native", case)
                # These values are deliberately synthetic and deterministic.
                for index, sample in enumerate(report["samples"]):
                    sample["phase_ns"]["whole_ns"] = 1_000_000 + cycle * 10_000 + slot * 1_000 + index
                name = f"native-c{cycle}-{case}-{slot}-{variant}.json"
                command = ["taskset", "-c", "12", binary(variant, False), "--case", case,
                           "--input", read(packet / "cases.json")[0 if case == "docfloat" else 1]["path"],
                           "--operation", "format", "--samples", "50", "--warmups", "3"]
                emit(name, report, command, "native", cycle, case, variant, slot)
    case_rows = {row["case"]: row for row in read(packet / "cases.json")}
    for case in cases:
        for slot, variant in enumerate(allocation_slots):
            report = report_template(packet, "allocation", case)
            name = f"allocation-c0-{case}-{slot}-{variant}.json"
            command = ["taskset", "-c", "12", binary(variant, True), "--case", case,
                       "--input", case_rows[case]["path"], "--operation", "format",
                       "--samples", "1", "--warmups", "0"]
            emit(name, report, command, "allocation", 0, case, variant, slot)
    write(captures / "manifest.json", {"mode": "synthetic-preflight", "status": "complete",
                                        "runs": runs})


def rebase_freeze(packet: Path) -> None:
    """Point packet-owned frozen bindings at the isolated packet copy."""
    path = packet / "freeze.json"
    frozen = read(path)
    bindings = frozen.get("bindings", frozen)
    original_prefix = str(P) + "/"
    rebased = {}
    for key, value in bindings.items():
        rebased[key.replace(original_prefix, str(packet) + "/", 1)
               if key.startswith(original_prefix) else key] = value
    if "bindings" in frozen:
        frozen["bindings"] = rebased
    else:
        frozen = rebased
    write(path, frozen)
    for key in rebased:
        assert not key.startswith(original_prefix), f"stale packet freeze binding: {key}"


def main() -> None:
    analyzer_spec = importlib.util.spec_from_file_location("change0730_analyzer", P / "analyze.py")
    audit_spec = importlib.util.spec_from_file_location("change0730_audit", P / "audit.py")
    assert analyzer_spec and analyzer_spec.loader and audit_spec and audit_spec.loader
    analyzer = importlib.util.module_from_spec(analyzer_spec)
    audit = importlib.util.module_from_spec(audit_spec)
    analyzer_spec.loader.exec_module(analyzer)
    audit_spec.loader.exec_module(audit)
    with tempfile.TemporaryDirectory(prefix="litchi-0730-preflight-") as temporary:
        packet = Path(temporary)
        clone_packet(packet)
        rebase_freeze(packet)
        synthetic_captures(packet)
        analyzer.P = packet
        analyzer.ROOT = ROOT
        analyzer.CAPTURES = packet / "captures"
        analyzer.BEFORE = packet / "before"
        with contextlib.redirect_stdout(io.StringIO()):
            analyzer.main()
        audit.PACKET = packet
        audit.ROOT = ROOT
        audit.CAPTURES = packet / "captures"
        audit.BEFORE = packet / "before"
        with contextlib.redirect_stdout(io.StringIO()):
            assert audit.main() == 0
        analysis = read(packet / "analysis.json")
        assert analysis["disposition"].startswith("bounded validated-render handoff")
        assert len(analysis["processes"]) == 44
        receipt = {"status": "passed", "kind": "synthetic schema integration only; no measurement",
                   "analyzer_sha256": digest(P / "analyze.py"),
                   "auditor_sha256": digest(P / "audit.py"),
                   "contract_sha256": digest(P / "oracle-contract.json"),
                   "script_sha256": digest(Path(__file__))}
        write(P / "preflight.json", receipt)
    print("PASS synthetic 44-process analyzer/audit integration; no measurement evidence created")


if __name__ == "__main__":
    main()
