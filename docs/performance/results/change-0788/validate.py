"""Fail-closed validator for the 0788 diagnostic replay packet.

Validation is intentionally offline.  It replays :mod:`analyze`, checks the
separate phase parser when its deterministic output is present, and verifies
the final restoration/seal witnesses.  It never launches a benchmark,
profiler, compiler, or native child.
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path
from typing import Any

import analyze


PACKET = analyze.PACKET
RECEIPT_SCHEMA = "litchi.cached-part-memory-receipts.0788.v1"
MEMORY_ANALYSIS_SCHEMA = "litchi.cached-part-memory-analysis.v1"


def require(condition: bool, message: str) -> None:
    analyze.require(condition, message)


def _artifact_identity(value: Any, label: str, *, packet_bound: bool = True) -> Path:
    if not packet_bound:
        cleanup, verified = analyze.load_cleanup()
        analyze.validate_binary(value, label, cleanup, verified)
        return Path(value["path"])
    return analyze.artifact_path(value, label, packet_bound=True)


def _receipt_file(path: Path, label: str) -> dict[str, Any]:
    value = analyze.read_json(path)
    require(isinstance(value, dict), f"{label} is malformed")
    require(value.get("schema") == RECEIPT_SCHEMA, f"{label} schema changed")
    rows = value.get("receipts")
    require(isinstance(rows, list), f"{label} receipt list is missing")
    return value


def validate_receipt_custody() -> dict[str, Any]:
    lanes = {
        "native": 120,
        "qualification-before": 10,
        "qualification-after": 10,
        "memory": 64,
        "heaptrack": 16,
    }
    result: dict[str, Any] = {}
    for lane, count in lanes.items():
        directory = PACKET / lane
        receipts_path = directory / "receipts.json"
        value = _receipt_file(receipts_path, f"{lane}/receipts.json")
        rows = value["receipts"]
        require(len(rows) == count, f"{lane} receipt count changed")
        for index, row in enumerate(rows):
            require(isinstance(row, dict), f"{lane} receipt {index} is malformed")
            require(row.get("schema") == RECEIPT_SCHEMA,
                    f"{lane} receipt {index} schema changed")
            require(row.get("exit_code") == 0, f"{lane} receipt {index} failed")
            for name in ("binary", "log", "report"):
                require(name in row, f"{lane} receipt {index} lacks {name}")
                _artifact_identity(row[name], f"{lane} receipt {index} {name}",
                                   packet_bound=name != "binary")
            if lane in ("native", "qualification-before", "qualification-after"):
                require("time" in row and "rss" in row,
                        f"{lane} receipt {index} lacks process controls")
                _artifact_identity(row["time"], f"{lane} receipt {index} time")
                _artifact_identity(row["rss"], f"{lane} receipt {index} rss")
            elif lane == "memory":
                require("rss" in row and "mode" in row,
                        f"memory receipt {index} lacks process controls")
                _artifact_identity(row["rss"], f"memory receipt {index} rss")
                snapshots = row.get("snapshots")
                if row.get("mode") in ("on", True):
                    require(snapshots is not None,
                            f"memory receipt {index} lacks phase snapshots")
                    _artifact_identity(snapshots, f"memory receipt {index} snapshots")
                else:
                    require(snapshots is None, f"memory receipt {index} off control has snapshots")
            else:
                for name in ("trace", "summary", "flamegraph", "time", "rss"):
                    require(name in row, f"heaptrack receipt {index} lacks {name}")
                    _artifact_identity(row[name], f"heaptrack receipt {index} {name}",
                                       packet_bound=name != "binary")
                _artifact_identity(row.get("raw_trace", row["trace"]),
                                   f"heaptrack receipt {index} raw trace")
                require(row.get("print_exit_code") == 0,
                        f"heaptrack receipt {index} print failed")
                command = analyze.normalise_command(row.get("print_command"))
                require("-m" in command and command[command.index("-m") + 1] == "0",
                        f"heaptrack receipt {index} merged backtraces")
        complete = directory / "complete.json"
        require(complete.is_file() and not complete.is_symlink(),
                f"{lane}/complete.json is missing")
        complete_value = analyze.read_json(complete)
        require(isinstance(complete_value, dict)
                and complete_value.get("schema") == RECEIPT_SCHEMA
                and complete_value.get("children") == count,
                f"{lane} completion witness changed")
        for field in ("receipts", "source"):
            require(field in complete_value, f"{lane} completion lacks {field}")
            _artifact_identity(complete_value[field], f"{lane} completion {field}")
        result[lane] = {"receipts": count, "complete": analyze.file_identity(complete)}
    return result


def validate_memory_analysis(*, required: bool) -> dict[str, Any] | None:
    path = PACKET / "memory-analysis.json"
    if not path.is_file():
        require(not required, "memory-analysis.json is missing")
        return None
    value = analyze.read_json(path)
    require(isinstance(value, dict) and value.get("schema") == MEMORY_ANALYSIS_SCHEMA,
            "memory-analysis schema changed")
    require(value.get("counts", {}).get("receipts") == 64
            and value.get("counts", {}).get("samples") == 992,
            "memory-analysis counts changed")
    script = PACKET / "memory_analysis.py"
    require(script.is_file() and not script.is_symlink(),
            "memory_analysis.py is missing")
    try:
        completed = subprocess.run([sys.executable, "-B", str(script), "--check"],
                                   cwd=PACKET, stdout=subprocess.PIPE,
                                   stderr=subprocess.PIPE, text=True, check=False)
    except OSError as error:
        analyze.fail(f"cannot replay memory analysis: {error}")
    require(completed.returncode == 0,
            f"memory analysis replay failed: {completed.stderr.strip()}")
    return {"path": analyze.rel(path), "sha256": analyze.sha256(path),
            "schema": value["schema"], "replayed": True}


def validate_heap_analysis(*, required: bool) -> dict[str, Any] | None:
    path = PACKET / "heap-analysis.json"
    if not path.is_file():
        require(not required, "heap-analysis.json is missing")
        return None
    script = PACKET / "heap_analysis.py"
    require(script.is_file() and not script.is_symlink(), "heap_analysis.py is missing")
    try:
        completed = subprocess.run([sys.executable, "-B", str(script), "--check"],
                                   cwd=PACKET, stdout=subprocess.PIPE,
                                   stderr=subprocess.PIPE, text=True, check=False)
    except OSError as error:
        analyze.fail(f"cannot replay heap analysis: {error}")
    require(completed.returncode == 0,
            f"heap analysis replay failed: {completed.stderr.strip()}")
    value = analyze.read_json(path)
    require(isinstance(value, dict)
            and value.get("settings", {}).get("merge_backtraces") is False
            and value.get("settings", {}).get("native_timing_or_rss_pooled") is False
            and len(value.get("reports", [])) == 16,
            "heap-analysis scope or count changed")
    return {"path": analyze.rel(path), "sha256": analyze.sha256(path),
            "reports": len(value["reports"]), "replayed": True}


def _current_production_files() -> dict[str, str]:
    revision = analyze.origin()["base"]
    names = analyze.production_names(revision)
    result: dict[str, str] = {}
    for name in names:
        path = analyze.ROOT / name
        require(path.is_file() and not path.is_symlink(),
                f"current production file is missing or symlinked: {name}")
        result[name] = analyze.sha256(path)
    return result


def validate_restoration(*, required: bool) -> dict[str, Any] | None:
    disposition_path = PACKET / "disposition.json"
    restored_path = PACKET / "restored-source.json"
    cleanup_path = PACKET / "cleanup.json"
    if not any(path.is_file() for path in (disposition_path, restored_path, cleanup_path)):
        require(not required, "final restoration witnesses are missing")
        return None
    result: dict[str, Any] = {}
    if restored_path.is_file():
        restored = analyze.read_json(restored_path)
        require(isinstance(restored, dict), "restored-source.json is malformed")
        files = restored.get("files")
        require(isinstance(files, dict) and len(files) == analyze.EXPECTED_PRODUCTION_FILES,
                "restored production census changed")
        baseline = analyze.read_json(PACKET / "build-before" / "source.json")["production"]["files"]
        require(files == baseline, "restored source is not the qualified baseline")
        require(files == _current_production_files(),
                "current production bytes do not match restoration witness")
        result["restored_source"] = analyze.file_identity(restored_path)
        result["production_restored"] = True
    else:
        require(not required, "restored-source.json is missing")
    if disposition_path.is_file():
        disposition = analyze.read_json(disposition_path)
        require(isinstance(disposition, dict), "disposition is malformed")
        require(disposition.get("decision") in ("diagnostic-only", "reject", "not-evaluated"),
                "diagnostic disposition changed")
        require(disposition.get("candidate_retained_only_in_archive") is True,
                "candidate retention witness changed")
        require(disposition.get("previous_0787_rejection_remains_authoritative") is True
                or "0787" in str(disposition.get("reason", "")),
                "0787 rejection authority witness is missing")
        require(disposition.get("schema") == "litchi.cached-part-memory-disposition.0788.v1",
                "diagnostic disposition schema changed")
        for key, name in (("analysis", "analysis.json"),
                          ("memory_analysis", "memory-analysis.json"),
                          ("heap_analysis", "heap-analysis.json"),
                          ("restored_source", "restored-source.json"),
                          ("cleanup", "cleanup.json")):
            linked = analyze.artifact_path(disposition.get(key), f"disposition {key}")
            require(linked.resolve() == (PACKET / name).resolve(),
                    f"disposition {key} points to a different artifact")
        result["disposition"] = analyze.file_identity(disposition_path)
    else:
        require(not required, "disposition.json is missing")
    if cleanup_path.is_file():
        cleanup = analyze.read_json(cleanup_path)
        require(isinstance(cleanup, dict) and cleanup.get("verified") is True,
                "cleanup witness is not verified")
        require(cleanup.get("target_absent_after_removal") is True,
                "owned target was not removed")
        result["cleanup"] = analyze.file_identity(cleanup_path)
    else:
        require(not required, "cleanup.json is missing")
    return result


def validate_seal(*, required: bool) -> dict[str, Any] | None:
    path = PACKET / "seal.json"
    if not path.is_file():
        require(not required, "final seal is missing")
        return None
    seal = analyze.read_json(path)
    require(isinstance(seal, dict), "seal is malformed")
    require(seal.get("schema") in ("litchi.cached-part-memory-seal.0788.v1",
                                    "litchi.execution-scaling-seal.v1"),
            "seal schema changed")
    files = seal.get("files")
    require(isinstance(files, dict) and files, "seal file set is missing")
    actual: dict[str, str] = {}
    for item in PACKET.rglob("*"):
        if item == path or "__pycache__" in item.parts:
            continue
        require(not item.is_symlink(), f"symlink in sealed packet: {item}")
        if item.is_file():
            actual[str(item.relative_to(PACKET))] = analyze.sha256(item)
    require(actual == files, "seal file set or hash changed")
    require(seal.get("payload_count") == len(actual), "seal payload count changed")
    return {"path": analyze.rel(path), "sha256": analyze.sha256(path),
            "payload_count": len(actual), "verified": True}


def validate(*, require_final_seal: bool = False) -> dict[str, Any]:
    result = analyze.analyze(check=True)
    require(result.get("schema") == analyze.ANALYSIS_SCHEMA,
            "analysis schema changed")
    require(result.get("diagnostic_only") is True
            and result.get("adoption_decision") == "not evaluated"
            and result.get("previous_0787_rejection_remains_authoritative") is True,
            "analysis diagnostic disposition changed")
    receipts = validate_receipt_custody()
    memory = validate_memory_analysis(required=require_final_seal)
    heap = validate_heap_analysis(required=require_final_seal)
    restoration = validate_restoration(required=require_final_seal)
    seal = validate_seal(required=require_final_seal)
    return {"schema": "litchi.cached-part-memory-validator.0788.v1",
            "analysis": analyze.file_identity(PACKET / "analysis.json"),
            "paired": analyze.file_identity(PACKET / "paired.csv"),
            "receipts": receipts, "memory_analysis": memory, "heap_analysis": heap,
            "restoration": restoration, "seal": seal,
            "verified": True}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--require-final-seal", action="store_true",
                        help="require restoration witnesses, memory replay, and seal")
    args = parser.parse_args(argv)
    try:
        result = validate(require_final_seal=args.require_final_seal)
    except analyze.ReplayError as error:
        print(f"0788 validation failed: {error}", file=sys.stderr)
        return 1
    print(json.dumps(result, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
