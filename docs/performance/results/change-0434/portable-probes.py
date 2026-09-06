#!/usr/bin/env python3
"""Exercise portable 0434 verification and fail-closed mutation probes.

The source bundle is copied into a shallow temporary directory.  All probes
mutate only that copy, refresh its sealed inventory, run one expected gate,
and restore the original bytes before the next probe.  No seal operation is
performed: the probe inventory is deliberately the only metadata refreshed.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
from typing import Any, Callable


ROOT = Path(__file__).resolve().parent
STAGES = ("precleanup", "aftercleanup", "final")


class ProbeError(RuntimeError):
    """The sealed bundle or a mutation probe did not meet its contract."""


def load(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise ProbeError(f"invalid JSON: {path}: {error}") from error


def write_json(path: Path, value: Any) -> None:
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def relative(bundle: Path, path: Path) -> str:
    return path.relative_to(bundle).as_posix()


def refresh_inventory(bundle: Path) -> None:
    """Refresh only the copied sealed inventory, without running seal.py."""

    inventory = bundle / "SHA256SUMS"
    if not inventory.is_file():
        raise ProbeError("sealed bundle has no SHA256SUMS")
    files = sorted(
        path
        for path in bundle.rglob("*")
        if path.is_file() and path.name != "SHA256SUMS"
    )
    inventory.write_text(
        "".join(f"{sha(path)}  {relative(bundle, path)}\n" for path in files),
        encoding="utf-8",
    )


def run_command(command: list[str], cwd: Path) -> dict[str, Any]:
    result = subprocess.run(
        command,
        cwd=cwd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        check=False,
    )
    return {
        "exit_code": result.returncode,
        "output": result.stdout,
    }


def verify_command(bundle: Path, stage: str) -> list[str]:
    return [
        sys.executable,
        "-B",
        str(bundle / "verify.py"),
        "--portable-check",
        "--require-inventory",
        "--stage",
        stage,
    ]


def lifecycle_command(bundle: Path, stage: str) -> list[str]:
    return [sys.executable, "-B", str(bundle / "lifecycle.py"), "--stage", stage]


def require_pass(result: dict[str, Any], label: str, marker: str) -> dict[str, Any]:
    if result["exit_code"] != 0 or marker not in result["output"]:
        raise ProbeError(
            f"{label} did not pass: exit={result['exit_code']}\n"
            f"{result['output'][-2000:]}"
        )
    return {
        "status": "pass",
        "exit_code": result["exit_code"],
        "evidence": result["output"][-2000:],
    }


def require_reject(
    result: dict[str, Any], label: str, needle: str
) -> dict[str, Any]:
    if result["exit_code"] == 0 or needle not in result["output"]:
        raise ProbeError(
            f"{label} was not rejected with {needle!r}: "
            f"exit={result['exit_code']}\n{result['output'][-2000:]}"
        )
    return {
        "status": "rejected",
        "exit_code": result["exit_code"],
        "needle": needle,
        "evidence": result["output"][-2000:],
    }


def first_formal_receipt(bundle: Path) -> Path:
    paths = sorted((bundle / "runs").glob("*/formal/*-receipt.json"))
    for path in paths:
        row = load(path)
        if isinstance(row, dict) and row.get("status") == "pass":
            return path
    raise ProbeError("sealed bundle has no passing formal capture receipt")


def first_a1_pair(bundle: Path) -> tuple[Path, Path]:
    index = load(bundle / "runs" / "A1" / "formal" / "capture-index.json")
    if not isinstance(index, list) or len(index) < 2:
        raise ProbeError("A1 capture index has fewer than two formal receipts")
    paths = [bundle / value for value in index[:2] if isinstance(value, str)]
    if len(paths) != 2 or any(not path.is_file() for path in paths):
        raise ProbeError("A1 capture index does not point to two receipts")
    return paths[0], paths[1]


def mutate_cpu(bundle: Path) -> tuple[Path, Callable[[Path], None], str]:
    path = first_formal_receipt(bundle)

    def apply(target: Path) -> None:
        row = load(target)
        argv = row.get("argv")
        if not isinstance(argv, list) or "-c" not in argv:
            raise ProbeError(f"{relative(bundle, target)} has no taskset CPU argument")
        index = argv.index("-c") + 1
        if index >= len(argv) or argv[index] == "3":
            raise ProbeError(f"{relative(bundle, target)} has no mutable CPU argument")
        argv[index] = "3"
        write_json(target, row)

    return path, apply, "argv[2]"


def mutate_phase_timestamp(bundle: Path) -> tuple[Path, Callable[[Path], None], str]:
    first, second = first_a1_pair(bundle)

    def apply(target: Path) -> None:
        row = load(target)
        first_row = load(first)
        started = first_row.get("started_utc")
        if not isinstance(started, str):
            raise ProbeError("first A1 receipt has no timestamp")
        row["started_utc"] = started
        write_json(target, row)

    return second, apply, "formal lanes overlap or are out of capture order"


def mutate_record_event(bundle: Path) -> tuple[Path, Callable[[Path], None], str]:
    path = bundle / "profiles" / "before-streaming" / "record" / "receipt.json"

    def apply(target: Path) -> None:
        row = load(target)
        if row.get("record_event") != "cycles:u":
            raise ProbeError("profile record receipt has no frozen cycles event")
        row["record_event"] = "instructions:u"
        write_json(target, row)

    return path, apply, "perf-record sampling/call-graph binding is stale"


def mutate_summary_latency(bundle: Path) -> tuple[Path, Callable[[Path], None], str]:
    path = bundle / "summary.json"

    def apply(target: Path) -> None:
        row = load(target)
        rows = row.get("rows")
        if not isinstance(rows, list) or not rows:
            raise ProbeError("summary has no rows")
        elapsed = rows[0].get("elapsed_ns")
        if not isinstance(elapsed, dict) or not isinstance(elapsed.get("p50"), (int, float)):
            raise ProbeError("summary first row has no numeric elapsed p50")
        elapsed["p50"] += 1
        write_json(target, row)

    return path, apply, "summary.json: retained summary differs from independent derivation"


def mutate_missing_required_receipt(bundle: Path) -> tuple[Path, Callable[[Path], None], str]:
    path = bundle / "checks" / "ods-tests-final.json"

    def apply(target: Path) -> None:
        if not target.is_file():
            raise ProbeError(f"required receipt is already missing: {relative(bundle, target)}")
        target.unlink()

    return path, apply, "checks/ods-tests-final.json: file is missing"


def mutate_compression_size(bundle: Path) -> tuple[Path, Callable[[Path], None], str]:
    path = bundle / "compression.json"

    def apply(target: Path) -> None:
        rows = load(target)
        if not isinstance(rows, dict) or not rows:
            raise ProbeError("compression metadata has no records")
        name = sorted(rows)[0]
        record = rows[name]
        if not isinstance(record, dict) or not isinstance(record.get("original_bytes"), int):
            raise ProbeError("first compression record has no original size")
        record["original_bytes"] += 1
        write_json(target, rows)

    return path, apply, "original artifact binding differs"


def mutation_cases(bundle: Path) -> list[tuple[str, Path, Callable[[Path], None], str, str]]:
    cases = (
        ("formal_receipt_cpu", mutate_cpu, "verify"),
        ("phase_timestamp_abba", mutate_phase_timestamp, "verify"),
        ("profile_record_event", mutate_record_event, "verify"),
        ("summary_latency", mutate_summary_latency, "lifecycle"),
        ("missing_required_test_receipt", mutate_missing_required_receipt, "lifecycle"),
        ("compression_original_size", mutate_compression_size, "verify"),
    )
    result: list[tuple[str, Path, Callable[[Path], None], str, str]] = []
    for name, factory, validator in cases:
        path, apply, needle = factory(bundle)
        if not path.exists() and name != "missing_required_test_receipt":
            raise ProbeError(f"mutation target is missing: {relative(bundle, path)}")
        result.append((name, path, apply, validator, needle))
    return result


def run(stage: str) -> dict[str, Any]:
    if not (ROOT / "SHA256SUMS").is_file():
        raise ProbeError("seal the source bundle before running portable probes")
    with tempfile.TemporaryDirectory(prefix="litchi-0434-probes-") as temporary:
        bundle = Path(temporary) / "b"
        shutil.copytree(ROOT, bundle)
        inventory = bundle / "SHA256SUMS"
        original_inventory = inventory.read_bytes()

        baseline_verify = require_pass(
            run_command(verify_command(bundle, stage), bundle.parent),
            "portable verify",
            "VALID",
        )
        baseline_lifecycle = require_pass(
            run_command(lifecycle_command(bundle, stage), bundle.parent),
            "portable lifecycle",
            '"summary_rederived": true',
        )
        mutations: list[dict[str, Any]] = []
        for name, path, apply, validator, needle in mutation_cases(bundle):
            original = path.read_bytes() if path.exists() else None
            restored = False
            try:
                apply(path)
                refresh_inventory(bundle)
                command = verify_command(bundle, stage) if validator == "verify" else lifecycle_command(bundle, stage)
                result = require_reject(run_command(command, bundle.parent), name, needle)
                mutations.append({
                    "name": name,
                    "path": relative(bundle, path),
                    "validator": validator,
                    **result,
                })
            finally:
                if original is None:
                    if path.exists():
                        path.unlink()
                else:
                    path.write_bytes(original)
                inventory.write_bytes(original_inventory)
                if original is None:
                    if path.exists():
                        raise ProbeError(f"failed to restore {relative(bundle, path)}")
                elif not path.is_file() or path.read_bytes() != original:
                    raise ProbeError(f"failed to restore {relative(bundle, path)}")
                if inventory.read_bytes() != original_inventory:
                    raise ProbeError("failed to restore SHA256SUMS")
                restored = True
            mutations[-1]["restored"] = restored

        final_verify = require_pass(
            run_command(verify_command(bundle, stage), bundle.parent),
            "restored portable verify",
            "VALID",
        )
        final_lifecycle = require_pass(
            run_command(lifecycle_command(bundle, stage), bundle.parent),
            "restored portable lifecycle",
            '"summary_rederived": true',
        )
    return {
        "status": "pass",
        "stage": stage,
        "baseline": {"verify": baseline_verify, "lifecycle": baseline_lifecycle},
        "mutations": mutations,
        "restored": {"verify": final_verify, "lifecycle": final_lifecycle},
        "scope": "Copied sealed bundle only; SHA256SUMS refreshed per mutation; source bundle untouched.",
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--stage", choices=STAGES, default="final")
    args = parser.parse_args()
    try:
        print(json.dumps(run(args.stage), indent=2, sort_keys=True))
    except (OSError, ProbeError, ValueError, KeyError, TypeError, AssertionError) as error:
        print(json.dumps({"status": "failed", "error": str(error)}, sort_keys=True))
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
