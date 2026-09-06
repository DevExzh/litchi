#!/usr/bin/env python3
"""Exercise copied-bundle 0438 ODP verification and fail-closed probes.

The complete sealed bundle is copied into a shallow temporary directory.
Each probe mutates only that copy, refreshes only the copied SHA256SUMS, runs
the appropriate verifier, and restores the exact original bytes before the
next probe.  The seal operation is never invoked and the source bundle is
never written.
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
CHANGE = 438
STAGES = ("precleanup", "aftercleanup", "final")


class ProbeError(RuntimeError):
    """The sealed bundle or one of its mutation probes violated its contract."""


def load(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise ProbeError(f"invalid JSON: {path}: {error}") from error


def protocol(bundle: Path) -> dict[str, Any]:
    value = load(bundle / "protocol.json")
    if not isinstance(value, dict) or value.get("change") != CHANGE:
        raise ProbeError("copied bundle has no frozen 0438 protocol")
    return value


def protocol_needle(bundle: Path, key: str, fallback: str) -> str:
    value = protocol(bundle).get("probe_needles")
    if isinstance(value, dict) and isinstance(value.get(key), str) and value[key]:
        return value[key]
    return fallback


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
    """Refresh only the copied inventory; never run seal.py for a probe."""

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
    return {"exit_code": result.returncode, "output": result.stdout}


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


def require_reject(result: dict[str, Any], label: str, needle: str) -> dict[str, Any]:
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


def formal_receipts(bundle: Path) -> list[Path]:
    result: list[Path] = []
    # 0438 capture emits runs/<phase>/<attempt>/<lane>-receipt.json.  Keep
    # this path contract separate from the older runs/<phase>/formal layout.
    for path in sorted((bundle / "runs").glob("*/*/*-receipt.json")):
        row = load(path)
        if isinstance(row, dict) and row.get("status") == "pass":
            result.append(path)
    return result


def first_formal_receipt(bundle: Path) -> Path:
    values = formal_receipts(bundle)
    if not values:
        raise ProbeError("sealed bundle has no passing formal capture receipt")
    return values[0]


def first_capture_pair(bundle: Path) -> tuple[Path, Path]:
    indexes = sorted((bundle / "runs").glob("*/*/capture-index.json"))
    if not indexes:
        raise ProbeError("formal runs have no capture index")
    index = load(indexes[0])
    if not isinstance(index, list) or len(index) < 2:
        raise ProbeError("first capture index has fewer than two formal receipts")
    values = [value for value in index[:2] if isinstance(value, str)]
    if any(Path(value).is_absolute() or ".." in Path(value).parts for value in values):
        raise ProbeError("first capture index contains an unsafe receipt path")
    paths = [bundle / value for value in values]
    if len(paths) != 2 or any(not path.is_file() for path in paths):
        raise ProbeError("first capture index does not point to two receipts")
    return paths[0], paths[1]


def mutate_cpu(bundle: Path) -> tuple[Path, Callable[[Path], None], str]:
    path = first_formal_receipt(bundle)

    def apply(target: Path) -> None:
        row = load(target)
        argv = row.get("argv")
        if not isinstance(argv, list) or len(argv) < 3:
            raise ProbeError(f"{relative(bundle, target)} has no taskset CPU argument")
        try:
            cpu_index = next(index for index, value in enumerate(argv[:-1])
                             if value == "-c" and argv[index - 1] == "taskset")
        except (StopIteration, IndexError):
            raise ProbeError(f"{relative(bundle, target)} has no taskset CPU argument")
        if not str(argv[cpu_index + 1]).isdigit():
            raise ProbeError(f"{relative(bundle, target)} has a nonnumeric CPU argument")
        original_cpu = int(argv[cpu_index + 1])
        probe_cpu = original_cpu + 1
        if str(probe_cpu) == str(argv[cpu_index + 1]):
            raise ProbeError(f"{relative(bundle, target)} already has the probe CPU")
        argv[cpu_index + 1] = str(probe_cpu)
        write_json(target, row)

    return path, apply, protocol_needle(bundle, "cpu_affinity", "CPU affinity")


def mutate_phase_timestamp(bundle: Path) -> tuple[Path, Callable[[Path], None], str]:
    first, second = first_capture_pair(bundle)

    def apply(target: Path) -> None:
        row = load(target)
        first_row = load(first)
        started = first_row.get("started_utc")
        if not isinstance(started, str):
            raise ProbeError("first capture receipt has no timestamp")
        row["started_utc"] = started
        write_json(target, row)

    return second, apply, protocol_needle(
        bundle, "phase_timestamp", "formal lanes overlap or are out of capture order"
    )


def profile_record_receipt(bundle: Path) -> Path:
    candidates = sorted((bundle / "profiles").glob("**/receipt.json"))
    preferred = [path for path in candidates if "after-streaming" in path.parts or path.parts[-3:-2] == ("after",)]
    for path in preferred + candidates:
        row = load(path)
        if isinstance(row, dict) and "record_event" in row:
            return path
    raise ProbeError("bundle has no record profile receipt")


def mutate_record_event(bundle: Path) -> tuple[Path, Callable[[Path], None], str]:
    path = profile_record_receipt(bundle)

    def apply(target: Path) -> None:
        row = load(target)
        if row.get("record_event") != "cycles:u":
            raise ProbeError("after-streaming profile has no frozen cycles event")
        row["record_event"] = "instructions:u"
        write_json(target, row)

    return path, apply, protocol_needle(
        bundle, "record_event", "perf-record sampling/call-graph binding is stale"
    )


def mutate_summary_latency(bundle: Path) -> tuple[Path, Callable[[Path], None], str]:
    value = protocol(bundle).get("summary_path", "summary.json")
    if not isinstance(value, str) or not value or Path(value).is_absolute() or ".." in Path(value).parts:
        raise ProbeError("protocol.summary_path is malformed")
    path = bundle / value

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

    return path, apply, protocol_needle(
        bundle, "summary", "summary.json: retained summary differs from independent derivation"
    )


def mutate_missing_required_receipt(bundle: Path) -> tuple[Path, Callable[[Path], None], str]:
    frozen = protocol(bundle)
    contract = frozen.get("pilot_contract")
    candidates = contract.get("required_receipts") if isinstance(contract, dict) else None
    if not isinstance(candidates, list):
        raise ProbeError("protocol pilot contract has no required receipts")
    for name in candidates:
        path = bundle / name
        if path.is_file():
            return path, lambda target: target.unlink(), protocol_needle(bundle, "missing_receipt", "file is missing")
    raise ProbeError("protocol pilot contract has no removable required receipt")


def report_path_for_receipt(bundle: Path, receipt: Path) -> Path:
    row = load(receipt)
    artifacts = row.get("artifacts")
    if not isinstance(artifacts, dict) or not isinstance(artifacts.get("report"), dict):
        raise ProbeError(f"{relative(bundle, receipt)} has no report artifact")
    value = artifacts["report"].get("path")
    if not isinstance(value, str) or not value:
        raise ProbeError(f"{relative(bundle, receipt)} report path is malformed")
    path = bundle / value
    if not path.is_file():
        raise ProbeError(f"{relative(bundle, path)} report artifact is missing")
    return path


def mutate_report_semantic_digest(bundle: Path) -> tuple[Path, Callable[[Path], None], str]:
    path = report_path_for_receipt(bundle, first_formal_receipt(bundle))

    def semantic_node(value: Any) -> dict[str, Any] | None:
        if isinstance(value, dict):
            if isinstance(value.get("semantic_sha256"), str):
                return value
            for child in value.values():
                found = semantic_node(child)
                if found is not None:
                    return found
        elif isinstance(value, list):
            for child in value:
                found = semantic_node(child)
                if found is not None:
                    return found
        return None

    def apply(target: Path) -> None:
        row = load(target)
        node = semantic_node(row)
        if node is None:
            raise ProbeError("formal ODP report has no semantic digest")
        node["semantic_sha256"] = "0" * 64
        write_json(target, row)

    return path, apply, protocol_needle(bundle, "semantic_digest", "semantic_sha256")


def oracle_verifier(bundle: Path) -> Path:
    oracle = protocol(bundle).get("oracle")
    if not isinstance(oracle, dict):
        raise ProbeError("protocol has no semantic oracle binding")
    value = oracle.get("verifier_path", oracle.get("path"))
    if not isinstance(value, str) or not value or Path(value).is_absolute() or ".." in Path(value).parts:
        raise ProbeError("protocol.oracle.verifier_path is unsafe or missing")
    path = (bundle / value).resolve()
    if not path.is_relative_to(bundle.resolve()):
        raise ProbeError("protocol.oracle.verifier_path escapes the copied bundle")
    if not path.is_file():
        raise ProbeError(f"copied oracle verifier is missing: {relative(bundle, path)}")
    return path


def oracle_role(bundle: Path, role: str) -> str:
    value = protocol(bundle)
    roles = value.get("roles")
    if isinstance(roles, dict):
        spec = roles.get(role)
        if isinstance(spec, dict) and isinstance(spec.get("oracle_role"), str):
            return spec["oracle_role"]
    mapped = value.get("oracle_roles")
    if isinstance(mapped, dict) and isinstance(mapped.get(role), str):
        return mapped[role]
    oracle = value.get("oracle")
    if isinstance(oracle, dict) and isinstance(oracle.get("roles"), dict) and isinstance(oracle["roles"].get(role), str):
        return oracle["roles"][role]
    # The capture verifier's final fallback is the production role name.  Do
    # the same here; an oracle-level ``role`` field is not part of the receipt
    # contract and must not be consulted.
    return role


def mutate_build_protocol(bundle: Path) -> tuple[Path, Callable[[Path], None], str]:
    value = protocol(bundle).get("candidate_build_descriptor", "after/build.json")
    if not isinstance(value, str) or not value:
        raise ProbeError("protocol.candidate_build_descriptor is malformed")
    path = bundle / value

    def apply(target: Path) -> None:
        row = load(target)
        if not isinstance(row.get("protocol_sha256"), str):
            raise ProbeError("after build descriptor has no protocol binding")
        row["protocol_sha256"] = "0" * 64
        write_json(target, row)

    return path, apply, protocol_needle(bundle, "build_protocol", "protocol binding is stale")


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

    return path, apply, protocol_needle(bundle, "compression", "original artifact binding differs")


def mutation_cases(bundle: Path) -> list[tuple[str, Path, Callable[[Path], None], str, str]]:
    cases = (
        ("formal_receipt_cpu", mutate_cpu, "verify"),
        ("phase_timestamp_order", mutate_phase_timestamp, "verify"),
        ("profile_record_event", mutate_record_event, "verify"),
        ("summary_latency", mutate_summary_latency, "lifecycle"),
        ("missing_required_check", mutate_missing_required_receipt, "lifecycle"),
        ("report_semantic_digest", mutate_report_semantic_digest, "oracle"),
        ("build_protocol_binding", mutate_build_protocol, "verify"),
        ("compression_original_size", mutate_compression_size, "verify"),
    )
    result: list[tuple[str, Path, Callable[[Path], None], str, str]] = []
    for name, factory, validator in cases:
        path, apply, needle = factory(bundle)
        if not path.is_file():
            raise ProbeError(f"mutation target is missing: {relative(bundle, path)}")
        result.append((name, path, apply, validator, needle))
    return result


def run(stage: str) -> dict[str, Any]:
    if not (ROOT / "SHA256SUMS").is_file():
        raise ProbeError("seal the source bundle before running portable probes")
    with tempfile.TemporaryDirectory(prefix="litchi-0438-probes-") as temporary:
        bundle = Path(temporary) / "bundle"
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
            '"formal_summary_rederived": true',
        )
        mutations: list[dict[str, Any]] = []
        for name, path, apply, validator, needle in mutation_cases(bundle):
            original = path.read_bytes()
            try:
                apply(path)
                refresh_inventory(bundle)
                if validator == "oracle":
                    receipt = load(first_formal_receipt(bundle))
                    receipt_role = receipt.get("role")
                    if not isinstance(receipt_role, str) or not receipt_role:
                        raise ProbeError(f"{relative(bundle, receipt)} has no role binding")
                    command = [sys.executable, "-B", str(oracle_verifier(bundle)),
                        "--report", str(path), "--mode", receipt["lane"]["mode"],
                        "--shape", receipt["lane"]["shape"], "--role", oracle_role(bundle, receipt_role)]
                else:
                    command = verify_command(bundle, stage) if validator == "verify" else lifecycle_command(bundle, stage)
                result = require_reject(run_command(command, bundle.parent), name, needle)
                mutations.append(
                    {
                        "name": name,
                        "path": relative(bundle, path),
                        "validator": validator,
                        **result,
                    }
                )
            finally:
                path.write_bytes(original)
                inventory.write_bytes(original_inventory)
                if path.read_bytes() != original:
                    raise ProbeError(f"failed to restore {relative(bundle, path)}")
                if inventory.read_bytes() != original_inventory:
                    raise ProbeError("failed to restore SHA256SUMS")
            mutations[-1]["restored"] = True

        final_verify = require_pass(
            run_command(verify_command(bundle, stage), bundle.parent),
            "restored portable verify",
            "VALID",
        )
        final_lifecycle = require_pass(
            run_command(lifecycle_command(bundle, stage), bundle.parent),
            "restored portable lifecycle",
            '"formal_summary_rederived": true',
        )
    return {
        "status": "pass",
        "change": CHANGE,
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
        print(json.dumps({"status": "failed", "change": CHANGE, "error": str(error)}, sort_keys=True))
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
