#!/usr/bin/env python3
"""Capture excluded whole-process CPU and syscall profiles for both routes.

This driver is intentionally separate from the formal timing capture.  It
uses the accepted normal binary for one materialized and one bounded run at
the largest corpus size, with the same warmup/sample shape as the formal
matrix.  Each profiler command is handed to ``gate.py`` independently so the
gate owns its normal source snapshot and CPU lock.  The coordinator must
serialize this entire driver with a different outer lock; this file must not
acquire ``common.TEMP/cpu.lock`` around the gate calls.

The driver records raw profiler output and custody metadata.  It does not
interpret profiler counters or manufacture metrics when a profiler is
unavailable or fails; the resulting profile manifest is ``partial`` and
retains the failed gate receipt where one exists.
"""

from __future__ import annotations

import argparse
import gzip
import hashlib
import os
import shutil
from pathlib import Path
import subprocess
import sys
from typing import Any

from common import ENV, ROOT, REPO, TEMP, meta, now, read, sha, write


SCHEMA = "docx-tail-append-process-profiles-v1"
BINARY_NAME = "docx_bounded_tail_append_compare"
COUNT = 131_072
SAMPLES = 30
WARMUPS = 3
CPU = 2
ROUTES = ("materialized", "bounded")
WORKLOAD_KINDS = ("perf-stat", "perf-record", "strace")
POSTPROCESS_KINDS = ("perf-report", "perf-script")
EVENTS = (
    "cycles",
    "instructions",
    "branches",
    "branch-misses",
    "cache-misses",
    "page-faults",
)
RECORD_FREQUENCY = 99
PERF = "/usr/bin/perf"
STRACE = "/usr/bin/strace"


def token(value: str) -> str:
    if (
        not value
        or value in {".", ".."}
        or "/" in value
        or "\\" in value
        or any(character.isspace() for character in value)
    ):
        raise SystemExit("--attempt must be a non-empty path-safe token")
    return value


def parse() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--attempt", required=True)
    return parser.parse_args()


def relative(path: Path) -> str:
    return path.relative_to(ROOT).as_posix()


def file_ref(path: Path) -> dict[str, Any] | None:
    if not path.is_file() or path.is_symlink():
        return None
    return {"path": relative(path), **meta(path)}


def binary_binding(attempt: str) -> dict[str, Any]:
    custody_path = ROOT / "builds" / f"binaries-{attempt}.json"
    if not custody_path.is_file():
        raise RuntimeError(f"accepted binary custody is missing: {custody_path}")
    custody = read(custody_path)
    if (
        not isinstance(custody, dict)
        or custody.get("schema") != "docx-tail-append-binaries-v1"
        or custody.get("attempt") != attempt
        or custody.get("binary_name") != BINARY_NAME
    ):
        raise RuntimeError(f"binary custody shape differs: {custody_path}")
    binaries = custody.get("binaries")
    if not isinstance(binaries, dict) or set(binaries) != {"normal", "allocator"}:
        raise RuntimeError("binary custody must contain normal and allocator identities")
    normal = binaries["normal"]
    if not isinstance(normal, dict):
        raise RuntimeError("normal binary identity is malformed")
    required = {
        "path",
        "bytes",
        "sha256",
        "build_path",
        "build_sha256",
        "source_manifest_sha256",
    }
    if set(normal) != required:
        raise RuntimeError("normal binary identity fields differ")
    binary_path = Path(normal["path"])
    actual = meta(binary_path)
    expected = {key: normal[key] for key in ("bytes", "sha256")}
    if actual != expected:
        raise RuntimeError(f"normal binary changed after build custody: {binary_path}")
    if not isinstance(normal["source_manifest_sha256"], str) or len(normal["source_manifest_sha256"]) != 64:
        raise RuntimeError("normal source manifest binding is malformed")
    return {
        "custody": {"path": relative(custody_path), **meta(custody_path)},
        "custody_attempt": attempt,
        "source_manifest_sha256": normal["source_manifest_sha256"],
        "binary": {
            "name": BINARY_NAME,
            "path": str(binary_path),
            "bytes": actual["bytes"],
            "sha256": actual["sha256"],
        },
        "build": {
            "path": normal["build_path"],
            "sha256": normal["build_sha256"],
        },
    }


def workload(binary: Path, route: str, report: Path) -> list[str]:
    return [
        str(binary),
        "--route",
        route,
        "--counts",
        str(COUNT),
        "--samples",
        str(SAMPLES),
        "--warmups",
        str(WARMUPS),
        "--json",
        str(report),
    ]


def taskset(command: list[str]) -> list[str]:
    return ["/usr/bin/taskset", "-c", str(CPU), *command]


def gate_one(attempt: str, label: str, argv: list[str]) -> dict[str, Any]:
    """Run one profiler command through the ordinary per-command gate.

    A missing executable can raise before gate.py writes its final receipt.
    That condition is retained as an invocation error and does not create any
    counter value.  Normal command failures still retain the gate receipt.
    """
    receipt_path = ROOT / "validation" / f"{label}-{attempt}.json"
    command = [
        sys.executable,
        "-B",
        str(ROOT / "gate.py"),
        "--attempt",
        attempt,
        label,
        *argv,
    ]
    try:
        completed = subprocess.run(command, cwd=REPO, env=ENV, check=False)
    except OSError as error:
        return {
            "status": "failed_to_start",
            "label": label,
            "argv": argv,
            "exit_code": None,
            "error": repr(error),
            "gate_command": command,
        }
    result: dict[str, Any] = {
        "status": "failed",
        "label": label,
        "argv": argv,
        "exit_code": completed.returncode,
        "gate_command": command,
    }
    if receipt_path.is_file():
        receipt = read(receipt_path)
        artifacts = receipt.get("artifacts") if isinstance(receipt, dict) else None
        result.update(
            {
                "status": "pass" if receipt.get("exit_code") == 0 and receipt.get("source_unchanged") is True else "failed",
                "gate_receipt": {"path": relative(receipt_path), **meta(receipt_path)},
                "gate_exit_code": receipt.get("exit_code"),
                "source_unchanged": receipt.get("source_unchanged"),
                "source_before": receipt.get("source_before"),
                "source_after": receipt.get("source_after"),
                "gate_artifacts": artifacts if isinstance(artifacts, dict) else None,
                "started_utc": receipt.get("started_utc"),
                "finished_utc": receipt.get("finished_utc"),
            }
        )
    else:
        result["error"] = "gate.py produced no final receipt"
    return result


def pack_profile(raw: Path, destination: Path) -> dict[str, Any]:
    if not raw.is_file() or raw.is_symlink():
        raise RuntimeError(f"perf record did not retain a regular data file: {raw}")
    if destination.exists():
        raise RuntimeError(f"refusing to replace compressed profile: {destination}")
    raw_identity = meta(raw)
    with raw.open("rb") as source, destination.open("xb") as target:
        with gzip.GzipFile(fileobj=target, mode="wb", filename="", mtime=0) as packed:
            shutil.copyfileobj(source, packed)
    digest = hashlib.sha256()
    size = 0
    with gzip.open(destination, "rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
            size += len(block)
    if {"bytes": size, "sha256": digest.hexdigest()} != raw_identity:
        raise RuntimeError(f"compressed profile failed round-trip verification: {destination}")
    raw.unlink()
    return {
        "status": "pass",
        "path": relative(destination),
        "compressed": meta(destination),
        "uncompressed": raw_identity,
    }


def output_record(path: Path) -> dict[str, Any] | None:
    value = file_ref(path)
    return value


def tool_available(path: str) -> bool:
    executable = Path(path)
    return executable.is_file() and os.access(executable, os.X_OK)


def unsupported_record(
    kind: str,
    route: str,
    argv: list[str],
    tool: str,
    binding: dict[str, Any],
) -> dict[str, Any]:
    """Describe an unavailable profiler without asking gate.py to start it."""

    return {
        "status": "unsupported",
        "kind": kind,
        "route": route,
        "argv": argv,
        "tool": tool,
        "reason": f"profiler executable is unavailable: {tool}",
        "exit_code": None,
        "gate_receipt": None,
        "binary": binding["binary"],
        "source_manifest_sha256": binding["source_manifest_sha256"],
        "harness_report": None,
        "workload_sample_count": None,
    }


def report_sample_count(path: Path, route: str) -> int | None:
    """Read the retained workload report's actual measured sample count.

    Only the shallow report shape emitted by this helper is accepted.  A
    missing or malformed report contributes no sample to the actual count.
    """

    if not path.is_file() or path.is_symlink():
        return None
    try:
        report = read(path)
    except (OSError, UnicodeError, ValueError, TypeError):
        return None
    if not isinstance(report, dict):
        return None
    configuration = report.get("config")
    if not isinstance(configuration, dict):
        return None
    if (
        configuration.get("counts") != [COUNT]
        or configuration.get("samples") != SAMPLES
        or configuration.get("warmups") != WARMUPS
    ):
        return None
    expected_name = (
        "materialized_paragraph_copy"
        if route == "materialized"
        else "bounded_plain_text_tail_append"
    )
    if configuration.get("route", configuration.get("operation")) not in {route, expected_name}:
        return None
    cases = report.get("cases")
    if not isinstance(cases, list) or len(cases) != 1 or not isinstance(cases[0], dict):
        return None
    case = cases[0]
    if case.get("count") != COUNT:
        return None
    routes = case.get("routes")
    if not isinstance(routes, list) or len(routes) != 1 or not isinstance(routes[0], dict):
        return None
    route_record = routes[0]
    if route_record.get("route") != expected_name:
        return None
    samples = route_record.get("samples")
    if not isinstance(samples, list):
        return None
    return len(samples)


def route_profile(attempt: str, route: str, binding: dict[str, Any], route_dir: Path) -> dict[str, Any]:
    binary = Path(binding["binary"]["path"])
    records: dict[str, Any] = {}
    raw = TEMP / attempt / "profiles" / route / "perf-record.data"
    raw.parent.mkdir(parents=True, exist_ok=True)
    if raw.exists():
        raise RuntimeError(f"refusing to replace existing temporary profile: {raw}")
    compressed = route_dir / "perf-record.data.gz"

    workload_reports = {
        "perf-stat": route_dir / "perf-stat.report.json",
        "perf-record": route_dir / "perf-record.report.json",
        "strace": route_dir / "strace.report.json",
    }
    csv_path = route_dir / "perf-stat.csv"
    strace_path = route_dir / "strace.txt"
    commands = {
        "perf-stat": taskset([
            PERF,
            "stat",
            "-x,",
            "-o",
            str(csv_path),
            "-e",
            ",".join(EVENTS),
            "--",
            *workload(binary, route, workload_reports["perf-stat"]),
        ]),
        "perf-record": taskset([
            PERF,
            "record",
            "-e",
            "cycles",
            "--call-graph",
            "fp",
            "-F",
            str(RECORD_FREQUENCY),
            "-o",
            str(raw),
            "--",
            *workload(binary, route, workload_reports["perf-record"]),
        ]),
        "strace": taskset([
            STRACE,
            "-c",
            "-o",
            str(strace_path),
            "--",
            *workload(binary, route, workload_reports["strace"]),
        ]),
    }

    for kind in WORKLOAD_KINDS:
        label = f"profile-{route}-{kind}"
        tool = STRACE if kind == "strace" else PERF
        if tool_available(tool):
            result = gate_one(attempt, label, commands[kind])
        else:
            result = unsupported_record(kind, route, commands[kind], tool, binding)
        result["kind"] = kind
        result["route"] = route
        result["binary"] = binding["binary"]
        result["source_manifest_sha256"] = binding["source_manifest_sha256"]
        result["harness_report"] = output_record(workload_reports[kind])
        result.setdefault("workload_sample_count", None)
        if kind == "perf-stat":
            result["counter_output"] = output_record(csv_path)
            if result["status"] == "pass" and result["counter_output"] is None:
                result["status"] = "failed"
                result["error"] = "perf stat gate passed but counter output is missing"
        elif kind == "strace":
            result["syscall_output"] = output_record(strace_path)
            if result["status"] == "pass" and result["syscall_output"] is None:
                result["status"] = "failed"
                result["error"] = "strace gate passed but syscall output is missing"
        elif kind == "perf-record" and result["status"] == "pass" and not raw.is_file():
            result["status"] = "failed"
            result["error"] = "perf record gate passed but raw profile is missing"
        if result["status"] == "pass":
            result["workload_sample_count"] = report_sample_count(workload_reports[kind], route)
            if result["workload_sample_count"] is None:
                result["status"] = "failed"
                result["error"] = "profiler gate passed but harness report is missing or malformed"
        records[kind] = result

    # Report and script are postprocessors over the retained perf.data.  They
    # are gated separately so their stdout/stderr and source snapshots remain
    # independently reviewable.  If perf record failed or produced no data,
    # retain an explicit skip instead of fabricating report metrics.
    if raw.is_file():
        postprocessors = {
            "perf-report": [
                PERF,
                "report",
                "--stdio",
                "--no-children",
                "--percent-limit",
                "0",
                "-i",
                str(raw),
            ],
            "perf-script": [PERF, "script", "-i", str(raw)],
        }
        for kind, command in postprocessors.items():
            label = f"profile-{route}-{kind}"
            if tool_available(PERF):
                result = gate_one(attempt, label, command)
            else:
                result = unsupported_record(kind, route, command, PERF, binding)
            result.update({
                "kind": kind,
                "route": route,
                "binary": binding["binary"],
                "source_manifest_sha256": binding["source_manifest_sha256"],
            })
            records[kind] = result
        try:
            records["raw_profile"] = pack_profile(raw, compressed)
        except (OSError, RuntimeError) as error:
            records["raw_profile"] = {"status": "failed", "error": repr(error)}
    else:
        for kind in POSTPROCESS_KINDS:
            records[kind] = {
                "status": "skipped",
                "kind": kind,
                "route": route,
                "reason": "perf-record produced no raw perf.data",
                "binary": binding["binary"],
                "source_manifest_sha256": binding["source_manifest_sha256"],
                "argv": (
                    [PERF, "report", "--stdio", "--no-children", "--percent-limit", "0", "-i", str(raw)]
                    if kind == "perf-report" else [PERF, "script", "-i", str(raw)]
                ),
                "tool": PERF,
                "gate_receipt": None,
                "workload_sample_count": None,
            }
        records["raw_profile"] = {
            "status": "skipped",
            "reason": "perf-record produced no raw perf.data",
        }

    statuses = [value.get("status") for key, value in records.items() if key != "raw_profile"]
    raw_status = records.get("raw_profile", {}).get("status")
    status = "pass" if statuses and all(value == "pass" for value in statuses) and raw_status == "pass" else "partial"
    workload_records = [records[kind] for kind in WORKLOAD_KINDS]
    successful_workload_samples = sum(
        value["workload_sample_count"]
        for value in workload_records
        if value.get("workload_sample_count") is not None
    )
    return {
        "route": route,
        "route_name": (
            "materialized_paragraph_copy"
            if route == "materialized"
            else "bounded_plain_text_tail_append"
        ),
        "count": COUNT,
        "samples": SAMPLES,
        "warmups": WARMUPS,
        "binary": binding["binary"],
        "source_manifest_sha256": binding["source_manifest_sha256"],
        "successful_workload_runs": sum(
            value.get("workload_sample_count") is not None
            for value in workload_records
        ),
        "successful_workload_samples": successful_workload_samples,
        "records": records,
        "status": status,
    }


def main() -> None:
    options = parse()
    attempt = token(options.attempt)
    binding = binary_binding(attempt)
    profile_root = ROOT / "profiles" / attempt
    if profile_root.exists():
        raise RuntimeError(f"refusing to replace existing profile attempt: {profile_root}")
    labels = [
        f"profile-{route}-{kind}-{attempt}"
        for route in ROUTES
        for kind in (*WORKLOAD_KINDS, *POSTPROCESS_KINDS)
    ]
    for label in labels:
        prefix = ROOT / "validation" / label
        if any(
            path.exists()
            for path in (
                prefix.with_suffix(".started.json"),
                prefix.with_suffix(".stdout"),
                prefix.with_suffix(".stderr"),
                prefix.with_suffix(".json"),
            )
        ):
            raise RuntimeError(f"refusing to replace existing validation attempt: {label}")
    profile_root.parent.mkdir(exist_ok=True)
    profile_root.mkdir()
    started = now()
    routes: dict[str, Any] = {}
    for route in ROUTES:
        route_dir = profile_root / route
        route_dir.mkdir()
        routes[route] = route_profile(attempt, route, binding, route_dir)

    status = "pass" if all(row["status"] == "pass" for row in routes.values()) else "partial"
    manifest = {
        "schema": SCHEMA,
        "attempt": attempt,
        "started_utc": started,
        "finished_utc": now(),
        "status": status,
        "formal_samples": False,
        "excluded_samples": {
            "planned_workload_runs": len(ROUTES) * len(WORKLOAD_KINDS),
            "planned_workload_samples": len(ROUTES) * len(WORKLOAD_KINDS) * SAMPLES,
            "successful_workload_runs": sum(
                row["successful_workload_runs"] for row in routes.values()
            ),
            "successful_workload_samples": sum(
                row["successful_workload_samples"] for row in routes.values()
            ),
            "unit": "retained harness samples from successful profiler workload reports",
        },
        "scope": (
            "separate whole-process normal runs including corpus construction, "
            "untimed independent oracles, warmups, 30 harness lifecycles, "
            "report serialization and teardown; no transaction-only attribution"
        ),
        "binary": binding,
        "configuration": {
            "cpu": CPU,
            "count": COUNT,
            "samples": SAMPLES,
            "warmups": WARMUPS,
            "routes": list(ROUTES),
            "stat_events": list(EVENTS),
            "record_event": "cycles",
            "record_frequency_hz": RECORD_FREQUENCY,
            "call_graph": "fp",
            "same_normal_binary_for_both_routes": True,
            "instrumented_timing_excluded_from_formal_evidence": True,
        },
        "drivers": {
            "profile.py": {"path": relative(Path(__file__)), "sha256": sha(Path(__file__))},
            "common.py": {"path": relative(ROOT / "common.py"), "sha256": sha(ROOT / "common.py")},
            "gate.py": {"path": relative(ROOT / "gate.py"), "sha256": sha(ROOT / "gate.py")},
        },
        "routes": routes,
        "status_note": (
            "A partial status retains failed tool receipts and omits unavailable "
            "metrics; no profiler failure is converted into a value."
        ),
    }
    write(profile_root / "profile.json", manifest)
    print(f"wrote {relative(profile_root / 'profile.json')} ({status})")


if __name__ == "__main__":
    main()
