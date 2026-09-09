#!/usr/bin/env python3
"""Freeze and execute the 0483 DOCX materialized/bounded total matrix.

The coordinator runs this driver after both instrumentation binaries and the
actual harness report schema have been reviewed.  ``--freeze`` is the only
operation that creates ``protocol.json``; all measurement receipts are
exclusive and retain their exact command, binary and environment identities.
"""

from __future__ import annotations

import argparse
import fcntl
from pathlib import Path
import subprocess
import sys
from typing import Any

from common import (
    DRIVER_SCRIPTS,
    ENV,
    ENV_KEYS,
    ROOT,
    REPO,
    TEMP,
    environment,
    meta,
    now,
    read,
    sha,
    write,
)


COUNTS = (64, 8_192, 131_072)
INSTRUMENTATIONS = ("normal", "allocator")
ARMS = (
    ("a1", 1, "materialized"),
    ("b1", 1, "bounded"),
    ("b2", 2, "bounded"),
    ("a2", 2, "materialized"),
)
SAMPLES = 30
WARMUPS = 3
CPU = 2
BINARY_NAME = "docx_bounded_tail_append_compare"
ROUTE_FLAG = "--route"


def token(value: str, field: str) -> str:
    if (
        not value
        or value in {".", ".."}
        or "/" in value
        or "\\" in value
        or any(character.isspace() for character in value)
    ):
        raise SystemExit(f"{field} must be a non-empty path-safe token")
    return value


def expected_captures(attempt: str) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for arm, repeat, route in ARMS:
        counts = COUNTS if repeat == 1 else tuple(reversed(COUNTS))
        for instrumentation in INSTRUMENTATIONS:
            for count in counts:
                label = f"{arm}-{instrumentation}-{count}-{route}"
                rows.append({
                    "label": label,
                    "arm": arm,
                    "repeat": repeat,
                    "route": route,
                    "route_name": (
                        "materialized_paragraph_copy"
                        if route == "materialized"
                        else "bounded_plain_text_tail_append"
                    ),
                    "instrumentation": instrumentation,
                    "count": count,
                    "attempt": attempt,
                    "argv": capture_argv(attempt, instrumentation, route, count, label),
                })
    return rows


def capture_argv(
    attempt: str,
    instrumentation: str,
    route: str,
    count: int,
    label: str,
) -> list[str]:
    output = ROOT / "captures" / label
    binary = TEMP / attempt / instrumentation / BINARY_NAME
    return [
        "/usr/bin/time",
        "-v",
        "-o",
        str(output.with_suffix(".resource")),
        "/usr/bin/taskset",
        "-c",
        str(CPU),
        str(binary),
        ROUTE_FLAG,
        route,
        "--counts",
        str(count),
        "--samples",
        str(SAMPLES),
        "--warmups",
        str(WARMUPS),
        "--json",
        str(output.with_suffix(".report.json")),
    ]


def parse() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--attempt")
    parser.add_argument("--arm", choices=[item[0] for item in ARMS])
    parser.add_argument("--freeze", action="store_true")
    return parser.parse_args()


def freeze(attempt: str) -> None:
    protocol_path = ROOT / "protocol.json"
    if protocol_path.exists():
        raise RuntimeError(f"refusing to replace frozen protocol: {protocol_path}")
    validation_plan_path = ROOT / "validation-plan.json"
    fuzz_plan_path = ROOT / "fuzz-plan.json"
    if not validation_plan_path.is_file() or not fuzz_plan_path.is_file():
        raise RuntimeError(
            "freeze requires validation-plan.json and fuzz-plan.json containing "
            "actual accepted labels/argv/receipt references"
        )
    validation = read(validation_plan_path)
    fuzz = read(fuzz_plan_path)
    if (
        not isinstance(validation, dict)
        or not isinstance(validation.get("required_labels"), list)
        or not isinstance(validation.get("pilot_labels"), list)
        or not isinstance(validation.get("argv"), dict)
        or not isinstance(validation.get("pilot_reports"), dict)
        or not isinstance(validation.get("developmental"), dict)
        or not isinstance(fuzz, dict)
    ):
        raise RuntimeError("validation/fuzz custody plan shape is invalid")
    labels = validation["required_labels"] + validation["pilot_labels"]
    if (
        not labels
        or not all(isinstance(label, str) for label in labels)
        or len(set(labels)) != len(labels)
        or set(validation["argv"]) != set(labels)
        or not all(token(label, "validation label") == label for label in labels)
    ):
        raise RuntimeError("validation custody plan labels/argv are incomplete or unsafe")
    if not isinstance(fuzz.get("required_labels"), list) or not fuzz["required_labels"]:
        raise RuntimeError("fuzz custody plan must name accepted receipt labels")
    protocol = {
        "schema": "docx-tail-append-comparison-v1",
        "attempt": attempt,
        "frozen_utc": now(),
        "samples": SAMPLES,
        "warmups": WARMUPS,
        "cpu": CPU,
        "arms": ["a1", "b1", "b2", "a2"],
        "routes": ["materialized", "bounded"],
        "instrumentations": list(INSTRUMENTATIONS),
        "counts": list(COUNTS),
        "comparison": "A1 materialized / B1 bounded / B2 bounded / A2 materialized; same executable and fresh process per capture",
        "scope": "source-backed DOCX tail append total lifecycle including source/package admission, edit preparation, commit, publication, sink finalization and drops; corpus/oracles excluded",
        "normal_and_allocator_timings_separate": True,
        "process_rss_includes_setup_oracles_and_teardown": True,
        "allocator_counters_are_operation_scoped": True,
        "phase_samples": False,
        "regression_review_percent": 5,
        "performance_claim": "none",
        "route_flag": ROUTE_FLAG,
        "scripts": {name: sha(ROOT / name) for name in DRIVER_SCRIPTS},
        "environment": environment(),
        "validation": validation,
        "fuzz": fuzz,
        "plan_files": {
            "validation-plan.json": sha(validation_plan_path),
            "fuzz-plan.json": sha(fuzz_plan_path),
        },
        "evidence_inputs": {
            name: meta(ROOT / name)
            for name in (
                "workspace-iwork-exclusion.json", "machine.json", "available-tools.json",
                f"profiles/{attempt}/profile.json",
            )
        },
        "profile_scope": "Separate whole-process diagnostic profiles; instrumented timings are excluded from formal latency and RSS comparisons.",
        "captures": expected_captures(attempt),
    }
    # Run the read-only custody subset before creating the immutable protocol.
    # This verifies accepted build/gate records, every classified validation
    # receipt and pilot report, and the fuzz seed/receipt plan without running
    # Cargo, the harness, or any measurement.
    try:
        import analyze
        import verify

        verify.check_protocol(protocol)
        binaries = verify.check_builds(protocol)
        validation_result = verify.check_validation_receipts(protocol, binaries)
        fuzz_result = verify.check_fuzz_custody(protocol)
        verify.check_pre_freeze_chronology(protocol, binaries, validation_result, fuzz_result)
    except (verify.VerificationError, analyze.AnalysisError) as error:
        raise RuntimeError(f"pre-freeze custody verification failed: {error}") from error
    write(protocol_path, protocol)
    print(f"Frozen {len(protocol['captures'])} total-lifecycle captures.")


def check_binary(spec: dict[str, Any]) -> None:
    path = Path(spec["path"])
    if meta(path) != {key: spec[key] for key in ("bytes", "sha256")}:
        raise RuntimeError(f"binary changed after custody: {path}")


def run_arm(attempt: str, arm: str) -> None:
    protocol_path = ROOT / "protocol.json"
    protocol = read(protocol_path)
    if protocol.get("attempt") != attempt:
        raise RuntimeError("capture attempt does not match frozen protocol")
    if protocol.get("scripts") != {name: sha(ROOT / name) for name in protocol["scripts"]}:
        raise RuntimeError("capture scripts changed after protocol freeze")
    binaries = read(ROOT / "builds" / f"binaries-{attempt}.json")
    if binaries.get("attempt") != attempt:
        raise RuntimeError("binary custody attempt differs")
    for value in binaries["binaries"].values():
        check_binary(value)
    rows = [row for row in protocol["captures"] if row["arm"] == arm]
    if len(rows) != len(INSTRUMENTATIONS) * len(COUNTS):
        raise RuntimeError(f"protocol has an incomplete arm: {arm}")
    (ROOT / "captures").mkdir(exist_ok=True)
    for capture in rows:
        label = capture["label"]
        output = ROOT / "captures" / label
        paths = (
            output.with_suffix(".started.json"),
            output.with_suffix(".stdout"),
            output.with_suffix(".stderr"),
            output.with_suffix(".json"),
        )
        if any(path.exists() for path in paths):
            raise RuntimeError(f"capture already exists: {label}")
        binary = binaries["binaries"][capture["instrumentation"]]
        check_binary(binary)
        started = {
            "schema": "docx-tail-append-capture-v1",
            "capture": capture,
            "cwd": str(REPO),
            "started_utc": now(),
            "protocol_sha256": sha(protocol_path),
            "binary": binary,
            "environment": environment(),
        }
        write(output.with_suffix(".started.json"), started)
        with output.with_suffix(".stdout").open("xb") as out, output.with_suffix(".stderr").open("xb") as err:
            completed = subprocess.run(capture["argv"], cwd=REPO, env=ENV, stdout=out, stderr=err)
        artifacts = {}
        for suffix in (".stdout", ".stderr", ".resource", ".report.json"):
            path = output.with_suffix(suffix)
            if path.is_file():
                artifacts[path.name] = meta(path)
        receipt = dict(
            started,
            exit_code=completed.returncode,
            finished_utc=now(),
            artifacts=artifacts,
        )
        write(output.with_suffix(".json"), receipt)
        print(label, completed.returncode, flush=True)
        if completed.returncode:
            raise SystemExit(completed.returncode)


def main() -> None:
    options = parse()
    if options.freeze:
        if not options.attempt:
            raise SystemExit("--freeze requires --attempt")
        freeze(token(options.attempt, "--attempt"))
        return
    if not options.arm or not options.attempt:
        raise SystemExit("capture.py requires --attempt ATTEMPT --arm a1|b1|b2|a2")
    attempt = token(options.attempt, "--attempt")
    TEMP.mkdir(parents=True, exist_ok=True)
    with (TEMP / "cpu.lock").open("a") as lock:
        fcntl.flock(lock, fcntl.LOCK_EX)
        run_arm(attempt, options.arm)


if __name__ == "__main__":
    main()
