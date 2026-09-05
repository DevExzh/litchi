#!/usr/bin/env python3
"""Capture the change-0416 ZIP strict-read guard in process-isolated ABBA order.

Each ordinary fixture is run once per reader mode and per leg in the fixed
``A1, B1, B2, A2`` order.  The command records the exact executable, source
fixture identity, process affinity, and ``/usr/bin/time -v`` report.  A
capability fixture may be supplied separately; its control refusal and
candidate admission are recorded as statuses and are never reduced to a
latency percentage.

Example::

    python3 capture.py \
      --control /tmp/control/zip-index-probe \
      --candidate /tmp/candidate/zip-index-probe \
      --control-tree /tmp/control-tree --candidate-tree /tmp/candidate-tree \
      --fixture zip32-tiny4=/tmp/zip32-tiny4.zip \
      --fixture zip32-many256=/tmp/zip32-many256.zip \
      --fixture central-zip64-tiny4=/tmp/central-zip64-tiny4.zip \
      --fixture central-zip64-many256=/tmp/central-zip64-many256.zip \
      --indexed-fixture native-docx=/tmp/document-properties-litchi.docx \
      --capability-fixture local-only=/tmp/local-only.zip \
      --output /tmp/change-0416-capture
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
from pathlib import Path
import re
import subprocess
from typing import Any


LEGS = ("A1", "B1", "B2", "A2")
MODES = ("borrowed", "indexed")
ROLE_FOR_LEG = {"A1": "control", "B1": "candidate", "B2": "candidate", "A2": "control"}
LABEL_RE = re.compile(r"^[A-Za-z0-9_.-]+$")


def fail(message: str) -> "NoReturn":
    raise SystemExit(f"error: {message}")


def iso_now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def parse_assignment(value: str, *, where: str) -> tuple[str, Path]:
    label, separator, raw_path = value.partition("=")
    if not separator or not label or not raw_path:
        fail(f"{where} must be LABEL=PATH")
    if not LABEL_RE.fullmatch(label):
        fail(f"{where} label {label!r} contains unsupported characters")
    path = Path(raw_path).expanduser().resolve()
    if not path.is_file():
        fail(f"{where} path is not a regular file: {path}")
    return label, path


def parse_assignments(values: list[str] | None, *, where: str, minimum: int) -> dict[str, Path]:
    if not values or len(values) < minimum:
        fail(f"at least {minimum} --{where.replace('_', '-')} assignment(s) are required")
    result: dict[str, Path] = {}
    for value in values:
        label, path = parse_assignment(value, where=f"--{where.replace('_', '-')}")
        if label in result:
            fail(f"duplicate {where} label {label!r}")
        result[label] = path
    return dict(sorted(result.items()))


def sha256_file(path: Path) -> tuple[str, int]:
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as source:
        while True:
            block = source.read(1024 * 1024)
            if not block:
                break
            digest.update(block)
            size += len(block)
    return digest.hexdigest(), size


def binary_identity(binary: Path, tree: Path) -> dict[str, Any]:
    status = subprocess.check_output(
        ["git", "status", "--porcelain"], cwd=tree, text=True
    )
    if status:
        fail(f"worktree is dirty: {tree}\n{status}")
    digest, size = sha256_file(binary)
    revision = subprocess.check_output(
        ["git", "rev-parse", "HEAD"], cwd=tree, text=True
    ).strip()
    return {
        "binary": str(binary),
        "sha256": digest,
        "bytes": size,
        "tree": str(tree),
        "revision": revision,
    }


def fixture_identity(fixtures: dict[str, Path]) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for label, path in fixtures.items():
        digest, size = sha256_file(path)
        result[label] = {"path": str(path), "sha256": digest, "bytes": size}
    return result


def write_manifest(output: Path, payload: dict[str, Any]) -> None:
    (output / "capture.json").write_text(
        json.dumps(payload, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )


def run_process(
    *,
    role: str,
    identity: dict[str, Any],
    tree: Path,
    argv: list[str],
    report: Path,
    time_report: Path,
    expected_exit: int | None,
) -> dict[str, Any]:
    command = [
        "taskset",
        "-c",
        str(CPU),
        "/usr/bin/time",
        "-v",
        "-o",
        str(time_report),
        identity["binary"],
        *argv,
    ]
    started = iso_now()
    with report.open("w", encoding="utf-8") as stdout:
        result = subprocess.run(
            command,
            cwd=tree,
            stdout=stdout,
            stderr=subprocess.PIPE,
            text=True,
            check=False,
        )
    finished = iso_now()
    record: dict[str, Any] = {
        "role": role,
        "argv": command,
        "started": started,
        "finished": finished,
        "exit_code": result.returncode,
        "stderr": result.stderr,
        "report": str(report),
        "time_v": str(time_report),
        "expected_exit": expected_exit,
    }
    if expected_exit is not None and result.returncode != expected_exit:
        fail(f"{report.name}: exit code {result.returncode}, expected {expected_exit}")
    return record


def load_status(report: Path) -> str | None:
    try:
        value = json.loads(report.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    if isinstance(value, dict) and isinstance(value.get("status"), str):
        return value["status"]
    return None


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--control", type=Path, required=True)
    parser.add_argument("--candidate", type=Path, required=True)
    parser.add_argument("--control-tree", type=Path, required=True)
    parser.add_argument("--candidate-tree", type=Path, required=True)
    parser.add_argument(
        "--fixture",
        action="append",
        required=True,
        metavar="LABEL=PATH",
        help="ordinary guard fixture; supply at least four assignments",
    )
    parser.add_argument(
        "--capability-fixture",
        action="append",
        metavar="LABEL=PATH",
        help="optional local-only capability fixture, kept outside ordinary rows",
    )
    parser.add_argument(
        "--indexed-fixture",
        action="append",
        metavar="LABEL=PATH",
        help="optional fixture measured only through the indexed preservation path",
    )
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--samples", type=int, default=300)
    parser.add_argument("--warmups", type=int, default=30)
    parser.add_argument("--cpu", type=int, default=2)
    args = parser.parse_args()

    if args.samples <= 0:
        fail("--samples must be positive")
    if args.warmups < 0:
        fail("--warmups must be non-negative")
    if args.cpu < 0:
        fail("--cpu must be non-negative")

    global CPU
    CPU = args.cpu
    control = args.control.resolve()
    candidate = args.candidate.resolve()
    control_tree = args.control_tree.resolve()
    candidate_tree = args.candidate_tree.resolve()
    for path, label in (
        (control, "--control"),
        (candidate, "--candidate"),
        (control_tree, "--control-tree"),
        (candidate_tree, "--candidate-tree"),
    ):
        if not path.exists():
            fail(f"{label} path does not exist: {path}")
    fixtures = parse_assignments(args.fixture, where="fixture", minimum=4)
    indexed_fixtures = parse_assignments(
        args.indexed_fixture, where="indexed_fixture", minimum=0
    ) if args.indexed_fixture else {}
    overlap = set(fixtures).intersection(indexed_fixtures)
    if overlap:
        fail(f"fixture labels cannot be shared by ordinary and indexed-only sets: {sorted(overlap)}")
    capability = parse_assignments(
        args.capability_fixture, where="capability_fixture", minimum=0
    ) if args.capability_fixture else {}

    output = args.output.resolve()
    (output / "guards").mkdir(parents=True, exist_ok=True)
    (output / "capability").mkdir(parents=True, exist_ok=True)
    identities = {
        "control": binary_identity(control, control_tree),
        "candidate": binary_identity(candidate, candidate_tree),
    }
    payload: dict[str, Any] = {
        "schema_version": 1,
        "tool": "litchi-goal-0416-zip-strict-capture",
        "abba_order": list(LEGS),
        "modes": list(MODES),
        "samples": args.samples,
        "warmups": args.warmups,
        "cpu": args.cpu,
        "identities": identities,
        "fixtures": fixture_identity(fixtures),
        "indexed_fixtures": fixture_identity(indexed_fixtures),
        "capability_fixtures": fixture_identity(capability),
        "runs": [],
        "capability_runs": [],
    }
    write_manifest(output, payload)

    for leg in LEGS:
        role = ROLE_FOR_LEG[leg]
        identity = identities[role]
        tree = control_tree if role == "control" else candidate_tree
        for fixture_label, fixture_path in fixtures.items():
            for mode in MODES:
                stem = f"{leg}-{fixture_label}-{mode}"
                report = output / "guards" / f"{stem}.json"
                time_report = output / "guards" / f"{stem}.time.txt"
                run = run_process(
                    role=role,
                    identity=identity,
                    tree=tree,
                    argv=[
                        "index",
                        mode,
                        str(fixture_path),
                        str(args.samples),
                        str(args.warmups),
                    ],
                    report=report,
                    time_report=time_report,
                    expected_exit=0,
                )
                run.update({"leg": leg, "fixture": fixture_label, "mode": mode})
                payload["runs"].append(run)
                write_manifest(output, payload)
                print(stem, flush=True)

        for fixture_label, fixture_path in indexed_fixtures.items():
            mode = "indexed"
            stem = f"{leg}-{fixture_label}-{mode}"
            report = output / "guards" / f"{stem}.json"
            time_report = output / "guards" / f"{stem}.time.txt"
            run = run_process(
                role=role,
                identity=identity,
                tree=tree,
                argv=[
                    "index",
                    mode,
                    str(fixture_path),
                    str(args.samples),
                    str(args.warmups),
                ],
                report=report,
                time_report=time_report,
                expected_exit=0,
            )
            run.update({"leg": leg, "fixture": fixture_label, "mode": mode, "indexed_only": True})
            payload["runs"].append(run)
            write_manifest(output, payload)
            print(stem, flush=True)

    for fixture_label, fixture_path in capability.items():
        for role, identity, tree in (
            ("control", identities["control"], control_tree),
            ("candidate", identities["candidate"], candidate_tree),
        ):
            for mode in MODES:
                stem = f"{role}-{fixture_label}-{mode}"
                report = output / "capability" / f"{stem}.json"
                time_report = output / "capability" / f"{stem}.time.txt"
                run = run_process(
                    role=role,
                    identity=identity,
                    tree=tree,
                    argv=["capability", mode, str(fixture_path)],
                    report=report,
                    time_report=time_report,
                    expected_exit=0,
                )
                status = load_status(report)
                run.update(
                    {
                        "fixture": fixture_label,
                        "mode": mode,
                        "status": status,
                        "expected_status": "error" if role == "control" else "ok",
                    }
                )
                payload["capability_runs"].append(run)
                write_manifest(output, payload)
                expected_status = run["expected_status"]
                if status != expected_status:
                    fail(
                        f"{stem}: capability status {status!r}, expected {expected_status!r}"
                    )
                print(f"capability/{stem}", flush=True)


CPU = 2


if __name__ == "__main__":
    main()
